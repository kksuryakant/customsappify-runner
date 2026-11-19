"""HMM ATA module (ported, pier2pier-based).

Reads a container Excel (``Container No`` + optional ``Liner``/``BL``), tracks
each HMM container through the third-party pier2pier.com tracking page, and
writes the per-port ETAs back to the ETA columns.

The HMM carrier site (hmm21.com) serves a firewall challenge to automated
browsers, so pier2pier's static snapshot of the HMM tracking page is used
instead. The flow is plain HTTP: GET the tracking URL twice (the first
response sets a cookie), read the ``Temp_Files`` iframe URL from the shell
page, then fetch that iframe and parse the Shipment Schedule / movement
tables. pier2pier resolves by container number only.
"""

from __future__ import annotations

import logging
import re
import time
from html import unescape as html_unescape
from http.cookiejar import CookieJar
from urllib.request import HTTPCookieProcessor, Request, build_opener

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule

logger = logging.getLogger(__name__)

TRACKING_COLUMNS = ["Container No", "BL", "Liner", "Date", "Location", "Description"]

# Output column -> location keywords found in a port name.
LOCATION_COLUMNS = [
    ("JEBEL ALI ETA", ("JEBEL ALI", "JEBELALI", "JEA")),
    ("KHORFKHAN ETA", ("FAKKAN", "KHORFKHAN", "KHOR FKAN", "KFK")),
    ("SALLALAH ETA", ("SALALAH", "SALLALAH", "SLL")),
    ("ABU DHABI ETC", ("ABU DHABI", "ABUDHABI", "AUH")),
    ("FUJAIRAH ETA", ("FUJAIRAH", "FUJAYRAH")),
    ("JEDDAH ETA", ("JEDDAH", "JEDDA")),
]


P2P_TRACKING_URL = (
    "https://www.pier2pier.com/links/tracking2.php?ID={container}&Type=CONT&Company=P2P"
)
P2P_IFRAME_RE = re.compile(r"src=['\"]([^'\"]*Temp_Files[^'\"]+\.html)['\"]")
P2P_USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
)

CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$")
DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}")

MODULE = Module(
    key="HMM_ATA",
    name="HMM ATA",
    description="Fetch HMM ETA per destination port via pier2pier.com (container number).",
    icon="🚢",
    category="Arrival Status",
    requires_browser=False,
    requires_credentials=False,
    ported=True,
    notes="HMM's own site blocks automated access; pier2pier's snapshot is used instead.",
)


class HMMATAModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        excel_files = list(ctx.uploaded.get("excel_file", []))
        if not excel_files:
            raise ModuleError("An Excel file with BL/container data is required.")
        excel_path = excel_files[0]
        if not excel_path.exists():
            raise ModuleError(f"Uploaded file not found: {excel_path.name}")

        skip_done = bool(ctx.params.get("skip_done", True))
        ctx.log(f"Reading shipments from {excel_path.name}")
        try:
            df = pd.read_excel(excel_path)
        except Exception as exc:
            raise ModuleError(f"Could not read the Excel file: {exc}") from exc

        has_ctr = "Container No" in df.columns
        bl_col = "BL Number" if "BL Number" in df.columns else ("BL" if "BL" in df.columns else None)
        if not has_ctr and not bl_col:
            raise ModuleError("The Excel file must contain a 'Container No' column.")
        has_liner = "Liner" in df.columns
        for col, _ in LOCATION_COLUMNS:
            if col not in df.columns:
                df[col] = ""
            df[col] = df[col].astype(object).where(df[col].notna(), "")

        total = len(df)
        if total == 0:
            raise ModuleError("The Excel file contains no rows.")
        ctx.log(f"Found {total} row(s)")

        ctx.progress(2, "Preparing pier2pier tracking")
        try:
            processed = 0
            for index, row in df.iterrows():
                container_no = self._cell(row, "Container No") if has_ctr else ""
                bl_no = self._cell(row, bl_col) if bl_col else ""
                liner = self._cell(row, "Liner") if has_liner else ""
                key = container_no or bl_no
                if not key:
                    continue

                if skip_done and self._all_filled(row):
                    ctx.log(f"Skipping {key} - all ETA columns already filled")
                    continue

                if has_liner and liner and "hmm" not in liner.lower():
                    ctx.log(f"{key}: not HMM (liner={liner})")
                    continue

                progress = 5 + int((index + 1) / total * 90)
                ctx.progress(progress, f"Tracking {key}")
                ctx.log(f"[HMM] Processing {key}")

                try:
                    events = self._track_p2p(ctx, container_no)
                except Exception as exc:
                    ctx.log(f"[HMM] {key}: {exc}")
                    events = []

                locations = self._map_events(events)
                for col, _ in LOCATION_COLUMNS:
                    if locations.get(col):
                        df.at[index, col] = locations[col]
                processed += 1
                summary = ", ".join(
                    f"{c.split()[0]}={locations[c]}" for c, _ in LOCATION_COLUMNS if locations.get(c)
                ) or "no ETA found"
                ctx.log(f"{key}: {summary}")
                self._checkpoint(df, ctx, excel_path.stem)

            ctx.progress(98, "Saving results")
            out_path = ctx.work_dir / f"{excel_path.stem}_ETA.xlsx"
            df.to_excel(out_path, index=False, engine="openpyxl")
            ctx.log(f"Saved results to {out_path.name}")
            ctx.progress(100, "Completed")
            ctx.log(f"Done. Processed {processed} shipment(s).")
        except Exception:
            raise

    @staticmethod
    def _cell(row, column) -> str:
        if not column or column not in row:
            return ""
        value = row.get(column)
        if pd.isna(value):
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    def _all_filled(self, row) -> bool:
        for col, _ in LOCATION_COLUMNS:
            value = self._cell(row, col)
            if not value or value.lower() in ("nan", "none"):
                return False
        return True

    @staticmethod
    def _map_events(events) -> dict:
        """Map each port column to its latest event date.

        A port can appear in several rows (schedule + vessel movement). The
        latest date always wins, regardless of milestone.
        """
        from datetime import datetime

        best: dict = {}

        def _key(date: str):
            try:
                return datetime.strptime(date, "%m/%d/%Y")
            except Exception:
                return None

        for event in events or []:
            date = event.get("date") or ""
            if not date:
                continue
            text = f"{event.get('location', '')} {event.get('description', '')}".upper()
            for col, keywords in LOCATION_COLUMNS:
                if not any(k in text for k in keywords):
                    continue
                current = best.get(col)
                date_key = _key(date)
                if current is None:
                    best[col] = (date, date_key)
                elif date_key is not None and current[1] is not None and date_key > current[1]:
                    best[col] = (date, date_key)
        return {col: value[0] for col, value in best.items()}

    # ------------------------------------------------------------- tracking
    def _track_p2p(self, ctx: ModuleContext, container_no: str) -> list:
        clean = "".join(container_no.split()) if container_no else ""
        if not clean:
            ctx.log("[HMM] pier2pier resolves by container number only")
            return []
        if not CONTAINER_RE.match(clean.upper()):
            ctx.log(f"[HMM] not a valid container number: {clean}")
            return []

        opener = build_opener(HTTPCookieProcessor(CookieJar()))
        iframe_url = ""
        for _ in range(3):
            page = self._http_get(opener, P2P_TRACKING_URL.format(container=clean.upper()))
            match = P2P_IFRAME_RE.search(page or "")
            if match:
                iframe_url = match.group(1)
                if iframe_url.startswith("//"):
                    iframe_url = "https:" + iframe_url
                break
            time.sleep(1.5)
        if not iframe_url:
            ctx.log(f"[HMM] no tracking snapshot on pier2pier for {clean}")
            return []

        iframe = self._http_get(opener, iframe_url)
        events = self._parse_movements(ctx, clean, iframe or "")
        ctx.log(f"[HMM] {clean}: {len(events)} movement rows")
        return events

    @staticmethod
    def _http_get(opener, url: str) -> str:
        req = Request(
            url,
            headers={
                "User-Agent": P2P_USER_AGENT,
                "Accept-Language": "en-US,en;q=0.9",
            },
        )
        with opener.open(req, timeout=30) as resp:
            return resp.read().decode("utf-8", errors="replace")

    def _parse_movements(self, ctx: ModuleContext, container_no: str, html: str) -> list:
        """Parse the HMM tracking snapshot into location+date events.

        The iframe contains, among other tables:

        * a "Shipment Schedule" table whose columns are Origin / Loading Port /
          Discharging Port / Destination, with an ``Arrival(ETB)`` row holding
          the per-port ETA and a ``Location`` row holding the port names;
        * a "Vessel Movement" list of ``vessel | POL | date | POD | date`` rows;
        * a "Shipment History" table of actual moves (date | time | location |
          status | ...).

        Every row that pairs a location with a date becomes an event so the
        shared latest-per-port mapping can fill the ETA columns.
        """
        body = re.sub(r"<script.*?</script>", " ", html, flags=re.S | re.I)
        body = re.sub(r"</tr>", "\n", body, flags=re.I)
        body = re.sub(r"</t[dh]>", " | ", body, flags=re.I)
        body = re.sub(r"<[^>]+>", " ", body)
        body = html_unescape(body)
        lines = [ln.strip() for ln in body.splitlines() if ln.strip()]

        events: list[dict] = []

        # ---- Shipment Schedule: Location row + Arrival(ETB) row ----
        schedule_idx = next(
            (i for i, ln in enumerate(lines) if "SHIPMENT SCHEDULE" in ln.upper() and "ORIGIN" in ln.upper()),
            None,
        )
        if schedule_idx is not None:
            headers = [c.strip().upper() for c in lines[schedule_idx].split("|")]
            loc_row = None
            arr_row = None
            for ln in lines[schedule_idx + 1:]:
                cells = [c.strip() for c in ln.split("|")]
                if len(cells) < len(headers):
                    continue
                label = cells[0].upper()
                if "LOCATION" in label:
                    loc_row = cells
                elif "ARRIVAL" in label or "ETB" in label:
                    arr_row = cells
                    break
            if loc_row and arr_row:
                for i in range(1, len(arr_row)):
                    col_name = headers[i] if i < len(headers) else ""
                    location = loc_row[i] if i < len(loc_row) else ""
                    date = self._format_date(arr_row[i])
                    if location and date:
                        events.append({
                            "container_no": container_no,
                            "date": date,
                            "location": location,
                            "description": f"Schedule {col_name} arrival",
                        })

        # ---- Vessel Movement rows: vessel | voyage | POL | pol_date | POD | pod_date ----
        for ln in lines:
            if any(ln.upper().startswith(prefix) for prefix in
                   ("VESSEL /", "VESSEL/V", "VESSEL MOVEMENT")):
                continue
            parts = [p.strip() for p in ln.split("|") if p.strip()]
            if len(parts) < 5:
                continue
            # locate two dates inside the parts
            date_idx = [i for i, p in enumerate(parts) if DATE_RE.search(p)]
            if len(date_idx) < 2:
                continue
            for i in date_idx:
                date = self._format_date(parts[i])
                # location = nearest non-date token left of the date
                location = ""
                for p in reversed(parts[:i]):
                    if not DATE_RE.search(p) and len(p) > 3:
                        location = p
                        break
                if date and location:
                    events.append({
                        "container_no": container_no,
                        "date": date,
                        "location": location,
                        "description": "Vessel movement",
                    })

        # ---- Shipment History rows: date | time | location | status ... ----
        history_idx = next(
            (i for i, ln in enumerate(lines) if "SHIPMENT HISTORY" in ln.upper()),
            None,
        )
        if history_idx is not None:
            for ln in lines[history_idx + 1:]:
                if "TRACKING RESULTS" in ln.upper() or "SHIPMENT SCHEDULE" in ln.upper():
                    break
                parts = [p.strip() for p in ln.split("|") if p.strip()]
                if len(parts) < 3:
                    continue
                m = DATE_RE.search(parts[0])
                if not m:
                    continue
                date = self._format_date(m.group(0))
                location = parts[2] if len(parts) > 2 else ""
                status = parts[3] if len(parts) > 3 else ""
                if date and location:
                    events.append({
                        "container_no": container_no,
                        "date": date,
                        "location": location,
                        "description": status or "Shipment history",
                    })

        return events

    @staticmethod
    def _format_date(date_text: str) -> str:
        match = DATE_RE.search(date_text or "")
        if not match:
            return ""
        year, month, day = match.group(0).split("-")
        return f"{month}/{day}/{year}"

    def _checkpoint(self, df, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_ETA.xlsx"
            df.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")