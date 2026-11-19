"""Hapag-Lloyd tracking module (ported).

Reads a container Excel (``Container No`` + ``Liner``), tracks each
Hapag-Lloyd container, and writes the per-port ETAs back to the ETA columns.

Source order (best → fallback):
  1. Official Hapag-Lloyd DCSA Track & Trace API (env `HL_API_CLIENT_ID` /
     `HL_API_CLIENT_SECRET`; free subscription at api-portal.hlag.com).
  2. pier2pier.com via the runner's local Selenium browser (survives the
     Turnstile challenge; no server-side browser needed).
  3. pier2pier via Playwright (opt-in, `HAPAG_USE_PLAYWRIGHT=1`).
  4. pier2pier via plain HTTP (legacy path; Turnstile usually blocks it).
  5. Direct Hapag-Lloyd booking page for BL-only rows (usually Cloudflare-
     blocked from datacenter networks).
"""

from __future__ import annotations

import json
import logging
import os
import re
import time
from html import unescape as html_unescape
from http.cookiejar import CookieJar
from urllib.request import HTTPCookieProcessor, Request, build_opener

try:  # optional: Playwright-based pier2pier fetch (env-gated)
    from playwright.sync_api import sync_playwright  # type: ignore
    _PLAYWRIGHT_AVAILABLE = True
except Exception:  # pragma: no cover
    _PLAYWRIGHT_AVAILABLE = False

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit

logger = logging.getLogger(__name__)

TRACKING_COLUMNS = ["Container No", "BL", "Liner", "Date", "Location", "Description"]

# Output column -> location keywords found in an event's location/description.
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

HAPAG_BL_URL = ("https://www.hapag-lloyd.com/en/online-business/track/"
                "track-by-booking-solution.html?blno={bl}")
CLOUDFLARE_INDICATORS = [
    "just a moment", "security check", "checking your browser",
    "cf-browser-verification", "cdn-cgi/challenge-platform", "cloudflare",
    "managed challenge", "turnstile",
]

CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$")
DATE_RE = re.compile(r"(\d{2}[/-]\d{2}[/-]\d{4}|\d{4}-\d{2}-\d{2})")

# Option 2: official Hapag-Lloyd DCSA Track & Trace API (api.hlag.com).
# Requires a free subscription on api-portal.hlag.com; credentials go in env:
#   HL_API_CLIENT_ID / HL_API_CLIENT_SECRET
HL_API_URL = "https://api.hlag.com/hlag/external/v2/events/"
HL_API_ENV_ID = "HL_API_CLIENT_ID"
HL_API_ENV_SECRET = "HL_API_CLIENT_SECRET"

# Option 1: Playwright fallback for pier2pier (needs browser installed).
HAPAG_PW_ENV = "HAPAG_USE_PLAYWRIGHT"

MODULE = Module(
    key="HapagLloyd",
    name="Hapag-Lloyd",
    description="Fetch Hapag-Lloyd ETA per destination port via pier2pier.com.",
    icon="🚢",
    category="Arrival Status",
    requires_browser=True,
    requires_credentials=False,
    ported=True,
)


class HapagLloydModule(WebModule):
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

        bl_col = "BL Number" if "BL Number" in df.columns else ("BL" if "BL" in df.columns else None)
        has_ctr = "Container No" in df.columns
        if not bl_col and not has_ctr:
            raise ModuleError("The Excel file must contain a 'BL Number'/'BL' or 'Container No' column.")
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
                bl_no = self._cell(row, bl_col) if bl_col else ""
                container_no = self._cell(row, "Container No") if has_ctr else ""
                liner = self._cell(row, "Liner") if has_liner else ""
                key = bl_no or container_no
                if not key:
                    continue

                if skip_done and self._all_filled(row):
                    ctx.log(f"Skipping {key} - all ETA columns already filled")
                    continue

                if has_liner and liner and "hapag" not in liner.lower():
                    ctx.log(f"{key}: not Hapag-Lloyd (liner={liner})")
                    continue

                progress = 5 + int((index + 1) / total * 90)
                ctx.progress(progress, f"Tracking {key}")
                ctx.log(f"[Hapag] Processing {key}")

                try:
                    events = self._track_hybrid(ctx, container_no, bl_no)
                except Exception as exc:
                    ctx.log(f"[Hapag] {key}: {exc}")
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

        A port can appear in several movement rows (arrival, discharge, load,
        departure...). The user wants the *latest* date per port, so a later
        date always replaces an earlier one, regardless of milestone.
        """
        from datetime import datetime

        best: dict = {}

        def _key(date: str):
            for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(date, fmt)
                except Exception:
                    continue
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

    # ------------------------------------------------------------- helpers
    def _format_date(self, date_text: str) -> str:
        try:
            date_text = str(date_text).strip()
            if not date_text or not any(c.isdigit() for c in date_text):
                return date_text
            if "/" in date_text:
                parts = date_text.split("/")
                if len(parts) == 3:
                    a, b, year = [p.strip() for p in parts]
                    if year.isdigit() and len(year) == 4:
                        return f"{b.zfill(2)}/{a.zfill(2)}/{year}"
                return date_text
            if "-" in date_text:
                parts = date_text.split("-")
                if len(parts) == 3:
                    a, b, c = [p.strip() for p in parts]
                    if len(a) == 4 and a.isdigit() and int(a) > 1900:
                        return f"{b}/{c}/{a}"
                    if len(c) == 4 and c.isdigit() and int(c) > 1900:
                        return f"{b}/{a}/{c}"
                return date_text
            try:
                from dateutil.parser import parse

                return parse(date_text).strftime("%m/%d/%Y")
            except Exception:
                return date_text
        except Exception:
            return str(date_text)

    # ------------------------------------------------------------- tracking
    def _track_hybrid(self, ctx: ModuleContext, container_no: str, bl_no: str) -> list:
        """Container first via pier2pier; BL-only rows fall back to the direct
        Hapag-Lloyd booking page (which may be Cloudflare-blocked from this
        network).

        Source order (best → fallback):
          1. Official HL DCSA API          (if HL_API_CLIENT_ID/SECRET env set)
          2. pier2pier via Selenium        (runner's local browser; survives Turnstile)
          3. pier2pier via Playwright      (if HAPAG_USE_PLAYWRIGHT=1)
          4. pier2pier via plain HTTP      (existing path)
          5. Direct HL booking page         (BL-only rows)
        """
        clean = "".join((container_no or "").split())

        api_id = os.getenv(HL_API_ENV_ID, "").strip()
        api_secret = os.getenv(HL_API_ENV_SECRET, "").strip()
        if clean and api_id and api_secret:
            events = self._track_carrier_api(ctx, clean)
            if events:
                return events
            ctx.log("[Hapag] carrier API returned no events; falling back")
        elif clean and api_id != api_secret:
            ctx.log("[Hapag] carrier API env incomplete (need both HL_API_CLIENT_ID and HL_API_CLIENT_SECRET)")

        if clean:
            events = self._track_p2p_selenium(ctx, clean)
            if events:
                return events
            ctx.log("[Hapag] Selenium pier2pier returned nothing; trying next source")

        if clean and _PLAYWRIGHT_AVAILABLE and self._env_flag(HAPAG_PW_ENV):
            events = self._track_p2p_pw(ctx, clean)
            if events:
                return events
            ctx.log("[Hapag] Playwright pier2pier returned nothing; falling back to plain HTTP")

        if clean:
            return self._track_p2p(ctx, clean, bl_no)

        if bl_no and len(bl_no) > 5:
            ctx.log("[Hapag] BL-only row: trying direct Hapag-Lloyd booking page")
            events = self._track_direct_bl(ctx, bl_no)
            if events:
                return events
            ctx.log("[Hapag] direct BL blocked/no data; provide a Container No for pier2pier")
        return []

    def _track_carrier_api(self, ctx: ModuleContext, container_no: str) -> list:
        """Official Hapag-Lloyd DCSA Track & Trace API (option 2).

        Requires env `HL_API_CLIENT_ID` + `HL_API_CLIENT_SECRET` (free
        subscription on api-portal.hlag.com). No browser needed.
        """
        if not CONTAINER_RE.match(container_no):
            return []
        import urllib.request

        req = urllib.request.Request(
            HL_API_URL + "?equipmentReference=" + container_no,
            headers={
                "Accept": "application/json",
                "X-IBM-Client-Id": os.getenv(HL_API_ENV_ID, "").strip(),
                "X-IBM-Client-Secret": os.getenv(HL_API_ENV_SECRET, "").strip(),
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="replace"))
        except Exception as exc:
            ctx.log(f"[Hapag] carrier API failed: {str(exc)[:100]}")
            return []

        # DCSA T&T: {"events": [...]} or a list; each event has eventType,
        # eventDateTime, location.UNLocationCode/locationName, and milestone
        # codes (ARRI/DEPU/...). Keep every event that carries a date + place.
        raw = payload.get("events") if isinstance(payload, dict) else payload
        if not isinstance(raw, list):
            return []
        events: list[dict] = []
        for ev in raw:
            if not isinstance(ev, dict):
                continue
            when = ev.get("eventDateTime") or ev.get("eventCreatedDateTime") or ""
            date = self._format_date(when.split("T")[0] if when else "")
            loc = ev.get("location") or {}
            place = (loc.get("locationName") or loc.get("UNLocationCode") or "").strip()
            if not date and not place:
                continue
            events.append({
                "container_no": container_no,
                "date": date,
                "location": place,
                "description": str(ev.get("eventTypeCode") or ev.get("eventType") or ""),
            })
        return events

    def _track_p2p_selenium(self, ctx: ModuleContext, container_no: str) -> list:
        """pier2pier via the runner's local browser (Selenium).

        Runs on the user's machine where Chrome exists — no server-side
        browser needed. Waits for the Turnstile challenge to clear, then
        follows the ``Temp_Files`` iframe and parses the movement table.
        """
        driver = None
        try:
            from selenium.webdriver.common.by import By
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.webdriver.support.ui import WebDriverWait

            driver = create_driver(headless=True)
            ctx.log(f"[Hapag] Selenium: opening pier2pier for {container_no}")
            driver.get(P2P_TRACKING_URL.format(container=container_no))
            try:
                frame = WebDriverWait(driver, 25).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "iframe[src*='Temp_Files']")
                    )
                )
            except Exception:
                ctx.log(f"[Hapag] pier2pier Turnstile not cleared (Selenium) for {container_no}")
                return []
            iframe_url = frame.get_attribute("src") or ""
            if iframe_url.startswith("//"):
                iframe_url = "https:" + iframe_url
            driver.get(iframe_url)
            html = driver.page_source
            return self._parse_movements(ctx, container_no, html)
        except Exception as exc:
            ctx.log(f"[Hapag] Selenium pier2pier failed: {str(exc)[:100]}")
            return []
        finally:
            safe_quit(driver)

    def _track_p2p_pw(self, ctx: ModuleContext, container_no: str) -> list:
        """pier2pier via Playwright (option 1): drives a real browser past the
        Turnstile "Browser verification" page, then parses the same iframe
        content. Requires `playwright` + a Chromium browser on the host
        (e.g. `python -m playwright install chromium`).
        """
        url = P2P_TRACKING_URL.format(container=container_no)
        try:
            with sync_playwright() as pw:
                browser = pw.chromium.launch(
                    headless=True,
                    args=["--disable-blink-features=AutomationControlled"],
                )
                page = browser.new_page(user_agent=P2P_USER_AGENT)
                page.goto(url, timeout=45000, wait_until="domcontentloaded")
                # Turnstile may take a moment to clear; wait for the iframe.
                try:
                    page.wait_for_selector("iframe[src*='Temp_Files']", timeout=25000)
                except Exception:
                    browser.close()
                    ctx.log(f"[Hapag] pier2pier Turnstile not cleared for {container_no}")
                    return []
                iframe_url = page.eval_on_selector(
                    "iframe[src*='Temp_Files']", "el => el.src"
                )
                if iframe_url.startswith("//"):
                    iframe_url = "https:" + iframe_url
                page.goto(iframe_url, timeout=30000, wait_until="domcontentloaded")
                html = page.content()
                browser.close()
        except Exception as exc:
            ctx.log(f"[Hapag] Playwright pier2pier failed: {str(exc)[:100]}")
            return []
        return self._parse_movements(ctx, container_no, html)

    def _track_direct_bl(self, ctx: ModuleContext, bl_no: str) -> list:
        opener = build_opener(HTTPCookieProcessor(CookieJar()))
        page = ""
        try:
            page = self._http_get(opener, HAPAG_BL_URL.format(bl=bl_no))
        except Exception as exc:
            ctx.log(f"[Hapag] direct BL fetch failed: {str(exc)[:80]}")
            return []
        lower = (page or "").lower()
        if any(ind in lower for ind in CLOUDFLARE_INDICATORS):
            ctx.log("[Hapag] direct BL: Cloudflare challenge (blocked on this network)")
            return []

        body = re.sub(r"<script.*?</script>", " ", page, flags=re.S | re.I)
        body = re.sub(r"</tr>", "\n", body, flags=re.I)
        body = re.sub(r"</t[dh]>", " | ", body, flags=re.I)
        body = re.sub(r"<[^>]+>", " ", body)
        body = html_unescape(body)
        events: list[dict] = []
        for line in [ln.strip() for ln in body.splitlines() if ln.strip()]:
            match = DATE_RE.search(line)
            if not match:
                continue
            date = self._format_date(match.group(1))
            if not date:
                continue
            parts = [p.strip() for p in line.split("|") if p.strip()]
            location = ""
            for p in parts:
                if p != match.group(1) and len(p) > 2 and not self._is_date_only(p):
                    location = p
                    break
            events.append({
                "container_no": "",
                "date": date,
                "location": location,
                "description": "BL " + bl_no,
            })
        return events

    @staticmethod
    def _is_date_only(text: str) -> bool:
        return bool(DATE_RE.fullmatch(text.strip()))

    @staticmethod
    def _env_flag(name: str) -> bool:
        return os.getenv(name, "").strip().lower() in ("1", "true", "yes", "on")

    def _track_p2p(self, ctx: ModuleContext, container_no: str, bl_no: str) -> list:
        clean = "".join(container_no.split()) if container_no else ""
        if not clean:
            ctx.log("[Hapag] pier2pier resolves by container number only; BL/booking rows need a Container No")
            return []
        if not CONTAINER_RE.match(clean):
            ctx.log(f"[Hapag] not a valid container number: {clean}")
            return []

        opener = build_opener(HTTPCookieProcessor(CookieJar()))
        iframe_url = ""
        for _ in range(3):
            page = self._http_get(opener, P2P_TRACKING_URL.format(container=clean))
            match = P2P_IFRAME_RE.search(page or "")
            if match:
                iframe_url = match.group(1)
                if iframe_url.startswith("//"):
                    iframe_url = "https:" + iframe_url
                break
            time.sleep(1)
        if not iframe_url:
            ctx.log(f"[Hapag] no tracking data on pier2pier for {clean}")
            return []

        iframe = self._http_get(opener, iframe_url)
        events = self._parse_movements(ctx, clean, iframe or "")
        ctx.log(f"[Hapag] {clean}: {len(events)} movement rows")
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
        body = re.sub(r"<script.*?</script>", " ", html, flags=re.S | re.I)
        body = re.sub(r"</tr>", "\n", body, flags=re.I)
        body = re.sub(r"</t[dh]>", " | ", body, flags=re.I)
        body = re.sub(r"<[^>]+>", " ", body)
        body = html_unescape(body)
        lines = [ln.strip() for ln in body.splitlines() if ln.strip()]
        header_idx = next(
            (i for i, ln in enumerate(lines) if "Place of Activity" in ln and "Status" in ln),
            None,
        )
        if header_idx is None:
            return []
        events: list[dict] = []
        for line in lines[header_idx + 1:]:
            if "Bold data represents" in line:
                break
            parts = [p.strip() for p in line.split("|") if p.strip()]
            if len(parts) < 3:
                continue
            status, place, date_text = parts[0], parts[1], parts[2]
            match = DATE_RE.search(date_text)
            if not match:
                continue
            date = self._format_date(match.group(1))
            if not date:
                continue
            events.append({
                "container_no": container_no,
                "date": date,
                "location": place,
                "description": status,
            })
        return events

    def _checkpoint(self, df, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_ETA.xlsx"
            df.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")
