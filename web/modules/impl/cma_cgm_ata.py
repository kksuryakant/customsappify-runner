"""CMA CGM ATA module (ported, browser-based).

Reads a shipment Excel, opens the public CMA CGM tracking site in a *local*
browser on the runner machine and extracts the ETA per destination port
(Jebel Ali / Khor Fakkan / Salalah / Abu Dhabi) for each CMA CGM B/L or
container, writing the results back to an Excel artifact. Each row is tracked
by *container number first*, falling back to the B/L only if the container
lookup returns no ETA.

The site is protected by DataDome: the first load may show a CAPTCHA. The
browser is opened *visible* so the operator can solve it once; the session
cookie then carries through the run. This module does **not** use a proxy.
"""

from __future__ import annotations

import datetime
import logging
import os
import re
import time
from pathlib import Path
from typing import Optional

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit


def _cma_profile_dir() -> str:
    """Persistent Chrome profile so a solved DataDome cookie survives runs."""
    path = os.environ.get("APPIFY_CMA_PROFILE") or str(Path.home() / ".appify" / "cma_profile")
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass
    return path


logger = logging.getLogger(__name__)

CMA_TRACKING_URL = "https://www.cma-cgm.com/ebusiness/tracking"
DATADOME_MARKERS = ("captcha-delivery", "geo.captcha-delivery", "DataDome")

# Carrier modules normally run headless; CMA CGM is DataDome-gated, so we open a
# visible window by default and let the operator solve the CAPTCHA once.
FORCE_HEADLESS = os.environ.get("APPIFY_CMA_HEADLESS", "").lower() in ("1", "true", "yes")

MONTHS = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
    "jul": "07", "aug": "08", "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}

LOCATION_COLUMNS = [
    ("JEBEL ALI ETA", ("JEBEL ALI", "JEBELALI", "JEBBEL ALI", "SHARJAH", "DUBAI")),
    ("KHORFKHAN ETA", ("FAKKAN", "KHORFKHAN", "KHOR FKAN", "KFK")),
    ("SALLALAH ETA", ("SALALAH", "SALLALAH", "SLL")),
    ("ABU DHABI ETC", ("ABU DHABI", "ABUDHABI", "AUH")),
    ("FUJAIRAH ETA", ("FUJAIRAH", "FUJAYRAH", "FJR")),
    ("JEDDAH ETA", ("JEDDAH", "JEDDA")),
]

MODULE = Module(
    key="CMA_CGM_ATA",
    name="CMA CGM ATA",
    description="Fetch CMA CGM ETA per destination port by B/L or container.",
    icon="🛳️",
    category="Arrival Status",
    requires_browser=True,
    requires_credentials=False,
    ported=True,
)


class CMACGMATAModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        excel_files = list(ctx.uploaded.get("excel_file", []))
        if not excel_files:
            raise ModuleError("An Excel file with B/L or container data is required.")
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

        driver = None
        try:
            ctx.progress(2, "Starting local browser")
            profile_dir = _cma_profile_dir()
            try:
                driver = create_driver(headless=FORCE_HEADLESS, user_data_dir=profile_dir, no_proxy=True)
            except Exception as exc:
                raise ModuleError(f"Failed to start Chrome (profile locked? close any open Chrome using this profile): {exc}") from exc
            ctx.log("Local CMA CGM tracking session started")
            ctx.log(f"[CMA] Using persistent profile: {profile_dir}")


            processed = 0
            for index, row in df.iterrows():
                bl_no = self._cell(row, bl_col) if bl_col else ""
                container_no = self._cell(row, "Container No") if has_ctr else ""
                liner = self._cell(row, "Liner") if has_liner else ""
                key = container_no or bl_no
                if not key:
                    continue
                if has_liner and liner and "cma" not in liner.lower():
                    ctx.log(f"{key}: not CMA CGM (liner={liner})")
                    continue
                if skip_done and self._all_filled(row):
                    ctx.log(f"Skipping {key} - all ETA columns already filled")
                    continue

                # Container first; fall back to the B/L only if the container
                # lookup yields no ETA.
                candidates: list[tuple[str, str]] = []
                if container_no:
                    candidates.append(("container", container_no))
                if bl_no:
                    candidates.append(("bl", bl_no))

                progress = 5 + int((index + 1) / total * 90)
                locations: dict = {}
                for pos, (mode, ref) in enumerate(candidates):
                    ctx.progress(progress, f"Tracking {ref} ({mode})")
                    ctx.log(f"[CMA] Processing {ref} ({mode})")
                    try:
                        events = self._track(ctx, driver, ref, mode)
                    except Exception as exc:
                        msg = str(exc)
                        ctx.log(f"[CMA] {ref}: {msg}")
                        if "no such window" in msg or "chrome not reachable" in msg or "target window already closed" in msg:
                            raise ModuleError(
                                f"Browser session was closed ({ref}). Please run again - the persistent profile may retain the CAPTCHA cookie."
                            ) from exc
                        events = []
                    locations = self._map_events(events)
                    if locations:
                        break
                    if pos + 1 < len(candidates):
                        nxt = candidates[pos + 1][1]
                        ctx.log(f"[CMA] {ref}: no ETA - falling back to {nxt}")

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
            ctx.progress(100, "Done")
            ctx.log(f"Done. Processed {processed} shipment(s).")
        except Exception:
            raise
        finally:
            try:
                safe_quit(driver)
            except Exception:
                pass



    # ------------------------------------------------------------------ browser

    def _open_tracking(self, ctx: ModuleContext, driver) -> None:
        from selenium.webdriver.common.by import By

        driver.get(CMA_TRACKING_URL)
        # small human-like move if pyautogui available
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            pyautogui.moveTo(400 + int(time.time() % 50), 300 + int(time.time() % 30), duration=0.5)
        except Exception:
            pass
        deadline = time.time() + 150
        warned = False
        while time.time() < deadline:
            try:
                if driver.find_elements(By.ID, "Reference"):
                    try:
                        import pyautogui

                        pyautogui.FAILSAFE = False
                        pyautogui.moveTo(500 + int(time.time() % 40), 350 + int(time.time() % 25), duration=0.3)
                    except Exception:
                        pass
                    return
                if not warned:
                    html = driver.page_source or ""
                    if any(m in html for m in DATADOME_MARKERS):
                        ctx.log(
                            "[CMA] DataDome CAPTCHA detected - please solve it in the "
                            "browser window (waiting up to 150s)"
                        )
                        warned = True
            except Exception:
                pass
            time.sleep(3)
        ctx.log("[CMA] tracking page did not become ready (CAPTCHA not solved in time?)")

    def _track(self, ctx: ModuleContext, driver, key: str, mode: str) -> list:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        self._open_tracking(ctx, driver)
        if not driver.find_elements(By.ID, "Reference"):
            ctx.log(f"[CMA] {key}: tracking page not ready (CAPTCHA unsolved / blocked)")
            return []
        search_by = "BL" if mode == "bl" else "Container"
        try:
            # Set hidden field
            driver.execute_script(
                "var e=document.getElementById('SearchBy'); if(e){e.value=arguments[0]; e.dispatchEvent(new Event('change',{bubbles:true}));}"
            )
            # Click any toggle with data-searchby if present (live pages)
            driver.execute_script(
                "var sb=arguments[0];"
                "document.querySelectorAll('[data-searchby]').forEach(el=>{"
                "  if((el.getAttribute('data-searchby')||'').toUpperCase()===sb.toUpperCase()){el.click();}"
                "});", search_by
            )
        except Exception:
            pass

        field = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "Reference")))
        field.clear()
        field.send_keys(key)
        time.sleep(0.6)
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            pyautogui.moveTo(550, 420, duration=0.3)
            pyautogui.click()
        except Exception:
            pass
        try:
            driver.execute_script("document.getElementById('btnTracking').click();")
        except Exception:
            field.submit()

        result = self._wait_results(ctx, driver)
        # Try to expand "Display Previous Moves"/"Show Previous Moves"
        try:
            driver.execute_script("""
                const targets = document.querySelectorAll('button, a, span, div[role="button"]');
                for (let el of targets) {
                  const t = (el.textContent || '').toLowerCase();
                  if (t.includes('display previous moves') || t.includes('show previous moves') || t.includes('hide previous moves') || t.includes('previous moves')) {
                    el.click(); break;
                  }
                }
            """)
            time.sleep(1.8)
        except Exception:
            pass
        # Also click any element that might control moves expansion via class
        try:
            driver.execute_script("""
                document.querySelectorAll('[class*="moves"], [class*="toggle"], [data-toggle*="moves"]').forEach(el=>{
                  el.click();
                });
            """)
            time.sleep(1.0)
        except Exception:
            pass
        # Re-read results after expansion
        try:
            result = self._wait_results(ctx, driver, timeout=20) or result
        except Exception:
            pass
        if result.get("notFound"):
            ctx.log(f"[CMA] {key}: shipment not found")
            return []
        data = result.get("data")
        if not data:
            ctx.log(f"[CMA] {key}: no tracking data returned")
            return []
        return self._parse(ctx, key, data)

    def _wait_results(self, ctx: ModuleContext, driver, timeout: int = 60) -> dict:
        js = r"""
        const o = window.options || {};
        let data = null;
        try { data = JSON.parse(o.responseData || 'null'); } catch (e) {}
        const alert = document.querySelector('#trackingAlertError');
        const notFound = !!(alert && alert.offsetParent !== null &&
                            getComputedStyle(alert).display !== 'none');
        const hasResults = !!document.querySelector('.result-card') ||
            !!(data && (data.SearchCriteria ||
                 (data.PastMoves && data.PastMoves.length) ||
                 (data.CurrentMoves && data.CurrentMoves.length) ||
                 (data.ProvisionalMoves && data.ProvisionalMoves.length)));
        return { data: data, notFound: notFound, hasResults: hasResults };
        """
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                r = driver.execute_script(js)
                if r and (r.get("notFound") or r.get("hasResults")):
                    return r
            except Exception:
                pass
            time.sleep(1.5)
        try:
            return driver.execute_script(js) or {}
        except Exception:
            return {}

    # -------------------------------------------------------------------- parse

    def _parse(self, ctx: ModuleContext, key: str, data: dict) -> list:
        events = []
        for bucket in ("PastMoves", "CurrentMoves", "ProvisionalMoves"):
            for mv in data.get(bucket) or []:
                loc = (mv.get("Location") or "").strip()
                term = (mv.get("LocationTerminal") or "").strip()
                date = self._parse_date(mv.get("DateString"))
                if not date or not loc:
                    continue
                state = (mv.get("State") or "").upper()
                events.append({
                    "container_no": key,
                    "date": date,
                    "location": f"{loc} {term}".strip(),
                    "description": f"{mv.get('Vessel', '')} {mv.get('Voyage', '')} {state}".strip(),
                    "is_arrival": state == "DONE",
                })

        dest = (data.get("Destination") or "").strip()
        eta = (
            self._parse_date(data.get("ActualTimeOfArrival"))
            or self._parse_date(data.get("EstimatedTimeOfArrival"))
        )
        if dest and eta:
            events.append({
                "container_no": key,
                "date": eta,
                "location": dest,
                "description": "Destination",
                "is_arrival": True,
            })
        # Also scan rendered page for arrival-like text if data sparse
        try:
            events.extend(self._parse_from_dom(ctx.driver if hasattr(ctx, 'driver') else None, key))
        except Exception:
            pass
        return events

    @staticmethod
    def _parse_date(value) -> str:
        if value is None:
            return ""
        v = str(value).strip()
        if not v:
            return ""

        ms = re.search(r"/Date\((-?\d+)", v)
        if ms:
            try:
                ts = int(ms.group(1)) / 1000.0
                return datetime.datetime.fromtimestamp(ts).strftime("%m/%d/%Y")
            except Exception:
                return ""

        m = re.search(r"(\d{4})-(\d{2})-(\d{2})", v)
        if m:
            y, mo, d = m.groups()
            return f"{mo}/{d}/{y}"

        m = re.search(r"(\d{1,2})[-/ ]([A-Za-z]{3,})[-/ ](\d{4})", v)
        if m:
            d, mon, y = m.groups()
            mo = MONTHS.get(mon[:3].lower())
            if mo:
                return f"{mo}/{int(d):02d}/{y}"

        m = re.search(r"([A-Za-z]{3,})[-, ]+(\d{1,2}),?[- ]+(\d{4})", v)
        if m:
            mon, d, y = m.groups()
            mo = MONTHS.get(mon[:3].lower())
            if mo:
                return f"{mo}/{int(d):02d}/{y}"
        return ""

    def _map_events(self, events: list) -> dict:
        locations: dict = {}
        for ev in events:
            loc = (ev.get("location") or "").upper()
            date = ev.get("date") or ""
            if not loc or not date:
                continue
            for col, tokens in LOCATION_COLUMNS:
                if any(tok in loc for tok in tokens):
                    if not locations.get(col) or ev.get("is_arrival"):
                        locations[col] = date
        return locations

    def _parse_from_dom(self, driver, key: str) -> list:
        events = []
        if driver is None:
            return events
        try:
            from selenium.webdriver.common.by import By
            html = driver.page_source or ""
            # look for dates like DD-MMM-YYYY
            for m in re.finditer(r'(\d{1,2})-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(\d{4})', html, re.I):
                d, mon, y = m.groups()
                mo = MONTHS.get(mon[:3].lower())
                if not mo:
                    continue
                date = f"{mo}/{int(d):02d}/{y}"
                # get surrounding text snippet
                start = max(0, m.start()-40); end = min(len(html), m.end()+80)
                snip = html[start:end].lower()
                loc = ''
                for col, toks in LOCATION_COLUMNS:
                    for tok in toks:
                        t = tok.lower()
                        if t in snip or tok in html[m.start()-60:m.end()+80].upper():
                            loc = col  # store mapped column name
                            break
                    if loc:
                        break
                if loc and any(k in snip for k in ['discharged','arrival','arrive','delivery','fpd','vessel arrival']):
                    events.append({'container_no': key, 'date': date, 'location': loc, 'is_arrival': True})
        except Exception:
            pass
        return events

    @staticmethod
    def _cell(row, column) -> str:
        if not column:
            return ""
        value = row.get(column)
        if value is None:
            return ""
        try:
            if pd.isna(value):
                return ""
        except Exception:
            pass
        return str(value).strip()

    def _all_filled(self, row) -> bool:
        return all(self._cell(row, col) for col, _ in LOCATION_COLUMNS)

    def _checkpoint(self, df, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_ETA.xlsx"
            df.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")
