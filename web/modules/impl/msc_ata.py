"""MSC ATA module (ported, browser-based).

Reads a container Excel, tracks each MSC container on the public MSC
track-a-shipment site in a *local* browser on the runner machine and writes a
tracking-history sheet (Container No, BL, Liner, Date, Location, Description).

The desktop original used a thread pool of browsers; this port runs a single
browser sequentially so it is safe inside the runner process.
"""

from __future__ import annotations

import logging
import random
import re
import time
from pathlib import Path

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit

logger = logging.getLogger(__name__)

MSC_TRACKING_URL = "https://www.msc.com/en/track-a-shipment"
TRACKING_COLUMNS = ["Container No", "BL", "Liner", "Date", "Location", "Description"]
CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$", re.IGNORECASE)

# Output column -> location keywords found in an event's location/description.
LOCATION_COLUMNS = [
    ("JEBEL ALI ETA", ("JEBEL ALI", "JEBELALI", "JEBBEL ALI")),
    ("KHORFKHAN ETA", ("FAKKAN", "KHORFKHAN", "KHOR FKAN", "KFK")),
    ("SALLALAH ETA", ("SALALAH", "SALLALAH", "SLL")),
    ("ABU DHABI ETC", ("ABU DHABI", "ABUDHABI", "AUH")),
    ("FUJAIRAH ETA", ("FUJAIRAH", "FUJAYRAH")),
    ("JEDDAH ETA", ("JEDDAH", "JEDDA")),
]

MODULE = Module(
    key="MSC_ATA",
    name="MSC ATA",
    description="Fetch MSC ETA per destination port by BL or container.",
    icon="🗂️",
    category="Arrival Status",
    requires_browser=True,
    requires_credentials=False,
    ported=True,
)


class MSCATAModule(WebModule):
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

        driver = None
        try:
            ctx.progress(2, "Starting local browser")
            driver = create_driver(headless=True)
            ctx.log("Local MSC tracking session started")

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

                if has_liner and liner and "msc" not in liner.lower():
                    ctx.log(f"{key}: not MSC (liner={liner})")
                    continue

                progress = 5 + int((index + 1) / total * 90)
                ctx.progress(progress, f"Tracking {key}")
                ctx.log(f"[MSC] Processing {key}")

                try:
                    events = self._process_msc(ctx, driver, container_no or bl_no)
                except Exception as exc:
                    ctx.log(f"[MSC] {key}: {exc}")
                    events = []

                real_events = [e for e in events if "bill of lading" not in e.get("description", "").lower()]
                locations = self._map_events(real_events)
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
        finally:
            safe_quit(driver)

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
        best: dict = {}
        for event in events or []:
            date = event.get("date") or ""
            if not date:
                continue
            text = f"{event.get('location', '')} {event.get('description', '')}".upper()
            is_arrival = "ARRIVAL" in text
            for col, keywords in LOCATION_COLUMNS:
                if not any(k in text for k in keywords):
                    continue
                current = best.get(col)
                if current is None or (is_arrival and not current[1]):
                    best[col] = (date, is_arrival)
        return {col: date for col, (date, _) in best.items()}

    def _process_msc(self, ctx: ModuleContext, driver, container_no: str) -> list:
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        try:
            ctx.log(f"[MSC] Starting process for container: {container_no}")
            time.sleep(random.uniform(3, 7))

            driver.get(MSC_TRACKING_URL)
            ctx.log("[MSC] MSC tracking page opened")
            time.sleep(2)

            try:
                cookie_xpath = (
                    "//button[contains(text(), 'Accept') or contains(text(), 'AGREE') "
                    "or contains(text(), 'Accept All')]"
                )
                cookie_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, cookie_xpath))
                )
                cookie_btn.click()
                time.sleep(1)
            except Exception:
                pass

            time.sleep(1)

            try:
                radio = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//input[@type='radio' and @value='0']")
                    )
                )
                if not radio.is_selected():
                    driver.execute_script("arguments[0].click();", radio)
                    time.sleep(1)
            except Exception as exc:
                ctx.log(f"[MSC] Error selecting search-type radio: {exc}")

            try:
                input_el = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.ID, "trackingNumber"))
                )
                input_el.click()
                input_el.clear()
                time.sleep(0.5)
                for char in container_no:
                    input_el.send_keys(char)
                    time.sleep(0.05)
                time.sleep(1)
            except Exception as exc:
                ctx.log(f"[MSC] Error entering tracking number: {exc}")
                return []

            try:
                driver.execute_script(
                    "const f=document.querySelector('form.js-form'); if(f) f.requestSubmit();"
                )
                time.sleep(3)
            except Exception as exc:
                ctx.log(f"[MSC] Error submitting search: {exc}")
                return []

            try:
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".msc-flow-tracking__step"))
                )
                time.sleep(1)
            except TimeoutException:
                ctx.log(f"[MSC] No tracking result for {container_no}")
                return []

            return self._extract_tracking_events(ctx, driver, container_no)
        except Exception as exc:
            ctx.log(f"[MSC] Error processing container {container_no}: {exc}")
            return []

    def _extract_tracking_events(self, ctx: ModuleContext, driver, container_no: str) -> list:
        from selenium.webdriver.common.by import By

        events: list[dict] = []

        def cell(step, suffix: str) -> str:
            try:
                els = step.find_elements(
                    By.CSS_SELECTOR,
                    f".msc-flow-tracking__cell--{suffix} .data-value",
                )
                return els[0].text.strip() if els else ""
            except Exception:
                return ""

        try:
            for step in driver.find_elements(By.CSS_SELECTOR, ".msc-flow-tracking__step"):
                date_text = cell(step, "two")
                location = cell(step, "three")
                description = cell(step, "four")
                vessel = cell(step, "five")
                if not date_text and not location:
                    continue
                events.append({
                    "container_no": container_no,
                    "date": self._format_date(ctx, date_text),
                    "location": location,
                    "description": description,
                    "vessel_voyage": vessel,
                    "facility": "",
                })
        except Exception as exc:
            ctx.log(f"[MSC] Extraction error: {exc}")

        return events

    def _parse_event_cells(self, ctx: ModuleContext, cells, container_no: str):
        try:
            texts = [c.text.strip() for c in cells]
            event = {
                "container_no": container_no, "date": "", "location": "",
                "description": "", "vessel_voyage": "", "facility": "",
            }
            if texts:
                event["date"] = self._format_date(ctx, texts[0]) if "/" in texts[0] else texts[0]
            if len(texts) >= 2:
                event["location"] = texts[1]
            if len(texts) >= 3:
                event["description"] = texts[2]
            if len(texts) >= 4:
                event["vessel_voyage"] = texts[3]
            if len(texts) >= 5:
                event["facility"] = texts[4]
            return event
        except Exception:
            return None

    def _parse_text_event(self, ctx: ModuleContext, text_parts, container_no: str):
        try:
            event = {
                "container_no": container_no, "date": "", "location": "",
                "description": "", "vessel_voyage": "", "facility": "",
            }
            date_found = False
            non_date_parts = []
            for part in text_parts:
                if not date_found and "/" in part and len(part) >= 8 and any(c.isdigit() for c in part):
                    event["date"] = self._format_date(ctx, part)
                    date_found = True
                else:
                    non_date_parts.append(part)
            if date_found:
                if len(non_date_parts) >= 1:
                    event["location"] = non_date_parts[0]
                if len(non_date_parts) >= 2:
                    event["description"] = non_date_parts[1]
                if len(non_date_parts) >= 3:
                    event["vessel_voyage"] = non_date_parts[2]
                if len(non_date_parts) >= 4:
                    event["facility"] = non_date_parts[3]
            return event if date_found else None
        except Exception:
            return None

    def _format_date(self, ctx: ModuleContext, date_text: str) -> str:
        """Convert DD/MM/YYYY (MSC format) -> MM/DD/YYYY for consistency."""
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
        except Exception as exc:
            ctx.log(f"Error formatting date '{date_text}': {exc}")
            return date_text

    def _checkpoint(self, df, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_ETA.xlsx"
            df.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")
