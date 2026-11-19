"""Evergreen ATA module (ported, browser-based).

Reads a container Excel, tracks each Evergreen container on the public
shipmentlink cargo-tracking site in a *local* browser on the runner machine,
extracts the estimated/actual arrival date and writes the ``ATA`` column back
to an Excel artifact.
"""

from __future__ import annotations

import logging
import random
import re
import time
from datetime import datetime
from pathlib import Path

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit

logger = logging.getLogger(__name__)

EVERGREEN_TRACKING_URL = "https://ct.shipmentlink.com/servlet/TDB1_CargoTracking.do"
CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$", re.IGNORECASE)
DATE_RE = re.compile(r"([A-Z]{3}-\d{2}-\d{4})|(\d{2}[/-][A-Z]{3}[/-]\d{4})|(\d{4}-\d{2}-\d{2})")

LOCATION_COLUMNS = [
    ("JEBEL ALI ETA", ("JEBEL ALI", "JEBELALI", "JEBBEL ALI")),
    ("KHORFKHAN ETA", ("FAKKAN", "KHORFKHAN", "KHOR FKAN")),
    ("SALLALAH ETA", ("SALALAH", "SALLALAH", "SLL")),
    ("ABU DHABI ETC", ("ABU DHABI", "ABUDHABI", "AUH")),
    ("FUJAIRAH ETA", ("FUJAIRAH", "FUJAYRAH")),
    ("JEDDAH ETA", ("JEDDAH", "JEDDA")),
]


MODULE = Module(
    key="EVERGREEN_ATA",
    name="Evergreen ATA",
    description="Fetch Evergreen ETA per destination port by BL or container.",
    icon="🌲",
    category="Arrival Status",
    requires_browser=True,
    requires_credentials=False,
    ported=True,
)


class EvergreenATAModule(WebModule):
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
            ctx.log("Local Evergreen tracking session started")

            processed = 0
            for index, row in df.iterrows():
                bl_no = self._cell(row, bl_col) if bl_col else ""
                container_no = self._cell(row, "Container No") if has_ctr else ""
                liner = self._cell(row, "Liner") if has_liner else ""
                key = bl_no or container_no
                if not key:
                    continue
                if has_liner and liner and "evergreen" not in liner.lower():
                    ctx.log(f"{key}: not Evergreen (liner={liner})")
                    continue
                if skip_done and self._all_filled(row):
                    ctx.log(f"Skipping {key} - all ETA columns already filled")
                    continue

                mode = "bl" if bl_no else "container"
                progress = 5 + int((index + 1) / total * 90)
                ctx.progress(progress, f"Tracking {key} ({mode})")
                ctx.log(f"[Evergreen] Processing {key} ({mode})")

                try:
                    events = self._track_evergreen(ctx, driver, key, mode)
                except Exception as exc:
                    ctx.log(f"[Evergreen] {key}: {exc}")
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
            is_arrival = "ARRIV" in text
            for col, keywords in LOCATION_COLUMNS:
                if not any(k in text for k in keywords):
                    continue
                current = best.get(col)
                if current is None or (is_arrival and not current[1]):
                    best[col] = (date, is_arrival)
        return {col: date for col, (date, _) in best.items()}

    @staticmethod
    def _format_date(ctx, date_text: str) -> str:
        date_text = (date_text or "").strip()
        for fmt in ("%b-%d-%Y", "%d-%b-%Y", "%d/%b/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(date_text.upper(), fmt).strftime("%m/%d/%Y")
            except ValueError:
                continue
        return date_text

    def _track_evergreen(self, ctx: ModuleContext, driver, key: str, mode: str) -> list:
        from selenium.common.exceptions import TimeoutException, UnexpectedAlertPresentException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        radio_id = "s_bl" if mode == "bl" else "s_cntr"
        driver.get(EVERGREEN_TRACKING_URL)
        time.sleep(random.uniform(1, 3))
        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        except TimeoutException:
            pass

        try:
            radio = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.ID, radio_id))
            )
            driver.execute_script("arguments[0].click();", radio)
        except Exception:
            ctx.log("Could not find search radio, continuing")

        input_no = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.ID, "NO")))
        input_no.clear()
        time.sleep(0.4)
        for char in key:
            input_no.send_keys(char)
            time.sleep(random.uniform(0.05, 0.15))

        time.sleep(random.uniform(1, 2))
        driver.execute_script("frmSubmit(13,2);")
        time.sleep(3)

        try:
            WebDriverWait(driver, 2).until(EC.alert_is_present())
            text = driver.switch_to.alert.text
            driver.switch_to.alert.accept()
            ctx.log(f"[Evergreen] {key}: site says '{text}'")
            return []
        except Exception:
            pass

        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.globalpage, table.ec-table-sm"))
            )
            time.sleep(1)
        except TimeoutException:
            ctx.log(f"[Evergreen] no result table for {key}")
            return []

        return self._parse_evergreen(ctx, driver, key)

    def _parse_evergreen(self, ctx: ModuleContext, driver, key: str) -> list:
        events: list[dict] = []
        seen = set()
        try:
            rows = driver.execute_script(
                """
                const out=[];
                document.querySelectorAll('tr').forEach(tr=>{
                    const leaf=[...tr.children].filter(
                        c=>c.tagName==='TD' && !c.querySelector('td'));
                    const texts=leaf.map(td=>(td.textContent||'').replace(/\\s+/g,' ').trim());
                    if(texts.filter(t=>t).length>=4) out.push(texts);
                });
                return out;
                """
            )
            for cells in rows:
                date = ""
                date_idx = -1
                for idx, c in enumerate(cells):
                    if DATE_RE.search(c.upper()):
                        date = c
                        date_idx = idx
                        break
                if not date:
                    continue
                location = ""
                for c in cells:
                    cs = c.strip()
                    if re.match(r"^[A-Za-z]", cs) and re.search(r"[A-Za-z]\s*\([A-Z]{2}\)", cs.upper()):
                        location = c
                        break
                if not location and date_idx > 0:
                    location = cells[date_idx - 1]
                desc = cells[3] if len(cells) > 3 else (cells[1] if len(cells) > 1 else "")
                marker = (date, location)
                if marker in seen:
                    continue
                seen.add(marker)
                events.append({
                    "container_no": key,
                    "date": self._format_date(ctx, DATE_RE.search(date.upper()).group(0)),
                    "location": location,
                    "description": desc,
                    "vessel_voyage": "",
                    "facility": "",
                })
        except Exception as exc:
            ctx.log(f"[Evergreen] parse error: {exc}")
        return events

    def _process_evergreen(self, ctx: ModuleContext, driver, wait_timeout: int, container_no: str) -> str:
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        max_retries = 3
        operation_timeout = 45

        for attempt in range(max_retries):
            try:
                start = time.time()
                ctx.log(f"Processing Evergreen container {container_no} (attempt {attempt + 1}/{max_retries})")
                driver.get(EVERGREEN_TRACKING_URL)
                time.sleep(random.uniform(1, 3))

                try:
                    WebDriverWait(driver, 3).until(EC.alert_is_present())
                    driver.switch_to.alert.accept()
                    time.sleep(1)
                except TimeoutException:
                    pass

                WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

                try:
                    radio = WebDriverWait(driver, wait_timeout).until(
                        EC.element_to_be_clickable((By.ID, "s_cntr"))
                    )
                    driver.execute_script("arguments[0].click();", radio)
                except Exception:
                    ctx.log("Could not find container radio button, continuing")

                try:
                    input_no = WebDriverWait(driver, wait_timeout).until(
                        EC.element_to_be_clickable((By.ID, "NO"))
                    )
                    input_no.clear()
                    time.sleep(0.5)
                    for char in container_no:
                        input_no.send_keys(char)
                        time.sleep(random.uniform(0.05, 0.2))
                except Exception as exc:
                    ctx.log(f"Error entering container number: {exc}")
                    if attempt < max_retries - 1:
                        continue
                    return "Error: Input failed"

                time.sleep(random.uniform(1, 2))
                try:
                    driver.execute_script("frmSubmit(13,2);")
                except Exception as exc:
                    ctx.log(f"Form submission error: {exc}")
                    if attempt < max_retries - 1:
                        continue
                    return "Error: Submission failed"

                time.sleep(2)
                if time.time() - start > operation_timeout:
                    raise TimeoutException("Operation timed out")

                try:
                    eta_td = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//td[contains(text(), 'Estimated Date of Arrival')]")
                        )
                    )
                    date_text = eta_td.get_attribute("innerHTML")
                    if "<br>" in date_text:
                        date_text = date_text.split("<br>")[-1].strip()
                    if date_text:
                        try:
                            dt = datetime.strptime(date_text, "%b-%d-%Y")
                            return dt.strftime("%d/%m/%Y")
                        except ValueError:
                            return date_text
                    return "Not Found"
                except TimeoutException:
                    ctx.log("ETA element not found, checking status table")

                    try:
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, "//table[contains(@class, 'ec-table-sm')]")
                            )
                        )
                        xpath = (
                            "//table[contains(@class,'ec-table-sm')]"
                            f"//tr[td[1][normalize-space()='{container_no}']]/td[3]"
                        )
                        date_cell = driver.find_element(By.XPATH, xpath).text.strip()
                        if date_cell:
                            try:
                                dt = datetime.strptime(date_cell, "%b-%d-%Y")
                                return dt.strftime("%d/%m/%Y")
                            except ValueError:
                                return date_cell
                        return "Date not found in table"
                    except Exception as exc:
                        detail = next((ln for ln in str(exc).splitlines() if ln.strip()), exc.__class__.__name__)
                        ctx.log(f"Table XPath extraction error: {detail}")
                        if attempt < max_retries - 1:
                            continue
                        return "Error: Table extraction failed"
            except TimeoutException:
                if attempt < max_retries - 1:
                    time.sleep(3)
                    continue
                return "Error: Operation timed out"
            except Exception as exc:
                ctx.log(f"Unexpected error: {exc}")
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                return f"Error: {exc}"
        return "Error: Max retries exceeded"

    def _checkpoint(self, df, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_ETA.xlsx"
            df.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")
