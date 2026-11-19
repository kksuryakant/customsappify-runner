"""COSCO ATA module (ported, browser-based).

Reads a container Excel, opens the public COSCO tracking site in a *local*
browser on the runner machine, extracts the actual-time-of-arrival for each
COSCO container and writes the results back to an Excel artifact.
"""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path
from typing import Optional

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit

logger = logging.getLogger(__name__)

COSCO_TRACKING_URL = "https://elines.coscoshipping.com/scct/public/ct/containers"
CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{7}$", re.IGNORECASE)
DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}")

LOCATION_COLUMNS = [
    ("JEBEL ALI ETA", ("JEBEL ALI", "JEBELALI", "JEBBEL ALI")),
    ("KHORFKHAN ETA", ("FAKKAN", "KHORFKHAN", "KHOR FKAN")),
    ("SALLALAH ETA", ("SALALAH", "SALLALAH", "SLL")),
    ("ABU DHABI ETC", ("ABU DHABI", "ABUDHABI", "AUH")),
    ("FUJAIRAH ETA", ("FUJAIRAH", "FUJAYRAH")),
    ("JEDDAH ETA", ("JEDDAH", "JEDDA")),
]


MODULE = Module(
    key="COSCO_ATA",
    name="COSCO ATA",
    description="Fetch COSCO ETA per destination port by BL or container.",
    icon="🚢",
    category="Arrival Status",
    requires_browser=False,
    requires_credentials=False,
    ported=True,
)


class COSCOATAModule(WebModule):
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
            ctx.log("Local COSCO tracking session started")

            processed = 0
            for index, row in df.iterrows():
                bl_no = self._cell(row, bl_col) if bl_col else ""
                container_no = self._cell(row, "Container No") if has_ctr else ""
                liner = self._cell(row, "Liner") if has_liner else ""
                key = bl_no or container_no
                if not key:
                    continue
                if has_liner and liner and "cosco" not in liner.lower():
                    ctx.log(f"{key}: not COSCO (liner={liner})")
                    continue
                if skip_done and self._all_filled(row):
                    ctx.log(f"Skipping {key} - all ETA columns already filled")
                    continue

                mode = "bl" if bl_no else "container"
                progress = 5 + int((index + 1) / total * 90)
                ctx.progress(progress, f"Tracking {key} ({mode})")
                ctx.log(f"[COSCO] Processing {key} ({mode})")

                try:
                    events = self._track_cosco(ctx, driver, key, mode)
                except Exception as exc:
                    ctx.log(f"[COSCO] {key}: {exc}")
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
            ctx.params["_artifacts_dir"] = str(ctx.work_dir)
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
            is_arrival = bool(event.get("is_arrival")) or "ARRIV" in text or "到港" in text
            for col, keywords in LOCATION_COLUMNS:
                if not any(k in text for k in keywords):
                    continue
                current = best.get(col)
                if current is None or (is_arrival and not current[1]):
                    best[col] = (date, is_arrival)
        return {col: date for col, (date, _) in best.items()}

    @staticmethod
    def _format_date(date_text: str) -> str:
        match = DATE_RE.search(date_text or "")
        if not match:
            return ""
        year, month, day = match.group(0).split("-")
        return f"{month}/{day}/{year}"

    def _track_cosco(self, ctx: ModuleContext, driver, key: str, mode: str) -> list:
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        mode_text = "Container No." if mode == "container" else "B/L No."
        driver.get(COSCO_TRACKING_URL)
        WebDriverWait(driver, 12).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(2)

        self._switch_english(driver)
        self._select_mode(driver, mode_text)
        input_field = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "input.ant-input"))
        )
        input_field.click()
        input_field.clear()
        input_field.send_keys(key)
        time.sleep(0.6)
        input_field.send_keys(Keys.ENTER)

        try:
            WebDriverWait(driver, 20).until(
                lambda d: "无法找到结果" in d.find_element(By.TAG_NAME, "body").text
                or bool(d.find_elements(By.CSS_SELECTOR, "table"))
            )
        except TimeoutException:
            ctx.log(f"[COSCO] no result for {key}")
            return []
        time.sleep(1.5)

        body_text = driver.find_element(By.TAG_NAME, "body").text
        if "无法找到结果" in body_text:
            ctx.log(f"[COSCO] {key}: no result found")
            return []

        tables = driver.execute_script(
            """
            const out=[];
            document.querySelectorAll('table').forEach(t=>{
                const rows=[...t.querySelectorAll('tr')].map(
                    tr=>[...tr.children].map(td=>(td.textContent||'').replace(/\\s+/g,' ').trim()));
                if(rows.length) out.push(rows);
            });
            return out;
            """
        )
        chain = driver.execute_script(
            """
            const out=[];
            document.querySelectorAll('span').forEach(s=>{
                const t=(s.textContent||'').trim();
                if(t==='FND' || t==='Last POD'){
                    const vert=s.closest('.ant-space-vertical');
                    const val=vert ? vert.querySelector('span[style*="bold"]') : null;
                    if(val) out.push([t, val.textContent.trim()]);
                }
            });
            return out;
            """
        )
        return self._parse_cosco(ctx, key, tables, body_text, chain)

    def _parse_cosco(self, ctx: ModuleContext, key: str, tables, body_text: str, chain=None) -> list:
        events: list[dict] = []
        last_eta = ""
        for rows in tables or []:
            if not rows:
                continue
            header = " ".join(rows[0])
            if ("Arrival Date" in header and "POD" in header) or (
                "到港时间" in header and "卸货港" in header
            ):
                for cells in rows[1:]:
                    if len(cells) < 6:
                        continue
                    disch = cells[4]
                    date = self._arrival_date(cells[5])
                    if not date:
                        continue
                    last_eta = date
                    events.append({
                        "container_no": key, "date": date, "location": disch,
                        "description": f"{cells[0]} {cells[1]} 到港".strip(),
                        "vessel_voyage": cells[1], "facility": cells[2], "is_arrival": True,
                    })
            elif ("Event Location" in header and "Event Time" in header) or (
                "位置" in header and "时间" in header
            ):
                for cells in rows[1:]:
                    if len(cells) < 4:
                        continue
                    date = self._format_date(cells[1])
                    if not date:
                        continue
                    events.append({
                        "container_no": key, "date": date, "location": cells[2],
                        "description": f"{cells[0]} {cells[3]}".strip(),
                        "vessel_voyage": "", "facility": "", "is_arrival": False,
                    })

        def _fmt(d: str) -> str:
            if not d:
                return ""
            return f"{d[5:7]}/{d[8:10]}/{d[0:4]}"

        # "Estimated cargo available for pickup at destination" -> final destination ETA
        pickup_date = ""
        m = re.search(
            r"available for pickup at destination\s*\n?\s*(\d{4}-\d{2}-\d{2})",
            body_text or "",
        )
        if m:
            pickup_date = _fmt(m.group(1))

        # Last POD milestone ETA (standalone ETA label line, e.g. Fujairah 11/14)
        lastpod_eta = ""
        m = re.search(r"(?:^|\n)\s*ETA\s*\n?\s*(\d{4}-\d{2}-\d{2})", body_text or "")
        if m:
            lastpod_eta = _fmt(m.group(1))

        # Port names for Last POD / FND from the labels+values row of the route
        lastpod_name = ""
        fnd_name = ""
        m = re.search(r"Last POD\s*\n\s*FND\s*\n((?:[^\n]+\n){6})", body_text or "")
        if m:
            vals = [v.strip() for v in m.group(1).strip().split("\n")]
            if len(vals) >= 6:
                lastpod_name, fnd_name = vals[4], vals[5]
        for label, value in chain or []:
            if label == "Last POD" and value and not lastpod_name:
                lastpod_name = value
            if label == "FND" and value and not fnd_name:
                fnd_name = value

        if lastpod_name and lastpod_eta:
            events.append({
                "container_no": key, "date": lastpod_eta, "location": lastpod_name,
                "description": "Last POD ETA", "vessel_voyage": "", "facility": "",
                "is_arrival": True,
            })
        if fnd_name and pickup_date:
            events.append({
                "container_no": key, "date": pickup_date, "location": fnd_name,
                "description": "Estimated cargo available for pickup", "vessel_voyage": "",
                "facility": "", "is_arrival": True,
            })
        return events

    @staticmethod
    def _arrival_date(cell: str) -> str:
        actual = re.search(r"(?:实际|Actual)[:：]\s*(\d{4}-\d{2}-\d{2})", cell)
        estimate = re.search(r"(?:预计|Expected)[:：]\s*(\d{4}-\d{2}-\d{2})", cell)
        raw = (actual or estimate)
        if not raw:
            return ""
        year, month, day = raw.group(1).split("-")
        return f"{month}/{day}/{year}"

    @staticmethod
    def _find_field(body_text: str, label: str) -> str:
        match = re.search(rf"{label}\s*[:：]\s*\n?\s*([^\n]+)", body_text)
        return match.group(1).strip() if match else ""

    def _switch_english(self, driver) -> None:
        from selenium.webdriver.common.by import By

        try:
            for el in driver.find_elements(By.XPATH, "//*[normalize-space(text())='English']"):
                if el.rect.get("width", 0) > 0:
                    driver.execute_script("arguments[0].click();", el)
                    time.sleep(2)
                    return
        except Exception:
            pass

    def _select_mode(self, driver, mode_text: str) -> bool:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        try:
            dropdown = WebDriverWait(driver, 6).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.ant-select"))
            )
            try:
                current = dropdown.find_element(By.CSS_SELECTOR, "span.ant-select-selection-item")
                if current.text.strip() == mode_text:
                    return True
            except Exception:
                pass
            dropdown.click()
            time.sleep(0.8)
            for option in driver.find_elements(By.CSS_SELECTOR, "div.ant-select-item"):
                if (option.text or "").strip() == mode_text:
                    option.click()
                    time.sleep(0.6)
                    return True
            return False
        except Exception:
            return False

    def _track_container(self, ctx: ModuleContext, driver, container_no: str) -> str:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        try:
            driver.get(COSCO_TRACKING_URL)
            WebDriverWait(driver, 10).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            time.sleep(1)

            if not self._select_container_option(driver):
                ctx.log(f"{container_no}: could not select the container-number search mode")
                return "Not Found"

            input_field = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input.ant-input"))
            )
            input_field.clear()
            input_field.send_keys(container_no)
            search_button = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ant-btn"))
            )
            search_button.click()
            time.sleep(3)

            return self._extract_ata(driver)
        except Exception as exc:
            ctx.log(f"{container_no}: tracking error - {exc}")
            return "Not Found"

    def _select_container_option(self, driver) -> bool:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        try:
            dropdown = WebDriverWait(driver, 6).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.ant-select"))
            )
            try:
                current = dropdown.find_element(
                    By.CSS_SELECTOR, "span.ant-select-selection-item"
                )
                if "箱号" in current.text:
                    return True
            except Exception:
                pass
            dropdown.click()
            time.sleep(1)
            for selector in (
                "//div[contains(@class, 'ant-select-item') and contains(., '箱号')]",
                "//div[@role='option' and contains(., '箱号')]",
                "//div[contains(text(), '箱号')]",
            ):
                try:
                    option = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    if option.is_displayed():
                        option.click()
                        time.sleep(0.8)
                        return True
                except Exception:
                    continue
            return False
        except Exception:
            return False

    def _extract_ata(self, driver) -> str:
        from selenium.webdriver.common.by import By

        try:
            page_text = driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            return "Not Found"

        if "无法找到结果" in page_text or "No results" in page_text:
            return "Not Found"

        patterns = (
            r"实际到港.*?(\d{4}-\d{2}-\d{2})",
            r"预计到港.*?(\d{4}-\d{2}-\d{2})",
            r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})",
            r"(\d{4}-\d{2}-\d{2})",
        )
        for pattern in patterns:
            match = re.search(pattern, page_text)
            if match:
                raw = match.group(1)
                date_match = re.search(r"(\d{4})-(\d{2})-(\d{2})", raw)
                if date_match:
                    year, month, day = date_match.groups()
                    return f"{month}/{day}/{year}"
                return raw
        return "Not Found"

    def _checkpoint(self, df, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_ATA.xlsx"
            df.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")
