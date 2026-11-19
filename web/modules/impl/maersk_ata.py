"""Maersk ATA module (ported, browser-based).

Reads an Excel whose key column holds *either* a BL number (``BL Number``) or a
container number (``Container No``), tracks it on the public Maersk tracking site
in a *local* browser on the runner machine, and fills one ETA column per
destination port (Jebel Ali, Khorfakkan, Salalah, Abu Dhabi) from the shipment's
transport plan.
"""

from __future__ import annotations

import logging
import random
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Dict

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit

logger = logging.getLogger(__name__)

MAERSK_TRACKING_URL = "https://www.maersk.com/tracking/"

# Output column -> location keywords found in a transport-plan leg.
LOCATION_COLUMNS = [
    ("JEBEL ALI ETA", ("JEBEL ALI", "JEBELALI", "JEA")),
    ("KHORFKHAN ETA", ("FAKKAN", "KHORFKHAN", "KHOR FKAN")),
    ("SALLALAH ETA", ("SALALAH", "SALLALAH")),
    ("ABU DHABI ETC", ("ABU DHABI", "ABUDHABI", "AUH")),
    ("FUJAIRAH ETA", ("FUJAIRAH", "FUJAYRAH")),
    ("JEDDAH ETA", ("JEDDAH", "JEDDA")),
]

_DATE_RE = re.compile(r"(\d{1,2} \w{3} \d{4}(?:\s+\d{1,2}:\d{2})?)")

MODULE = Module(
    key="MAERSK_ATA",
    name="Maersk ATA",
    description="Fetch Maersk ETA per destination port by BL or container.",
    icon="📦",
    category="Arrival Status",
    requires_browser=True,
    requires_credentials=False,
    ported=True,
)


class MaerskATAModule(WebModule):
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

        has_bl = "BL Number" in df.columns
        has_ctr = "Container No" in df.columns
        if not (has_bl or has_ctr):
            raise ModuleError("The Excel file must contain a 'BL Number' or 'Container No' column.")
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
            ctx.log("Local Maersk tracking session started")

            processed = 0
            first = True
            for index, row in df.iterrows():
                bl = self._cell(row, "BL Number") if has_bl else ""
                container = self._cell(row, "Container No") if has_ctr else ""
                liner = self._cell(row, "Liner")
                key = bl or container

                if not key or key.lower() in ("nan", "none"):
                    continue

                if skip_done and self._all_filled(row):
                    ctx.log(f"Skipping {key} - all ETA columns already filled")
                    continue

                progress = 5 + int((index + 1) / total * 90)
                ctx.progress(progress, f"Tracking {key}")

                if liner and "maersk" not in liner.lower():
                    ctx.log(f"{key}: not Maersk (liner={liner})")
                    continue

                locations = self._track_maersk(ctx, driver, key, is_first=first)
                first = False
                for col, _ in LOCATION_COLUMNS:
                    if locations.get(col):
                        df.at[index, col] = locations[col]
                processed += 1
                summary = ", ".join(f"{c.split()[0]}={locations[c]}" for c, _ in LOCATION_COLUMNS if locations.get(c)) or "no ETA found"
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
    def _cell(row, column: str) -> str:
        if column not in row:
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

    def _handle_cookie_consent(self, ctx: ModuleContext, driver) -> bool:
        from selenium.common.exceptions import NoSuchElementException, TimeoutException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        selectors = [
            (By.CSS_SELECTOR, "button[data-test='coi-allow-all-button']"),
            (By.CSS_SELECTOR, "button.coi-banner__accept"),
            (By.XPATH, "//button[contains(text(), 'Allow all')]"),
            (By.XPATH, "//button[contains(@aria-label, 'Allow all')]"),
        ]
        accepted = False
        for selector_type, selector_value in selectors:
            try:
                button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((selector_type, selector_value))
                )
                if button.is_displayed():
                    driver.execute_script("arguments[0].click();", button)
                    ctx.log("Cookie consent accepted")
                    time.sleep(1)
                    accepted = True
                    break
            except (TimeoutException, NoSuchElementException):
                continue
        # Remove any lingering consent overlay so it cannot intercept clicks.
        driver.execute_script(
            """
            ['coiOverlay','coi-banner-wrapper'].forEach(id => {
                const e = document.getElementById(id);
                if (e) e.remove();
            });
            document.querySelectorAll('.coi-banner__wrapper,.coi-consent-banner').forEach(e => e.remove());
            document.documentElement.style.overflow = 'auto';
            document.body.style.overflow = 'auto';
            """
        )
        return accepted

    def _track_maersk(self, ctx: ModuleContext, driver, search_key: str, is_first: bool = True) -> Dict[str, str]:
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        max_retries = 3
        operation_timeout = 45

        for attempt in range(max_retries):
            try:
                start = time.time()
                ctx.log(f"Processing Maersk '{search_key}' (attempt {attempt + 1}/{max_retries})")

                # The new results view keeps the search input on the page, so
                # the old "reload when input is absent" logic never fires for
                # rows 2+ and a stale result (row 1) would be re-parsed over
                # and over. Always reload before every search.
                driver.get(MAERSK_TRACKING_URL)
                time.sleep(random.uniform(2, 4))
                self._handle_cookie_consent(ctx, driver)
                WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

                try:
                    mc_input = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "mc-input[data-test='track-input']")
                        )
                    )
                    # The lit <mc-input> ignores JS-set values; use the shadow-root
                    # API + real key events so the component registers the input.
                    input_field = mc_input.shadow_root.find_element(
                        By.CSS_SELECTOR, "input[data-cy='input']"
                    )
                    input_field.clear()
                    input_field.send_keys(search_key)
                    time.sleep(1)
                except Exception as exc:
                    ctx.log(f"Error entering search value: {exc}")
                    if attempt < max_retries - 1:
                        continue
                    return {}

                try:
                    mc_button = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "mc-button[data-test='track-button']")
                        )
                    )
                    button = mc_button.shadow_root.find_element(By.CSS_SELECTOR, "button")
                    driver.execute_script("arguments[0].click();", button)
                    time.sleep(5)
                except Exception as exc:
                    ctx.log(f"Error clicking search button: {exc}")
                    if attempt < max_retries - 1:
                        continue
                    return {}

                if time.time() - start > operation_timeout:
                    raise TimeoutException("Operation timed out")

                try:
                    # New Maersk tracking experience: results render under
                    # .ocean-design__bl-info with per-leg location blocks and
                    # event groups (no more ul.transport-plan__list).
                    WebDriverWait(driver, 30).until(
                        lambda d: d.find_elements(By.CSS_SELECTOR, ".container-accordion-header__journey")
                        or d.find_elements(By.CSS_SELECTOR, ".ocean-design__bl-info")
                        or "No results found" in d.find_element(By.TAG_NAME, "body").text
                    )
                    if "No results found" in driver.find_element(By.TAG_NAME, "body").text:
                        ctx.log(f"{search_key}: Maersk returned no results (not publicly trackable).")
                        return {}
                    self._dismiss_overlays(ctx, driver)
                    return self._parse_locations(driver)
                except Exception as exc:
                    detail = next((ln for ln in str(exc).splitlines() if ln.strip()), exc.__class__.__name__)
                    ctx.log(f"Timeline extraction error: {detail}")
                    if attempt < max_retries - 1:
                        continue
                    return {}
            except TimeoutException:
                if attempt < max_retries - 1:
                    time.sleep(3)
                    continue
                return {}
            except Exception as exc:
                ctx.log(f"Unexpected error: {exc}")
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                return {}
        return {}

    def _parse_locations(self, driver) -> Dict[str, str]:
        """Map each destination location to its ETA from the results page.

        Maersk's new tracking view renders one accordion group per location
        (``.container-event-group__content``). The accordion may be collapsed,
        so visible text is empty; parse the raw ``innerHTML`` of each group
        instead. Each group carries a city header plus event rows (name + date,
        e.g. "Feeder arrival", "Gate in"). The vessel arrival / gate-in is
        preferred over departures when a location has several events.
        """
        from selenium.webdriver.common.by import By

        found: Dict[str, tuple] = {}  # column -> (date, is_arrival)

        for group in driver.find_elements(By.CSS_SELECTOR, ".container-event-group__content"):
            html = group.get_attribute("innerHTML") or ""
            city_match = re.search(r'event-location-city[^>]*>\s*<p>(.*?)</p>', html, re.S)
            if not city_match:
                continue
            name = re.sub(r"<[^>]+>", "", city_match.group(1)).strip().upper()
            events = []
            for name_match in re.finditer(r'data-test="event-name"[^>]*>(.*?)</(?:span|div|header)>', html, re.S):
                event_name = re.sub(r"<[^>]+>", "", name_match.group(1)).strip()
                tail = html[name_match.end():name_match.end() + 800]
                date_match = re.search(r'data-test="event-date"[^>]*>(.*?)</', tail, re.S)
                if not date_match:
                    continue
                event_date = self._format_date(re.sub(r"<[^>]+>", "", date_match.group(1)).strip())
                if event_date:
                    events.append((event_name, event_date))
            if not events:
                continue
            for col, keywords in LOCATION_COLUMNS:
                if not any(k in name for k in keywords):
                    continue
                chosen = None
                # 1. vessel arrival / gate-in milestone
                for event_name, event_date in events:
                    if "arriv" in event_name.lower() or "gate in" in event_name.lower():
                        chosen = event_date
                        break
                # 2. anything that is not a departure
                if chosen is None:
                    for event_name, event_date in events:
                        if "depart" not in event_name.lower() and "gate out" not in event_name.lower():
                            chosen = event_date
                            break
                # 3. first event as a last resort
                if chosen is None:
                    chosen = events[0][1]
                found.setdefault(col, (chosen, True))

        # Journey block provides the overall arrival date for the single-date
        # wrapper used by the Arrival_Date dispatcher.
        for block in driver.find_elements(By.CSS_SELECTOR, ".container-accordion-header__location-block"):
            text = block.text or ""
            if "arriv" in text.lower():
                match = _DATE_RE.search(text)
                if match:
                    found.setdefault("_first", (self._format_date(match.group(1)), True))
                    break

        return {col: value[0] for col, value in found.items()}

    def _dismiss_overlays(self, ctx: ModuleContext, driver) -> None:
        """Close the Maersk new-tracking intro coach-mark overlay if present."""
        from selenium.webdriver.common.by import By

        for selector in ("button[data-test='skipButton']", ".coach__button--skip"):
            try:
                button = driver.find_element(By.CSS_SELECTOR, selector)
                if button.is_displayed():
                    driver.execute_script("arguments[0].click();", button)
                    ctx.log("Dismissed tracking intro overlay")
                    time.sleep(0.5)
                    return
            except Exception:
                continue

    @staticmethod
    def _format_date(raw: str) -> str:
        for fmt in ("%d %b %Y %H:%M", "%d %b %Y"):
            try:
                return datetime.strptime(raw, fmt).strftime("%m/%d/%Y")
            except ValueError:
                continue
        return raw

    def _checkpoint(self, df, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_ETA.xlsx"
            df.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")
