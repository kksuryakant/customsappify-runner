"""Fetch Value and MOT module (ported, browser-based).

Logs in to Dubai Trade, then for each declaration number in the Excel searches
the Mirsal2 declaration view and writes back the Cargo Channel (mode of
transport) and total invoice value. OTP is cleared manually in the local Chrome
window, so the browser is visible.
"""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit
from . import dt_common as dt

logger = logging.getLogger(__name__)

SEARCH_URL = ("https://mirsal2new.dubaitrade.ae/m2decext/declaration/dt/"
              "setupPrintDeclarationSearch.do?service_id=3047024")
DETAILS_URL = ("https://mirsal2new.dubaitrade.ae/m2decext/declaration/dt/"
               "setupViewCustomsDeclaration.do?value(declarationId)={dec_id}")

MODULE = Module(
    key="Fetch_declaration",
    name="Fetch Value and MOT",
    description="Fetch declaration value and mode of transport from the portal.",
    icon="🔍",
    category="Utilities",
    requires_browser=True,
    requires_credentials=True,
    ported=True,
)


class FetchDeclarationModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        excel_files = list(ctx.uploaded.get("excel_file", []))
        if not excel_files:
            raise ModuleError("An Excel file with declaration numbers is required.")
        excel_path = excel_files[0]
        dec_col = str(ctx.params.get("declaration_col") or "DeclarationNo").strip()

        try:
            df = pd.read_excel(excel_path)
        except Exception as exc:
            raise ModuleError(f"Could not read the Excel file: {exc}") from exc
        if dec_col not in df.columns:
            raise ModuleError(f"The Excel file must contain a '{dec_col}' column.")
        if "Mode_of_Transport" not in df.columns:
            df["Mode_of_Transport"] = ""
        if "Total_Invoice_Value" not in df.columns:
            df["Total_Invoice_Value"] = ""

        username = ctx.credentials.get("username", "")
        password = ctx.credentials.get("password", "")
        total = len(df)
        ctx.log(f"Found {total} declaration row(s)")

        driver = None
        processed = 0
        try:
            ctx.progress(5, "Launching visible browser (OTP may be required)")
            driver = create_driver(headless=False)
            wait = self._wait(driver)
            dt.login(driver, wait, username, password, ctx, success_fragment="portal2")

            for index, row in df.iterrows():
                dec_no = str(row.get(dec_col, "")).strip()
                if not dec_no or dec_no.lower() in ("nan", "none"):
                    continue
                progress = 5 + int((index + 1) / total * 90)
                ctx.progress(progress, f"Fetching {dec_no}")
                ctx.log(f"Processing {index + 1}/{total}: {dec_no}")

                existing_mot = str(row.get("Mode_of_Transport", "")).strip()
                existing_val = str(row.get("Total_Invoice_Value", "")).strip()
                if (existing_mot not in ("", "nan", "None", "Not Found", "Error")
                        and existing_val not in ("", "nan", "None", "Not Found", "Error")):
                    ctx.log(f"Skipping {dec_no} - already processed")
                    continue

                try:
                    mot, value = self._fetch_one(ctx, driver, wait, dec_no)
                    df.at[index, "Mode_of_Transport"] = mot
                    df.at[index, "Total_Invoice_Value"] = value
                    processed += 1
                    ctx.log(f"{dec_no}: MOT={mot}, Invoice Value={value}")
                except Exception as exc:
                    ctx.log(f"{dec_no}: {exc}")
                    df.at[index, "Mode_of_Transport"] = "Not Found"
                    df.at[index, "Total_Invoice_Value"] = "Not Found"

                self._checkpoint(df, ctx, excel_path.stem)

            ctx.progress(98, "Saving results")
            out = ctx.work_dir / f"{excel_path.stem}_with_mot.xlsx"
            df.to_excel(out, index=False, engine="openpyxl")
            ctx.log(f"Saved results to {out.name}")
            ctx.progress(100, "Completed")
            ctx.log(f"Done. Processed {processed}/{total} declaration(s).")
        finally:
            safe_quit(driver)

    def _wait(self, driver):
        from selenium.webdriver.support.ui import WebDriverWait

        return WebDriverWait(driver, 20)

    def _fetch_one(self, ctx, driver, wait, dec_no: str):
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        driver.get(SEARCH_URL)
        time.sleep(2)
        field = wait.until(EC.presence_of_element_located((By.NAME, "value(declarationNo)")))
        field.clear()
        field.send_keys(dec_no)
        wait.until(EC.element_to_be_clickable((By.ID, "search"))).click()
        time.sleep(3)

        try:
            link = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//a[contains(text(), '{dec_no}')]")))
        except Exception:
            link = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(@data-iframe-url, 'setupViewCustomsDeclaration.do')]")))
        iframe_url = link.get_attribute("data-iframe-url") or ""
        dec_id = iframe_url.split("=")[1] if "=" in iframe_url else "unknown"

        driver.get(DETAILS_URL.format(dec_id=dec_id))
        time.sleep(3)

        mot = "Not Found"
        try:
            cargo = wait.until(EC.presence_of_element_located((
                By.XPATH,
                "/html/body/div/div/div/div[3]/div[1]/div[2]/div/div/div/table/tbody/tr[4]/td[1]",
            )))
            mot = cargo.find_element(By.CLASS_NAME, "flat-data-content").text.strip()
        except Exception:
            try:
                for el in driver.find_elements(By.XPATH, "//td[contains(@data-title, 'CARGO CHANNEL')]"):
                    text = el.find_element(By.CLASS_NAME, "flat-data-content").text.strip()
                    if text:
                        mot = text
                        break
            except Exception:
                mot = "Error"

        value = "Not Found"
        try:
            cell = wait.until(EC.presence_of_element_located((
                By.XPATH,
                "/html/body/div/div/div/div[3]/div[3]/div[2]/div/div/div/table/tbody/tr[2]/td[10]",
            )))
            try:
                value = cell.find_element(By.TAG_NAME, "b").text.strip()
            except Exception:
                value = cell.text.strip()
        except Exception:
            try:
                for el in driver.find_elements(By.XPATH, "//td[contains(., 'AED')]"):
                    text = el.text.strip()
                    match = re.search(r"AED\s*([\d,\.]+)", text)
                    if match and "Total" in el.find_element(By.XPATH, "./..").text:
                        value = f"AED {match.group(1)}"
                        break
            except Exception:
                value = "Error"

        return mot or "Not Found", value or "Not Found"

    def _checkpoint(self, df, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_with_mot.xlsx"
            df.to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")
