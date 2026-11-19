"""HS Code Verify module (ported, browser-based).

Port of the desktop ``hscode_verify`` module. The source was an incomplete
sketch: it used absolute XPaths, a hard-coded iframe declaration id and always
typed the literal ``85171800`` instead of the row's HS code. The port keeps the
same navigation flow but submits each row's HS Code and records a per-row
result. Selectors remain portal-specific and may need re-confirmation.
"""

from __future__ import annotations

import logging
import time
from pathlib import Path

import pandas as pd

from ..base import Field, Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit
from . import dt_common as dt

logger = logging.getLogger(__name__)

HS_GATEWAY = ("https://eservices.dubaitrade.ae/portal2/sso/gateway.do"
              "?serviceId=3047021&agentType=DC&agentCode=AE-1000624&isMobile=false")

MODULE = Module(
    key="hscode_verify",
    name="HS Code Verify",
    description="Batch HS code verification against the portal.",
    icon="🔎",
    category="Utilities",
    requires_browser=True,
    requires_credentials=True,
    ported=True,
    notes="Portal-specific XPaths; may need re-confirmation.",
    fields=[
        Field("excel_file", "HS Code Excel", "file", required=True, accept=".xlsx,.xls",
              help="Must contain an 'HS Code' column"),
    ],
)


class HSCodeVerifyModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        excel_files = list(ctx.uploaded.get("excel_file", []))
        if not excel_files:
            raise ModuleError("An HS Code Excel file is required.")
        excel_path = excel_files[0]

        username = ctx.credentials.get("username", "")
        password = ctx.credentials.get("password", "")

        try:
            df = pd.read_excel(excel_path)
        except Exception as exc:
            raise ModuleError(f"Could not read the Excel file: {exc}") from exc
        if "HS Code" not in df.columns:
            raise ModuleError("The Excel file must contain an 'HS Code' column.")
        if "Result" not in df.columns:
            df["Result"] = ""

        total = len(df)
        processed = 0
        driver = None
        try:
            ctx.progress(5, "Launching visible browser (OTP may be required)")
            driver = create_driver(headless=False)
            wait = self._wait(driver)
            dt.login(driver, wait, username, password, ctx, success_fragment="portal2")

            ctx.progress(15, "Opening HS Code Verification page")
            self._navigate_to_hs_page(ctx, driver, wait)

            for index, row in df.iterrows():
                hs_code = str(row.get("HS Code", "")).strip()
                if not hs_code or hs_code.lower() in ("nan", "none"):
                    continue
                progress = 15 + int((index + 1) / max(total, 1) * 80)
                ctx.progress(progress, f"Verifying {hs_code}")
                ctx.log(f"Processing HS Code: {hs_code}")
                try:
                    self._process_hs_code(ctx, driver, wait, hs_code)
                    df.at[index, "Result"] = "Checked"
                    processed += 1
                except Exception as exc:
                    ctx.log(f"HS Code {hs_code} failed: {exc}")
                    df.at[index, "Result"] = f"Error: {exc}"

            ctx.progress(98, "Saving results")
            out = ctx.work_dir / f"{excel_path.stem}_verified.xlsx"
            df.to_excel(out, index=False, engine="openpyxl")
            ctx.log(f"Results saved to {out.name}")
            ctx.log(f"Done. Processed {processed}/{total} HS code(s).")
            ctx.progress(100, "Completed")
        finally:
            safe_quit(driver)

    def _wait(self, driver):
        from selenium.webdriver.support.ui import WebDriverWait

        return WebDriverWait(driver, 20)

    def _navigate_to_hs_page(self, ctx, driver, wait) -> None:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        driver.get(HS_GATEWAY)
        try:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div/div[2]/div[2]/div/div[2]/div/form/div/table/tbody/tr[1]/td[5]/button"))).click()
            time.sleep(2)
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div/div[2]/div[2]/div/div[2]/div/form/div[5]/button"))).click()
            time.sleep(2)
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div/div[2]/div[2]/div/div[2]/div/form/div[4]/button[2]"))).click()
            time.sleep(2)
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div/div[2]/div[2]/div/div[2]/div/div[2]/form/div[1]/div[2]/div/div/table/tbody/tr/td[8]/button[1]"))).click()
            time.sleep(2)
        except Exception as exc:
            raise ModuleError(f"Could not navigate to the HS Code page: {exc}") from exc

    def _process_hs_code(self, ctx, driver, wait, hs_code: str) -> None:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        search_button_xpath = ("/html/body/div/div[2]/div[2]/div/div[2]/div/form/div[2]"
                               "/div[2]/div/div/table/tbody/tr/td[5]/button[1]")
        wait.until(EC.element_to_be_clickable((By.XPATH, search_button_xpath))).click()
        time.sleep(2)

        popup_trigger_xpath = ("/html/body/div/div[2]/div[2]/div/div[2]/div/form[1]/div[2]"
                               "/div[2]/div/div/div/div[1]/div/div/span/button")
        wait.until(EC.element_to_be_clickable((By.XPATH, popup_trigger_xpath))).click()
        time.sleep(2)

        iframe = wait.until(EC.presence_of_element_located(
            (By.XPATH, '//*[contains(@src, "setupSearchHsCodes.do")]')))
        driver.switch_to.frame(iframe)
        try:
            search_box = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="searchHsCode"]')))
            search_box.clear()
            search_box.send_keys(hs_code)
            time.sleep(2)
        finally:
            driver.switch_to.default_content()

        driver.find_element(By.XPATH, '//*[@id="submitButton"]').click()
        time.sleep(1)
