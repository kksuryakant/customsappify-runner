"""OOCL ATA module (ported, browser-based).

Port of the desktop ``OOCL_ATA`` module. NOTE: the source module was shipped
with placeholder/TODO selectors against the OOCL business portal and read its
credentials from the ``OOCL_USERNAME`` / ``OOCL_PASSWORD`` environment variables.
Those selectors are preserved verbatim below and will need confirming against
the live portal. Credentials are still taken from the environment; the framework
only models Dubai Trade credentials, so OOCL credentials are not stored per
company.
"""

from __future__ import annotations

import logging
import os
import time
from pathlib import Path

import pandas as pd

from ..base import Field, Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit

logger = logging.getLogger(__name__)

# OOCL portal configuration (from the desktop module; selectors are placeholders).
OOCL_CONFIG = {
    "tracking_url": "https://www.oocl.com/eng/ourservices/eservices/Pages/cargotracking.aspx",
    "login_url": "https://www.oocl.com/eng/Pages/login.aspx",
    "username_env": "OOCL_USERNAME",
    "password_env": "OOCL_PASSWORD",
    "select_login_username": 'input[id="username"]',
    "select_login_password": 'input[id="password"]',
    "select_login_button": 'button[id="loginBtn"]',
    "select_tracking_input": 'input[id="trackInput"]',
    "select_tracking_submit": 'button[id="trackBtn"]',
    "select_ata_field": 'span[class*="ata"]',
}

MODULE = Module(
    key="OOCL_ATA",
    name="OOCL ATA",
    description="Fetch OOCL container arrival dates.",
    icon="⚓",
    category="Arrival Status",
    requires_browser=True,
    requires_credentials=True,
    ported=True,
    notes="Placeholder selectors; requires OOCL_USERNAME/OOCL_PASSWORD env vars.",
    fields=[
        Field("excel_file", "Container Excel", "file", required=True, accept=".xlsx,.xls",
              help="Container list with a 'Container No' column"),
    ],
)


class OOCLATAModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        excel_files = list(ctx.uploaded.get("excel_file", []))
        if not excel_files:
            raise ModuleError("A container Excel file is required.")
        excel_path = excel_files[0]

        username = os.getenv(OOCL_CONFIG["username_env"], "")
        password = os.getenv(OOCL_CONFIG["password_env"], "")
        if not username or not password:
            raise ModuleError(
                f"Missing {OOCL_CONFIG['username_env']} / {OOCL_CONFIG['password_env']} "
                "environment variables on the runner."
            )

        try:
            df = pd.read_excel(excel_path)
        except Exception as exc:
            raise ModuleError(f"Could not read the Excel file: {exc}") from exc
        if "Container No" not in df.columns:
            raise ModuleError("The Excel file must contain a 'Container No' column.")
        if "ATA" not in df.columns:
            df["ATA"] = ""

        total = len(df)
        processed = 0
        driver = None
        try:
            ctx.progress(5, "Launching browser")
            driver = create_driver(headless=True)
            self._login(ctx, driver)

            for index, row in df.iterrows():
                container_no = str(row.get("Container No", "")).strip()
                if not container_no or container_no.lower() in ("nan", "none"):
                    continue
                progress = 5 + int((index + 1) / max(total, 1) * 90)
                ctx.progress(progress, f"Processing {container_no}")
                ctx.log(f"Processing container: {container_no}")

                try:
                    ata_date = self._fetch_ata_date(ctx, driver, container_no)
                except Exception as exc:
                    ctx.log(f"Error fetching {container_no}: {exc}")
                    ata_date = None

                if ata_date:
                    df.at[index, "ATA"] = ata_date
                    ctx.log(f"ATA date for {container_no}: {ata_date}")
                else:
                    ctx.log(f"No ATA date found for {container_no}")
                processed += 1

            ctx.progress(98, "Saving results")
            out = ctx.work_dir / f"{excel_path.stem}_with_ata.xlsx"
            df.to_excel(out, index=False, engine="openpyxl")
            ctx.log(f"Results saved to {out.name}")
            ctx.log(f"Done. Processed {processed}/{total} container(s).")
            ctx.progress(100, "Completed")
        finally:
            safe_quit(driver)

    def _login(self, ctx, driver) -> None:
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        username = os.getenv(OOCL_CONFIG["username_env"], "")
        password = os.getenv(OOCL_CONFIG["password_env"], "")

        driver.get(OOCL_CONFIG["login_url"])
        time.sleep(2)
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, OOCL_CONFIG["select_login_username"]))
            )
            driver.find_element(By.CSS_SELECTOR, OOCL_CONFIG["select_login_username"]).send_keys(username)
            driver.find_element(By.CSS_SELECTOR, OOCL_CONFIG["select_login_password"]).send_keys(password)
            driver.find_element(By.CSS_SELECTOR, OOCL_CONFIG["select_login_button"]).click()
            time.sleep(3)
            ctx.log("Logged in to OOCL portal")
        except TimeoutException as exc:
            raise ModuleError("Not on the expected OOCL login page") from exc

    def _fetch_ata_date(self, ctx, driver, container_no: str):
        from selenium.common.exceptions import NoSuchElementException, TimeoutException
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        driver.get(OOCL_CONFIG["tracking_url"])
        time.sleep(2)
        try:
            track_input = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, OOCL_CONFIG["select_tracking_input"]))
            )
            track_input.clear()
            track_input.send_keys(container_no)
            driver.find_element(By.CSS_SELECTOR, OOCL_CONFIG["select_tracking_submit"]).click()
            time.sleep(3)
            ata_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, OOCL_CONFIG["select_ata_field"]))
            )
            return ata_element.text.strip()
        except (TimeoutException, NoSuchElementException):
            return None
