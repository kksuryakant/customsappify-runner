"""ReExport Temp module (ported, browser-based).

Logs in to Dubai Trade, opens the Re-Export claim page for a single ``M2``
declaration, uploads its PDFs, then walks the paginated outbound-detail table,
cross-references each row against the Excel (Invoice Number / LineItem Serial No
/ HSCODE) and fills the add-outbound form (customs authority, exit point,
departure date, quantities…) before submitting.

OTP is cleared manually in the local Chrome window, so the browser is visible.
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

FILE_INPUT_XPATH = "//input[@type='file' and @name='file' and contains(@class, 'formatter4')]"
UPLOAD_BUTTON_XPATH = ("//input[@type='submit' and @name='saveDepositAttachment' "
                       "and contains(@class, 'smpl-btn')]")
DETAILS_LINK_XPATH = ("/html/body/div[3]/div/div/table/tbody/tr[2]/td[2]/div/div/"
                      "form[2]/fieldset/table/tbody/tr/td[1]/div/table/tbody/tr/td[5]/a")
TOTAL_RECORDS_XPATH = ("/html/body/div[4]/div/div/form/div[2]/div[2]/div[1]/div[4]/"
                       "fieldset[2]/div[2]/div/div[1]/div[2]")

MODULE = Module(
    key="ReExport_Temp",
    name="ReExport Temp",
    description="Re-export documentation automation and template generation.",
    icon="♻️",
    category="Claim Submission",
    requires_browser=True,
    requires_credentials=True,
    ported=True,
)


class ReExportTempModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        excel_files = list(ctx.uploaded.get("excel_file", []))
        if not excel_files:
            raise ModuleError("A claim Excel file is required.")
        excel_path = excel_files[0]
        pdfs = list(ctx.uploaded.get("files", []))
        declaration_number = str(ctx.params.get("declaration_number", "")).strip()
        if not declaration_number:
            raise ModuleError("An M2 declaration number is required.")

        try:
            df = pd.read_excel(excel_path)
        except Exception as exc:
            raise ModuleError(f"Could not read the Excel file: {exc}") from exc
        if "M2 Declaration Number" not in df.columns:
            raise ModuleError("The Excel file must contain an 'M2 Declaration Number' column.")
        df["M2 Declaration Number"] = df["M2 Declaration Number"].astype(str).str.strip()
        if declaration_number not in df["M2 Declaration Number"].values:
            raise ModuleError(f"Declaration number {declaration_number} not found in the data.")

        username = ctx.credentials.get("username", "")
        password = ctx.credentials.get("password", "")
        compressed_dir = ctx.work_dir / "compressed"

        driver = None
        try:
            ctx.progress(5, "Launching visible browser (OTP may be required)")
            driver = create_driver(headless=False)
            wait = self._wait(driver, 35)

            dt.login(driver, wait, username, password, ctx)

            ctx.progress(20, "Opening claim page")
            driver.get(dt.gateway_url())
            self._select_claim_type(driver, wait)

            ctx.log(f"Processing declaration {declaration_number}")
            wait.until(self._by("id", "claimDeclarationNo")).send_keys(declaration_number)
            wait.until(self._clickable("id", "adddeclarationId")).click()
            time.sleep(2)

            self._upload_files(ctx, driver, wait, declaration_number, pdfs, compressed_dir)

            wait.until(self._clickable("id", "isSelected00")).click()
            time.sleep(0.5)
            wait.until(self._clickable("id", "exportStatus00")).click()
            wait.until(self._clickable("xpath", "//option[@value='4001']")).click()
            time.sleep(0.5)

            wait.until(self._clickable("name", "addDeclDtlsToList")).click()
            time.sleep(1)
            ctx.log("Declaration added to list")

            self._process_outbound_rows(ctx, driver, wait, df)

            ctx.progress(98, "Finalizing")
            wait.until(self._clickable("name", "addDeclDtlsToList")).click()
            ctx.log("Declaration submitted successfully")
            ctx.progress(100, "Completed")
        finally:
            safe_quit(driver)

    # ------------------------------------------------------------- helpers
    def _wait(self, driver, timeout: int):
        from selenium.webdriver.support.ui import WebDriverWait

        return WebDriverWait(driver, timeout)

    @staticmethod
    def _by(how: str, what: str):
        from selenium.webdriver.common.by import By

        return {"id": By.ID, "name": By.NAME, "xpath": By.XPATH,
                "css": By.CSS_SELECTOR, "class": By.CLASS_NAME}[how], what

    def _clickable(self, how: str, what: str):
        from selenium.webdriver.support import expected_conditions as EC

        return EC.element_to_be_clickable(self._by(how, what))

    def _select_claim_type(self, driver, wait) -> None:
        wait.until(self._clickable("id", "claimTypeListId")).click()
        driver.find_element(*self._by("xpath", "//option[@value='1']")).click()
        wait.until(self._clickable("id", "claimTypeChangeId")).click()

    def _upload_files(self, ctx, driver, wait, declaration_number, pdfs, compressed_dir) -> None:
        from selenium.webdriver.common.by import By

        matching = dt.match_declaration_pdfs(declaration_number, pdfs)
        if not matching:
            ctx.log(f"Warning: no PDFs matched for {declaration_number}")
        for i, src in enumerate(matching, start=1):
            try:
                upload = dt.compress_pdf(src, compressed_dir, 2.0, ctx)
                ctx.log(f"Uploading {i}/{len(matching)}: {upload.name}")
                file_input = wait.until(self._presence("xpath", FILE_INPUT_XPATH))
                file_input.send_keys(str(upload))
                wait.until(self._clickable("xpath", UPLOAD_BUTTON_XPATH)).click()
                try:
                    self._wait(driver, 30).until(
                        lambda d: not d.find_elements(By.XPATH, "//div[contains(@class, 'uploading')]"))
                except Exception:
                    pass
                errs = driver.find_elements(By.XPATH, "//div[contains(@class, 'error-message')]")
                if errs and errs[0].text.strip():
                    raise RuntimeError(errs[0].text.strip())
            except Exception as exc:
                ctx.log(f"Error uploading {src.name}: {exc}")

    def _presence(self, how: str, what: str):
        from selenium.webdriver.support import expected_conditions as EC

        return EC.presence_of_element_located(self._by(how, what))

    def _process_outbound_rows(self, ctx, driver, wait, df) -> None:
        from selenium.webdriver.common.by import By

        ctx.log("Opening declaration details")
        wait.until(self._clickable("xpath", DETAILS_LINK_XPATH)).click()
        time.sleep(1)

        records_text = driver.find_element(*self._by("xpath", TOTAL_RECORDS_XPATH)).text
        match = re.search(r"Total Records :(\d+)", records_text)
        if not match:
            raise ModuleError("Could not read the total records count.")
        total_records = int(match.group(1))
        rows_per_page = 10
        total_pages = (total_records // rows_per_page) + (1 if total_records % rows_per_page else 0)
        ctx.log(f"Found {total_records} record(s) across {total_pages} page(s)")

        for page_num in range(1, total_pages + 1):
            try:
                if page_num > 1:
                    ctx.log(f"Navigating to page {page_num}")
                    wait.until(self._clickable("xpath", f"//a[text()='{page_num}']")).click()
                    self._wait(driver, 20).until(self._presence(
                        "xpath",
                        "//form[@id='outboundFormId']//*[@id='outbounddetailTable']/tbody"))
                start_row = (page_num - 1) * rows_per_page + 1
                end_row = min(page_num * rows_per_page, total_records)
                for row_index in range(start_row, end_row + 1):
                    row_on_page = row_index - (page_num - 1) * rows_per_page
                    row_xpath = ("//form[@id='outboundFormId']//*[@id='outbounddetailTable']"
                                 f"/tbody//tr[{row_on_page}]")
                    try:
                        self._wait(driver, 20).until(self._presence("xpath", row_xpath))
                        driver.find_element(By.XPATH, f"{row_xpath}//td[1]//input[1]").click()
                        time.sleep(1)
                        invoice = driver.find_element(By.XPATH, f"{row_xpath}//td[2]").text
                        line_item = driver.find_element(By.XPATH, f"{row_xpath}//td[3]").text
                        hscode = driver.find_element(By.XPATH, f"{row_xpath}//td[4]").text
                        ctx.log(f"Row {row_index}: Invoice={invoice}, Line={line_item}, HS={hscode}")
                        driver.find_element(By.XPATH, f"{row_xpath}//td[6]//a[1]").click()
                        time.sleep(1)
                        matching = self._match_row(df, invoice, line_item, hscode)
                        if matching is None:
                            raise ValueError(f"No matching Excel row for row {row_index}")
                        self._fill_outbound_form(ctx, driver, matching)
                    except Exception as exc:
                        ctx.log(f"Error on row {row_index}: {exc}")
                        continue
            except Exception as exc:
                ctx.log(f"Error on page {page_num}: {exc}")
                continue

    @staticmethod
    def _match_row(df, invoice, line_item, hscode):
        for _, row in df.iterrows():
            r_inv = str(row["Invoice Number"]).strip() if pd.notna(row["Invoice Number"]) else ""
            r_item = str(row["LineItem Serial No"]).strip() if pd.notna(row["LineItem Serial No"]) else ""
            r_hs = str(row["HSCODE"]).strip() if pd.notna(row["HSCODE"]) else ""
            if r_inv == invoice and r_item == line_item and r_hs == hscode:
                return row
        return None

    def _fill_outbound_form(self, ctx, driver, row) -> None:
        from selenium.webdriver.common.by import By

        driver.find_element(
            By.XPATH,
            "//form[@id='outboundDeclarationDetailId']//fieldset[2]"
            "//div[@class='configlinks']//a[@class='blue']").click()
        time.sleep(1)

        authority = str(row.get("customsAuthorityId", "")).strip()
        try:
            dropdown = driver.find_element(
                By.XPATH, "//form[@id='addOutboundForm']//select[@id='customsAuthorityId']")
            from selenium.webdriver.support.ui import Select

            options = [o.text.strip() for o in Select(dropdown).options]
            if authority in options:
                Select(dropdown).select_by_visible_text(authority)
                ctx.log(f"Selected customs authority '{authority}'")
            else:
                ctx.log(f"Authority '{authority}' not in dropdown options")
        except Exception as exc:
            ctx.log(f"Error selecting customs authority: {exc}")
            raise

        driver.find_element(
            By.XPATH, "//form[@id='addOutboundForm']//input[@id='decId']"
        ).send_keys(str(row.get("Out Declaration No", "")))
        time.sleep(1)

        if authority == "Dubai Customs":
            driver.find_element(
                By.XPATH,
                "//form[@id='addOutboundForm']//input[@name='retrieveOutBoundDeclaration']").click()
            time.sleep(1)
        else:
            driver.find_element(
                By.XPATH,
                "//form[@id='addOutboundForm']//select[@id='outboundDecType']").send_keys(
                str(row.get("Declaration Type", "")))
            time.sleep(1)
            driver.find_element(
                By.XPATH,
                "//form[@id='addOutboundForm']//select[@id='outboundReExportDest']").send_keys(
                str(row.get("Re-Export To", "")))
            time.sleep(1)

        driver.find_element(
            By.XPATH, "//form[@id='addOutboundForm']//select[@id='exitPointId']"
        ).send_keys(str(row.get("Exit Point", "")))
        time.sleep(1)

        dep_date = pd.to_datetime(row["DepartureDate"]).strftime("%d/%m/%Y")
        driver.find_element(
            By.XPATH, "//form[@id='addOutboundForm']//input[@id='actualdepDateId']"
        ).send_keys(dep_date)
        time.sleep(1)
        driver.find_element(
            By.XPATH, "//form[@id='addOutboundForm']//input[@id='stsQtyId']"
        ).send_keys(str(row.get("Stat Quantity", "")))
        time.sleep(1)
        driver.find_element(
            By.XPATH, "//form[@id='addOutboundForm']//input[@id='weightId']"
        ).send_keys(str(row.get("Weight", "")))
        time.sleep(1)

        ctx.log("Submitting claim form")
        driver.find_element(By.XPATH, "/html/body/div[8]/div/div/form/div[2]/div[3]/div[4]/div/input[1]").click()
        time.sleep(1)
        driver.find_element(By.XPATH, "/html/body/div[6]/div/div/form/div[2]/div[3]/div[4]/div/input[1]").click()
        time.sleep(1)
