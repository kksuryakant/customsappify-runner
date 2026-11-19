"""Transit Deposits Claim module (ported, browser-based).

Logs in to Dubai Trade, opens the transit-deposit claim page, and for each
declaration in the Excel enters the declaration number, uploads the matching
PDFs (compressed), sets the departure date and adds it to the claim list.

OTP is cleared manually in the *local* Chrome window on the runner machine, so
this module launches a visible browser.
"""

from __future__ import annotations

import logging
import time
from pathlib import Path

import pandas as pd

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..browser import create_driver, safe_quit
from . import dt_common as dt

logger = logging.getLogger(__name__)

MODULE = Module(
    key="Transit_Deposits_Claim",
    name="Transit Deposits Claim",
    description="Transit deposit refund processing with automated calculations and evidence upload.",
    icon="🚚",
    category="Claim Submission",
    requires_browser=True,
    requires_credentials=True,
    ported=True,
)


class TransitDepositsClaimModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        excel_files = list(ctx.uploaded.get("excel_file", []))
        if not excel_files:
            raise ModuleError("A claim Excel file is required.")
        excel_path = excel_files[0]

        pdfs = list(ctx.uploaded.get("files", []))
        ctx.log(f"Claim file: {excel_path.name}; {len(pdfs)} PDF(s) uploaded")

        try:
            df = pd.read_excel(excel_path)
        except Exception as exc:
            raise ModuleError(f"Could not read the Excel file: {exc}") from exc
        df = dt.load_claim_rows(df, ["DeclarationNo", "DepartureDate"])

        username = ctx.credentials.get("username", "")
        password = ctx.credentials.get("password", "")
        total = len(df)
        compressed_dir = ctx.work_dir / "compressed"

        driver = None
        results: list[dict] = []
        try:
            ctx.progress(5, "Launching visible browser (OTP may be required)")
            driver = create_driver(headless=False)
            wait = self._wait(driver)

            dt.login(driver, wait, username, password, ctx)

            ctx.progress(20, "Opening claim page")
            ctx.log("Navigating to the transit-deposit claim page")
            driver.get(dt.gateway_url())
            self._select_claim_type(driver, wait, "1", ctx)

            processed = 0
            for index, row in df.iterrows():
                dec_no = str(row["DeclarationNo"]).split(".")[0].strip()
                if not dec_no or dec_no.lower() == "nan":
                    continue
                try:
                    dep_date = pd.to_datetime(row["DepartureDate"]).strftime("%d/%m/%Y")
                except Exception:
                    ctx.log(f"{dec_no}: invalid departure date, skipping")
                    results.append({"DeclarationNo": dec_no, "Status": "invalid date"})
                    continue

                progress = 20 + int((index + 1) / total * 70)
                ctx.progress(progress, f"Processing {dec_no}")
                ctx.log(f"Processing claim {index + 1}/{total}: {dec_no}")

                status = self._process_declaration(ctx, driver, wait, dec_no, dep_date,
                                                   pdfs, compressed_dir)
                results.append({"DeclarationNo": dec_no, "Status": status})
                if status == "submitted":
                    processed += 1
                self._checkpoint(results, ctx, excel_path.stem)

            ctx.progress(98, "Writing summary")
            out = ctx.work_dir / f"{excel_path.stem}_claim_summary.xlsx"
            pd.DataFrame(results).to_excel(out, index=False, engine="openpyxl")
            ctx.log(f"Saved summary to {out.name}")
            ctx.progress(100, "Completed")
            ctx.log(f"Done. Submitted {processed}/{total} claim(s).")
        finally:
            safe_quit(driver)

    # ------------------------------------------------------------- steps
    def _wait(self, driver):
        from selenium.webdriver.support.ui import WebDriverWait

        return WebDriverWait(driver, 30)

    def _select_claim_type(self, driver, wait, value: str, ctx: ModuleContext) -> None:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        try:
            wait.until(EC.element_to_be_clickable((By.ID, "claimTypeListId"))).click()
            driver.find_element(By.XPATH, f"//option[@value='{value}']").click()
            wait.until(EC.element_to_be_clickable((By.ID, "claimTypeChangeId"))).click()
        except Exception as exc:
            raise ModuleError(f"Could not open the claim form: {exc}") from exc

    def _process_declaration(self, ctx, driver, wait, dec_no, dep_date, pdfs,
                             compressed_dir) -> str:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        try:
            field = wait.until(EC.visibility_of_element_located((By.ID, "claimDeclarationNo")))
            field.clear()
            time.sleep(0.5)
            field.send_keys(dec_no)
            wait.until(EC.element_to_be_clickable((By.ID, "adddeclarationId"))).click()
            time.sleep(2)

            errors = driver.find_elements(
                By.CSS_SELECTOR, ".error, .err, .error-message, .validation-error"
            )
            if errors and errors[0].text.strip():
                ctx.log(f"Portal rejected {dec_no}: {errors[0].text.strip()}")
                return "rejected"
        except Exception as exc:
            ctx.log(f"Error entering {dec_no}: {exc}")
            return "error"

        matching = dt.match_declaration_pdfs(dec_no, pdfs)
        if not matching:
            ctx.log(f"No PDFs matched for {dec_no}")
        for i, src in enumerate(matching, start=1):
            try:
                upload = dt.compress_pdf(src, compressed_dir, 2.0, ctx)
                ctx.log(f"Uploading {i}/{len(matching)}: {upload.name}")
                file_input = wait.until(EC.presence_of_element_located((
                    By.XPATH, "//input[@type='file' and @name='file' and contains(@class, 'formatter4')]",
                )))
                driver.execute_script("arguments[0].style.display = 'block';", file_input)
                file_input.send_keys(str(upload))
                time.sleep(2)
                wait.until(EC.element_to_be_clickable((
                    By.XPATH, "//input[@type='submit' and @name='saveDepositAttachment' and contains(@class, 'smpl-btn')]",
                ))).click()
                time.sleep(2)
            except Exception as exc:
                ctx.log(f"Error uploading {src.name}: {exc}")

        try:
            wait.until(EC.element_to_be_clickable((By.ID, "isSelected00"))).click()
            time.sleep(1)
            date_field = wait.until(EC.visibility_of_element_located((By.ID, "date")))
            date_field.clear()
            date_field.send_keys("\b" * 10)
            date_field.clear()
            for char in dep_date:
                date_field.send_keys(char)
                time.sleep(0.1)
            time.sleep(1)
            wait.until(EC.element_to_be_clickable((By.NAME, "addDeclDtlsToList"))).click()
            time.sleep(2)
            ctx.log(f"Submitted {dec_no}")
            return "submitted"
        except Exception as exc:
            ctx.log(f"Error finalizing {dec_no}: {exc}")
            self._screenshot(driver, dec_no)
            return "error"

    def _screenshot(self, driver, dec_no: str) -> None:
        try:
            path = f"error_declaration_{dec_no}_{int(time.time())}.png"
            driver.save_screenshot(path)
        except Exception:
            pass

    def _checkpoint(self, results, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_claim_summary.xlsx"
            pd.DataFrame(results).to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")
