"""Non-Remittance (NR) Sea & Air claim modules (ported, browser-based).

Both flows are identical: log in to Dubai Trade, open the claim page, select
claim type 3, then for each declaration upload the matching PDFs (compressed)
and add it to the list. OTP is cleared manually in the local Chrome window.
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

FILE_INPUT_XPATH = "//input[@type='file' and @name='file' and contains(@class, 'formatter4')]"
UPLOAD_BUTTON_XPATH = ("//input[@type='submit' and @name='saveDepositAttachment' "
                       "and contains(@class, 'smpl-btn')]")
ADD_DECLARATION_XPATH = ("//input[@type='button' and @value='Add Declaration' "
                         "and @name='addDeclDtlsToList']")


class _NRClaimBase(WebModule):
    """Shared implementation for the NR Sea/Air claim flows."""

    claim_type_value = "3"

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
        df = dt.load_claim_rows(df, ["DeclarationNo"])

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
            driver.get(dt.gateway_url())
            self._select_claim_type(driver, wait, ctx)

            processed = 0
            for index, row in df.iterrows():
                dec_no = str(row["DeclarationNo"]).split(".")[0].strip()
                if not dec_no or dec_no.lower() == "nan":
                    continue
                progress = 20 + int((index + 1) / total * 70)
                ctx.progress(progress, f"Processing {dec_no}")
                ctx.log(f"Processing claim {index + 1}/{total}: {dec_no}")
                status = self._process(ctx, driver, wait, dec_no, pdfs, compressed_dir)
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

    def _wait(self, driver):
        from selenium.webdriver.support.ui import WebDriverWait

        return WebDriverWait(driver, 30)

    def _select_claim_type(self, driver, wait, ctx: ModuleContext) -> None:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        try:
            wait.until(EC.presence_of_element_located((By.ID, "claimTypeListId")))
            wait.until(EC.element_to_be_clickable((By.ID, "claimTypeListId"))).click()
            driver.find_element(By.XPATH, f"//option[@value='{self.claim_type_value}']").click()
            wait.until(EC.element_to_be_clickable((By.ID, "claimTypeChangeId"))).click()
        except Exception as exc:
            raise ModuleError(f"Could not open the claim form: {exc}") from exc

    def _process(self, ctx, driver, wait, dec_no, pdfs, compressed_dir) -> str:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC

        try:
            field = wait.until(EC.visibility_of_element_located((By.ID, "claimDeclarationNo")))
            field.clear()
            field.send_keys(dec_no)
            wait.until(EC.element_to_be_clickable((By.ID, "adddeclarationId"))).click()
            time.sleep(2)
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
                file_input = wait.until(EC.presence_of_element_located((By.XPATH, FILE_INPUT_XPATH)))
                file_input.send_keys(str(upload))
                wait.until(EC.element_to_be_clickable((By.XPATH, UPLOAD_BUTTON_XPATH))).click()
                try:
                    self._wait(driver).until(
                        lambda d: not d.find_elements(By.XPATH, "//div[contains(@class, 'uploading')]"),
                        timeout=30,
                    )
                except Exception:
                    pass
                errs = driver.find_elements(By.XPATH, "//div[contains(@class, 'error-message')]")
                if errs and errs[0].text.strip():
                    raise RuntimeError(errs[0].text.strip())
            except Exception as exc:
                ctx.log(f"Error uploading {src.name}: {exc}")

        try:
            time.sleep(1)
            wait.until(EC.element_to_be_clickable((By.XPATH, ADD_DECLARATION_XPATH))).click()
            time.sleep(2)
            ctx.log(f"Submitted {dec_no}")
            return "submitted"
        except Exception as exc:
            ctx.log(f"Error finalizing {dec_no}: {exc}")
            return "error"

    def _checkpoint(self, results, ctx: ModuleContext, stem: str) -> None:
        try:
            path = ctx.work_dir / f"{stem}_claim_summary.xlsx"
            pd.DataFrame(results).to_excel(path, index=False, engine="openpyxl")
        except Exception as exc:
            ctx.log(f"Checkpoint save failed: {exc}")


MODULE_SEA = Module(
    key="NR_Sea_Claim",
    name="NR Sea Claim",
    description="Non-Remittance sea freight claim submission with vessel-specific handling.",
    icon="🛳️",
    category="Claim Submission",
    requires_browser=True,
    requires_credentials=True,
    ported=True,
)


class NRSeaClaimModule(_NRClaimBase):
    module = MODULE_SEA


MODULE_AIR = Module(
    key="NR_Air_Claim",
    name="NR Air Claim",
    description="Non-Remittance air freight claim submission with secure upload.",
    icon="✈️",
    category="Claim Submission",
    requires_browser=True,
    requires_credentials=True,
    ported=True,
)


class NRAirClaimModule(_NRClaimBase):
    module = MODULE_AIR
