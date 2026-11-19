"""Shared Dubai Trade (DT) helpers for claim modules.

All Dubai Trade claim flows log in through the same SSO page, sit behind the
same OTP challenge, use the same service gateway and compress/upload PDFs the
same way. These helpers keep each claim module small while running on the
*runner* machine (the browser is local, never on the server).
"""

from __future__ import annotations

import glob
import os
from pathlib import Path
from typing import List, Optional, Sequence

from ..base import ModuleContext, ModuleError

DUBAI_LOGIN_URL = "https://eservices.dubaitrade.ae/portal2/sso/main.do"
DUBAI_GATEWAY_URL = (
    "https://eservices.dubaitrade.ae/portal2/sso/gateway.do"
    "?serviceId={service_id}&agentType=DC&agentCode={agent_code}&isMobile=false"
)
DEFAULT_SERVICE_ID = "27722067"
DEFAULT_AGENT_CODE = "AE-1000624"


def gateway_url(service_id: str = DEFAULT_SERVICE_ID,
                agent_code: str = DEFAULT_AGENT_CODE) -> str:
    return DUBAI_GATEWAY_URL.format(service_id=service_id, agent_code=agent_code)


def login(driver, wait, username: str, password: str, ctx: ModuleContext,
          otp_timeout: int = 300, success_fragment: str = "gateway.do") -> None:
    """Log in to Dubai Trade SSO, waiting for the human to clear any OTP.

    The runner must use a *visible* browser for this: OTP entry happens in the
    local Chrome window on the runner machine.
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC

    if not username or not password:
        raise ModuleError(
            "Dubai Trade credentials are not configured for your company. "
            "Add them under Company settings."
        )

    ctx.log("Opening Dubai Trade login page")
    driver.get(DUBAI_LOGIN_URL)
    try:
        wait.until(EC.visibility_of_element_located((By.ID, "username"))).send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        wait.until(EC.element_to_be_clickable((By.ID, "kc-login"))).click()
    except Exception as exc:
        raise ModuleError(f"Dubai Trade login failed: {exc}") from exc

    try:
        wait.until(EC.visibility_of_element_located((By.ID, "otp")))
        ctx.log("OTP required — enter it in the local Chrome window, then click Verify")
        wait.until(EC.url_contains(success_fragment), timeout=otp_timeout)
        ctx.log("OTP verified")
    except Exception:
        ctx.log("No OTP challenge, continuing")


def compress_pdf(src: Path, out_dir: Path, max_size_mb: float = 2.0,
                 ctx: Optional[ModuleContext] = None) -> Path:
    """Compress a PDF below *max_size_mb* using pikepdf, or return the original.

    Mirrors the desktop ``compress_pdf_with_pikepdf`` flow: try progressively
    lower compression levels and keep the first that fits; otherwise return the
    original path unchanged.
    """
    try:
        import pikepdf

        if not src.exists():
            raise FileNotFoundError(src)
        original_mb = src.stat().st_size / (1024 * 1024)
        if original_mb <= max_size_mb:
            if ctx:
                ctx.log(f"{src.name} already under {max_size_mb}MB ({original_mb:.2f}MB)")
            return src

        out_dir.mkdir(parents=True, exist_ok=True)
        dest = out_dir / f"{src.stem}_compressed.pdf"

        for level in (80, 60, 40, 20):
            try:
                with pikepdf.open(str(src)) as pdf:
                    pdf.save(str(dest), compress_streams=True,
                             object_stream_mode=pikepdf.ObjectStreamMode.generate,
                             compression_level=level)
                size_mb = dest.stat().st_size / (1024 * 1024)
                if size_mb <= max_size_mb:
                    if ctx:
                        ctx.log(f"Compressed {src.name}: {original_mb:.2f}MB -> {size_mb:.2f}MB (q{level})")
                    return dest
            except Exception as exc:
                if ctx:
                    ctx.log(f"Compression error for {src.name} at q{level}: {exc}")
                continue

        if ctx:
            ctx.log(f"Could not compress {src.name} below {max_size_mb}MB; using original")
    except Exception as exc:
        if ctx:
            ctx.log(f"PDF compression unavailable for {src.name}: {exc}")
    return src


def match_declaration_pdfs(declaration_no: str, files: Sequence[Path]) -> List[Path]:
    """Return the uploaded PDFs whose filename starts with *declaration_no*."""
    dec = str(declaration_no).split(".")[0].strip()
    matches: List[Path] = []
    for path in files:
        name = Path(path).name
        if name.lower().endswith(".pdf") and name.startswith(dec):
            matches.append(Path(path))
    return matches


def load_claim_rows(df, columns: Sequence[str]) -> "object":
    """Validate that *df* contains the required columns; return it unchanged."""
    import pandas as pd

    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ModuleError("Excel file is missing column(s): " + ", ".join(missing))
    if df.empty:
        raise ModuleError("Excel file is empty")
    return df
