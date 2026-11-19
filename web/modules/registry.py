"""Web module registry.

Catalog mirrors the desktop app's module catalog. Modules flagged
`ported=True` have a `WebModule.execute` implementation that runs inside the
local runner; the rest are listed so the UI can show what is available and
what still needs porting. To port another module: implement a `WebModule`
subclass, add its `Module` spec below and register the class in `IMPLEMENTATIONS`.
"""

from __future__ import annotations

import importlib
import logging
from typing import Dict, List, Optional

from .base import Field, Module, WebModule

logger = logging.getLogger(__name__)


CATALOG: List[Module] = [
    # ------------------------------------------------------------- claims
    Module(
        key="Transit_Deposits_Claim",
        name="Transit Deposits Claim",
        description="Transit deposit refund processing with automated calculations and evidence upload.",
        icon="🚚", category="Claim Submission", requires_browser=True, requires_credentials=True,
        ported=True,
        fields=[
            Field("excel_file", "Claim Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Claim data with 'DeclarationNo' and 'DepartureDate' columns"),
            Field("files", "Supporting PDFs", "file", required=True, accept=".pdf", multiple=True,
                  help="PDFs whose filename starts with the declaration number"),
        ],
    ),
    Module(
        key="NR_Sea_Claim",
        name="NR Sea Claim",
        description="Non-Remittance sea freight claim submission with vessel-specific handling.",
        icon="🛳️", category="Claim Submission", requires_browser=True, requires_credentials=True,
        ported=True,
        fields=[
            Field("excel_file", "Claim Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Claim data with a 'DeclarationNo' column"),
            Field("files", "Supporting PDFs", "file", required=True, accept=".pdf", multiple=True,
                  help="PDFs whose filename starts with the declaration number"),
        ],
    ),
    Module(
        key="NR_Air_Claim",
        name="NR Air Claim",
        description="Non-Remittance air freight claim submission with secure upload.",
        icon="✈️", category="Claim Submission", requires_browser=True, requires_credentials=True,
        ported=True,
        fields=[
            Field("excel_file", "Claim Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Claim data with a 'DeclarationNo' column"),
            Field("files", "Supporting PDFs", "file", required=True, accept=".pdf", multiple=True,
                  help="PDFs whose filename starts with the declaration number"),
        ],
    ),
    Module(
        key="ReExport_Temp",
        name="ReExport Temp",
        description="Re-export documentation automation and template generation.",
        icon="♻️", category="Claim Submission", requires_browser=True, requires_credentials=True,
        ported=True,
        fields=[
            Field("excel_file", "Claim Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Claim data with an 'M2 Declaration Number' column"),
            Field("files", "Supporting PDFs", "file", required=True, accept=".pdf", multiple=True,
                  help="PDFs whose filename starts with the declaration number"),
            Field("declaration_number", "M2 Declaration Number", "text", required=True),
        ],
    ),
    # ------------------------------------------------------------ arrivals
    Module(
        key="COSCO_ATA",
        name="COSCO ATA",
        description="Fetch COSCO arrival dates (per-location ETAs) from COSCO tracking.",
        icon="🚢", category="Arrival Status", requires_browser=True, requires_credentials=False,
        ported=True,
        fields=[
            Field("excel_file", "Container Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Must contain 'BL Number'/'BL' or 'Container No' and a 'Liner' column. "
                       "Use a BL No. for destination ETAs (Jebel Ali/Khor Fakkan/Salalah/Abu Dhabi) "
                       "— container mode only shows the latest location."),
            Field("skip_done", "Skip containers that already have an ATA", "checkbox", default=True),
        ],
    ),
    Module(
        key="CMA_CGM_ATA",
        name="CMA CGM ATA",
        description="Fetch CMA CGM arrival dates (per-location ETAs) from CMA CGM tracking.",
        icon="🛳️", category="Arrival Status", requires_browser=True, requires_credentials=False,
        ported=True,
        notes="Site is DataDome-protected; opens a visible browser so the operator can "
              "solve the CAPTCHA once. Set APPIFY_PROXY to keep the runner IP off the carrier.",
        fields=[
            Field("excel_file", "Shipment Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Must contain a 'BL Number'/'BL' or 'Container No' column and a 'Liner' column. "
                       "Outputs JEBEL ALI / KHORFKHAN / SALLALAH / ABU DHABI ETA columns."),
            Field("skip_done", "Skip rows whose ETA columns are already filled", "checkbox", default=True),
        ],
    ),
    Module(
        key="EVERGREEN_ATA",
        name="Evergreen ATA",
        description="Fetch Evergreen container arrival dates.",
        icon="🌲", category="Arrival Status", requires_browser=True, ported=True,
        fields=[
            Field("excel_file", "Container Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Must contain 'Container No' and 'Liner' columns"),
            Field("skip_done", "Skip containers that already have an ATA", "checkbox", default=True),
        ],
    ),
    Module(
        key="MSC_ATA",
        name="MSC ATA",
        description="Fetch MSC container arrival dates.",
        icon="🗂️", category="Arrival Status", requires_browser=True, ported=True,
        fields=[
            Field("excel_file", "Container Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Must contain a 'Container No' column"),
        ],
    ),
    Module(
        key="OOCL_ATA",
        name="OOCL ATA",
        description="Fetch OOCL container arrival dates.",
        icon="⚓", category="Arrival Status", requires_browser=True, requires_credentials=True,
        ported=True, notes="Placeholder selectors; requires OOCL_USERNAME/OOCL_PASSWORD env vars.",
        fields=[Field("excel_file", "Container Excel", "file", required=True, accept=".xlsx,.xls")],
    ),
    Module(
        key="MAERSK_ATA",
        name="Maersk ATA",
        description="Fetch Maersk ETA per destination port by BL or container.",
        icon="📦", category="Arrival Status", requires_browser=True, ported=True,
        fields=[
            Field("excel_file", "Shipment Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Must contain a 'BL Number' or 'Container No' column; optional 'Liner'. "
                       "Outputs JEBEL ALI / KHORFKHAN / SALLALAH / ABU DHABI ETA columns."),
            Field("skip_done", "Skip rows whose ETA columns are already filled", "checkbox", default=True),
        ],
    ),
    Module(
        key="HapagLloyd",
        name="Hapag-Lloyd",
        description="Fetch Hapag-Lloyd container arrival dates.",
        icon="🚢", category="Arrival Status", requires_browser=True, ported=True,
        fields=[
            Field("excel_file", "Container Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Must contain a 'Container No' column. Optional 'Liner' and 'BL' columns"),
        ],
    ),
    Module(
        key="HMM_ATA",
        name="HMM ATA",
        description="Fetch HMM ETA per destination port via pier2pier.com (container number).",
        icon="🚢", category="Arrival Status", requires_browser=False, ported=True,
        notes="HMM's own site blocks automated access; pier2pier's snapshot is used instead.",
        fields=[
            Field("excel_file", "Container Excel", "file", required=True, accept=".xlsx,.xls",
                  help="Must contain a 'Container No' column. Optional 'Liner' and 'BL' columns"),
            Field("skip_done", "Skip rows whose ETA columns are already filled", "checkbox", default=True),
        ],
    ),
    # ------------------------------------------------------------ utilities
    Module(
        key="pdf_compress",
        name="PDF Compressor",
        description="Compress PDFs locally on the runner with configurable quality and methods.",
        icon="📄", category="Utilities", requires_browser=False, requires_credentials=False,
        ported=True,
        fields=[
            Field("pdf_files", "PDF Files", "file", required=True, accept=".pdf", multiple=True,
                  help="One or more PDFs to compress"),
            Field("method", "Method", "select", default="auto",
                  options=[
                      {"value": "auto", "label": "Auto"},
                      {"value": "pymupdf", "label": "PyMuPDF (re-render)"},
                      {"value": "lossless", "label": "Lossless (recompress images)"},
                      {"value": "aggressive", "label": "Aggressive (smaller, lower quality)"},
                  ]),
            Field("quality", "Image Quality", "number", default=80, help="50-100"),
            Field("skip_small", "Skip files smaller than 100 KB", "checkbox", default=True),
        ],
    ),
    Module(
        key="Fetch_declaration",
        name="Fetch Value and MOT",
        description="Fetch declaration value and mode of transport from the portal.",
        icon="🔍", category="Utilities", requires_browser=True, requires_credentials=True,
        ported=True,
        fields=[
            Field("excel_file", "Excel File", "file", required=True, accept=".xlsx,.xls"),
            Field("declaration_col", "Declaration No Column", "text", default="DeclarationNo"),
        ],
    ),
    Module(
        key="hscode_verify",
        name="HS Code Verify",
        description="Batch HS code verification against the portal.",
        icon="🔎", category="Utilities", requires_browser=True, requires_credentials=True,
        ported=True, notes="Portal-specific XPaths; may need re-confirmation.",
        fields=[Field("excel_file", "HS Code Excel", "file", required=True, accept=".xlsx,.xls")],
    ),
    Module(
        key="BOE_Tracker",
        name="BOE Tracker",
        description="Bill of Entry tracking (source module was empty in the desktop build).",
        icon="⚠️", category="Utilities", requires_browser=False, ported=True,
        notes="Desktop source was empty; no tracking logic available.",
        fields=[],
    ),
]


# key -> dotted path to a WebModule subclass inside `web.modules.impl.*`
IMPLEMENTATIONS: Dict[str, str] = {
    "pdf_compress": "web.modules.impl.pdf_compress.PDFCompressModule",
    "COSCO_ATA": "web.modules.impl.cosco_ata.COSCOATAModule",
    "CMA_CGM_ATA": "web.modules.impl.cma_cgm_ata.CMACGMATAModule",
    "EVERGREEN_ATA": "web.modules.impl.evergreen_ata.EvergreenATAModule",
    "MSC_ATA": "web.modules.impl.msc_ata.MSCATAModule",
    "MAERSK_ATA": "web.modules.impl.maersk_ata.MaerskATAModule",
    "HapagLloyd": "web.modules.impl.hapag_lloyd.HapagLloydModule",
    "HMM_ATA": "web.modules.impl.hmm_ata.HMMATAModule",
    "Transit_Deposits_Claim": "web.modules.impl.transit_claim.TransitDepositsClaimModule",
    "NR_Sea_Claim": "web.modules.impl.nr_claims.NRSeaClaimModule",
    "NR_Air_Claim": "web.modules.impl.nr_claims.NRAirClaimModule",
    "Fetch_declaration": "web.modules.impl.fetch_declaration.FetchDeclarationModule",
    "ReExport_Temp": "web.modules.impl.re_export_temp.ReExportTempModule",
    "OOCL_ATA": "web.modules.impl.oocl_ata.OOCLATAModule",
    "hscode_verify": "web.modules.impl.hscode_verify.HSCodeVerifyModule",
    "BOE_Tracker": "web.modules.impl.boe_tracker.BOETrackerModule",
}


def list_modules() -> List[Module]:
    return CATALOG


def get_module(key: str) -> Optional[Module]:
    for module in CATALOG:
        if module.key == key:
            return module
    return None


def get_implementation(key: str) -> Optional[WebModule]:
    """Instantiate the runner-side implementation for *key*, if ported."""
    dotted = IMPLEMENTATIONS.get(key)
    if not dotted:
        return None
    module_path, class_name = dotted.rsplit(".", 1)
    try:
        module = importlib.import_module(module_path)
        cls = getattr(module, class_name)
        return cls()
    except Exception:
        logger.exception("Failed to load module implementation %s", dotted)
        return None
