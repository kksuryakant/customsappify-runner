"""PDF Compressor module (ported, no browser).

Compresses one or more PDFs on the runner's local machine and uploads the
results back to the server as artifacts.
"""

from __future__ import annotations

import gc
import io
import logging
import os
import shutil
import time
from pathlib import Path
from typing import List

from ..base import Module, ModuleContext, ModuleError, WebModule
from ..registry import IMPLEMENTATIONS  # noqa: F401  (registration marker)

logger = logging.getLogger(__name__)

MODULE = Module(
    key="pdf_compress",
    name="PDF Compressor",
    description="Compress PDFs locally on the runner with configurable quality and methods.",
    icon="📄",
    category="Utilities",
    requires_browser=False,
    requires_credentials=False,
    ported=True,
)


class PDFCompressModule(WebModule):
    module = MODULE

    def execute(self, ctx: ModuleContext) -> None:
        import fitz  # PyMuPDF
        from PIL import Image

        files: List[Path] = list(ctx.uploaded.get("pdf_files", []))
        if not files:
            raise ModuleError("No PDF files were uploaded.")

        method = str(ctx.params.get("method", "auto"))
        quality = self._clamp(int(ctx.params.get("quality", 80) or 80), 20, 100)
        skip_small = bool(ctx.params.get("skip_small", True))

        ctx.log(f"Compressing {len(files)} PDF file(s) with method={method}, quality={quality}")
        total_before = total_after = 0
        processed = skipped = failed = 0
        out_dir = ctx.work_dir / "compressed"
        out_dir.mkdir(parents=True, exist_ok=True)

        for index, pdf_path in enumerate(files, start=1):
            if not pdf_path.exists():
                ctx.log(f"Missing input file: {pdf_path.name}")
                failed += 1
                continue

            original_kb = pdf_path.stat().st_size / 1024.0
            if skip_small and pdf_path.stat().st_size < 102_400:
                ctx.log(f"Skipping {pdf_path.name} (under 100 KB)")
                skipped += 1
                ctx.progress(int(index / len(files) * 100), f"Skipped {pdf_path.name}")
                continue

            out_path = out_dir / f"{pdf_path.stem}_compressed.pdf"
            ctx.progress(int((index - 1) / len(files) * 100), f"Compressing {pdf_path.name}")
            ctx.log(f"[{index}/{len(files)}] {pdf_path.name} ({original_kb:.1f} KB)")

            try:
                used = self._compress(pdf_path, out_path, method, quality, fitz, Image)
            except Exception as exc:
                ctx.log(f"Failed to compress {pdf_path.name}: {exc}")
                failed += 1
                continue

            if not used or not out_path.exists():
                ctx.log(f"No reduction possible for {pdf_path.name}; keeping original")
                skipped += 1
                continue

            new_kb = out_path.stat().st_size / 1024.0
            reduction = 100 * (1 - new_kb / original_kb) if original_kb > 0 else 0
            ctx.log(f"✓ {pdf_path.name}: {original_kb:.1f} → {new_kb:.1f} KB "
                    f"({reduction:.1f}% reduction)")
            total_before += original_kb
            total_after += new_kb
            processed += 1
            gc.collect()

        ctx.progress(100, "Done")
        ctx.log("")
        ctx.log("=== Compression Summary ===")
        ctx.log(f"Compressed: {processed}   Skipped: {skipped}   Failed: {failed}")
        if total_before:
            ctx.log(f"Total size: {total_before:.1f} KB → {total_after:.1f} KB "
                    f"({100 * (1 - total_after / total_before):.1f}% saved)")
        ctx.log(f"Output folder: {out_dir}")

        # Hand the compressed folder back to the runner so it can upload artifacts.
        ctx.params["_artifacts_dir"] = str(out_dir)

    def _compress(self, src: Path, dst: Path, method: str, quality: int, fitz, Image) -> bool:
        if method == "lossless":
            return self._lossless(src, dst, quality, fitz, Image)
        if method == "aggressive":
            return self._rerender(src, dst, quality // 2, fitz)
        # auto / pymupdf
        return self._rerender(src, dst, quality, fitz)

    def _rerender(self, src: Path, dst: Path, quality: int, fitz) -> bool:
        doc = fitz.open(src)
        try:
            if len(doc) == 0:
                return False
            writer = fitz.open()
            for page in doc:
                rect = page.rect
                zoom = min(160 / 72, 2.0)
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img_bytes = pix.tobytes("jpeg", jpg_quality=max(20, min(quality, 100)))
                new_page = writer.new_page(width=rect.width, height=rect.height)
                new_page.insert_image(rect, stream=img_bytes)
                pix = None
            writer.save(dst, garbage=4, deflate=True, clean=True)
            writer.close()
        finally:
            doc.close()
        return dst.exists() and dst.stat().st_size > 0

    def _lossless(self, src: Path, dst: Path, quality: int, fitz, Image) -> bool:
        doc = fitz.open(src)
        try:
            for page in doc:
                for img in page.get_images(full=True):
                    try:
                        xref = img[0]
                        base = doc.extract_image(xref)
                        if not base:
                            continue
                        pil = Image.open(io.BytesIO(base["image"]))
                        out = io.BytesIO()
                        if pil.mode == "RGBA":
                            pil.save(out, format="PNG", optimize=True)
                        else:
                            if pil.mode != "RGB":
                                pil = pil.convert("RGB")
                            pil.save(out, format="JPEG", quality=quality, optimize=True,
                                     progressive=True)
                        doc.update_image(xref, stream=out.getvalue())
                        out.close()
                    except Exception:
                        continue
            doc.save(dst, garbage=4, deflate=True, clean=True)
        finally:
            doc.close()
        return dst.exists() and dst.stat().st_size > 0

    @staticmethod
    def _clamp(value: int, low: int, high: int) -> int:
        return max(low, min(high, value))
