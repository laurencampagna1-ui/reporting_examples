"""Converts a .pptx file to .pdf using headless LibreOffice."""
from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path


class PdfConversionError(RuntimeError):
    pass


def pptx_bytes_to_pdf_bytes(pptx_bytes: bytes, timeout: int = 90) -> bytes:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        src = tmp_path / "deck.pptx"
        src.write_bytes(pptx_bytes)

        profile_dir = tmp_path / "lo_profile"
        result = subprocess.run(
            [
                "soffice",
                "--headless",
                "--norestore",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp_path),
                str(src),
            ],
            capture_output=True,
            timeout=timeout,
        )
        pdf_path = tmp_path / "deck.pdf"
        if result.returncode != 0 or not pdf_path.exists():
            raise PdfConversionError(
                "LibreOffice failed to convert pptx to pdf: "
                + result.stderr.decode(errors="replace")
            )
        return pdf_path.read_bytes()
