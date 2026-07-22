"""Convert offer DOCX bytes to PDF using Word (Windows) or LibreOffice."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional


class PdfConversionError(RuntimeError):
    pass


def _find_libreoffice() -> Optional[str]:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    # Common Windows install paths
    candidates = [
        Path(os.environ.get("PROGRAMFILES", r"C:\Program Files"))
        / "LibreOffice"
        / "program"
        / "soffice.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"))
        / "LibreOffice"
        / "program"
        / "soffice.exe",
    ]
    for path in candidates:
        if path.exists():
            return str(path)
    return None


def _convert_with_libreoffice(docx_path: Path, out_dir: Path) -> Path:
    soffice = _find_libreoffice()
    if not soffice:
        raise PdfConversionError("LibreOffice nicht gefunden")
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(docx_path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    pdf_path = out_dir / f"{docx_path.stem}.pdf"
    if proc.returncode != 0 or not pdf_path.exists():
        detail = (proc.stderr or proc.stdout or "").strip()
        raise PdfConversionError(detail or "LibreOffice-Konvertierung fehlgeschlagen")
    return pdf_path


def _convert_with_word_com(docx_path: Path, pdf_path: Path) -> Path:
    """Windows: Microsoft Word via PowerShell COM (wdFormatPDF = 17)."""
    if sys.platform != "win32":
        raise PdfConversionError("Word-COM nur unter Windows verfügbar")

    # Escape single quotes for PowerShell single-quoted strings
    docx = str(docx_path).replace("'", "''")
    pdf = str(pdf_path).replace("'", "''")
    script = f"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {{
  $doc = $word.Documents.Open('{docx}')
  try {{
    $doc.SaveAs([ref] '{pdf}', [ref] 17)
  }} finally {{
    $doc.Close([ref] $false)
  }}
}} finally {{
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}}
"""
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        timeout=180,
    )
    if proc.returncode != 0 or not pdf_path.exists():
        detail = (proc.stderr or proc.stdout or "").strip()
        raise PdfConversionError(detail or "Word-PDF-Export fehlgeschlagen (ist Microsoft Word installiert?)")
    return pdf_path


def convert_docx_bytes_to_pdf(docx_bytes: bytes, basename: str = "angebot") -> bytes:
    """
    Build a PDF from DOCX bytes with identical layout to Word.

    Preference order:
      1) LibreOffice (cross-platform)
      2) Microsoft Word COM (Windows)
    """
    safe = "".join(c for c in basename if c.isalnum() or c in "-_") or "angebot"
    with tempfile.TemporaryDirectory(prefix="wamas-pdf-") as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / f"{safe}.docx"
        pdf_path = tmp_dir / f"{safe}.pdf"
        docx_path.write_bytes(docx_bytes)

        errors = []
        # LibreOffice first when available
        try:
            if _find_libreoffice():
                produced = _convert_with_libreoffice(docx_path, tmp_dir)
                return produced.read_bytes()
        except Exception as exc:  # noqa: BLE001
            errors.append(f"LibreOffice: {exc}")

        # Windows Word
        try:
            if sys.platform == "win32":
                produced = _convert_with_word_com(docx_path, pdf_path)
                return produced.read_bytes()
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Word: {exc}")

        hint = (
            "PDF 1:1 aus Word benötigt Microsoft Word (Windows) oder LibreOffice. "
            "Alternative: Word-Datei öffnen und dort als PDF speichern."
        )
        if errors:
            hint += " Details: " + " | ".join(errors)
        raise PdfConversionError(hint)
