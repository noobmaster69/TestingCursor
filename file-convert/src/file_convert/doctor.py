from __future__ import annotations

import importlib
import shutil

from rich.console import Console

from file_convert import __version__
from file_convert.handlers.heic_jpeg import _HEIC_AVAILABLE
from file_convert.handlers.office_pdf import find_soffice
from file_convert.handlers import build_registry

console = Console()


def _check_module(name: str, module: str) -> str:
    try:
        importlib.import_module(module)
        return "OK"
    except ImportError:
        return "MISSING"


def run_doctor() -> None:
    console.print(f"file-convert {__version__}\n")
    console.print("[bold]Core (Python)[/bold]")

    modules = [
        ("markdown", "markdown"),
        ("pymupdf", "fitz"),
        ("openpyxl", "openpyxl"),
        ("pandas", "pandas"),
        ("Pillow", "PIL"),
        ("img2pdf", "img2pdf"),
    ]
    for label, mod in modules:
        status = _check_module(label, mod)
        console.print(f"  {label:<16} {status}")

    console.print("\n[bold]Optional[/bold]")
    heic_status = "OK" if _HEIC_AVAILABLE else "MISSING  (pip install file-convert[heic])"
    console.print(f"  {'pillow-heif':<16} {heic_status}")

    soffice = find_soffice()
    if soffice:
        lo_status = f"OK  {soffice}"
    else:
        lo_status = "MISSING  (install LibreOffice for docx→pdf)"
    console.print(f"  {'LibreOffice':<16} {lo_status}")

    registry = build_registry()
    pairs = registry.list_pairs()
    available = len(pairs)
    if not _HEIC_AVAILABLE:
        available_note = " (heic→jpg unavailable without [heic])"
    else:
        available_note = ""
    console.print(f"\n[bold]Ready:[/bold] {available} conversion pair(s) available{available_note}")
