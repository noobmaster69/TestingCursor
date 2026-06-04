from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook


@pytest.fixture
def fixtures_dir() -> Path:
    return Path(__file__).parent / "fixtures"


@pytest.fixture
def sample_xlsx(tmp_path: Path, fixtures_dir: Path) -> Path:
    path = tmp_path / "sample.xlsx"
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Summary"
    ws1.append(["id", "amount"])
    ws1.append([1, 100])
    ws2 = wb.create_sheet("Q1")
    ws2.append(["month", "total"])
    ws2.append(["Jan", 50])
    wb.save(path)
    return path


@pytest.fixture
def sample_pdf(tmp_path: Path) -> Path:
    import fitz

    path = tmp_path / "sample.pdf"
    doc = fitz.open()
    for i in range(2):
        page = doc.new_page()
        page.insert_text((72, 72), f"Page {i + 1}")
    doc.save(path)
    doc.close()
    return path


@pytest.fixture
def sample_zip(tmp_path: Path) -> Path:
    import zipfile

    path = tmp_path / "sample.zip"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("readme.txt", "hello")
        zf.writestr("data/info.json", "{}")
    return path
