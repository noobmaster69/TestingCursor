from pathlib import Path

from file_convert.handlers.spreadsheet import SpreadsheetHandler
from file_convert.models import ConversionJob


def test_xlsx_to_csv(sample_xlsx: Path, tmp_path: Path):
    out = tmp_path / "out.csv"
    job = ConversionJob(
        source=sample_xlsx,
        source_format="xlsx",
        target_format="csv",
        options={"sheet": "Q1"},
    )
    SpreadsheetHandler().convert(job, out)
    text = out.read_text(encoding="utf-8")
    assert "month" in text
    assert "Jan" in text


def test_csv_to_xlsx(fixtures_dir: Path, tmp_path: Path):
    src = fixtures_dir / "sample.csv"
    out = tmp_path / "out.xlsx"
    job = ConversionJob(source=src, source_format="csv", target_format="xlsx")
    SpreadsheetHandler().convert(job, out)
    assert out.is_file()
    assert out.stat().st_size > 0
