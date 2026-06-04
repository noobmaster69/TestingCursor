from pathlib import Path

import pytest

from file_convert.errors import UserError
from file_convert.models import ConversionJob
from file_convert.output import resolve_output_path


def test_default_output_beside_source(tmp_path: Path):
    src = tmp_path / "report.md"
    src.write_text("# hi", encoding="utf-8")
    job = ConversionJob(source=src, source_format="md", target_format="html")
    out = resolve_output_path(job)
    assert out == tmp_path / "report.html"


def test_output_directory(tmp_path: Path):
    src = tmp_path / "data.csv"
    src.write_text("a\n", encoding="utf-8")
    out_dir = tmp_path / "out"
    job = ConversionJob(
        source=src,
        source_format="csv",
        target_format="xlsx",
        output=out_dir,
    )
    out = resolve_output_path(job)
    assert out == out_dir / "data.xlsx"


def test_collision_raises(tmp_path: Path):
    src = tmp_path / "a.md"
    src.write_text("x", encoding="utf-8")
    existing = tmp_path / "a.html"
    existing.write_text("<html></html>", encoding="utf-8")
    job = ConversionJob(source=src, source_format="md", target_format="html")
    with pytest.raises(UserError):
        resolve_output_path(job)
