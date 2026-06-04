from pathlib import Path

from file_convert.core import ConversionRequest, run_conversion, targets_for_source


def test_targets_for_md():
    assert "html" in targets_for_source("md")


def test_run_md_to_html(fixtures_dir: Path, tmp_path: Path):
    out = tmp_path / "out.html"
    req = ConversionRequest(
        source=fixtures_dir / "sample.md",
        target_format="html",
        output=out,
        options={"theme": "minimal"},
    )
    resp = run_conversion(req)
    assert resp.ok
    assert out.is_file()
