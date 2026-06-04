from pathlib import Path

from file_convert.handlers.markdown_html import MarkdownHtmlHandler
from file_convert.models import ConversionJob


def test_md_to_html(fixtures_dir: Path, tmp_path: Path):
    src = fixtures_dir / "sample.md"
    out = tmp_path / "out.html"
    job = ConversionJob(source=src, source_format="md", target_format="html", options={"theme": "github"})
    handler = MarkdownHtmlHandler()
    result = handler.convert(job, out)
    assert out.is_file()
    text = out.read_text(encoding="utf-8")
    assert "<h1>Sample Report</h1>" in text or "Sample Report" in text
    assert result.message
