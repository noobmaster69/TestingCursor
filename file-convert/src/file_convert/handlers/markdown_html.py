from __future__ import annotations

import importlib.resources
from pathlib import Path

import markdown

from file_convert.handlers.base import BaseHandler
from file_convert.models import ConversionJob


class MarkdownHtmlHandler(BaseHandler):
    @property
    def pairs(self) -> frozenset[tuple[str, str]]:
        return frozenset({("md", "html")})

    def _load_css(self, theme: str) -> str:
        if theme.endswith(".css") and Path(theme).is_file():
            return Path(theme).read_text(encoding="utf-8")
        name = theme if theme in ("github", "minimal") else "minimal"
        pkg = importlib.resources.files("file_convert") / "data" / "themes" / f"{name}.css"
        return pkg.read_text(encoding="utf-8")

    def _run(self, job: ConversionJob, output_path: Path) -> str | None:
        text = job.source.read_text(encoding="utf-8-sig")
        body = markdown.markdown(
            text,
            extensions=["fenced_code", "tables", "toc", "codehilite"],
        )
        theme = job.options.get("theme", "minimal")
        title = job.options.get("title", job.source.stem)
        css = self._load_css(theme)
        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{_escape(title)}</title>
  <style>{css}</style>
</head>
<body>
{body}
</body>
</html>
"""
        output_path.write_text(html, encoding="utf-8")
        return f"Wrote HTML ({len(html)} bytes)"


def _escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
