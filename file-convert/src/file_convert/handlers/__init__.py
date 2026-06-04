from __future__ import annotations

from file_convert.handlers.heic_jpeg import HeicJpegHandler
from file_convert.handlers.images_pdf import ImagesPdfHandler
from file_convert.handlers.markdown_html import MarkdownHtmlHandler
from file_convert.handlers.office_pdf import OfficePdfHandler
from file_convert.handlers.pdf_images import PdfImagesHandler
from file_convert.handlers.spreadsheet import SpreadsheetHandler
from file_convert.handlers.zip_list import ZipListHandler
from file_convert.registry import Registry


def build_registry() -> Registry:
    registry = Registry()
    handlers = [
        MarkdownHtmlHandler(),
        SpreadsheetHandler(),
        PdfImagesHandler(),
        ImagesPdfHandler(),
        HeicJpegHandler(),
        OfficePdfHandler(),
        ZipListHandler(),
    ]
    for handler in handlers:
        if handler.pairs:
            registry.register(handler)
    return registry
