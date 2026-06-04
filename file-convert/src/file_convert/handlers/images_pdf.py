from __future__ import annotations

from pathlib import Path

import img2pdf
from PIL import Image

from file_convert.handlers.base import BaseHandler
from file_convert.models import ConversionJob


class ImagesPdfHandler(BaseHandler):
    @property
    def pairs(self) -> frozenset[tuple[str, str]]:
        return frozenset(
            {
                ("png", "pdf"),
                ("jpg", "pdf"),
                ("jpeg", "pdf"),
                ("webp", "pdf"),
                ("tiff", "pdf"),
            }
        )

    def _run(self, job: ConversionJob, output_path: Path) -> str | None:
        with Image.open(job.source) as img:
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
                temp = output_path.with_suffix(".tmp.jpg")
                img.save(temp, format="JPEG", quality=95)
                try:
                    pdf_bytes = img2pdf.convert(str(temp))
                finally:
                    temp.unlink(missing_ok=True)
            else:
                pdf_bytes = img2pdf.convert(str(job.source))

        output_path.write_bytes(pdf_bytes)
        return "Wrote PDF"
