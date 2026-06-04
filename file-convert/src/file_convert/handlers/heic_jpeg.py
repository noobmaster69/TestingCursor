from __future__ import annotations

from pathlib import Path

from file_convert.errors import DependencyError
from file_convert.handlers.base import BaseHandler
from file_convert.models import ConversionJob

_HEIC_AVAILABLE = False

try:
    import pillow_heif

    pillow_heif.register_heif_opener()
    _HEIC_AVAILABLE = True
except ImportError:
    pass


class HeicJpegHandler(BaseHandler):
    @property
    def pairs(self) -> frozenset[tuple[str, str]]:
        if not _HEIC_AVAILABLE:
            return frozenset()
        return frozenset({("heic", "jpg")})

    def validate(self, job: ConversionJob) -> None:
        super().validate(job)
        if not _HEIC_AVAILABLE:
            raise DependencyError(
                "HEIC support requires pillow-heif. Install: pip install file-convert[heic]"
            )

    def _run(self, job: ConversionJob, output_path: Path) -> str | None:
        from PIL import Image

        quality = int(job.options.get("quality", 85))
        max_w = job.options.get("max_width")
        max_h = job.options.get("max_height")

        with Image.open(job.source) as img:
            img = _apply_exif_orientation(img)
            if max_w or max_h:
                img.thumbnail(
                    (int(max_w) if max_w else img.width, int(max_h) if max_h else img.height),
                    Image.Resampling.LANCZOS,
                )
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            img.save(output_path, format="JPEG", quality=quality)
        return "Wrote JPEG"


def _apply_exif_orientation(img):
    try:
        from PIL import ImageOps

        return ImageOps.exif_transpose(img)
    except Exception:
        return img
