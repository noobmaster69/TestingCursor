from __future__ import annotations

import zipfile
from pathlib import Path

from file_convert.errors import ValidationError
from file_convert.handlers.base import BaseHandler
from file_convert.models import ConversionJob, ConversionResult


class ZipListHandler(BaseHandler):
    """Zip listing is a special action (no output file)."""

    @property
    def pairs(self) -> frozenset[tuple[str, str]]:
        return frozenset({("zip", "list")})

    def convert(self, job: ConversionJob, output_path: Path) -> ConversionResult:
        self.validate(job)
        lines = self._format_listing(job.source)
        text = "\n".join(lines)
        if job.dry_run:
            return ConversionResult(
                output_path=job.source,
                duration_ms=0,
                message=f"[dry-run] Would list {job.source} ({len(lines) - 1} entries)",
            )
        print(text)
        return ConversionResult(
            output_path=job.source,
            duration_ms=0,
            message=f"Listed {len(lines) - 1} entries",
        )

    def _format_listing(self, path: Path) -> list[str]:
        header = f"{'Name':<40} {'Size':>12} {'Compressed':>12}  Dir"
        sep = "-" * len(header)
        rows = [header, sep]
        with zipfile.ZipFile(path, "r") as zf:
            for info in zf.infolist():
                name = info.filename[:40]
                rows.append(
                    f"{name:<40} {info.file_size:>12} {info.compress_size:>12}  "
                    f"{'yes' if info.is_dir() else 'no'}"
                )
        return rows

    def _run(self, job: ConversionJob, output_path: Path) -> str | None:
        raise NotImplementedError

    def validate(self, job: ConversionJob) -> None:
        super().validate(job)
        if not zipfile.is_zipfile(job.source):
            raise ValidationError(f"Not a valid zip file: {job.source}")
