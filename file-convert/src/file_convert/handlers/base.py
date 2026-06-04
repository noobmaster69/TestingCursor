from __future__ import annotations

import time
from abc import ABC, abstractmethod
from pathlib import Path

from file_convert.models import ConversionJob, ConversionResult


class BaseHandler(ABC):
    @property
    @abstractmethod
    def pairs(self) -> frozenset[tuple[str, str]]: ...

    def validate(self, job: ConversionJob) -> None:
        if not job.source.is_file():
            from file_convert.errors import ValidationError

            raise ValidationError(f"Input file not found: {job.source}")

    @abstractmethod
    def _run(self, job: ConversionJob, output_path: Path) -> str | None: ...

    def convert(self, job: ConversionJob, output_path: Path) -> ConversionResult:
        self.validate(job)
        if job.dry_run:
            msg = self._dry_run_message(job, output_path)
            return ConversionResult(output_path=output_path, duration_ms=0, message=msg)

        start = time.perf_counter()
        message = self._run(job, output_path)
        duration_ms = int((time.perf_counter() - start) * 1000)
        return ConversionResult(output_path=output_path, duration_ms=duration_ms, message=message)

    def _dry_run_message(self, job: ConversionJob, output_path: Path) -> str:
        return f"[dry-run] Would write: {output_path}"
