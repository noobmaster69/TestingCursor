from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from file_convert.errors import ConversionError, DependencyError
from file_convert.handlers.base import BaseHandler
from file_convert.models import ConversionJob


def find_soffice() -> Path | None:
    found = shutil.which("soffice")
    if found:
        return Path(found)
    # Common Windows install paths
    for pattern in (
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ):
        if pattern.is_file():
            return pattern
    return None


class OfficePdfHandler(BaseHandler):
    @property
    def pairs(self) -> frozenset[tuple[str, str]]:
        return frozenset({("docx", "pdf"), ("odt", "pdf")})

    def _run(self, job: ConversionJob, output_path: Path) -> str | None:
        soffice = find_soffice()
        if soffice is None:
            raise DependencyError(
                "LibreOffice not found. Install LibreOffice or run: convert --doctor"
            )

        cmd = [
            str(soffice),
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
        ]

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            cmd_outdir = str(tmp_path)
            full_cmd = cmd + [cmd_outdir, str(job.source.resolve())]

            try:
                subprocess.run(
                    full_cmd,
                    check=True,
                    capture_output=not job.options.get("verbose"),
                    timeout=job.timeout,
                )
            except subprocess.TimeoutExpired as exc:
                raise ConversionError(f"LibreOffice timed out after {job.timeout}s") from exc
            except subprocess.CalledProcessError as exc:
                raise ConversionError(f"LibreOffice failed: {exc}") from exc

            produced = tmp_path / f"{job.source.stem}.pdf"
            if not produced.is_file():
                pdfs = list(tmp_path.glob("*.pdf"))
                if not pdfs:
                    raise ConversionError("LibreOffice did not produce a PDF file")
                produced = pdfs[0]

            shutil.move(str(produced), str(output_path))

        return "Wrote PDF via LibreOffice"

    def _dry_run_message(self, job: ConversionJob, output_path: Path) -> str:
        soffice = find_soffice()
        binary = str(soffice) if soffice else "(not found)"
        return (
            f"[dry-run] LibreOffice: {binary}\n"
            f"  Input:  {job.source}\n"
            f"  Output: {output_path}"
        )
