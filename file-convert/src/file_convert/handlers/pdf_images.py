from __future__ import annotations

from pathlib import Path

import fitz
from rich.progress import Progress

from file_convert.errors import ConversionError, ValidationError
from file_convert.handlers.base import BaseHandler
from file_convert.models import ConversionJob, ConversionResult
from file_convert.output import resolve_multi_page_paths
from file_convert.utils import parse_pages, timer


class PdfImagesHandler(BaseHandler):
    @property
    def pairs(self) -> frozenset[tuple[str, str]]:
        return frozenset({("pdf", "png"), ("pdf", "jpg")})

    def _run(self, job: ConversionJob, output_path: Path) -> str | None:
        raise NotImplementedError("Use convert() for PDF image export")

    def validate(self, job: ConversionJob) -> None:
        super().validate(job)
        if job.target_format not in ("png", "jpg"):
            raise ValidationError(f"Unsupported image target: {job.target_format}")

    def convert(self, job: ConversionJob, output_path: Path) -> ConversionResult:
        """Override: multi-file output for multi-page PDFs."""
        self.validate(job)
        dpi = int(job.options.get("dpi", 150))
        quality = int(job.options.get("quality", 85))
        pages_spec = job.options.get("pages")

        doc = fitz.open(job.source)
        try:
            page_count = doc.page_count
            indices = parse_pages(pages_spec, page_count)
            out_paths = resolve_multi_page_paths(
                job,
                len(indices),
                single_page_index=indices[0] if len(indices) == 1 else None,
            )
            if job.dry_run:
                msg = f"[dry-run] Would render {len(indices)} page(s) to {len(out_paths)} file(s)"
                return ConversionResult(output_path=out_paths[0], duration_ms=0, message=msg)

            zoom = dpi / 72.0
            matrix = fitz.Matrix(zoom, zoom)
            use_progress = len(indices) > 5

            with timer() as elapsed, Progress() if use_progress else _null_context() as progress:
                task = None
                if use_progress:
                    task = progress.add_task("Rendering PDF pages...", total=len(indices))
                for out_path, page_idx in zip(out_paths, indices):
                    page = doc[page_idx]
                    pix = page.get_pixmap(matrix=matrix)
                    if job.target_format == "png":
                        pix.save(str(out_path))
                    else:
                        pix.save(str(out_path), jpg_quality=quality)
                    if task is not None:
                        progress.advance(task)

            message = f"Wrote {len(out_paths)} file(s)"
            return ConversionResult(
                output_path=out_paths[0],
                duration_ms=elapsed[0],
                message=message,
            )
        except ValueError as exc:
            raise ValidationError(str(exc)) from exc
        except Exception as exc:
            raise ConversionError(f"PDF conversion failed: {exc}") from exc
        finally:
            doc.close()


class _null_context:
    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass
