from __future__ import annotations

from pathlib import Path

from file_convert.errors import UserError
from file_convert.models import ConversionJob


def resolve_output_path(job: ConversionJob) -> Path:
    """Resolve final output file path for a single-file conversion."""
    source = job.source.resolve()
    stem = source.stem
    ext = f".{job.target_format}"

    if job.output is None:
        target = source.with_name(stem + ext)
    else:
        out = job.output.resolve()
        if out.is_dir() or (not out.exists() and out.suffix.lower() != ext):
            out.mkdir(parents=True, exist_ok=True)
            target = out / (stem + ext)
        else:
            if out.suffix.lower() != ext and not out.is_dir():
                target = out.with_suffix(ext)
            else:
                target = out

    if target.resolve() == source.resolve():
        raise UserError("Output path cannot be the same as the input file.")

    if target.exists() and not job.force:
        raise UserError(
            f"Output already exists: {target}. Use --force to overwrite or -o for another path."
        )

    target.parent.mkdir(parents=True, exist_ok=True)
    return target


def resolve_multi_page_paths(
    job: ConversionJob,
    page_count: int,
    *,
    single_page_index: int | None = None,
) -> list[Path]:
    """Resolve output paths for PDF → multi-page images."""
    source = job.source.resolve()
    stem = source.stem
    ext = f".{job.target_format}"

    if page_count == 1 and single_page_index is not None:
        paths = [_single_output(job, stem, ext)]
        _check_collisions(paths, job)
        return paths

    base_dir = _output_directory(job, source)
    base_dir.mkdir(parents=True, exist_ok=True)
    paths = [base_dir / f"{stem}_page_{i + 1:04d}{ext}" for i in range(page_count)]
    _check_collisions(paths, job)
    return paths


def _single_output(job: ConversionJob, stem: str, ext: str) -> Path:
    if job.output is None:
        return job.source.resolve().with_name(stem + ext)
    out = job.output.resolve()
    if out.is_dir():
        return out / (stem + ext)
    if out.suffix.lower() != ext:
        return out.with_suffix(ext.lstrip("."))
    return out


def _output_directory(job: ConversionJob, source: Path) -> Path:
    if job.output is None:
        return source.parent
    out = job.output.resolve()
    if out.is_dir():
        return out
    return out.parent


def _check_collisions(paths: list[Path], job: ConversionJob) -> None:
    existing = [p for p in paths if p.exists()]
    if existing and not job.force:
        raise UserError(
            f"{len(existing)} output file(s) already exist (e.g. {existing[0]}). "
            "Use --force to overwrite."
        )
