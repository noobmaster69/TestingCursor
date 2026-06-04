from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from file_convert.errors import ConvertError
from file_convert.formats import infer_format, normalize_format
from file_convert.handlers import build_registry
from file_convert.models import ConversionJob, ConversionResult
from file_convert.output import resolve_output_path
from file_convert.registry import Registry

_registry: Registry | None = None


def get_registry() -> Registry:
    global _registry
    if _registry is None:
        _registry = build_registry()
    return _registry


def list_all_pairs() -> list[tuple[str, str]]:
    return get_registry().list_pairs()


def targets_for_source(source_format: str | None) -> list[str]:
    if not source_format:
        return []
    src = normalize_format(source_format)
    targets = sorted({tgt for s, tgt in get_registry().list_pairs() if s == src})
    return targets


@dataclass
class ConversionRequest:
    source: Path
    target_format: str
    output: Path | None = None
    from_format: str | None = None
    force: bool = False
    dry_run: bool = False
    timeout: int = 120
    options: dict[str, Any] = field(default_factory=dict)


@dataclass
class ConversionResponse:
    ok: bool
    message: str
    output_path: Path | None = None
    exit_code: int = 0
    result: ConversionResult | None = None


def run_conversion(request: ConversionRequest) -> ConversionResponse:
    source = request.source.resolve()
    if not source.is_file():
        return ConversionResponse(ok=False, message=f"File not found: {source}", exit_code=1)

    src_fmt = (
        normalize_format(request.from_format)
        if request.from_format
        else infer_format(source)
    )
    if not src_fmt:
        return ConversionResponse(
            ok=False,
            message="Could not detect file type. Choose a supported file.",
            exit_code=1,
        )

    tgt_fmt = normalize_format(request.target_format)
    registry = get_registry()

    job = ConversionJob(
        source=source,
        source_format=src_fmt,
        target_format=tgt_fmt,
        output=request.output,
        options=dict(request.options),
        dry_run=request.dry_run,
        force=request.force,
        timeout=request.timeout,
    )
    if job.options.get("title") is None:
        job.options["title"] = source.stem

    try:
        handler = registry.resolve(src_fmt, tgt_fmt)
        handler.validate(job)

        if tgt_fmt == "list":
            result = handler.convert(job, source)
            text = result.message or "Done."
            return ConversionResponse(
                ok=True,
                message=text,
                output_path=source,
                exit_code=0,
                result=result,
            )

        out_path = resolve_output_path(job)
        result = handler.convert(job, out_path)
        msg = result.message or f"Saved to {result.output_path}"
        return ConversionResponse(
            ok=True,
            message=msg,
            output_path=result.output_path,
            exit_code=0,
            result=result,
        )
    except ConvertError as exc:
        return ConversionResponse(ok=False, message=exc.message, exit_code=exc.exit_code)
