from __future__ import annotations

from pathlib import Path
from typing import Annotated, Optional

import typer
from rich.console import Console

from file_convert import __version__
from file_convert.doctor import run_doctor
from file_convert.errors import ConvertError
from file_convert.formats import infer_format, normalize_format
from file_convert.handlers import build_registry
from file_convert.models import ConversionJob
from file_convert.output import resolve_output_path

console = Console()
_registry = None


def _get_registry():
    global _registry
    if _registry is None:
        _registry = build_registry()
    return _registry


def main(
    input_path: Annotated[
        Optional[Path],
        typer.Argument(help="Input file path", exists=True, dir_okay=False, readable=True),
    ] = None,
    to: Annotated[Optional[str], typer.Option("--to", help="Target format (pdf, html, csv, ...)")] = None,
    output: Annotated[Optional[Path], typer.Option("-o", "--output", help="Output file or directory")] = None,
    from_format: Annotated[
        Optional[str], typer.Option("--from-format", help="Override source format")
    ] = None,
    list_zip: Annotated[bool, typer.Option("--list", help="List zip contents")] = False,
    list_formats_flag: Annotated[
        bool, typer.Option("--list-formats", help="Show supported conversion pairs")
    ] = False,
    doctor_flag: Annotated[bool, typer.Option("--doctor", help="Check dependencies")] = False,
    version_flag: Annotated[bool, typer.Option("--version", help="Show version")] = False,
    dry_run: Annotated[bool, typer.Option("--dry-run", help="Preview without writing")] = False,
    force: Annotated[bool, typer.Option("--force", help="Overwrite existing output")] = False,
    verbose: Annotated[bool, typer.Option("--verbose", help="Verbose errors")] = False,
    timeout: Annotated[int, typer.Option("--timeout", help="Subprocess timeout (seconds)")] = 120,
    sheet: Annotated[Optional[str], typer.Option("--sheet", help="Sheet name or index")] = None,
    encoding: Annotated[str, typer.Option("--encoding", help="CSV encoding")] = "utf-8",
    delimiter: Annotated[str, typer.Option("--delimiter", help="CSV delimiter")] = ",",
    theme: Annotated[str, typer.Option("--theme", help="MD theme: github, minimal, or .css path")] = "minimal",
    title: Annotated[Optional[str], typer.Option("--title", help="HTML title")] = None,
    pages: Annotated[Optional[str], typer.Option("--pages", help="PDF pages spec")] = None,
    dpi: Annotated[int, typer.Option("--dpi", help="PDF DPI")] = 150,
    quality: Annotated[int, typer.Option("--quality", help="JPEG quality 1-100")] = 85,
    max_width: Annotated[Optional[int], typer.Option("--max-width")] = None,
    max_height: Annotated[Optional[int], typer.Option("--max-height")] = None,
) -> None:
    """Local-first file format converter."""
    if version_flag:
        console.print(__version__)
        raise typer.Exit(0)

    if list_formats_flag:
        registry = _get_registry()
        for src, tgt in registry.list_pairs():
            console.print(f"  {src} -> {tgt}")
        raise typer.Exit(0)

    if doctor_flag:
        run_doctor()
        raise typer.Exit(0)

    if input_path is None:
        console.print("Usage: convert INPUT --to FORMAT  (or convert --list-formats)")
        raise typer.Exit(1)

    registry = _get_registry()
    src_fmt = normalize_format(from_format) if from_format else infer_format(input_path)
    if not src_fmt:
        console.print("[red]Could not detect input format. Use --from-format.[/red]")
        raise typer.Exit(1)

    if list_zip or (to and normalize_format(to) == "list"):
        tgt_fmt = "list"
    elif to:
        tgt_fmt = normalize_format(to)
    else:
        console.print("[red]Missing --to FORMAT (or --list for zip files).[/red]")
        raise typer.Exit(1)

    options = {
        "sheet": sheet,
        "encoding": encoding,
        "delimiter": delimiter,
        "theme": theme,
        "title": title or input_path.stem,
        "pages": pages,
        "dpi": dpi,
        "quality": quality,
        "max_width": max_width,
        "max_height": max_height,
        "verbose": verbose,
    }

    job = ConversionJob(
        source=input_path,
        source_format=src_fmt,
        target_format=tgt_fmt,
        output=output,
        options=options,
        dry_run=dry_run,
        force=force,
        timeout=timeout,
    )

    try:
        handler = registry.resolve(src_fmt, tgt_fmt)
        handler.validate(job)

        if tgt_fmt == "list":
            result = handler.convert(job, input_path)
        else:
            out_path = resolve_output_path(job)
            result = handler.convert(job, out_path)

        if result.message:
            style = "yellow" if dry_run else "green"
            console.print(f"[{style}]{result.message}[/{style}]")
        if tgt_fmt != "list":
            console.print(f"Output: {result.output_path}")
        raise typer.Exit(0)

    except ConvertError as exc:
        console.print(f"[red]{exc.message}[/red]")
        if verbose:
            console.print_exception()
        raise typer.Exit(exc.exit_code) from exc
    except KeyboardInterrupt:
        raise typer.Exit(130) from None


def app() -> None:
    """Console script entry point."""
    typer.run(main)


if __name__ == "__main__":
    app()
