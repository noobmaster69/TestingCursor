from __future__ import annotations

from pathlib import Path
from typing import Annotated, Optional

import typer
from rich.console import Console

from file_convert import __version__
from file_convert.core import ConversionRequest, get_registry, run_conversion
from file_convert.doctor import run_doctor
from file_convert.errors import ConvertError
from file_convert.formats import infer_format, normalize_format

console = Console()


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
    gui_flag: Annotated[bool, typer.Option("--gui", help="Open graphical interface")] = False,
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
    if gui_flag:
        from file_convert.gui import run as run_gui

        run_gui()
        raise typer.Exit(0)

    if version_flag:
        console.print(__version__)
        raise typer.Exit(0)

    if list_formats_flag:
        for src, tgt in get_registry().list_pairs():
            console.print(f"  {src} -> {tgt}")
        raise typer.Exit(0)

    if doctor_flag:
        run_doctor()
        raise typer.Exit(0)

    if input_path is None:
        console.print("Usage: file-convert INPUT --to FORMAT")
        console.print("       file-convert --gui")
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

    request = ConversionRequest(
        source=input_path,
        target_format=tgt_fmt,
        output=output,
        from_format=from_format,
        dry_run=dry_run,
        force=force,
        timeout=timeout,
        options=options,
    )

    try:
        response = run_conversion(request)
        if not response.ok:
            console.print(f"[red]{response.message}[/red]")
            raise typer.Exit(response.exit_code)

        if tgt_fmt == "list":
            console.print(response.message)
        else:
            style = "yellow" if dry_run else "green"
            console.print(f"[{style}]{response.message}[/{style}]")
            if response.output_path:
                console.print(f"Output: {response.output_path}")
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
