"""Format name normalization and aliases."""

from __future__ import annotations

from pathlib import Path

ALIASES: dict[str, str] = {
    "jpeg": "jpg",
    "tif": "tiff",
    "markdown": "md",
    "htm": "html",
}


def normalize_format(fmt: str) -> str:
    f = fmt.lower().lstrip(".")
    return ALIASES.get(f, f)


def infer_format(path: Path) -> str | None:
    if not path.suffix:
        return None
    return normalize_format(path.suffix[1:])
