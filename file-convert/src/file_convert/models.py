from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol


@dataclass(frozen=True)
class ConversionJob:
    source: Path
    source_format: str
    target_format: str
    output: Path | None = None
    options: dict[str, Any] = field(default_factory=dict)
    dry_run: bool = False
    force: bool = False
    timeout: int = 120


@dataclass(frozen=True)
class ConversionResult:
    output_path: Path
    duration_ms: int
    message: str | None = None


class Handler(Protocol):
    @property
    def pairs(self) -> frozenset[tuple[str, str]]: ...

    def validate(self, job: ConversionJob) -> None: ...

    def convert(self, job: ConversionJob, output_path: Path) -> ConversionResult: ...
