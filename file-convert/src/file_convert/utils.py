from __future__ import annotations

import time
from contextlib import contextmanager
from typing import Iterator


@contextmanager
def timer() -> Iterator[list[int]]:
    start = time.perf_counter()
    elapsed: list[int] = [0]
    try:
        yield elapsed
    finally:
        elapsed[0] = int((time.perf_counter() - start) * 1000)


def parse_pages(spec: str | None, page_count: int) -> list[int]:
    """Parse 1-based page spec into sorted unique 0-based indices."""
    if not spec or spec.strip().lower() == "all":
        return list(range(page_count))

    indices: set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start_s, end_s = part.split("-", 1)
            start = int(start_s.strip())
            end = int(end_s.strip())
            if start < 1 or end < 1 or start > end:
                raise ValueError(f"Invalid page range: {part}")
            for p in range(start, end + 1):
                if p > page_count:
                    raise ValueError(f"Page {p} out of range (document has {page_count} pages)")
                indices.add(p - 1)
        else:
            p = int(part)
            if p < 1 or p > page_count:
                raise ValueError(f"Page {p} out of range (document has {page_count} pages)")
            indices.add(p - 1)

    return sorted(indices)
