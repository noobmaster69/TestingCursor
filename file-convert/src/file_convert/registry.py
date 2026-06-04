from __future__ import annotations

from file_convert.errors import UserError
from file_convert.models import Handler


class Registry:
    def __init__(self) -> None:
        self._handlers: list[Handler] = []
        self._pair_index: dict[tuple[str, str], Handler] = {}

    def register(self, handler: Handler) -> None:
        for pair in handler.pairs:
            if pair in self._pair_index:
                raise RuntimeError(f"Duplicate handler registration for {pair[0]}→{pair[1]}")
            self._pair_index[pair] = handler
        self._handlers.append(handler)

    def resolve(self, src_fmt: str, tgt_fmt: str) -> Handler:
        key = (src_fmt, tgt_fmt)
        handler = self._pair_index.get(key)
        if handler is None:
            supported = self.list_pairs()
            pairs_str = ", ".join(f"{a}→{b}" for a, b in supported[:12])
            extra = " ..." if len(supported) > 12 else ""
            raise UserError(
                f"Cannot convert .{src_fmt} → .{tgt_fmt}. "
                f"Supported: {pairs_str}{extra}. Run: convert --list-formats"
            )
        return handler

    def list_pairs(self) -> list[tuple[str, str]]:
        return sorted(self._pair_index.keys())
