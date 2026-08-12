from __future__ import annotations

from typing import cast

from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

from xpyxl.engines._xlsxwriter_streaming import ConstantMemoryTracker


class _UnknownWorksheet:
    """Small stand-in for an XlsxWriter version with different internals."""

    def __init__(self) -> None:
        self.previous_row = -1
        self.blank_writes: list[tuple[int, int]] = []

    def write_blank(
        self,
        row: int,
        col: int,
        value: object,
        cell_format: object,
    ) -> None:
        self.blank_writes.append((row, col))
        self.previous_row = row


def test_unknown_xlsxwriter_internals_use_public_spacer_fallback() -> None:
    worksheet = _UnknownWorksheet()
    tracker = ConstantMemoryTracker(cast(Worksheet, worksheet))

    tracker.advance_to(4, cast(Format, object()))
    tracker.compact()

    assert worksheet.blank_writes == [(4, 0)]
