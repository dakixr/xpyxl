"""Compatibility boundary for XlsxWriter constant-memory bookkeeping."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlsxwriter.format import Format
    from xlsxwriter.worksheet import Worksheet


class ConstantMemoryTracker:
    """Release metadata after XlsxWriter has flushed a worksheet row.

    Correctness uses public XlsxWriter methods. Private attributes only bound
    otherwise-retained bookkeeping; cleanup becomes a no-op when an unknown
    XlsxWriter version does not expose the expected internals.
    """

    _REQUIRED_ATTRIBUTES = (
        "previous_row",
        "set_rows",
        "row_sizes",
        "merged_cells",
        "dim_colmin",
        "dim_colmax",
        "_write_single_row",
    )

    def __init__(self, worksheet: Worksheet) -> None:
        self._worksheet = worksheet
        self._supported = all(
            hasattr(worksheet, attribute)
            for attribute in self._REQUIRED_ATTRIBUTES
        )
        self._released_rows = 0

    @property
    def current_row(self) -> int:
        return int(getattr(self._worksheet, "previous_row", -1))

    def advance_to(self, row: int, fallback_format: Format | None = None) -> None:
        """Advance the optimized row stream before writing a new row."""
        previous_row = self.current_row
        if row <= previous_row:
            return
        if self._supported:
            self._worksheet._write_single_row(row)  # pyright: ignore[reportPrivateUsage]
        elif fallback_format is not None:
            self._worksheet.write_blank(row, 0, None, fallback_format)
        self._release(previous_row)

    def _release(self, previous_row: int) -> None:
        if not self._supported or self.current_row <= previous_row:
            return
        worksheet = self._worksheet
        worksheet.set_rows.pop(previous_row, None)
        worksheet.row_sizes.pop(previous_row, None)
        min_col = worksheet.dim_colmin or 0
        max_col = worksheet.dim_colmax or min_col
        for col_idx in range(min_col, max_col + 1):
            worksheet.merged_cells.pop((previous_row, col_idx), None)
        self._released_rows += 1
        if self._released_rows % 1024 == 0:
            self.compact()

    def compact(self) -> None:
        """Return capacity from dictionaries whose entries were released."""
        if not self._supported:
            return
        for metadata in (
            self._worksheet.set_rows,
            self._worksheet.row_sizes,
            self._worksheet.merged_cells,
        ):
            retained = dict(metadata)
            metadata.clear()
            metadata.update(retained)
