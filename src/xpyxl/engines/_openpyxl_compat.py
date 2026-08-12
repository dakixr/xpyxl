"""Compatibility boundary for performance-sensitive OpenPyXL internals."""

from __future__ import annotations

import copy
from collections.abc import Iterable, Mapping
from typing import TYPE_CHECKING, Generic, Hashable, TypeVar

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.styles.cell_style import StyleArray

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet

_Key = TypeVar("_Key", bound=Hashable)


class CompiledStyleCache(Generic[_Key]):
    """Cache destination-workbook styles without sharing mutable records."""

    def __init__(self) -> None:
        self._styles: dict[_Key, StyleArray] = {}

    def apply(self, cell: Cell, key: _Key) -> bool:
        compiled = self._styles.get(key)
        if compiled is None or not hasattr(cell, "_style"):
            return False
        cell._style = copy.copy(compiled)
        return True

    def capture(self, cell: Cell, key: _Key) -> None:
        style = getattr(cell, "_style", None)
        if isinstance(style, StyleArray):
            self._styles[key] = copy.copy(style)


def source_style_key(cell: Cell) -> tuple[int, ...] | None:
    """Return a workbook-local style identity when OpenPyXL exposes one."""
    style = getattr(cell, "_style", None)
    if not isinstance(style, StyleArray):
        return None
    return tuple(style)


def populated_cells(sheet: Worksheet) -> Iterable[Cell | MergedCell]:
    """Iterate stored cells without materializing a sparse rectangular range."""
    cells = getattr(sheet, "_cells", None)
    if isinstance(cells, Mapping):
        return cells.values()
    return (cell for row in sheet.iter_rows() for cell in row)
