from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from ._workbook import Workbook
from .nodes import (
    CellNode,
    ColumnNode,
    HorizontalStackNode,
    RowNode,
    SheetComponent,
    SheetItem,
    SheetNode,
    SpacerNode,
    TableNode,
    VerticalStackNode,
    WorkbookNode,
)
from .styles import Style

__all__ = [
    "cell",
    "col",
    "row",
    "space",
    "vstack",
    "hstack",
    "sheet",
    "table",
    "workbook",
]


Node = (
    CellNode
    | RowNode
    | ColumnNode
    | TableNode
    | SpacerNode
    | VerticalStackNode
    | HorizontalStackNode
)


def _as_tuple(values: Any) -> tuple[Any, ...]:
    if isinstance(values, tuple):
        return values
    if isinstance(values, list):
        return tuple(values)
    if isinstance(values, Sequence) and not isinstance(values, (str, bytes, bytearray)):
        return tuple(values)
    return (values,)


def _ensure_cell(value: Any) -> CellNode:
    if isinstance(value, CellNode):
        return value
    if isinstance(value, (RowNode, ColumnNode, TableNode)):
        msg = "Cannot nest row/column/table directly inside a cell"
        raise TypeError(msg)
    return CellNode(value=value)


def _ensure_component(value: Any) -> SheetComponent:
    if isinstance(value, Node):
        return value
    msg = "Layouts accept composed nodes. Call the primitive builder before nesting."
    raise TypeError(msg)


def _coerce_row(value: Any, extra_styles: Sequence[Style] = ()) -> RowNode:
    if isinstance(value, RowNode):
        if not extra_styles:
            return value
        return RowNode(cells=value.cells, styles=tuple(extra_styles) + value.styles)
    cells = tuple(_ensure_cell(item) for item in _as_tuple(value))
    return RowNode(cells=cells, styles=tuple(extra_styles))


class _BuilderBase:
    def __init__(self, *, styles: Sequence[Style] | None = None) -> None:
        self._styles: tuple[Style, ...] = tuple(styles or ())


class CellBuilder(_BuilderBase):
    def __getitem__(self, value: Any) -> CellNode:
        return CellNode(value=value, styles=self._styles)


class RowBuilder(_BuilderBase):
    def __getitem__(self, values: Sequence[Any]) -> RowNode:
        cells = tuple(_ensure_cell(item) for item in _as_tuple(values))
        return RowNode(cells=cells, styles=self._styles)


class ColumnBuilder(_BuilderBase):
    def __getitem__(self, values: Sequence[Any]) -> ColumnNode:
        cells = tuple(_ensure_cell(item) for item in _as_tuple(values))
        return ColumnNode(cells=cells, styles=self._styles)


class TableBuilder(_BuilderBase):
    def __init__(
        self,
        *,
        header: Any | None = None,
        styles: Sequence[Style] | None = None,
        header_style: Sequence[Style] | None = None,
    ) -> None:
        super().__init__(styles=styles)
        self._header_raw = header
        self._header_styles: tuple[Style, ...] = tuple(header_style or ())

    def __getitem__(self, rows: Sequence[RowNode] | Sequence[list]) -> TableNode:
        row_nodes = tuple(_coerce_row(row) for row in _as_tuple(rows))
        header_node = None
        if self._header_raw is not None:
            header_node = _coerce_row(
                self._header_raw, extra_styles=self._header_styles
            )
        return TableNode(rows=row_nodes, styles=self._styles, header=header_node)


class SheetBuilder:
    def __init__(self, name: str) -> None:
        self._name = name

    def __getitem__(self, items: Any) -> SheetNode:
        entries: list[SheetItem] = []
        for item in _as_tuple(items):
            if isinstance(item, Node):
                entries.append(item)
            else:
                msg = (
                    "Sheets accept rows, columns, tables, spacers, or layout stacks. "
                    "Call the builder before nesting."
                )
                raise TypeError(msg)
        return SheetNode(name=self._name, items=tuple(entries))


class WorkbookBuilder:
    def __getitem__(self, sheets: Any) -> Workbook:
        sheet_nodes: list[SheetNode] = []
        for item in _as_tuple(sheets):
            if isinstance(item, SheetNode):
                sheet_nodes.append(item)
            else:
                raise TypeError(
                    "Workbooks accept sheet builders that have been indexed"
                )
        node = WorkbookNode(sheets=tuple(sheet_nodes))
        return Workbook(node)


def cell(*, style: Sequence[Style] | None = None) -> CellBuilder:
    return CellBuilder(styles=style)


def row(*, style: Sequence[Style] | None = None) -> RowBuilder:
    return RowBuilder(styles=style)


def col(*, style: Sequence[Style] | None = None) -> ColumnBuilder:
    return ColumnBuilder(styles=style)


def table(
    *,
    header: Any | None = None,
    style: Sequence[Style] | None = None,
    header_style: Sequence[Style] | None = None,
) -> TableBuilder:
    return TableBuilder(header=header, styles=style, header_style=header_style)


def sheet(name: str) -> SheetBuilder:
    return SheetBuilder(name)


def space(rows: int = 1, *, height: float | None = None) -> SpacerNode:
    if rows < 1:
        msg = "Spacer rows must be >= 1"
        raise ValueError(msg)
    return SpacerNode(rows=rows, height=height)


def vstack(*items: Any, gap: int = 0, style: Sequence[Style] | None = None) -> VerticalStackNode:
    if not items:
        raise ValueError("Vertical stack requires at least one item")
    if gap < 0:
        raise ValueError("Vertical stack gap must be >= 0")
    components = tuple(_ensure_component(item) for item in items)
    return VerticalStackNode(items=components, gap=gap, styles=tuple(style or ()))


def hstack(*items: Any, gap: int = 0, style: Sequence[Style] | None = None) -> HorizontalStackNode:
    if not items:
        raise ValueError("Horizontal stack requires at least one item")
    if gap < 0:
        raise ValueError("Horizontal stack gap must be >= 0")
    components = tuple(_ensure_component(item) for item in items)
    return HorizontalStackNode(items=components, gap=gap, styles=tuple(style or ()))


def workbook() -> WorkbookBuilder:
    return WorkbookBuilder()
