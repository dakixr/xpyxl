from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import TypeAlias, cast

from ._workbook import Workbook
from .nodes import (
    CellNode,
    CellValue,
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


ColumnKey: TypeAlias = str
CellSource: TypeAlias = CellValue | CellNode
Node = (
    CellNode
    | RowNode
    | ColumnNode
    | TableNode
    | SpacerNode
    | VerticalStackNode
    | HorizontalStackNode
)


def _as_tuple(values: object) -> tuple[object, ...]:
    if isinstance(values, tuple):
        return values
    if isinstance(values, list):
        return tuple(values)
    if isinstance(values, Sequence) and not isinstance(values, (str, bytes, bytearray)):
        return tuple(values)
    return (values,)


def _ensure_cell(value: CellSource) -> CellNode:
    if isinstance(value, CellNode):
        return value
    if isinstance(value, (RowNode, ColumnNode, TableNode)):
        msg = "Cannot nest row/column/table directly inside a cell"
        raise TypeError(msg)
    return CellNode(value=value)


def _ensure_component(value: object) -> SheetComponent:
    if isinstance(value, Node):
        return value
    msg = "Layouts accept composed nodes. Call the primitive builder before nesting."
    raise TypeError(msg)


def _coerce_row(value: object, extra_styles: Sequence[Style] = ()) -> RowNode:
    if isinstance(value, RowNode):
        if not extra_styles:
            return value
        return RowNode(cells=value.cells, styles=tuple(extra_styles) + value.styles)
    cells = tuple(_ensure_cell(item) for item in _as_tuple(value))
    return RowNode(cells=cells, styles=tuple(extra_styles))


def _rows_from_records(
    records: Sequence[Mapping[ColumnKey, CellSource]],
    *,
    header_styles: Sequence[Style],
    column_order: Sequence[ColumnKey] | None,
) -> tuple[tuple[RowNode, ...], RowNode | None]:
    if not records and not column_order:
        return (), None

    columns: list[ColumnKey] = list(column_order or ())
    seen = set(columns)
    for record in records:
        for key in record.keys():
            if key not in seen:
                seen.add(key)
                columns.append(key)

    header_node = _coerce_row(columns, extra_styles=header_styles) if columns else None
    body_rows = tuple(
        RowNode(cells=tuple(_ensure_cell(record.get(col)) for col in columns))
        for record in records
    )
    return body_rows, header_node


def _rows_from_dict_of_lists(
    table_data: Mapping[ColumnKey, Sequence[CellSource]],
    *,
    header_styles: Sequence[Style],
    column_order: Sequence[ColumnKey] | None,
) -> tuple[tuple[RowNode, ...], RowNode | None]:
    columns: list[ColumnKey] = list(column_order or ())
    seen = set(columns)
    for name in table_data.keys():
        if name not in seen:
            seen.add(name)
            columns.append(name)

    if not columns:
        return (), None

    normalized: dict[ColumnKey, Sequence[CellSource]] = {}
    lengths: set[int] = set()
    for key, values in table_data.items():
        if isinstance(values, Mapping) or isinstance(values, (str, bytes, bytearray)):
            msg = (
                "Dict-of-lists table data must map column names to sequences of values"
            )
            raise TypeError(msg)
        if not isinstance(values, Sequence):
            values = tuple(values)
        normalized[key] = values
        lengths.add(len(values))

    if len(lengths) > 1:
        msg = f"All columns must be the same length; got lengths {sorted(lengths)}"
        raise ValueError(msg)

    row_count = lengths.pop() if lengths else 0
    body_rows: list[RowNode] = []
    for idx in range(row_count):
        body_rows.append(
            RowNode(
                cells=tuple(
                    _ensure_cell(
                        normalized[col][idx]
                        if col in normalized and idx < len(normalized[col])
                        else None
                    )
                    for col in columns
                )
            )
        )

    header_node = _coerce_row(columns, extra_styles=header_styles)
    return tuple(body_rows), header_node


class _BuilderBase:
    def __init__(self, *, styles: Sequence[Style] | None = None) -> None:
        self._styles: tuple[Style, ...] = tuple(styles or ())


class CellBuilder(_BuilderBase):
    def __getitem__(self, value: CellSource) -> CellNode:
        return CellNode(value=value, styles=self._styles)


class RowBuilder(_BuilderBase):
    def __getitem__(self, values: Sequence[CellSource] | CellSource) -> RowNode:
        cells = tuple(_ensure_cell(item) for item in _as_tuple(values))
        return RowNode(cells=cells, styles=self._styles)


class ColumnBuilder(_BuilderBase):
    def __getitem__(self, values: Sequence[CellSource] | CellSource) -> ColumnNode:
        cells = tuple(_ensure_cell(item) for item in _as_tuple(values))
        return ColumnNode(cells=cells, styles=self._styles)


class TableBuilder(_BuilderBase):
    def __init__(
        self,
        *,
        styles: Sequence[Style] | None = None,
        header_style: Sequence[Style] | None = None,
        columns: Sequence[ColumnKey] | None = None,
    ) -> None:
        super().__init__(styles=styles)
        self._header_styles: tuple[Style, ...] = tuple(header_style or ())
        self._columns: tuple[ColumnKey, ...] | None = (
            tuple(columns) if columns is not None else None
        )

    def __getitem__(
        self,
        rows: Sequence[RowNode]
        | Sequence[Sequence[CellSource]]
        | Sequence[Mapping[ColumnKey, CellSource]]
        | Mapping[ColumnKey, Sequence[CellSource]],
    ) -> TableNode:
        column_order = self._columns

        derived_header: RowNode | None = None
        if isinstance(rows, Mapping):
            row_nodes, derived_header = _rows_from_dict_of_lists(
                rows,
                header_styles=self._header_styles,
                column_order=column_order,
            )
        else:
            tupled_rows = _as_tuple(rows)
            if tupled_rows and all(isinstance(item, Mapping) for item in tupled_rows):
                typed_records = cast(
                    tuple[Mapping[ColumnKey, CellSource], ...], tupled_rows
                )
                row_nodes, derived_header = _rows_from_records(
                    typed_records,
                    header_styles=self._header_styles,
                    column_order=column_order,
                )
            else:
                row_nodes = tuple(_coerce_row(row) for row in tupled_rows)

        return TableNode(rows=row_nodes, styles=self._styles, header=derived_header)


class SheetBuilder:
    def __init__(self, name: str) -> None:
        self._name = name

    def __getitem__(
        self, items: SheetComponent | Sequence[SheetComponent]
    ) -> SheetNode:
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
    def __getitem__(self, sheets: SheetNode | Sequence[SheetNode]) -> Workbook:
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
    style: Sequence[Style] | None = None,
    header_style: Sequence[Style] | None = None,
    column_order: Sequence[ColumnKey] | None = None,
) -> TableBuilder:
    return TableBuilder(styles=style, header_style=header_style, columns=column_order)


def sheet(name: str) -> SheetBuilder:
    return SheetBuilder(name)


def space(rows: int = 1, *, height: float | None = None) -> SpacerNode:
    if rows < 1:
        msg = "Spacer rows must be >= 1"
        raise ValueError(msg)
    return SpacerNode(rows=rows, height=height)


def vstack(
    *items: SheetComponent, gap: int = 0, style: Sequence[Style] | None = None
) -> VerticalStackNode:
    if not items:
        raise ValueError("Vertical stack requires at least one item")
    if gap < 0:
        raise ValueError("Vertical stack gap must be >= 0")
    components = tuple(_ensure_component(item) for item in items)
    return VerticalStackNode(items=components, gap=gap, styles=tuple(style or ()))


def hstack(
    *items: SheetComponent, gap: int = 0, style: Sequence[Style] | None = None
) -> HorizontalStackNode:
    if not items:
        raise ValueError("Horizontal stack requires at least one item")
    if gap < 0:
        raise ValueError("Horizontal stack gap must be >= 0")
    components = tuple(_ensure_component(item) for item in items)
    return HorizontalStackNode(items=components, gap=gap, styles=tuple(style or ()))


def workbook() -> WorkbookBuilder:
    return WorkbookBuilder()
