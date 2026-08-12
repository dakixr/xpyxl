from __future__ import annotations

from collections.abc import Iterator, Mapping, Sequence
from pathlib import Path
from typing import BinaryIO, TypeAlias, cast, overload

from ._workbook import Workbook
from .nodes import (
    CellNode,
    CellValue,
    ColumnNode,
    HorizontalStackNode,
    ImportedSheetNode,
    RowNode,
    SheetComponent,
    SheetItem,
    SheetNode,
    SpacerNode,
    TableNode,
    VerticalStackNode,
    WorkbookNode,
)
from .styles import Style, normalize_hex

__all__ = [
    "cell",
    "col",
    "import_sheet",
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

_INVALID_SHEET_NAME_CHARS = frozenset("[]:*?/\\")
_MAX_SHEET_NAME_LENGTH = 31


def _validate_sheet_name(name: str) -> None:
    """Reject sheet names Excel cannot represent (backend-agnostic errors)."""
    if not name:
        msg = "Sheet name cannot be empty"
        raise ValueError(msg)
    if len(name) > _MAX_SHEET_NAME_LENGTH:
        msg = f"Sheet name must be <= 31 characters, got {len(name)}: {name!r}"
        raise ValueError(msg)
    invalid = sorted(set(name) & _INVALID_SHEET_NAME_CHARS)
    if invalid:
        msg = f"Sheet name {name!r} contains invalid characters: {' '.join(invalid)}"
        raise ValueError(msg)
    if name.startswith("'") or name.endswith("'"):
        msg = f"Sheet name cannot begin or end with an apostrophe: {name!r}"
        raise ValueError(msg)


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


class _RecordRows(Sequence[RowNode]):
    """Create row nodes on demand from a re-iterable record sequence."""

    __slots__ = ("_columns", "_records")

    def __init__(
        self,
        records: Sequence[Mapping[ColumnKey, CellSource]],
        columns: Sequence[ColumnKey],
    ) -> None:
        self._records = records
        self._columns = tuple(columns)

    def __len__(self) -> int:
        return len(self._records)

    @overload
    def __getitem__(self, index: int) -> RowNode: ...

    @overload
    def __getitem__(self, index: slice) -> Sequence[RowNode]: ...

    def __getitem__(self, index: int | slice) -> RowNode | Sequence[RowNode]:
        if isinstance(index, slice):
            return tuple(self[idx] for idx in range(*index.indices(len(self))))
        record = self._records[index]
        return RowNode(
            cells=tuple(_ensure_cell(record.get(column)) for column in self._columns)
        )

    def __iter__(self) -> Iterator[RowNode]:
        for record in self._records:
            yield RowNode(
                cells=tuple(
                    _ensure_cell(record.get(column)) for column in self._columns
                )
            )


class _SheetItems(Sequence[SheetItem]):
    """Validate repeatable sheet components lazily as they are consumed."""

    __slots__ = ("_items",)

    def __init__(self, items: Sequence[object]) -> None:
        self._items = items

    def __len__(self) -> int:
        return len(self._items)

    @overload
    def __getitem__(self, index: int) -> SheetItem: ...

    @overload
    def __getitem__(self, index: slice) -> Sequence[SheetItem]: ...

    def __getitem__(self, index: int | slice) -> SheetItem | Sequence[SheetItem]:
        if isinstance(index, slice):
            return tuple(self[item] for item in range(*index.indices(len(self))))
        return self._validate(self._items[index])

    def __iter__(self) -> Iterator[SheetItem]:
        for item in self._items:
            yield self._validate(item)

    @staticmethod
    def _validate(item: object) -> SheetItem:
        if isinstance(item, Node):
            return item
        msg = (
            "Sheets accept rows, columns, tables, spacers, or layout stacks. "
            "Call the builder before nesting."
        )
        raise TypeError(msg)


def _record_rows(
    rows: object,
    *,
    header_styles: Sequence[Style],
    column_order: Sequence[ColumnKey] | None,
) -> tuple[Sequence[RowNode], RowNode | None] | None:
    """Map repeatable records to row nodes without duplicating every cell."""
    if not isinstance(rows, Sequence) or isinstance(rows, (str, bytes, bytearray)):
        return None

    columns: list[ColumnKey] = list(column_order or ())
    seen = set(columns)
    for record in rows:
        if not isinstance(record, Mapping):
            return None
        for key in record:
            if key not in seen:
                seen.add(key)
                columns.append(key)

    typed_rows = cast(Sequence[Mapping[ColumnKey, CellSource]], rows)
    header = _coerce_row(columns, extra_styles=header_styles) if columns else None
    return _RecordRows(typed_rows, columns), header


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
    def __init__(
        self,
        *,
        styles: Sequence[Style] | None = None,
        colspan: int = 1,
        rowspan: int = 1,
    ) -> None:
        super().__init__(styles=styles)
        if colspan < 1:
            msg = "Cell colspan must be >= 1"
            raise ValueError(msg)
        if rowspan < 1:
            msg = "Cell rowspan must be >= 1"
            raise ValueError(msg)
        self._colspan = colspan
        self._rowspan = rowspan

    def __getitem__(self, value: CellSource) -> CellNode:
        return CellNode(
            value=value,
            styles=self._styles,
            colspan=self._colspan,
            rowspan=self._rowspan,
        )


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
            record_rows = _record_rows(
                rows,
                header_styles=self._header_styles,
                column_order=column_order,
            )
            if record_rows is not None:
                row_nodes, derived_header = record_rows
                return TableNode(
                    rows=row_nodes,
                    styles=self._styles,
                    header=derived_header,
                )
            tupled_rows = _as_tuple(rows)
            row_nodes = tuple(_coerce_row(row) for row in tupled_rows)

        return TableNode(rows=row_nodes, styles=self._styles, header=derived_header)


class SheetBuilder:
    def __init__(
        self,
        name: str,
        *,
        background_color: str | None = None,
        show_gridlines: bool = True,
    ) -> None:
        _validate_sheet_name(name)
        self._name = name
        self._background_color = (
            normalize_hex(background_color) if background_color else None
        )
        self._show_gridlines = show_gridlines

    def __getitem__(
        self, items: SheetComponent | Sequence[SheetComponent]
    ) -> SheetNode:
        if isinstance(items, Sequence) and not isinstance(items, (Node, list, tuple)):
            return SheetNode(
                name=self._name,
                items=_SheetItems(cast(Sequence[object], items)),
                background_color=self._background_color,
                show_gridlines=self._show_gridlines,
            )
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
        return SheetNode(
            name=self._name,
            items=tuple(entries),
            background_color=self._background_color,
            show_gridlines=self._show_gridlines,
        )


class WorkbookBuilder:
    def __getitem__(
        self,
        sheets: SheetNode | ImportedSheetNode | Sequence[SheetNode | ImportedSheetNode],
    ) -> Workbook:
        sheet_nodes: list[SheetNode | ImportedSheetNode] = []
        for item in _as_tuple(sheets):
            if isinstance(item, (SheetNode, ImportedSheetNode)):
                sheet_nodes.append(item)
            else:
                raise TypeError(
                    "Workbooks accept sheet builders that have been indexed"
                )
        seen_names: set[str] = set()
        for sheet_node in sheet_nodes:
            _validate_sheet_name(sheet_node.name)
            key = sheet_node.name.casefold()
            if key in seen_names:
                msg = (
                    f"Duplicate sheet name {sheet_node.name!r}: sheet names must "
                    "be unique (case-insensitive) across the workbook"
                )
                raise ValueError(msg)
            seen_names.add(key)
        node = WorkbookNode(sheets=tuple(sheet_nodes))
        return Workbook(node)


def cell(
    *,
    style: Sequence[Style] | None = None,
    colspan: int = 1,
    rowspan: int = 1,
) -> CellBuilder:
    return CellBuilder(styles=style, colspan=colspan, rowspan=rowspan)


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


def sheet(
    name: str,
    *,
    background_color: str | None = None,
    show_gridlines: bool = True,
) -> SheetBuilder:
    return SheetBuilder(
        name,
        background_color=background_color,
        show_gridlines=show_gridlines,
    )


def import_sheet(
    source: str | Path | bytes | BinaryIO,
    sheet_name: str,
    *,
    name: str | None = None,
    show_gridlines: bool | None = None,
) -> ImportedSheetNode:
    """Import an existing sheet from a workbook without translating it."""

    dest_name = name or sheet_name
    _validate_sheet_name(dest_name)
    return ImportedSheetNode(
        name=dest_name,
        source=source,
        source_sheet=sheet_name,
        show_gridlines=show_gridlines,
    )


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
