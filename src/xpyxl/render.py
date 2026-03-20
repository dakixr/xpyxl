from __future__ import annotations

import math
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Literal, assert_never

from .engines.base import EffectiveStyle, Engine
from .nodes import (
    CellNode,
    ColumnNode,
    HorizontalStackNode,
    ImportedSheetNode,
    RowNode,
    SheetComponent,
    SheetNode,
    SpacerNode,
    TableNode,
    VerticalStackNode,
)
from .styles import (
    DEFAULT_BORDER_STYLE_NAME,
    BorderStyleName,
    Style,
    align_middle,
    bold,
    combine_styles,
    normalize_hex,
    text_center,
)

__all__ = ["render_sheet"]


DEFAULT_FONT_NAME = "Calibri"
DEFAULT_FONT_SIZE = 11.0
DEFAULT_MONO_FONT = "Consolas"
DEFAULT_TEXT_COLOR = normalize_hex("#000000")
DEFAULT_BORDER_COLOR = normalize_hex("#000000")
DEFAULT_BORDER_STYLE: BorderStyleName = DEFAULT_BORDER_STYLE_NAME
DEFAULT_ROW_HEIGHT = 16.0
DEFAULT_TABLE_HEADER_BG = None
DEFAULT_TABLE_HEADER_TEXT = None
DEFAULT_TABLE_STRIPE_COLOR = normalize_hex("#F2F4F7")
DEFAULT_TABLE_COMPACT_HEIGHT = 18.0
DEFAULT_BACKGROUND_MIN_ROWS = 200
DEFAULT_BACKGROUND_MIN_COLS = 80

_Axis = Literal["vertical", "horizontal"]


@dataclass(frozen=True)
class _PlacedCell:
    row: int
    col: int
    value: object
    styles: tuple[Style, ...]
    colspan: int
    rowspan: int
    prefer_height: float | None = None


@dataclass(frozen=True)
class _PlacedSpacer:
    row: int
    col: int
    rows: int
    height: float | None
    direction: _Axis


class _GridPlan:
    __slots__ = ("cells", "spacers", "_occupied", "max_row", "max_col")

    def __init__(self) -> None:
        self.cells: list[_PlacedCell] = []
        self.spacers: list[_PlacedSpacer] = []
        self._occupied: set[tuple[int, int]] = set()
        self.max_row = 0
        self.max_col = 0

    def is_clear(self, row: int, col: int, rowspan: int, colspan: int) -> bool:
        for row_idx in range(row, row + rowspan):
            for col_idx in range(col, col + colspan):
                if (row_idx, col_idx) in self._occupied:
                    return False
        return True

    def add_cell(
        self,
        row: int,
        col: int,
        value: object,
        styles: tuple[Style, ...],
        *,
        colspan: int = 1,
        rowspan: int = 1,
        prefer_height: float | None = None,
    ) -> None:
        if not self.is_clear(row, col, rowspan, colspan):
            msg = "Merged cells cannot overlap existing content"
            raise ValueError(msg)

        self.cells.append(
            _PlacedCell(
                row=row,
                col=col,
                value=value,
                styles=styles,
                colspan=colspan,
                rowspan=rowspan,
                prefer_height=prefer_height,
            )
        )
        for row_idx in range(row, row + rowspan):
            for col_idx in range(col, col + colspan):
                self._occupied.add((row_idx, col_idx))
        self.max_row = max(self.max_row, row + rowspan - 1)
        self.max_col = max(self.max_col, col + colspan - 1)

    def add_spacer(
        self,
        row: int,
        col: int,
        *,
        rows: int,
        height: float | None,
        direction: _Axis,
    ) -> None:
        self.spacers.append(
            _PlacedSpacer(
                row=row,
                col=col,
                rows=rows,
                height=height,
                direction=direction,
            )
        )
        if direction == "horizontal":
            self.max_row = max(self.max_row, row)
            self.max_col = max(self.max_col, col + rows - 1)
        else:
            self.max_row = max(self.max_row, row + rows - 1)
            self.max_col = max(self.max_col, col)

    def merge(self, other: _GridPlan, *, row: int, col: int) -> None:
        for placement in other.cells:
            self.add_cell(
                row + placement.row - 1,
                col + placement.col - 1,
                placement.value,
                placement.styles,
                colspan=placement.colspan,
                rowspan=placement.rowspan,
                prefer_height=placement.prefer_height,
            )
        for spacer in other.spacers:
            self.add_spacer(
                row + spacer.row - 1,
                col + spacer.col - 1,
                rows=spacer.rows,
                height=spacer.height,
                direction=spacer.direction,
            )
        if other.max_row > 0:
            self.max_row = max(self.max_row, row + other.max_row - 1)
        if other.max_col > 0:
            self.max_col = max(self.max_col, col + other.max_col - 1)


def _resolve(styles: Sequence[Style]) -> EffectiveStyle:
    base_style = Style(
        font_name=DEFAULT_FONT_NAME,
        font_size=DEFAULT_FONT_SIZE,
        text_color=DEFAULT_TEXT_COLOR,
    )
    merged = combine_styles(styles, base=base_style)

    font_name = merged.font_name or DEFAULT_FONT_NAME
    if merged.mono:
        font_name = DEFAULT_MONO_FONT
    font_size = merged.font_size if merged.font_size is not None else DEFAULT_FONT_SIZE
    if merged.font_size_delta is not None:
        font_size += merged.font_size_delta

    bold_flag = merged.bold if merged.bold is not None else False
    italic_flag = merged.italic if merged.italic is not None else False

    text_color = normalize_hex(merged.text_color or DEFAULT_TEXT_COLOR)
    fill_color = normalize_hex(merged.fill_color) if merged.fill_color else None
    border_color = normalize_hex(merged.border_color) if merged.border_color else None
    shrink_to_fit = merged.shrink_to_fit if merged.shrink_to_fit is not None else False
    auto_width = merged.auto_width if merged.auto_width is not None else True
    row_height = merged.row_height
    row_width = merged.row_width
    border_top = merged.border_top if merged.border_top is not None else False
    border_bottom = merged.border_bottom if merged.border_bottom is not None else False
    border_left = merged.border_left if merged.border_left is not None else False
    border_right = merged.border_right if merged.border_right is not None else False

    return EffectiveStyle(
        font_name=font_name,
        font_size=font_size,
        bold=bold_flag,
        italic=italic_flag,
        text_color=text_color,
        fill_color=fill_color,
        horizontal_align=merged.horizontal_align,
        vertical_align=merged.vertical_align,
        indent=merged.indent,
        wrap_text=merged.wrap_text if merged.wrap_text is not None else False,
        shrink_to_fit=shrink_to_fit,
        auto_width=auto_width,
        row_height=row_height,
        row_width=row_width,
        number_format=merged.number_format,
        border=merged.border,
        border_color=border_color,
        border_top=border_top,
        border_bottom=border_bottom,
        border_left=border_left,
        border_right=border_right,
    )


def _default_row_height() -> float:
    return DEFAULT_ROW_HEIGHT


def _estimate_wrap_lines(text: str) -> int:
    wrap_line_length = 30
    if not text:
        return 1
    lines = 0
    for raw_line in text.splitlines() or [text]:
        length = max(len(raw_line), 1)
        lines += max(1, math.ceil(length / wrap_line_length))
    return max(lines, 1)


def _update_dimensions(
    *,
    col_widths: dict[int, float],
    row_heights: dict[int, float],
    column_index: int,
    row_index: int,
    value: object,
    style: EffectiveStyle,
    colspan: int = 1,
    rowspan: int = 1,
    prefer_height: float | None = None,
) -> None:
    text = "" if value is None else str(value)
    font_scale = style.font_size / DEFAULT_FONT_SIZE if style.font_size else 1.0
    width_hint = max(len(text), 1.0)
    existing_total_width = sum(
        col_widths.get(col_idx, 0.0)
        for col_idx in range(column_index, column_index + colspan)
    )
    if style.row_width is not None:
        total_width = style.row_width
    elif not style.auto_width:
        total_width = existing_total_width if existing_total_width else 8.0 * colspan
    elif style.wrap_text:
        total_width = existing_total_width or 8.0 * colspan
    else:
        total_width = width_hint * font_scale + 1.0
    per_column_width = total_width / colspan
    for col_idx in range(column_index, column_index + colspan):
        col_widths[col_idx] = max(col_widths.get(col_idx, 0.0), per_column_width)

    if style.row_height is not None:
        per_row_height = style.row_height
    else:
        base_height = (
            prefer_height if prefer_height is not None else _default_row_height()
        )
        if style.wrap_text:
            base_height *= _estimate_wrap_lines(text)
        base_height *= font_scale
        base_height += 2.0
        per_row_height = base_height / rowspan
    for row_idx in range(row_index, row_index + rowspan):
        row_heights[row_idx] = max(row_heights.get(row_idx, 0.0), per_row_height)


def _find_next_clear_col(
    plan: _GridPlan,
    *,
    row: int,
    start_col: int,
    rowspan: int,
    colspan: int,
) -> int:
    candidate = max(start_col, 1)
    while not plan.is_clear(row, candidate, rowspan, colspan):
        candidate += 1
    return candidate


def _find_next_clear_row(
    plan: _GridPlan,
    *,
    start_row: int,
    col: int,
    rowspan: int,
    colspan: int,
) -> int:
    candidate = max(start_row, 1)
    while not plan.is_clear(candidate, col, rowspan, colspan):
        candidate += 1
    return candidate


def _place_row_node(
    plan: _GridPlan,
    node: RowNode,
    *,
    row: int,
    extra_styles: tuple[Style, ...],
) -> None:
    cursor = 1
    for cell_node in node.cells:
        column_index = _find_next_clear_col(
            plan,
            row=row,
            start_col=cursor,
            rowspan=cell_node.rowspan,
            colspan=cell_node.colspan,
        )
        plan.add_cell(
            row,
            column_index,
            cell_node.value,
            (*extra_styles, *node.styles, *cell_node.styles),
            colspan=cell_node.colspan,
            rowspan=cell_node.rowspan,
        )
        cursor = column_index + cell_node.colspan


def _place_column_node(
    plan: _GridPlan,
    node: ColumnNode,
    *,
    col: int,
    extra_styles: tuple[Style, ...],
) -> None:
    cursor = 1
    for cell_node in node.cells:
        row_index = _find_next_clear_row(
            plan,
            start_row=cursor,
            col=col,
            rowspan=cell_node.rowspan,
            colspan=cell_node.colspan,
        )
        plan.add_cell(
            row_index,
            col,
            cell_node.value,
            (*extra_styles, *node.styles, *cell_node.styles),
            colspan=cell_node.colspan,
            rowspan=cell_node.rowspan,
        )
        cursor = row_index + 1


def _place_single_cell(
    plan: _GridPlan,
    node: CellNode,
    *,
    row: int,
    extra_styles: tuple[Style, ...],
) -> None:
    column_index = _find_next_clear_col(
        plan,
        row=row,
        start_col=1,
        rowspan=node.rowspan,
        colspan=node.colspan,
    )
    plan.add_cell(
        row,
        column_index,
        node.value,
        (*extra_styles, *node.styles),
        colspan=node.colspan,
        rowspan=node.rowspan,
    )


def _table_has_merged_cells(node: TableNode) -> bool:
    if node.header and any(
        cell.colspan > 1 or cell.rowspan > 1 for cell in node.header.cells
    ):
        return True
    return any(
        cell.colspan > 1 or cell.rowspan > 1
        for row in node.rows
        for cell in row.cells
    )


def _build_table_plan(node: TableNode, extra_styles: tuple[Style, ...]) -> _GridPlan:
    if _table_has_merged_cells(node):
        msg = "Merged cells are not supported inside tables"
        raise ValueError(msg)

    plan = _GridPlan()
    table_style = combine_styles((*extra_styles, *node.styles))
    banded = table_style.table_banded if table_style.table_banded is not None else False
    bordered = (
        table_style.table_bordered if table_style.table_bordered is not None else True
    )
    compact = (
        table_style.table_compact if table_style.table_compact is not None else False
    )
    border_color = (
        table_style.border_color
        if table_style.border_color is not None
        else DEFAULT_BORDER_COLOR
    )
    border_style = (
        table_style.border if table_style.border is not None else DEFAULT_BORDER_STYLE
    )

    table_border_style = (
        Style(border=border_style, border_color=border_color) if bordered else None
    )
    stripe_style = Style(fill_color=DEFAULT_TABLE_STRIPE_COLOR) if banded else None
    compact_height = DEFAULT_TABLE_COMPACT_HEIGHT if compact else None

    current_row = 1

    def add_row(
        row_node: RowNode,
        *,
        extras: Sequence[Style] = (),
        prefer_height: float | None = None,
        extras_first: bool = False,
    ) -> None:
        for column_offset, cell_node in enumerate(row_node.cells, start=1):
            base_chain = (*extra_styles, *node.styles)
            if extras_first:
                style_chain = (*base_chain, *extras, *row_node.styles, *cell_node.styles)
            else:
                style_chain = (*base_chain, *row_node.styles, *extras, *cell_node.styles)
            if table_border_style:
                style_chain = (*style_chain, table_border_style)
            plan.add_cell(
                current_row,
                column_offset,
                cell_node.value,
                style_chain,
                prefer_height=prefer_height,
            )

    if node.header:
        header_extras: list[Style] = [bold, text_center, align_middle]
        if DEFAULT_TABLE_HEADER_BG:
            header_extras.append(Style(fill_color=DEFAULT_TABLE_HEADER_BG))
        if DEFAULT_TABLE_HEADER_TEXT:
            header_extras.append(Style(text_color=DEFAULT_TABLE_HEADER_TEXT))
        add_row(
            node.header,
            extras=header_extras,
            prefer_height=compact_height,
            extras_first=True,
        )
        current_row += 1

    for idx, row_node in enumerate(node.rows):
        extras: list[Style] = []
        if stripe_style and idx % 2 == 1:
            extras.append(stripe_style)
        add_row(row_node, extras=extras, prefer_height=compact_height)
        current_row += 1

    return plan


def _build_horizontal_plan(
    items: Sequence[SheetComponent],
    *,
    extra_styles: tuple[Style, ...],
    gap: int,
) -> _GridPlan:
    plan = _GridPlan()
    col_cursor = 1
    for idx, child in enumerate(items):
        if isinstance(child, SpacerNode):
            plan.add_spacer(
                1,
                col_cursor,
                rows=child.rows,
                height=child.height,
                direction="horizontal",
            )
            col_cursor += child.rows
        else:
            child_plan = _build_item_plan(child, extra_styles=extra_styles)
            plan.merge(child_plan, row=1, col=col_cursor)
            col_cursor += _logical_width(child)
        if idx < len(items) - 1:
            col_cursor += gap
    return plan


def _build_vertical_plan(
    items: Sequence[SheetComponent],
    *,
    extra_styles: tuple[Style, ...],
    gap: int,
) -> _GridPlan:
    plan = _GridPlan()
    row_cursor = 1
    for idx, child in enumerate(items):
        if isinstance(child, CellNode):
            _place_single_cell(plan, child, row=row_cursor, extra_styles=extra_styles)
            row_cursor += 1
        elif isinstance(child, RowNode):
            _place_row_node(plan, child, row=row_cursor, extra_styles=extra_styles)
            row_cursor += 1
        elif isinstance(child, SpacerNode):
            plan.add_spacer(
                row_cursor,
                1,
                rows=child.rows,
                height=child.height,
                direction="vertical",
            )
            row_cursor += child.rows
        else:
            child_plan = _build_item_plan(child, extra_styles=extra_styles)
            plan.merge(child_plan, row=row_cursor, col=1)
            row_cursor += _logical_height(child)
        if idx < len(items) - 1:
            row_cursor += gap
    return plan


def _logical_width(item: SheetComponent) -> int:
    if isinstance(item, CellNode):
        return 1
    if isinstance(item, RowNode):
        return len(item.cells)
    if isinstance(item, ColumnNode):
        return 1
    if isinstance(item, TableNode):
        width = 0
        if item.header:
            width = max(width, len(item.header.cells))
        for row in item.rows:
            width = max(width, len(row.cells))
        return width
    if isinstance(item, SpacerNode):
        return 1
    if isinstance(item, VerticalStackNode):
        return max(_logical_width(child) for child in item.items)
    if isinstance(item, HorizontalStackNode):
        total = sum(_logical_width(child) for child in item.items)
        total += item.gap * (len(item.items) - 1)
        return total
    assert_never(item)


def _logical_height(item: SheetComponent) -> int:
    if isinstance(item, CellNode):
        return 1
    if isinstance(item, RowNode):
        return 1
    if isinstance(item, ColumnNode):
        return len(item.cells)
    if isinstance(item, TableNode):
        return len(item.rows) + (1 if item.header else 0)
    if isinstance(item, SpacerNode):
        return item.rows
    if isinstance(item, VerticalStackNode):
        total = sum(_logical_height(child) for child in item.items)
        total += item.gap * (len(item.items) - 1)
        return total
    if isinstance(item, HorizontalStackNode):
        return max(_logical_height(child) for child in item.items)
    assert_never(item)


def _build_item_plan(
    item: SheetComponent,
    *,
    extra_styles: tuple[Style, ...] = (),
) -> _GridPlan:
    if isinstance(item, CellNode):
        plan = _GridPlan()
        plan.add_cell(
            1,
            1,
            item.value,
            (*extra_styles, *item.styles),
            colspan=item.colspan,
            rowspan=item.rowspan,
        )
        return plan
    if isinstance(item, RowNode):
        plan = _GridPlan()
        _place_row_node(plan, item, row=1, extra_styles=extra_styles)
        return plan
    if isinstance(item, ColumnNode):
        plan = _GridPlan()
        _place_column_node(plan, item, col=1, extra_styles=extra_styles)
        return plan
    if isinstance(item, TableNode):
        return _build_table_plan(item, extra_styles)
    if isinstance(item, SpacerNode):
        plan = _GridPlan()
        plan.add_spacer(
            1,
            1,
            rows=item.rows,
            height=item.height,
            direction="vertical",
        )
        return plan
    if isinstance(item, VerticalStackNode):
        return _build_vertical_plan(
            item.items,
            extra_styles=extra_styles + item.styles,
            gap=item.gap,
        )
    if isinstance(item, HorizontalStackNode):
        return _build_horizontal_plan(
            item.items,
            extra_styles=extra_styles + item.styles,
            gap=item.gap,
        )
    assert_never(item)


def _apply_dimensions(
    engine: Engine,
    col_widths: Mapping[int, float],
    row_heights: Mapping[int, float],
) -> None:
    for column_index, width in col_widths.items():
        engine.set_column_width(column_index, width)
    for row_index, height in row_heights.items():
        engine.set_row_height(row_index, height)


def render_sheet(engine: Engine, node: SheetNode | ImportedSheetNode) -> None:
    """Render a sheet node using the given engine."""
    if isinstance(node, ImportedSheetNode):
        engine.copy_sheet(node.source, node.source_sheet, node.name)
        return

    engine.create_sheet(node.name)

    col_widths: dict[int, float] = {}
    row_heights: dict[int, float] = {}
    plan = _build_vertical_plan(node.items, extra_styles=(), gap=0)

    max_row = plan.max_row
    max_col = plan.max_col

    if node.background_color:
        normalized = normalize_hex(node.background_color)
        target_max_row = max(max_row, DEFAULT_BACKGROUND_MIN_ROWS)
        target_max_col = max(max_col, DEFAULT_BACKGROUND_MIN_COLS)
        engine.fill_background(normalized, target_max_row, target_max_col)

    for placement in sorted(plan.cells, key=lambda cell: (cell.row, cell.col)):
        effective = _resolve(placement.styles)
        if placement.colspan == 1 and placement.rowspan == 1:
            engine.write_cell(
                placement.row,
                placement.col,
                placement.value,
                effective,
                DEFAULT_BORDER_COLOR,
            )
        else:
            engine.write_merged_cell(
                placement.row,
                placement.col,
                placement.rowspan,
                placement.colspan,
                placement.value,
                effective,
                DEFAULT_BORDER_COLOR,
            )
        _update_dimensions(
            col_widths=col_widths,
            row_heights=row_heights,
            column_index=placement.col,
            row_index=placement.row,
            value=placement.value,
            style=effective,
            colspan=placement.colspan,
            rowspan=placement.rowspan,
            prefer_height=placement.prefer_height,
        )

    for spacer in plan.spacers:
        if spacer.direction == "horizontal":
            continue
        height = spacer.height if spacer.height is not None else _default_row_height()
        for offset in range(spacer.rows):
            row_index = spacer.row + offset
            row_heights[row_index] = max(row_heights.get(row_index, 0.0), height)

    _apply_dimensions(engine, col_widths, row_heights)
