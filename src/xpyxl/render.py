from __future__ import annotations

import math
from collections.abc import Mapping, Sequence
from typing import Literal, assert_never

from .engines.base import EffectiveStyle, Engine
from .nodes import (
    CellNode,
    ColumnNode,
    HorizontalStackNode,
    RenderableItem,
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


class _Size:
    __slots__ = ("width", "height")

    def __init__(self, width: int, height: int) -> None:
        self.width = width
        self.height = height


_Axis = Literal["vertical", "horizontal"]


class _Placement:
    __slots__ = ("row", "col", "item", "styles", "size", "direction")

    def __init__(
        self,
        row: int,
        col: int,
        item: RenderableItem,
        styles: tuple[Style, ...],
        size: _Size,
        direction: _Axis,
    ) -> None:
        self.row = row
        self.col = col
        self.item = item
        self.styles = styles
        self.size = size
        self.direction = direction


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
    WRAP_LINE_LENGTH = 30
    if not text:
        return 1
    lines = 0
    for raw_line in text.splitlines() or [text]:
        length = max(len(raw_line), 1)
        lines += max(1, math.ceil(length / WRAP_LINE_LENGTH))
    return max(lines, 1)


def _update_dimensions(
    *,
    col_widths: dict[int, float],
    row_heights: dict[int, float],
    column_index: int,
    row_index: int,
    value: object,
    style: EffectiveStyle,
    prefer_height: float | None = None,
) -> None:
    text = "" if value is None else str(value)
    font_scale = style.font_size / DEFAULT_FONT_SIZE if style.font_size else 1.0
    width_hint = max(len(text), 1.0)
    existing_width = col_widths.get(column_index, 0.0)
    if style.row_width is not None:
        width_hint = style.row_width
    elif not style.auto_width:
        width_hint = existing_width if existing_width else 8.0
    elif style.wrap_text:
        width_hint = existing_width or 8.0
    width_hint *= font_scale
    width_hint += 1.0
    col_widths[column_index] = max(existing_width, width_hint)

    if style.row_height is not None:
        base_height = style.row_height
    else:
        base_height = (
            prefer_height if prefer_height is not None else _default_row_height()
        )
        if style.wrap_text:
            base_height *= _estimate_wrap_lines(text)
        base_height *= font_scale
        base_height += 2.0
    row_heights[row_index] = max(row_heights.get(row_index, 0.0), base_height)


def _render_row(
    engine: Engine,
    node: RowNode,
    start_row: int,
    start_col: int,
    col_widths: dict[int, float],
    row_heights: dict[int, float],
    extra_styles: tuple[Style, ...] = (),
) -> None:
    row_index = start_row
    for column_offset, cell_node in enumerate(node.cells, start=1):
        styles = (*extra_styles, *node.styles, *cell_node.styles)
        effective = _resolve(styles)
        column_index = start_col + column_offset - 1
        engine.write_cell(
            row_index, column_index, cell_node.value, effective, DEFAULT_BORDER_COLOR
        )
        _update_dimensions(
            col_widths=col_widths,
            row_heights=row_heights,
            column_index=column_index,
            row_index=row_index,
            value=cell_node.value,
            style=effective,
        )


def _render_column(
    engine: Engine,
    node: ColumnNode,
    start_row: int,
    start_col: int,
    col_widths: dict[int, float],
    row_heights: dict[int, float],
    extra_styles: tuple[Style, ...] = (),
) -> None:
    row_index = start_row
    for cell_node in node.cells:
        styles = (*extra_styles, *node.styles, *cell_node.styles)
        effective = _resolve(styles)
        engine.write_cell(
            row_index, start_col, cell_node.value, effective, DEFAULT_BORDER_COLOR
        )
        _update_dimensions(
            col_widths=col_widths,
            row_heights=row_heights,
            column_index=start_col,
            row_index=row_index,
            value=cell_node.value,
            style=effective,
        )
        row_index += 1


def _render_cell(
    engine: Engine,
    node: CellNode,
    row_index: int,
    column_index: int,
    col_widths: dict[int, float],
    row_heights: dict[int, float],
    extra_styles: tuple[Style, ...] = (),
) -> None:
    effective = _resolve((*extra_styles, *node.styles))
    engine.write_cell(
        row_index, column_index, node.value, effective, DEFAULT_BORDER_COLOR
    )
    _update_dimensions(
        col_widths=col_widths,
        row_heights=row_heights,
        column_index=column_index,
        row_index=row_index,
        value=node.value,
        style=effective,
    )


def _render_table(
    engine: Engine,
    node: TableNode,
    start_row: int,
    start_col: int,
    col_widths: dict[int, float],
    row_heights: dict[int, float],
    extra_styles: tuple[Style, ...] = (),
) -> None:
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

    current_row = start_row

    def render(
        row_node: RowNode,
        *,
        extra: Sequence[Style] = (),
        prefer_height: float | None = None,
        extra_first: bool = False,
    ) -> None:
        for column_offset, cell_node in enumerate(row_node.cells, start=1):
            base_chain = (*extra_styles, *node.styles)
            if extra_first:
                style_chain = (*base_chain, *extra, *row_node.styles, *cell_node.styles)
            else:
                style_chain = (*base_chain, *row_node.styles, *extra, *cell_node.styles)
            if table_border_style:
                style_chain = (*style_chain, table_border_style)
            effective = _resolve(style_chain)
            column_index = start_col + column_offset - 1
            engine.write_cell(
                current_row, column_index, cell_node.value, effective, border_color
            )
            _update_dimensions(
                col_widths=col_widths,
                row_heights=row_heights,
                column_index=column_index,
                row_index=current_row,
                value=cell_node.value,
                style=effective,
                prefer_height=prefer_height,
            )

    if node.header:
        header_extras: list[Style] = [bold, text_center, align_middle]
        if DEFAULT_TABLE_HEADER_BG:
            header_extras.append(Style(fill_color=DEFAULT_TABLE_HEADER_BG))
        if DEFAULT_TABLE_HEADER_TEXT:
            header_extras.append(Style(text_color=DEFAULT_TABLE_HEADER_TEXT))
        render(
            node.header,
            extra=header_extras,
            prefer_height=compact_height,
            extra_first=True,
        )
        current_row += 1

    for idx, row_node in enumerate(node.rows):
        extras: list[Style] = []
        if stripe_style and idx % 2 == 1:
            extras.append(stripe_style)
        render(row_node, extra=extras, prefer_height=compact_height)
        current_row += 1


def _table_size(node: TableNode) -> _Size:
    width = 0
    height = 0
    if node.header:
        width = max(width, len(node.header.cells))
        height += 1
    for row in node.rows:
        width = max(width, len(row.cells))
        height += 1
    return _Size(width=width, height=height)


def _layout_item(
    item: SheetComponent,
    start_row: int,
    start_col: int,
    inherited_styles: tuple[Style, ...] = (),
    direction: _Axis = "vertical",
) -> tuple[list[_Placement], _Size]:
    if isinstance(item, CellNode):
        size = _Size(width=1, height=1)
        return (
            [
                _Placement(
                    row=start_row,
                    col=start_col,
                    item=item,
                    styles=inherited_styles,
                    size=size,
                    direction=direction,
                )
            ],
            size,
        )
    elif isinstance(item, RowNode):
        size = _Size(width=len(item.cells), height=1)
        return (
            [
                _Placement(
                    row=start_row,
                    col=start_col,
                    item=item,
                    styles=inherited_styles,
                    size=size,
                    direction=direction,
                )
            ],
            size,
        )
    elif isinstance(item, ColumnNode):
        size = _Size(width=1, height=len(item.cells))
        return (
            [
                _Placement(
                    row=start_row,
                    col=start_col,
                    item=item,
                    styles=inherited_styles,
                    size=size,
                    direction=direction,
                )
            ],
            size,
        )
    elif isinstance(item, TableNode):
        size = _table_size(item)
        return (
            [
                _Placement(
                    row=start_row,
                    col=start_col,
                    item=item,
                    styles=inherited_styles,
                    size=size,
                    direction=direction,
                )
            ],
            size,
        )
    elif isinstance(item, SpacerNode):
        if direction == "horizontal":
            size = _Size(width=item.rows, height=1)
        else:
            size = _Size(width=1, height=item.rows)
        return (
            [
                _Placement(
                    row=start_row,
                    col=start_col,
                    item=item,
                    styles=inherited_styles,
                    size=size,
                    direction=direction,
                )
            ],
            size,
        )
    elif isinstance(item, VerticalStackNode):
        combined_styles = inherited_styles + item.styles
        placements: list[_Placement] = []  # pyright: ignore[reportRedeclaration]
        row_cursor = start_row
        max_width = 0
        for idx, child in enumerate(item.items):
            child_placements, child_size = _layout_item(
                child, row_cursor, start_col, combined_styles, direction="vertical"
            )
            placements.extend(child_placements)
            row_cursor += child_size.height
            if idx < len(item.items) - 1:
                row_cursor += item.gap
            max_width = max(max_width, child_size.width)
        height = row_cursor - start_row
        return placements, _Size(width=max_width, height=height)
    elif isinstance(item, HorizontalStackNode):
        combined_styles = inherited_styles + item.styles
        placements: list[_Placement] = []
        col_cursor = start_col
        max_height = 0
        for idx, child in enumerate(item.items):
            child_placements, child_size = _layout_item(
                child, start_row, col_cursor, combined_styles, direction="horizontal"
            )
            placements.extend(child_placements)
            col_cursor += child_size.width
            if idx < len(item.items) - 1:
                col_cursor += item.gap
            max_height = max(max_height, child_size.height)
        width = col_cursor - start_col
        return placements, _Size(width=width, height=max_height)
    else:
        assert_never(item)


def _apply_dimensions(
    engine: Engine, col_widths: Mapping[int, float], row_heights: Mapping[int, float]
) -> None:
    for column_index, width in col_widths.items():
        engine.set_column_width(column_index, width)
    for row_index, height in row_heights.items():
        engine.set_row_height(row_index, height)


def render_sheet(engine: Engine, node: SheetNode) -> None:
    """Render a sheet node using the given engine.

    Args:
        engine: The rendering engine to use
        node: The sheet node to render
    """
    engine.create_sheet(node.name)

    col_widths: dict[int, float] = {}
    row_heights: dict[int, float] = {}
    placements: list[_Placement] = []
    row_cursor = 1
    max_col = 0

    for item in node.items:
        item_placements, size = _layout_item(item, row_cursor, 1, direction="vertical")
        placements.extend(item_placements)
        row_cursor += size.height
        max_col = max(max_col, size.width)

    max_row = 0
    for placement in placements:
        max_row = max(max_row, placement.row + placement.size.height - 1)
        max_col = max(max_col, placement.col + placement.size.width - 1)

    if node.background_color:
        normalized = normalize_hex(node.background_color)
        target_max_row = max(max_row, DEFAULT_BACKGROUND_MIN_ROWS)
        target_max_col = max(max_col, DEFAULT_BACKGROUND_MIN_COLS)
        engine.fill_background(normalized, target_max_row, target_max_col)

    for placement in placements:
        target = placement.item
        if isinstance(target, CellNode):
            _render_cell(
                engine,
                target,
                placement.row,
                placement.col,
                col_widths,
                row_heights,
                placement.styles,
            )
        elif isinstance(target, RowNode):
            _render_row(
                engine,
                target,
                placement.row,
                placement.col,
                col_widths,
                row_heights,
                placement.styles,
            )
        elif isinstance(target, ColumnNode):
            _render_column(
                engine,
                target,
                placement.row,
                placement.col,
                col_widths,
                row_heights,
                placement.styles,
            )
        elif isinstance(target, TableNode):
            _render_table(
                engine,
                target,
                placement.row,
                placement.col,
                col_widths,
                row_heights,
                placement.styles,
            )
        elif isinstance(target, SpacerNode):
            if placement.direction == "horizontal":
                continue
            height = (
                target.height if target.height is not None else _default_row_height()
            )
            for offset in range(target.rows):
                row_index = placement.row + offset
                row_heights[row_index] = max(row_heights.get(row_index, 0.0), height)
        else:
            assert_never(target)

    _apply_dimensions(engine, col_widths, row_heights)
