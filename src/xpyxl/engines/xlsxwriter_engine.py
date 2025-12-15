"""XlsxWriter rendering engine implementation."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING, Any, BinaryIO

import xlsxwriter

from ..styles import normalize_hex
from .base import EffectiveStyle, Engine, SaveTarget

if TYPE_CHECKING:
    from xlsxwriter.format import Format
    from xlsxwriter.worksheet import Worksheet

__all__ = ["XlsxWriterEngine"]


# Map openpyxl border style names to xlsxwriter border style indices
# https://xlsxwriter.readthedocs.io/format.html#set_border
BORDER_STYLE_MAP: dict[str, int] = {
    "none": 0,
    "thin": 1,
    "medium": 2,
    "dashed": 3,
    "dotted": 4,
    "thick": 5,
    "double": 6,
    "hair": 7,
    "mediumDashed": 8,
    "dashDot": 9,
    "mediumDashDot": 10,
    "dashDotDot": 11,
    "mediumDashDotDot": 12,
    "slantDashDot": 13,
}


class XlsxWriterEngine(Engine):
    """Rendering engine using xlsxwriter."""

    def __init__(self) -> None:
        super().__init__()
        self._buffer = BytesIO()
        # Convert NaN/INF to Excel errors to avoid xlsxwriter write_number() failures
        workbook_options = {"in_memory": True, "nan_inf_to_errors": True}
        self._workbook = xlsxwriter.Workbook(self._buffer, workbook_options)
        self._current_sheet: Worksheet | None = None
        # Cache format objects to avoid duplicates
        self._format_cache: dict[tuple[Any, ...], Format] = {}
        self._closed = False

    def create_sheet(self, name: str) -> None:
        self._current_sheet = self._workbook.add_worksheet(name)

    def _get_format(
        self, style: EffectiveStyle, border_fallback_color: str
    ) -> "Format":
        """Get or create a format object for the given style."""
        # Create a hashable key from style properties
        cache_key = (
            style.font_name,
            style.font_size,
            style.bold,
            style.italic,
            style.text_color,
            style.fill_color,
            style.horizontal_align,
            style.vertical_align,
            style.indent,
            style.wrap_text,
            style.shrink_to_fit,
            style.number_format,
            style.border,
            style.border_color or border_fallback_color if style.border else None,
            style.border_top,
            style.border_bottom,
            style.border_left,
            style.border_right,
        )

        if cache_key in self._format_cache:
            return self._format_cache[cache_key]

        fmt = self._workbook.add_format()

        # Font settings
        fmt.set_font_name(style.font_name)
        fmt.set_font_size(int(style.font_size))
        if style.bold:
            fmt.set_bold()
        if style.italic:
            fmt.set_italic()

        # Text color (convert from #RRGGBB to xlsxwriter format)
        text_color = normalize_hex(style.text_color)
        fmt.set_font_color(text_color)

        # Fill color
        if style.fill_color:
            fill_color = normalize_hex(style.fill_color)
            # XlsxWriter requires a fill pattern for background colors to render.
            # Use a solid pattern and set both foreground/background for compatibility.
            fmt.set_pattern(1)
            fmt.set_fg_color(fill_color)
            fmt.set_bg_color(fill_color)

        # Alignment
        if style.horizontal_align:
            fmt.set_align(style.horizontal_align)  # type: ignore[arg-type]
        if style.vertical_align:
            # xlsxwriter uses 'vcenter' instead of 'center' for vertical
            valign = style.vertical_align
            if valign == "center":
                valign = "vcenter"
            fmt.set_align(valign)  # type: ignore[arg-type]
        elif style.horizontal_align or style.wrap_text or style.shrink_to_fit:
            # Default vertical alignment to bottom like openpyxl
            fmt.set_align("bottom")

        if style.indent is not None:
            fmt.set_indent(style.indent)
        if style.wrap_text:
            fmt.set_text_wrap()
        if style.shrink_to_fit:
            fmt.set_shrink()

        # Number format
        if style.number_format:
            fmt.set_num_format(style.number_format)

        # Borders
        if style.border and style.border != "none":
            border_style = BORDER_STYLE_MAP.get(style.border, 1)  # default to thin
            border_color = normalize_hex(style.border_color or border_fallback_color)

            explicit = (
                style.border_top
                or style.border_bottom
                or style.border_left
                or style.border_right
            )

            if explicit:
                if style.border_top:
                    fmt.set_top(border_style)
                    fmt.set_top_color(border_color)
                if style.border_bottom:
                    fmt.set_bottom(border_style)
                    fmt.set_bottom_color(border_color)
                if style.border_left:
                    fmt.set_left(border_style)
                    fmt.set_left_color(border_color)
                if style.border_right:
                    fmt.set_right(border_style)
                    fmt.set_right_color(border_color)
            else:
                # Apply to all sides
                fmt.set_border(border_style)
                fmt.set_border_color(border_color)

        self._format_cache[cache_key] = fmt
        return fmt

    def write_cell(
        self,
        row: int,
        col: int,
        value: object,
        style: EffectiveStyle,
        border_fallback_color: str,
    ) -> None:
        if self._current_sheet is None:
            raise RuntimeError("No sheet created. Call create_sheet() first.")

        # Convert from 1-based to 0-based indexing
        row_idx = row - 1
        col_idx = col - 1

        fmt = self._get_format(style, border_fallback_color)

        # xlsxwriter has different write methods for different types
        if value is None:
            self._current_sheet.write_blank(row_idx, col_idx, None, fmt)
        elif isinstance(value, bool):
            self._current_sheet.write_boolean(row_idx, col_idx, value, fmt)
        elif isinstance(value, (int, float)):
            self._current_sheet.write_number(row_idx, col_idx, value, fmt)
        else:
            self._current_sheet.write_string(row_idx, col_idx, str(value), fmt)

    def set_column_width(self, col: int, width: float) -> None:
        if self._current_sheet is None:
            raise RuntimeError("No sheet created. Call create_sheet() first.")

        # Convert from 1-based to 0-based indexing
        col_idx = col - 1
        self._current_sheet.set_column(col_idx, col_idx, max(width, 8.0))

    def set_row_height(self, row: int, height: float) -> None:
        if self._current_sheet is None:
            raise RuntimeError("No sheet created. Call create_sheet() first.")

        # Convert from 1-based to 0-based indexing
        row_idx = row - 1
        self._current_sheet.set_row(row_idx, height)

    def fill_background(
        self,
        color: str,
        max_row: int,
        max_col: int,
    ) -> None:
        if self._current_sheet is None:
            raise RuntimeError("No sheet created. Call create_sheet() first.")

        # Create a format with just the background color
        fill_color = normalize_hex(color)
        bg_fmt = self._workbook.add_format({"bg_color": fill_color})

        # Fill all cells in the range (0-based indexing)
        for row_idx in range(max_row):
            for col_idx in range(max_col):
                self._current_sheet.write_blank(row_idx, col_idx, None, bg_fmt)

    def copy_sheet(
        self, source: SaveTarget | bytes | BinaryIO, sheet_name: str, dest_name: str
    ) -> None:
        msg = "Importing/copying existing sheets is only supported by the openpyxl engine."
        raise NotImplementedError(msg)

    def save(self, target: SaveTarget | None = None) -> bytes | None:
        if not self._closed:
            self._workbook.close()
            self._closed = True

        data = self._buffer.getvalue()
        if target is None:
            return data

        if isinstance(target, (str, Path)):
            Path(target).write_bytes(data)
        else:
            target.write(data)
            target.flush()
        return None
