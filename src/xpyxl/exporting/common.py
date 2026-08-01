"""Shared geometry and output helpers for document renderers."""

from __future__ import annotations

import math
from datetime import date, datetime, time
from pathlib import Path
from typing import BinaryIO, Literal

from ..engines.base import SaveTarget
from ..styles import BorderStyleName
from .model import SheetLayout, WorkbookLayout

__all__ = [
    "ExportFormat",
    "border_line_style",
    "border_width",
    "column_width",
    "format_value",
    "offsets",
    "row_height",
    "select_sheets",
    "sheet_size",
    "validate_scale",
    "write_output",
]

ExportFormat = Literal["pdf", "png"]

COLUMN_WIDTH_PX = 7.0
DEFAULT_COLUMN_WIDTH = 8.0
DEFAULT_ROW_HEIGHT_PT = 16.0
PT_TO_PX = 96.0 / 72.0


def select_sheets(
    workbook: WorkbookLayout,
    sheet: str | int | None,
    format: ExportFormat,
) -> tuple[SheetLayout, ...]:
    """Select sheets with the same rules for every document renderer."""
    if not workbook.sheets:
        raise ValueError("Cannot export a workbook with no sheets")
    if sheet is None:
        if format == "png" and len(workbook.sheets) > 1:
            raise ValueError("PNG export requires sheet= for multi-sheet workbooks")
        return workbook.sheets if format == "pdf" else (next(iter(workbook.sheets)),)
    if isinstance(sheet, int):
        try:
            return (workbook.sheets[sheet],)
        except IndexError as error:
            raise ValueError(f"Sheet index out of range: {sheet}") from error
    for candidate in workbook.sheets:
        if candidate.name == sheet:
            return (candidate,)
    raise ValueError(f"Sheet not found: {sheet}")


def validate_scale(scale: float) -> None:
    if not math.isfinite(scale) or scale <= 0:
        raise ValueError("scale must be a finite number greater than zero")


def sheet_size(sheet: SheetLayout) -> tuple[float, float]:
    """Return the finite sheet canvas in CSS pixels."""
    width = sum(
        column_width(sheet, col) for col in range(1, max(sheet.max_col, 1) + 1)
    )
    height = sum(
        row_height(sheet, row) for row in range(1, max(sheet.max_row, 1) + 1)
    )
    return width, height


def column_width(sheet: SheetLayout, column: int) -> float:
    return sheet.column_widths.get(column, DEFAULT_COLUMN_WIDTH) * COLUMN_WIDTH_PX


def row_height(sheet: SheetLayout, row: int) -> float:
    return sheet.row_heights.get(row, DEFAULT_ROW_HEIGHT_PT) * PT_TO_PX


def border_width(border: BorderStyleName) -> float:
    if border in {"medium", "mediumDashed", "mediumDashDot", "mediumDashDotDot"}:
        return 2.0
    if border in {"thick", "double"}:
        return 3.0
    return 1.0


def border_line_style(border: BorderStyleName) -> str:
    if border in {"dashed", "mediumDashed", "dashDot", "mediumDashDot"}:
        return "dashed"
    if border in {"dotted", "dashDotDot", "mediumDashDotDot"}:
        return "dotted"
    if border == "double":
        return "double"
    return "solid"


def offsets(sizes: list[float]) -> list[float]:
    result = [0.0]
    for size in sizes[:-1]:
        result.append(result[-1] + size)
    return result


def format_value(value: object, number_format: str | None) -> str:
    """Format a normalized cell value consistently across document renderers."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, (date, time)):
        return value.isoformat()
    if number_format and isinstance(value, (int, float)):
        if "%" in number_format:
            return f"{value * 100:.2f}%"
        if "$" in number_format:
            return f"${value:,.2f}"
        if "€" in number_format:
            return f"€{value:,.2f}"
        if "#,##0" in number_format:
            return f"{value:,.2f}" if ".00" in number_format else f"{value:,.0f}"
    return str(value)


def write_output(rendered: bytes, target: SaveTarget | None) -> bytes | None:
    if target is None:
        return rendered
    if isinstance(target, (str, Path)):
        Path(target).write_bytes(rendered)
    else:
        stream: BinaryIO = target
        stream.write(rendered)
        if hasattr(stream, "flush"):
            stream.flush()
    return None
