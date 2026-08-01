"""In-process PDF and PNG rendering with ReportLab and Pillow."""

from __future__ import annotations

import importlib.util
import math
from collections.abc import Callable
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING

from ..engines.base import EffectiveStyle, SaveTarget
from .common import (
    ExportFormat,
    border_line_style,
    border_width,
    column_width,
    format_value,
    offsets,
    row_height,
    select_sheets,
    sheet_size,
    validate_scale,
    write_output,
)
from .model import CellLayout, SheetLayout, WorkbookLayout

if TYPE_CHECKING:
    from PIL.Image import Image
    from PIL.ImageDraw import ImageDraw
    from PIL.ImageFont import FreeTypeFont
    from reportlab.pdfgen.canvas import Canvas

__all__ = ["ReportLabRenderer"]

_PX_TO_PT = 72.0 / 96.0
_PT_TO_PX = 96.0 / 72.0
_HORIZONTAL_PADDING_PX = 6.0
_VERTICAL_PADDING_PX = 2.0
_GRIDLINE_COLOR = "#D9D9D9"


@dataclass(frozen=True)
class _CellBox:
    cell: CellLayout | None
    left: float
    top: float
    width: float
    height: float


@dataclass(frozen=True)
class _LineSpec:
    width: float
    style: str
    color: str
    double: bool = False


class ReportLabRenderer:
    """Render PDF and PNG files without launching an external process."""

    def __init__(self) -> None:
        _require_reportlab()

    def render(
        self,
        workbook: WorkbookLayout,
        target: SaveTarget | None = None,
        *,
        format: ExportFormat = "pdf",
        sheet: str | int | None = None,
        scale: float = 1.0,
    ) -> bytes | None:
        if format not in ("pdf", "png"):
            raise ValueError(f"Unsupported ReportLab export format: {format}")
        validate_scale(scale)

        sheets = select_sheets(workbook, sheet, format)
        if format == "pdf":
            rendered = _render_pdf(sheets)
        else:
            rendered = _render_png(next(iter(sheets)), scale)
        return write_output(rendered, target)


def _require_reportlab() -> None:
    if importlib.util.find_spec("reportlab") is None:
        raise ModuleNotFoundError(
            "The ReportLab renderer requires the optional dependency. "
            "Install it with: pip install 'xpyxl[reportlab]'"
        )


def _render_pdf(sheets: tuple[SheetLayout, ...]) -> bytes:
    from reportlab.pdfgen.canvas import Canvas

    buffer = BytesIO()
    first_width, first_height = sheet_size(next(iter(sheets)))
    canvas = Canvas(
        buffer,
        pagesize=(first_width * _PX_TO_PT, first_height * _PX_TO_PT),
        pageCompression=1,
    )
    canvas.setTitle("xpyxl report")
    for sheet in sheets:
        width, height = sheet_size(sheet)
        canvas.setPageSize((width * _PX_TO_PT, height * _PX_TO_PT))
        _draw_pdf_sheet(canvas, sheet, width, height)
        canvas.showPage()
    canvas.save()
    return buffer.getvalue()


def _draw_pdf_sheet(
    canvas: Canvas,
    sheet: SheetLayout,
    width_px: float,
    height_px: float,
) -> None:
    page_width = width_px * _PX_TO_PT
    page_height = height_px * _PX_TO_PT
    canvas.setFillColor(_pdf_color(sheet.background_color or "#FFFFFF"))
    canvas.rect(0, 0, page_width, page_height, stroke=0, fill=1)
    boxes = _cell_boxes(sheet)

    for box in boxes:
        if box.cell is None or box.cell.style.fill_color is None:
            continue
        left, bottom, width, height = _pdf_box(box, page_height)
        canvas.setFillColor(_pdf_color(box.cell.style.fill_color))
        canvas.rect(left, bottom, width, height, stroke=0, fill=1)

    for box in boxes:
        for side, line in _line_specs(box.cell, sheet.show_gridlines).items():
            _draw_pdf_line(canvas, box, side, line, page_height)

    for box in boxes:
        if box.cell is not None:
            _draw_pdf_text(canvas, box, page_height)


def _draw_pdf_line(
    canvas: Canvas,
    box: _CellBox,
    side: str,
    line: _LineSpec,
    page_height: float,
) -> None:
    left, bottom, width, height = _pdf_box(box, page_height)
    if side == "top":
        start = (left, bottom + height)
        end = (left + width, bottom + height)
    elif side == "bottom":
        start = (left, bottom)
        end = (left + width, bottom)
    elif side == "left":
        start = (left, bottom)
        end = (left, bottom + height)
    else:
        start = (left + width, bottom)
        end = (left + width, bottom + height)

    canvas.saveState()
    canvas.setStrokeColor(_pdf_color(line.color))
    canvas.setLineWidth(line.width * _PX_TO_PT)
    _set_pdf_dash(canvas, line.style, line.width * _PX_TO_PT)
    canvas.line(*start, *end)
    if line.double:
        inset = max(line.width * _PX_TO_PT, 1.5)
        if side == "top":
            canvas.line(start[0], start[1] - inset, end[0], end[1] - inset)
        elif side == "bottom":
            canvas.line(start[0], start[1] + inset, end[0], end[1] + inset)
        elif side == "left":
            canvas.line(start[0] + inset, start[1], end[0] + inset, end[1])
        else:
            canvas.line(start[0] - inset, start[1], end[0] - inset, end[1])
    canvas.restoreState()


def _draw_pdf_text(canvas: Canvas, box: _CellBox, page_height: float) -> None:
    from reportlab.pdfbase import pdfmetrics

    assert box.cell is not None
    cell = box.cell
    text = format_value(cell.value, cell.style.number_format)
    if not text:
        return
    left, bottom, width, height = _pdf_box(box, page_height)
    font_name = _pdf_font(cell.style)
    font_size = cell.style.font_size
    horizontal_padding = (
        _HORIZONTAL_PADDING_PX + (cell.style.indent or 0) * 8
    ) * _PX_TO_PT
    vertical_padding = _VERTICAL_PADDING_PX * _PX_TO_PT
    available_width = max(width - horizontal_padding - _HORIZONTAL_PADDING_PX * _PX_TO_PT, 0)
    lines = _wrap_text(
        text,
        available_width,
        cell.style.wrap_text,
        lambda value: pdfmetrics.stringWidth(value, font_name, font_size),
    )
    line_height = font_size * 1.2
    block_height = len(lines) * line_height
    ascent = pdfmetrics.getAscent(font_name) * font_size / 1000
    top = bottom + height
    if cell.style.vertical_align == "bottom":
        first_baseline = bottom + vertical_padding + block_height - line_height + ascent
    elif cell.style.vertical_align == "center":
        block_bottom = bottom + (height - block_height) / 2
        first_baseline = block_bottom + block_height - line_height + ascent
    else:
        first_baseline = top - vertical_padding - ascent

    path = canvas.beginPath()
    path.rect(left, bottom, width, height)
    canvas.saveState()
    canvas.clipPath(path, stroke=0, fill=0)
    canvas.setFont(font_name, font_size)
    canvas.setFillColor(_pdf_color(cell.style.text_color))
    for index, line in enumerate(lines):
        line_width = pdfmetrics.stringWidth(line, font_name, font_size)
        x = _aligned_x(
            cell.style.horizontal_align,
            left,
            width,
            horizontal_padding,
            line_width,
        )
        canvas.drawString(x, first_baseline - index * line_height, line)
    canvas.restoreState()


def _render_png(sheet: SheetLayout, scale: float) -> bytes:
    from PIL import Image, ImageDraw

    width, height = sheet_size(sheet)
    pixel_width = max(1, math.ceil(width * scale))
    pixel_height = max(1, math.ceil(height * scale))
    image = Image.new("RGB", (pixel_width, pixel_height), _rgb(sheet.background_color))
    draw = ImageDraw.Draw(image)
    boxes = _cell_boxes(sheet)

    for box in boxes:
        if box.cell is None or box.cell.style.fill_color is None:
            continue
        draw.rectangle(_png_box(box, scale), fill=_rgb(box.cell.style.fill_color))

    for box in boxes:
        for side, line in _line_specs(box.cell, sheet.show_gridlines).items():
            _draw_png_line(draw, box, side, line, scale)

    for box in boxes:
        if box.cell is not None:
            _draw_png_text(image, box, scale)

    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def _draw_png_line(
    draw: ImageDraw,
    box: _CellBox,
    side: str,
    line: _LineSpec,
    scale: float,
) -> None:
    left, top, right, bottom = _png_box(box, scale)
    if side == "top":
        start, end = (left, top), (right, top)
    elif side == "bottom":
        start, end = (left, bottom), (right, bottom)
    elif side == "left":
        start, end = (left, top), (left, bottom)
    else:
        start, end = (right, top), (right, bottom)
    stroke_width = max(1, round(line.width * scale))
    _draw_patterned_png_line(draw, start, end, _rgb(line.color), stroke_width, line.style)
    if line.double:
        inset = max(stroke_width, round(1.5 * scale))
        if side == "top":
            start, end = (start[0], start[1] + inset), (end[0], end[1] + inset)
        elif side == "bottom":
            start, end = (start[0], start[1] - inset), (end[0], end[1] - inset)
        elif side == "left":
            start, end = (start[0] + inset, start[1]), (end[0] + inset, end[1])
        else:
            start, end = (start[0] - inset, start[1]), (end[0] - inset, end[1])
        draw.line((start, end), fill=_rgb(line.color), width=stroke_width)


def _draw_patterned_png_line(
    draw: ImageDraw,
    start: tuple[int, int],
    end: tuple[int, int],
    color: tuple[int, int, int],
    width: int,
    style: str,
) -> None:
    if style == "solid":
        draw.line((start, end), fill=color, width=width)
        return
    length = abs(end[0] - start[0]) + abs(end[1] - start[1])
    dash = max(width * (1 if style == "dotted" else 4), 1)
    gap = max(width * 2, 1)
    horizontal = start[1] == end[1]
    direction = 1 if (end[0] + end[1]) >= (start[0] + start[1]) else -1
    position = 0
    while position < length:
        segment_end = min(position + dash, length)
        if horizontal:
            segment = (
                (start[0] + direction * position, start[1]),
                (start[0] + direction * segment_end, start[1]),
            )
        else:
            segment = (
                (start[0], start[1] + direction * position),
                (start[0], start[1] + direction * segment_end),
            )
        draw.line(segment, fill=color, width=width)
        position += dash + gap


def _draw_png_text(image: Image, box: _CellBox, scale: float) -> None:
    from PIL import Image, ImageDraw

    assert box.cell is not None
    cell = box.cell
    text = format_value(cell.value, cell.style.number_format)
    if not text:
        return
    pixel_width = max(1, math.ceil(box.width * scale))
    pixel_height = max(1, math.ceil(box.height * scale))
    layer = Image.new("RGBA", (pixel_width, pixel_height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    font = _pillow_font(cell.style, scale)
    horizontal_padding = round(
        (_HORIZONTAL_PADDING_PX + (cell.style.indent or 0) * 8) * scale
    )
    right_padding = round(_HORIZONTAL_PADDING_PX * scale)
    vertical_padding = round(_VERTICAL_PADDING_PX * scale)
    available_width = max(pixel_width - horizontal_padding - right_padding, 0)
    lines = _wrap_text(
        text,
        float(available_width),
        cell.style.wrap_text,
        lambda value: draw.textlength(value, font=font),
    )
    line_height = max(1, round(cell.style.font_size * _PT_TO_PX * 1.2 * scale))
    block_height = len(lines) * line_height
    if cell.style.vertical_align == "bottom":
        top = pixel_height - vertical_padding - block_height
    elif cell.style.vertical_align == "center":
        top = (pixel_height - block_height) / 2
    else:
        top = vertical_padding
    for index, line in enumerate(lines):
        line_width = draw.textlength(line, font=font)
        x = _aligned_x(
            cell.style.horizontal_align,
            0,
            pixel_width,
            horizontal_padding,
            line_width,
            right_padding=right_padding,
        )
        draw.text(
            (x, top + index * line_height),
            line,
            font=font,
            fill=_rgb(cell.style.text_color),
            anchor="lt",
        )
    image.paste(
        layer,
        (round(box.left * scale), round(box.top * scale)),
        layer,
    )


def _cell_boxes(sheet: SheetLayout) -> list[_CellBox]:
    max_row = max(sheet.max_row, 1)
    max_col = max(sheet.max_col, 1)
    column_widths = [column_width(sheet, col) for col in range(1, max_col + 1)]
    row_heights = [row_height(sheet, row) for row in range(1, max_row + 1)]
    column_offsets = offsets(column_widths)
    row_offsets = offsets(row_heights)
    boxes: list[_CellBox] = []
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            if (row, col) in sheet.covered_cells:
                continue
            cell = sheet.cells.get((row, col))
            rowspan = cell.rowspan if cell else 1
            colspan = cell.colspan if cell else 1
            boxes.append(
                _CellBox(
                    cell=cell,
                    left=column_offsets[col - 1],
                    top=row_offsets[row - 1],
                    width=sum(column_widths[col - 1 : col - 1 + colspan]),
                    height=sum(row_heights[row - 1 : row - 1 + rowspan]),
                )
            )
    return boxes


def _line_specs(
    cell: CellLayout | None,
    show_gridlines: bool,
) -> dict[str, _LineSpec]:
    lines: dict[str, _LineSpec] = {}
    if show_gridlines:
        gridline = _LineSpec(1.0, "solid", _GRIDLINE_COLOR)
        lines.update({"right": gridline, "bottom": gridline})
    if cell is None:
        return lines
    style = cell.style
    if style.border == "none":
        return {}
    if not style.border and not any(
        (style.border_top, style.border_bottom, style.border_left, style.border_right)
    ):
        return lines

    border = style.border or "thin"
    custom = _LineSpec(
        width=border_width(border),
        style="solid" if border == "double" else border_line_style(border),
        color=style.border_color or cell.border_fallback_color,
        double=border == "double",
    )
    sides = {
        "top": style.border_top,
        "bottom": style.border_bottom,
        "left": style.border_left,
        "right": style.border_right,
    }
    if any(sides.values()):
        lines.update({side: custom for side, enabled in sides.items() if enabled})
    else:
        lines.update({side: custom for side in sides})
    return lines


def _wrap_text(
    text: str,
    available_width: float,
    wrap: bool,
    measure: Callable[[str], float],
) -> list[str]:
    if not wrap:
        return [text.replace("\n", " ")]
    lines: list[str] = []
    for paragraph in text.splitlines() or [""]:
        if not paragraph:
            lines.append("")
            continue
        current: list[str] = []
        current_width = 0.0
        for character in paragraph:
            character_width = measure(character)
            if current and current_width + character_width > available_width:
                lines.append("".join(current).rstrip())
                if character.isspace():
                    current = []
                    current_width = 0.0
                else:
                    current = [character]
                    current_width = character_width
            else:
                current.append(character)
                current_width += character_width
        lines.append("".join(current))
    return lines or [""]


def _aligned_x(
    alignment: str | None,
    left: float,
    width: float,
    left_padding: float,
    line_width: float,
    *,
    right_padding: float | None = None,
) -> float:
    right_padding = left_padding if right_padding is None else right_padding
    if line_width > width - left_padding - right_padding:
        # Chromium anchors overflowing, non-wrapped cell text at the leading
        # content edge before clipping it to the cell rectangle.
        return left + left_padding
    if alignment == "center":
        return left + (width - line_width) / 2
    if alignment == "right":
        return left + width - right_padding - line_width
    return left + left_padding


def _pdf_box(box: _CellBox, page_height: float) -> tuple[float, float, float, float]:
    width = box.width * _PX_TO_PT
    height = box.height * _PX_TO_PT
    return (
        box.left * _PX_TO_PT,
        page_height - (box.top + box.height) * _PX_TO_PT,
        width,
        height,
    )


def _png_box(box: _CellBox, scale: float) -> tuple[int, int, int, int]:
    return (
        round(box.left * scale),
        round(box.top * scale),
        round((box.left + box.width) * scale) - 1,
        round((box.top + box.height) * scale) - 1,
    )


def _pdf_color(value: str):
    from reportlab.lib.colors import HexColor

    return HexColor(value)


def _rgb(value: str | None) -> tuple[int, int, int]:
    normalized = (value or "#FFFFFF").lstrip("#")
    return (
        int(normalized[0:2], 16),
        int(normalized[2:4], 16),
        int(normalized[4:6], 16),
    )


def _set_pdf_dash(canvas: Canvas, style: str, width: float) -> None:
    if style == "dotted":
        canvas.setDash(max(width, 0.5), max(width * 2, 1))
    elif style == "dashed":
        canvas.setDash(max(width * 4, 2), max(width * 2, 1))
    else:
        canvas.setDash()


def _pdf_font(style: EffectiveStyle) -> str:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    path = _font_file(style.font_name, style.bold, style.italic)
    registered_name = "xpyxl-" + path.stem
    if registered_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(registered_name, str(path)))
    return registered_name


def _pillow_font(style: EffectiveStyle, scale: float) -> FreeTypeFont:
    from PIL import ImageFont

    size = max(1, round(style.font_size * _PT_TO_PX * scale))
    return ImageFont.truetype(
        str(_font_file(style.font_name, style.bold, style.italic)),
        size=size,
    )


def _font_file(_font_name: str, bold: bool, italic: bool) -> Path:
    import reportlab

    filename = {
        (False, False): "Vera.ttf",
        (True, False): "VeraBd.ttf",
        (False, True): "VeraIt.ttf",
        (True, True): "VeraBI.ttf",
    }[(bold, italic)]
    return Path(reportlab.__file__).parent / "fonts" / filename
