"""Chromium-backed PDF and PNG rendering."""

from __future__ import annotations

import json
import math
import os
import shutil
import subprocess
import tempfile
from html import escape
from pathlib import Path

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

__all__ = ["ChromiumRenderer"]

class ChromiumRenderer:
    """Render a resolved workbook with a locally installed Chromium browser."""

    def __init__(self, executable: str | Path | None = None) -> None:
        self.executable = _find_chromium(executable)

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
            raise ValueError(f"Unsupported Chromium export format: {format}")
        validate_scale(scale)

        sheets = select_sheets(workbook, sheet, format)
        with tempfile.TemporaryDirectory(prefix="xpyxl-chromium-") as directory:
            working_directory = Path(directory)
            html_path = working_directory / "report.html"
            output_path = working_directory / f"report.{format}"
            html_path.write_text(_render_document(sheets), encoding="utf-8")
            command = self._command(
                html_path,
                output_path,
                format=format,
                sheet=sheets[0] if format == "png" else None,
                scale=scale,
                user_data_directory=working_directory / "profile",
            )
            completed = _run_chromium(command, output_path)
            if completed.returncode != 0 or not output_path.exists():
                details = (completed.stderr or completed.stdout).strip()
                raise RuntimeError(
                    f"Chromium failed to render {format.upper()}: "
                    f"{details or 'no output was produced'}"
                )
            rendered = output_path.read_bytes()

        return write_output(rendered, target)

    def _command(
        self,
        html_path: Path,
        output_path: Path,
        *,
        format: ExportFormat,
        sheet: SheetLayout | None,
        scale: float,
        user_data_directory: Path,
    ) -> list[str]:
        command = [
            self.executable,
            "--headless=new",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--hide-scrollbars",
            f"--user-data-dir={user_data_directory}",
        ]
        if getattr(os, "geteuid", lambda: -1)() == 0:
            command.append("--no-sandbox")
        if format == "pdf":
            command.extend(
                [
                    "--no-pdf-header-footer",
                    "--print-to-pdf-no-header",
                    f"--print-to-pdf={output_path}",
                ]
            )
        else:
            assert sheet is not None
            width, height = sheet_size(sheet)
            viewport_width = max(1, math.ceil(width))
            viewport_height = max(1, math.ceil(height))
            command.extend(
                [
                    f"--force-device-scale-factor={scale}",
                    f"--window-size={viewport_width},{viewport_height}",
                    f"--screenshot={output_path}",
                ]
            )
        command.append(html_path.resolve().as_uri())
        return command


def _run_chromium(command: list[str], output_path: Path) -> subprocess.CompletedProcess[str]:
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        check=False,
    )
    if (
        (completed.returncode != 0 or not output_path.exists())
        and "--no-sandbox" not in command
        and "no usable sandbox" in (completed.stderr + completed.stdout).lower()
    ):
        retry_command = [*command[:-1], "--no-sandbox", command[-1]]
        completed = subprocess.run(
            retry_command,
            capture_output=True,
            text=True,
            check=False,
        )
    return completed


def _find_chromium(executable: str | Path | None) -> str:
    configured = executable or os.environ.get("XPYXL_CHROMIUM_PATH")
    if configured:
        path = shutil.which(str(configured))
        if path:
            return path
        candidate = Path(configured)
        if candidate.is_file():
            return str(candidate)
        raise FileNotFoundError(f"Chromium executable not found: {configured}")

    for name in ("chromium", "chromium-browser", "google-chrome", "chrome"):
        path = shutil.which(name)
        if path:
            return path
    raise FileNotFoundError(
        "Chromium was not found. Install Chromium/Google Chrome or set "
        "XPYXL_CHROMIUM_PATH."
    )


def _render_document(sheets: tuple[SheetLayout, ...]) -> str:
    page_rules: list[str] = []
    sheet_markup: list[str] = []
    for index, sheet in enumerate(sheets):
        width, height = sheet_size(sheet)
        page_name = f"sheet{index}"
        page_rules.append(
            f"@page {page_name} {{ size: {width:.3f}px {height:.3f}px; margin: 0; }}"
        )
        sheet_markup.append(_render_sheet(sheet, index, page_name, width, height))

    return """<!doctype html>
<html><head><meta charset="utf-8"><style>
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; background: white; }
body { overflow: hidden; }
.sheet { position: relative; overflow: hidden; break-after: page; }
.sheet:last-child { break-after: auto; }
.cell { position: absolute; display: flex; padding: 2px 6px; overflow: hidden; }
.value { width: 100%; overflow: visible; }
""" + "\n".join(page_rules) + "\n</style></head><body>" + "".join(sheet_markup) + "</body></html>"


def _render_sheet(
    sheet: SheetLayout,
    index: int,
    page_name: str,
    width: float,
    height: float,
) -> str:
    max_row = max(sheet.max_row, 1)
    max_col = max(sheet.max_col, 1)
    column_widths = [column_width(sheet, col) for col in range(1, max_col + 1)]
    row_heights = [row_height(sheet, row) for row in range(1, max_row + 1)]
    column_offsets = offsets(column_widths)
    row_offsets = offsets(row_heights)
    cells: list[str] = []

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            if (row, col) in sheet.covered_cells:
                continue
            cell = sheet.cells.get((row, col))
            rowspan = cell.rowspan if cell else 1
            colspan = cell.colspan if cell else 1
            cell_width = sum(column_widths[col - 1 : col - 1 + colspan])
            cell_height = sum(row_heights[row - 1 : row - 1 + rowspan])
            style = _cell_css(cell, sheet.show_gridlines)
            value = format_value(cell.value, cell.style.number_format) if cell else ""
            cells.append(
                '<div class="cell" style="left:{left:.3f}px;top:{top:.3f}px;'
                'width:{width:.3f}px;height:{height:.3f}px;{style}">'
                '<div class="value">{value}</div></div>'.format(
                    left=column_offsets[col - 1],
                    top=row_offsets[row - 1],
                    width=cell_width,
                    height=cell_height,
                    style=escape(style, quote=True),
                    value=escape(value),
                )
            )

    background = sheet.background_color or "#FFFFFF"
    return (
        f'<section class="sheet sheet-{index}" aria-label="{escape(sheet.name)}" '
        f'style="page:{page_name};width:{width:.3f}px;height:{height:.3f}px;'
        f'background:{background}">{"".join(cells)}</section>'
    )


def _cell_css(cell: CellLayout | None, show_gridlines: bool) -> str:
    parts: list[str] = []
    if show_gridlines:
        parts.extend(["border-right:1px solid #D9D9D9", "border-bottom:1px solid #D9D9D9"])
    if cell is None:
        return ";".join(parts)

    style = cell.style
    font_name = json.dumps(style.font_name, ensure_ascii=False)
    parts.extend(
        [
            f"font-family:{font_name},Arial,sans-serif",
            f"font-size:{style.font_size}pt",
            f"color:{style.text_color}",
            f"font-weight:{'700' if style.bold else '400'}",
            f"font-style:{'italic' if style.italic else 'normal'}",
            f"white-space:{'pre-wrap' if style.wrap_text else 'pre'}",
            f"overflow-wrap:{'anywhere' if style.wrap_text else 'normal'}",
        ]
    )
    if style.fill_color:
        parts.append(f"background:{style.fill_color}")
    if style.horizontal_align:
        parts.append(
            "justify-content:"
            + {"left": "flex-start", "center": "center", "right": "flex-end"}.get(
                style.horizontal_align, "flex-start"
            )
        )
        parts.append(
            "text-align:"
            + {"left": "left", "center": "center", "right": "right"}.get(
                style.horizontal_align, "left"
            )
        )
    parts.append(
        "align-items:"
        + {"top": "flex-start", "center": "center", "bottom": "flex-end"}.get(
            style.vertical_align or "top", "flex-start"
        )
    )
    if style.indent:
        parts.append(f"padding-left:{6 + style.indent * 8}px")
    parts.extend(_border_css(style, cell.border_fallback_color))
    return ";".join(parts)


def _border_css(style: EffectiveStyle, fallback_color: str) -> list[str]:
    if style.border == "none":
        return ["border:none"]
    if not style.border and not any(
        (style.border_top, style.border_bottom, style.border_left, style.border_right)
    ):
        return []
    border = style.border or "thin"
    declaration = (
        f"{border_width(border):g}px {border_line_style(border)} "
        f"{style.border_color or fallback_color}"
    )
    sides = {
        "top": style.border_top,
        "bottom": style.border_bottom,
        "left": style.border_left,
        "right": style.border_right,
    }
    if any(sides.values()):
        return [f"border-{side}:{declaration}" for side, enabled in sides.items() if enabled]
    return [f"border:{declaration}"]
