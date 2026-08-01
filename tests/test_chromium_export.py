"""Tests for the dedicated Chromium document exporter."""

from __future__ import annotations

import shutil
import tempfile
from io import BytesIO
from pathlib import Path

import pytest
from PIL import Image

import xpyxl as x
from xpyxl.exporting.chromium import _render_document
from xpyxl.exporting.model import build_workbook_layout


def _browser() -> str | None:
    for name in ("chromium", "chromium-browser", "google-chrome", "chrome"):
        if path := shutil.which(name):
            return path
    return None


def test_print_document_is_standalone_and_does_not_use_html_engine() -> None:
    workbook = x.workbook()[
        x.sheet("Summary", show_gridlines=False)[
            x.row(style=[x.bold, x.bg_primary, x.text_white])["Name", "Value"],
            x.row()["Revenue", 42],
        ]
    ]
    layout = build_workbook_layout(workbook._node)
    html = _render_document(layout.sheets)

    assert "Revenue" in html
    assert "position: absolute" in html
    assert "cdn.tailwindcss.com" not in html
    assert "data-sheet" not in html
    assert "font-family:&quot;Calibri&quot;" in html


def test_print_document_escapes_font_names_as_css_strings() -> None:
    workbook = x.workbook()[
        x.sheet("S")[
            x.row(style=[x.Style(font_name='Bad\n";background:red')])["Safe"]
        ]
    ]
    layout = build_workbook_layout(workbook._node)
    html = _render_document(layout.sheets)

    assert 'font-family:&quot;Bad\\n\\&quot;;background:red&quot;' in html
    assert 'font-family:&quot;Bad\n' not in html


def test_export_format_is_inferred_from_path() -> None:
    browser = _browser()
    if browser is None:
        pytest.skip("Chromium is not installed")
    with tempfile.TemporaryDirectory() as directory:
        output = Path(directory) / "report.pdf"
        workbook = x.workbook()[x.sheet("S")[x.row()["PDF"]]]
        result = workbook.export(output, chromium_path=browser)

        assert result is None
        assert output.read_bytes().startswith(b"%PDF")


def test_export_png_bytes() -> None:
    browser = _browser()
    if browser is None:
        pytest.skip("Chromium is not installed")
    workbook = x.workbook()[
        x.sheet("S")[x.row()[x.cell(style=[x.bg_primary, x.text_white])["PNG"]]]
    ]
    result = workbook.export(format="png", chromium_path=browser)

    assert isinstance(result, bytes)
    assert result.startswith(b"\x89PNG\r\n\x1a\n")
    with Image.open(BytesIO(result)).convert("RGB") as image:
        assert image.getpixel((2, 2)) == (37, 99, 235)


def test_export_writes_binary_stream() -> None:
    browser = _browser()
    if browser is None:
        pytest.skip("Chromium is not installed")
    stream = BytesIO()
    workbook = x.workbook()[x.sheet("S")[x.row()["PDF"]]]

    result = workbook.export(stream, chromium_path=browser)

    assert result is None
    assert stream.getvalue().startswith(b"%PDF")


def test_png_requires_sheet_for_multiple_sheets() -> None:
    browser = _browser()
    if browser is None:
        pytest.skip("Chromium is not installed")
    workbook = x.workbook()[
        x.sheet("One")[x.row()[1]],
        x.sheet("Two")[x.row()[2]],
    ]

    with pytest.raises(ValueError, match="requires sheet"):
        workbook.export(format="png", chromium_path=browser)


def test_invalid_export_suffix_is_rejected_before_browser_lookup() -> None:
    workbook = x.workbook()[x.sheet("S")[x.row()[1]]]

    with pytest.raises(ValueError, match=r"\.pdf or \.png"):
        workbook.export("report.jpg")
