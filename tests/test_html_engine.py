"""Tests for HTML rendering engine."""

from __future__ import annotations

import tempfile
from io import BytesIO
from pathlib import Path

import pytest

import xpyxl as x


def test_save_returns_bytes_when_target_none() -> None:
    workbook = x.workbook()[x.sheet("Test")[x.row()["Hello", "World"]]]
    result = workbook.save(engine="html")

    assert result is not None
    assert isinstance(result, bytes)
    assert b"<!DOCTYPE html>" in result
    assert b"Hello" in result
    assert b"World" in result


def test_save_writes_file_when_path_given() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "output.html"

        workbook = x.workbook()[x.sheet("Test")[x.row()["Data"]]]
        result = workbook.save(output_path, engine="html")

        assert result is None
        assert output_path.exists()
        content = output_path.read_text()
        assert "<!DOCTYPE html>" in content
        assert "Data" in content


def test_save_writes_to_stream() -> None:
    buffer = BytesIO()

    workbook = x.workbook()[x.sheet("Test")[x.row()["Stream"]]]
    result = workbook.save(buffer, engine="html")

    assert result is None
    content = buffer.getvalue().decode("utf-8")
    assert "<!DOCTYPE html>" in content
    assert "Stream" in content


def test_multiple_sheets_with_tabs() -> None:
    workbook = x.workbook()[
        x.sheet("Sheet1")[x.row()["A"]],
        x.sheet("Sheet2")[x.row()["B"]],
        x.sheet("Sheet3")[x.row()["C"]],
    ]
    result = workbook.save(engine="html")
    assert isinstance(result, bytes)
    html = result.decode("utf-8")

    assert "Sheet1" in html
    assert "Sheet2" in html
    assert "Sheet3" in html
    assert "data-sheet" in html
    assert html.count("class=\"sheet-content\"") == 3


def test_sheet_names_escaped() -> None:
    workbook = x.workbook()[x.sheet("<Script>")[x.row()["Safe"]]]
    result = workbook.save(engine="html")
    assert isinstance(result, bytes)
    html = result.decode("utf-8")

    assert "&lt;Script&gt;" in html
    assert "<Script>" not in html


def test_bold_italic_styles() -> None:
    workbook = x.workbook()[
        x.sheet("Test")[
            x.row()[x.cell(style=[x.bold])["Bold"]],
            x.row()[x.cell(style=[x.italic])["Italic"]],
        ]
    ]
    result = workbook.save(engine="html")
    assert isinstance(result, bytes)
    html = result.decode("utf-8")

    assert "font-bold" in html
    assert "italic" in html


def test_text_alignment() -> None:
    workbook = x.workbook()[
        x.sheet("Test")[
            x.row()[x.cell(style=[x.text_left])["Left"]],
            x.row()[x.cell(style=[x.text_center])["Center"]],
            x.row()[x.cell(style=[x.text_right])["Right"]],
        ]
    ]
    result = workbook.save(engine="html")
    assert isinstance(result, bytes)
    html = result.decode("utf-8")

    assert "text-left" in html
    assert "text-center" in html
    assert "text-right" in html


def test_background_color() -> None:
    workbook = x.workbook()[
        x.sheet("Test")[x.row()[x.cell(style=[x.bg_primary])["Colored"]]]
    ]
    result = workbook.save(engine="html")
    assert isinstance(result, bytes)
    html = result.decode("utf-8")

    assert "background-color:" in html


def test_borders() -> None:
    workbook = x.workbook()[
        x.sheet("Test")[x.row()[x.cell(style=[x.border_all])["Bordered"]]]
    ]
    result = workbook.save(engine="html")
    assert isinstance(result, bytes)
    html = result.decode("utf-8")

    assert "border:" in html or "border-top:" in html


def test_copy_sheet_from_xlsx() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        source_path = Path(tmpdir) / "source.xlsx"
        source_wb = x.workbook()[
            x.sheet("Source")[x.row()["Imported", "Data"]]
        ]
        source_wb.save(source_path, engine="openpyxl")

        workbook = x.workbook()[x.import_sheet(source_path, "Source", name="Imported")]
        result = workbook.save(engine="html")
        assert isinstance(result, bytes)
        html = result.decode("utf-8")

        assert "Imported" in html
        assert "Data" in html


def test_copy_sheet_missing_sheet_raises() -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        source_path = Path(tmpdir) / "source.xlsx"
        source_wb = x.workbook()[x.sheet("Exists")[x.row()["Data"]]]
        source_wb.save(source_path, engine="openpyxl")

        workbook = x.workbook()[
            x.import_sheet(source_path, "NotExists", name="Imported")
        ]

        with pytest.raises(ValueError, match="not found"):
            workbook.save(engine="html")


def test_with_table() -> None:
    table = x.table()[
        [
            ["Name", "Value"],
            ["A", 1],
            ["B", 2],
        ]
    ]
    workbook = x.workbook()[x.sheet("Table")[table]]
    result = workbook.save(engine="html")
    assert isinstance(result, bytes)
    html = result.decode("utf-8")

    assert "Name" in html
    assert "Value" in html
    assert "<table" in html


def test_tailwind_cdn_included() -> None:
    workbook = x.workbook()[x.sheet("Test")[x.row()["Data"]]]
    result = workbook.save(engine="html")
    assert isinstance(result, bytes)
    html = result.decode("utf-8")

    assert "cdn.tailwindcss.com" in html
