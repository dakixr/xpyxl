"""Tests for builder-level validation."""

from __future__ import annotations

import pytest

import xpyxl as x


@pytest.mark.parametrize("bad_name", ["bad/name", "bad:name", "a[b]", "q?", "s\\t"])
def test_sheet_rejects_invalid_characters(bad_name: str) -> None:
    with pytest.raises(ValueError, match="invalid characters"):
        x.sheet(bad_name)


def test_sheet_rejects_empty_name() -> None:
    with pytest.raises(ValueError, match="cannot be empty"):
        x.sheet("")


def test_sheet_rejects_long_name() -> None:
    with pytest.raises(ValueError, match="<= 31 characters"):
        x.sheet("x" * 32)


def test_import_sheet_validates_dest_name() -> None:
    with pytest.raises(ValueError, match="invalid characters"):
        x.import_sheet("template.xlsx", "Cover", name="bad/name")


@pytest.mark.parametrize("bad_name", ["'quoted", "quoted'", "'both'"])
def test_sheet_rejects_edge_apostrophes(bad_name: str) -> None:
    with pytest.raises(ValueError, match="apostrophe"):
        x.sheet(bad_name)


def test_workbook_rejects_duplicate_sheet_names() -> None:
    with pytest.raises(ValueError, match="Duplicate sheet name"):
        x.workbook()[
            x.sheet("Data")[x.row()["a"]],
            x.sheet("DATA")[x.row()["b"]],
        ]


def test_workbook_rejects_duplicate_import_names() -> None:
    with pytest.raises(ValueError, match="Duplicate sheet name"):
        x.workbook()[
            x.sheet("Data")[x.row()["a"]],
            x.import_sheet("template.xlsx", "Cover", name="data"),
        ]


def test_workbook_validates_direct_sheet_nodes() -> None:
    with pytest.raises(ValueError, match="invalid characters"):
        x.workbook()[x.SheetNode(name="bad/name", items=())]


def test_workbook_validates_direct_imported_sheet_nodes() -> None:
    with pytest.raises(ValueError, match="apostrophe"):
        x.workbook()[x.ImportedSheetNode(name="'bad", source="t.xlsx", source_sheet="S")]
