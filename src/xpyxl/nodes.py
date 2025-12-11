from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, TypeAlias

from .styles import Style

__all__ = [
    "CellNode",
    "RowNode",
    "ColumnNode",
    "TableNode",
    "SpacerNode",
    "VerticalStackNode",
    "HorizontalStackNode",
    "SheetNode",
    "ImportedSheetNode",
    "WorkbookNode",
    "SheetItem",
    "RenderableItem",
    "CellValue",
]

CellValue: TypeAlias = object


@dataclass(frozen=True)
class CellNode:
    value: CellValue
    styles: tuple[Style, ...] = ()


@dataclass(frozen=True)
class RowNode:
    cells: tuple[CellNode, ...]
    styles: tuple[Style, ...] = ()


@dataclass(frozen=True)
class ColumnNode:
    cells: tuple[CellNode, ...]
    styles: tuple[Style, ...] = ()


@dataclass(frozen=True)
class TableNode:
    rows: tuple[RowNode, ...]
    styles: tuple[Style, ...] = ()
    header: RowNode | None = None


@dataclass(frozen=True)
class SpacerNode:
    rows: int = 1
    height: float | None = None


@dataclass(frozen=True)
class VerticalStackNode:
    items: tuple["SheetComponent", ...]
    gap: int = 0
    styles: tuple[Style, ...] = ()


@dataclass(frozen=True)
class HorizontalStackNode:
    items: tuple["SheetComponent", ...]
    gap: int = 0
    styles: tuple[Style, ...] = ()


SheetComponent = (
    CellNode
    | RowNode
    | ColumnNode
    | TableNode
    | SpacerNode
    | VerticalStackNode
    | HorizontalStackNode
)


RenderableItem = CellNode | RowNode | ColumnNode | TableNode | SpacerNode


SheetItem = SheetComponent


@dataclass(frozen=True)
class SheetNode:
    name: str
    items: tuple[SheetItem, ...]
    background_color: str | None = None


@dataclass(frozen=True)
class ImportedSheetNode:
    """Reference to an existing Excel sheet to be copied as-is."""

    name: str
    source: str | Path | bytes | BinaryIO
    source_sheet: str


@dataclass(frozen=True)
class WorkbookNode:
    sheets: tuple[SheetNode | ImportedSheetNode, ...]
