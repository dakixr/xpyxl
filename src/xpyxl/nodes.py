from __future__ import annotations

from dataclasses import dataclass
from typing import Any

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
    "WorkbookNode",
    "SheetItem",
    "RenderableItem",
]


@dataclass(frozen=True)
class CellNode:
    value: Any
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


@dataclass(frozen=True)
class WorkbookNode:
    sheets: tuple[SheetNode, ...]
