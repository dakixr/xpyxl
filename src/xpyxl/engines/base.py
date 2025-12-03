"""Abstract base class for rendering engines."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, BinaryIO

if TYPE_CHECKING:
    from ..styles import BorderStyleName

__all__ = ["Engine", "EffectiveStyle", "SaveTarget"]


@dataclass
class EffectiveStyle:
    """Resolved style with all defaults applied."""

    font_name: str
    font_size: float
    bold: bool
    italic: bool
    text_color: str
    fill_color: str | None
    horizontal_align: str | None
    vertical_align: str | None
    indent: int | None
    wrap_text: bool
    shrink_to_fit: bool
    auto_width: bool
    row_height: float | None
    row_width: float | None
    number_format: str | None
    border: "BorderStyleName | None"
    border_color: str | None
    border_top: bool
    border_bottom: bool
    border_left: bool
    border_right: bool


SaveTarget = str | Path | BinaryIO


class Engine(ABC):
    """Abstract base class for Excel rendering engines.

    Engines handle the low-level details of writing to Excel files,
    including cell values, styles, and dimensions.

    Note on indexing:
    - The engine interface uses 1-based indexing for rows and columns
      (matching Excel conventions and openpyxl).
    - Implementations that use 0-based indexing (like xlsxwriter) must
      convert internally.
    """

    def __init__(self) -> None:
        """Initialize engine state without binding to an output target."""

    @abstractmethod
    def create_sheet(self, name: str) -> None:
        """Create a new worksheet with the given name."""
        ...

    @abstractmethod
    def write_cell(
        self,
        row: int,
        col: int,
        value: object,
        style: EffectiveStyle,
        border_fallback_color: str,
    ) -> None:
        """Write a value to a cell with the given style.

        Args:
            row: 1-based row index
            col: 1-based column index
            value: The cell value
            style: Resolved style to apply
            border_fallback_color: Default border color if not specified in style
        """
        ...

    @abstractmethod
    def set_column_width(self, col: int, width: float) -> None:
        """Set the width of a column.

        Args:
            col: 1-based column index
            width: Column width in characters
        """
        ...

    @abstractmethod
    def set_row_height(self, row: int, height: float) -> None:
        """Set the height of a row.

        Args:
            row: 1-based row index
            height: Row height in points
        """
        ...

    @abstractmethod
    def fill_background(
        self,
        color: str,
        max_row: int,
        max_col: int,
    ) -> None:
        """Fill the background of a range of cells.

        Args:
            color: Hex color code (e.g., "#FFFFFF")
            max_row: Maximum row to fill (1-based)
            max_col: Maximum column to fill (1-based)
        """
        ...

    @abstractmethod
    def save(self, target: SaveTarget | None = None) -> bytes | None:
        """Finalize and persist the workbook.

        Args:
            target: Destination to write the workbook to. If None, the engine
                should return the workbook as bytes. When provided, the engine
                writes to the given path or binary file-like object and should
                return None.
        """
        ...
