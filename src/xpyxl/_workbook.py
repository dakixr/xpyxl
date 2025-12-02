from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from openpyxl import Workbook as _OpenpyxlWorkbook

from .engines import EngineName, get_engine
from .nodes import WorkbookNode
from .render import render_sheet

if TYPE_CHECKING:
    pass

__all__ = ["Workbook"]


class Workbook:
    """Immutable workbook aggregate with a `.save()` convenience."""

    def __init__(self, node: WorkbookNode) -> None:
        self._node = node

    def save(self, path: str | Path, *, engine: EngineName = "openpyxl") -> None:
        """Save the workbook to a file.

        Args:
            path: The file path to save to.
            engine: The rendering engine to use. Options are "openpyxl" (default)
                or "xlsxwriter".
        """
        engine_instance = get_engine(engine, path)
        for sheet in self._node.sheets:
            render_sheet(engine_instance, sheet)
        engine_instance.save()

    def to_openpyxl(self) -> _OpenpyxlWorkbook:
        """Convert to an openpyxl Workbook object.

        This method is provided for backward compatibility and advanced use cases
        where direct access to the openpyxl workbook is needed.
        """
        from .engines.openpyxl_engine import OpenpyxlEngine

        # Create a temporary path - we won't actually save to it
        engine = OpenpyxlEngine(Path("/tmp/temp.xlsx"))
        for sheet in self._node.sheets:
            render_sheet(engine, sheet)
        return engine._workbook
