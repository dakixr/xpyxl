from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, BinaryIO

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

    def save(
        self,
        target: str | Path | BinaryIO | None = None,
        *,
        engine: EngineName = "openpyxl",
    ) -> bytes | None:
        """Save the workbook to a file or binary stream.

        Args:
            target: File path or binary buffer to write to. Pass None to receive
                the rendered workbook as bytes.
            engine: The rendering engine to use. Options are "openpyxl" (default)
                or "xlsxwriter".
        """
        engine_instance = get_engine(engine)
        for sheet in self._node.sheets:
            render_sheet(engine_instance, sheet)
        return engine_instance.save(target)

    def to_openpyxl(self) -> _OpenpyxlWorkbook:
        """Convert to an openpyxl Workbook object.

        This method is provided for backward compatibility and advanced use cases
        where direct access to the openpyxl workbook is needed.
        """
        from .engines.openpyxl_engine import OpenpyxlEngine

        # Render with the openpyxl engine without persisting to disk
        engine = OpenpyxlEngine()
        for sheet in self._node.sheets:
            render_sheet(engine, sheet)
        return engine._workbook
