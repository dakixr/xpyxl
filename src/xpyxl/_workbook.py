from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import TYPE_CHECKING, BinaryIO

from openpyxl import Workbook as _OpenpyxlWorkbook
from openpyxl import load_workbook as _load_workbook

from .engines import EngineName, get_engine
from .nodes import ImportedSheetNode, SheetNode, WorkbookNode
from .render import render_sheet

if TYPE_CHECKING:
    pass

__all__ = ["Workbook"]


class Workbook:
    """Immutable workbook aggregate with a `.save()` convenience."""

    def __init__(self, node: WorkbookNode) -> None:
        self._node = node

    def _has_imported_sheets(self) -> bool:
        """Check if any sheets are ImportedSheetNode."""
        return any(isinstance(s, ImportedSheetNode) for s in self._node.sheets)

    def _save_hybrid_xlsxwriter(
        self, target: str | Path | BinaryIO | None
    ) -> bytes | None:
        """Hybrid save: render SheetNodes with xlsxwriter, merge ImportedSheetNodes with openpyxl.

        This allows using the fast xlsxwriter engine for generated sheets while
        still supporting import_sheet() via openpyxl post-processing.
        """
        from .engines.openpyxl_engine import OpenpyxlEngine
        from .engines.xlsxwriter_engine import XlsxWriterEngine

        # Phase A: Render only SheetNode entries with xlsxwriter
        sheet_nodes = [s for s in self._node.sheets if isinstance(s, SheetNode)]

        if sheet_nodes:
            xw_engine = XlsxWriterEngine()
            for sheet in sheet_nodes:
                render_sheet(xw_engine, sheet)
            xlsx_bytes = xw_engine.save(None)
            assert xlsx_bytes is not None

            # Load the xlsxwriter output with openpyxl
            merged_wb = _load_workbook(
                BytesIO(xlsx_bytes),
                data_only=False,
                rich_text=True,
            )
        else:
            # No SheetNodes, create empty workbook
            merged_wb = _OpenpyxlWorkbook()
            default_sheet = merged_wb.active
            if default_sheet is not None:
                merged_wb.remove(default_sheet)

        # Phase B: Copy imported sheets using OpenpyxlEngine
        openpyxl_engine = OpenpyxlEngine.from_workbook(merged_wb)

        for sheet in self._node.sheets:
            if isinstance(sheet, ImportedSheetNode):
                openpyxl_engine.copy_sheet(sheet.source, sheet.source_sheet, sheet.name)

        # Phase C: Reorder sheets to match the original declaration order
        self._reorder_sheets(merged_wb)

        # Save the final workbook
        return openpyxl_engine.save(target)

    def _reorder_sheets(self, workbook: _OpenpyxlWorkbook) -> None:
        """Reorder workbook sheets to match self._node.sheets declaration order."""
        expected_order = [s.name for s in self._node.sheets]

        # Access internal _sheets list (openpyxl doesn't expose a public reorder API).
        # Reorder in-place rather than replacing the list, to avoid breaking internal
        # workbook invariants that Excel is sensitive to.
        sheets = workbook._sheets  # type: ignore[attr-defined]
        title_to_index = {ws.title: i for i, ws in enumerate(sheets)}

        insert_at = 0
        for title in expected_order:
            idx = title_to_index.get(title)
            if idx is None:
                continue

            if idx != insert_at:
                ws = sheets.pop(idx)
                sheets.insert(insert_at, ws)

                # Update indices for the moved slice.
                start = min(insert_at, idx)
                end = max(insert_at, idx)
                for j in range(start, end + 1):
                    title_to_index[sheets[j].title] = j

            insert_at += 1

        # Ensure the active sheet index is valid after reordering.
        try:
            if sheets:
                workbook.active = 0
        except Exception:
            pass

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
        # Hybrid path: xlsxwriter + imported sheets requires openpyxl post-processing
        if engine == "xlsxwriter" and self._has_imported_sheets():
            return self._save_hybrid_xlsxwriter(target)

        # Standard path: use the selected engine directly
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
