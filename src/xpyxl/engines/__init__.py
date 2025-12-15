"""Pluggable rendering engines for xpyxl."""

from __future__ import annotations

from typing import Literal

from .base import Engine
from .hybrid_engine import HybridEngine
from .openpyxl_engine import OpenpyxlEngine
from .xlsxwriter_engine import XlsxWriterEngine

__all__ = [
    "Engine",
    "HybridEngine",
    "OpenpyxlEngine",
    "XlsxWriterEngine",
    "EngineName",
    "get_engine",
]

EngineName = Literal["openpyxl", "xlsxwriter", "hybrid"]


def get_engine(name: EngineName) -> Engine:
    """Create an engine instance for the given name.

    Available engines:
    - "hybrid" (default): Combines xlsxwriter speed for generated sheets with
      openpyxl for importing sheets. Best balance of speed and features.
    - "openpyxl": Full-featured engine using openpyxl. Supports all features
      including import_sheet.
    - "xlsxwriter": Fast generation engine using xlsxwriter. Does NOT support
      import_sheet - use "hybrid" or "openpyxl" for that.
    """
    if name == "openpyxl":
        return OpenpyxlEngine()
    elif name == "xlsxwriter":
        return XlsxWriterEngine()
    elif name == "hybrid":
        return HybridEngine()
    else:
        raise ValueError(f"Unknown engine: {name}")
