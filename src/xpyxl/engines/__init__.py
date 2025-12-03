"""Pluggable rendering engines for xpyxl."""

from __future__ import annotations

from typing import Literal

from .base import Engine
from .openpyxl_engine import OpenpyxlEngine
from .xlsxwriter_engine import XlsxWriterEngine

__all__ = [
    "Engine",
    "OpenpyxlEngine",
    "XlsxWriterEngine",
    "EngineName",
    "get_engine",
]

EngineName = Literal["openpyxl", "xlsxwriter"]


def get_engine(name: EngineName) -> Engine:
    """Create an engine instance for the given name."""
    if name == "openpyxl":
        return OpenpyxlEngine()
    elif name == "xlsxwriter":
        return XlsxWriterEngine()
    else:
        raise ValueError(f"Unknown engine: {name}")
