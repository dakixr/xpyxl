"""Pluggable rendering engines for xpyxl."""

from __future__ import annotations

from typing import TYPE_CHECKING, Literal

from .base import Engine
from .openpyxl_engine import OpenpyxlEngine
from .xlsxwriter_engine import XlsxWriterEngine

if TYPE_CHECKING:
    from pathlib import Path

__all__ = [
    "Engine",
    "OpenpyxlEngine",
    "XlsxWriterEngine",
    "EngineName",
    "get_engine",
]

EngineName = Literal["openpyxl", "xlsxwriter"]


def get_engine(name: EngineName, path: str | Path) -> Engine:
    """Create an engine instance for the given name and output path."""
    if name == "openpyxl":
        return OpenpyxlEngine(path)
    elif name == "xlsxwriter":
        return XlsxWriterEngine(path)
    else:
        raise ValueError(f"Unknown engine: {name}")


