"""Tests for style helpers."""

from __future__ import annotations

import pytest

import xpyxl as x


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("#fff", "#FFFFFF"),
        ("fff", "#FFFFFF"),
        ("#a1b2c3", "#A1B2C3"),
        ("  #A1B2C3  ", "#A1B2C3"),
    ],
)
def test_normalize_hex_accepts_valid_colors(raw: str, expected: str) -> None:
    assert x.normalize_hex(raw) == expected


@pytest.mark.parametrize("raw", ["#GGGGGG", "12345z", "#12 456"])
def test_normalize_hex_rejects_non_hex_digits(raw: str) -> None:
    with pytest.raises(ValueError, match="Invalid hex color"):
        x.normalize_hex(raw)


def test_normalize_hex_rejects_wrong_length() -> None:
    with pytest.raises(ValueError, match="Expected 6 hex characters"):
        x.normalize_hex("#1234")


def test_normalize_hex_rejects_empty() -> None:
    with pytest.raises(ValueError, match="cannot be empty"):
        x.normalize_hex("   ")
