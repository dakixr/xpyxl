from __future__ import annotations

import tempfile
from collections.abc import Iterator, Mapping, Sequence
from pathlib import Path
from typing import overload

import openpyxl
import pytest

import xpyxl as x


class _Records(Sequence[Mapping[str, object]]):
    def __init__(self, prefix: str, count: int) -> None:
        self.prefix = prefix
        self.row_count = count
        self.generated = 0

    def __len__(self) -> int:
        return self.row_count

    @overload
    def __getitem__(self, index: int) -> Mapping[str, object]: ...

    @overload
    def __getitem__(self, index: slice) -> Sequence[Mapping[str, object]]: ...

    def __getitem__(
        self, index: int | slice
    ) -> Mapping[str, object] | Sequence[Mapping[str, object]]:
        if isinstance(index, slice):
            return tuple(self[item] for item in range(*index.indices(self.row_count)))
        if index < 0:
            index += self.row_count
        if index < 0 or index >= self.row_count:
            raise IndexError(index)
        self.generated += 1
        return {"Name": f"{self.prefix}-{index}", "Value": index}

    def __iter__(self) -> Iterator[Mapping[str, object]]:
        for index in range(self.row_count):
            yield self[index]


def test_record_sequence_table_rows_remain_lazy() -> None:
    records = _Records("row", 4)

    table = x.table()[records]

    assert not isinstance(table.rows, tuple)
    assert records.generated == 4  # One pass discovers the complete column set.
    assert table.rows[2].cells[0].value == "row-2"
    assert records.generated == 5


@pytest.mark.parametrize("engine", ["xlsxwriter", "hybrid"])
def test_streamed_tables_render_in_row_order_when_side_by_side(engine: str) -> None:
    left = x.table(style=[x.table_banded])[_Records("left", 3)]
    right = x.table(style=[x.table_bordered])[_Records("right", 3)]
    workbook = x.workbook()[x.sheet("Data")[x.hstack(left, right, gap=1)]]

    with tempfile.TemporaryDirectory() as tmpdir:
        output = Path(tmpdir) / "streamed.xlsx"
        workbook.save(output, engine=engine)  # type: ignore[arg-type]
        sheet = openpyxl.load_workbook(output, data_only=False)["Data"]

    assert sheet["A1"].value == "Name"
    assert sheet["D1"].value == "Name"
    assert sheet["A4"].value == "left-2"
    assert sheet["D4"].value == "right-2"
    assert sheet["E4"].value == 2


@pytest.mark.parametrize("engine", ["xlsxwriter", "hybrid"])
def test_streamed_row_height_is_set_before_rows_are_flushed(engine: str) -> None:
    workbook = x.workbook()[
        x.sheet("Data")[x.row(style=[x.row_height(42)])["tall", "row"]]
    ]

    with tempfile.TemporaryDirectory() as tmpdir:
        output = Path(tmpdir) / "height.xlsx"
        workbook.save(output, engine=engine)  # type: ignore[arg-type]
        sheet = openpyxl.load_workbook(output)["Data"]

    assert sheet.row_dimensions[1].height == 42
