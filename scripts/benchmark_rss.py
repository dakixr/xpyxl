"""Large, process-isolated benchmark for xpyxl wall time and peak RSS.

The default workload writes 2.4 million data cells across three richly styled
worksheets.  Run the benchmark in a fresh child process so the operating
system's peak-RSS counter is not polluted by earlier benchmark iterations::

    uv run python scripts/benchmark_rss.py

Use smaller inputs while iterating, for example ``--rows-per-sheet 5000``.
The final line can be emitted as JSON for automated before/after comparisons.
"""

from __future__ import annotations

import argparse
import gc
import json
import resource
import statistics
import subprocess
import sys
import tempfile
import time
import zipfile
from collections.abc import Iterator, Mapping, Sequence
from dataclasses import asdict, dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import cast, overload

import xpyxl as x
from xpyxl.engines import EngineName

__all__ = ["main"]

_MIB = 1024 * 1024
_COLUMNS = (
    "Order ID",
    "Booked",
    "Region",
    "Country",
    "Account",
    "Owner",
    "Segment",
    "Product",
    "Units",
    "Unit Price",
    "Discount",
    "Net Revenue",
    "Cost",
    "Margin",
    "Approved",
    "Forecast",
)
_REGIONS = ("AMER", "EMEA", "APAC", "LATAM")
_COUNTRIES = ("US", "DE", "SG", "BR", "GB", "ES", "JP", "CA")
_SEGMENTS = ("Enterprise", "Mid-market", "SMB", "Public sector")
_PRODUCTS = ("Atlas", "Beacon", "Cirrus", "Delta", "Ember", "Flux")


class GeneratedRecords(Sequence[Mapping[str, object]]):
    """Deterministic, re-iterable records without retaining source dictionaries."""

    def __init__(self, row_count: int, sheet_number: int) -> None:
        self._row_count = row_count
        self._sheet_number = sheet_number

    def __len__(self) -> int:
        return self._row_count

    @overload
    def __getitem__(self, index: int) -> Mapping[str, object]: ...

    @overload
    def __getitem__(self, index: slice) -> Sequence[Mapping[str, object]]: ...

    def __getitem__(
        self, index: int | slice
    ) -> Mapping[str, object] | Sequence[Mapping[str, object]]:
        if isinstance(index, slice):
            return tuple(self[idx] for idx in range(*index.indices(self._row_count)))
        if index < 0:
            index += self._row_count
        if index < 0 or index >= self._row_count:
            raise IndexError(index)

        sequence = self._sheet_number * self._row_count + index + 1
        units = 1 + sequence % 250
        unit_price = 8.5 + (sequence % 700) / 10
        discount = (sequence % 6) * 0.025
        revenue = units * unit_price * (1 - discount)
        cost = revenue * (0.48 + (sequence % 17) / 100)
        excel_row = index + 7
        return {
            "Order ID": f"ORD-{sequence:09d}",
            "Booked": date(2024, 1, 1) + timedelta(days=sequence % 730),
            "Region": _REGIONS[sequence % len(_REGIONS)],
            "Country": _COUNTRIES[sequence % len(_COUNTRIES)],
            "Account": f"Account {sequence % 12_000:05d}",
            "Owner": f"Rep {sequence % 480:03d}",
            "Segment": _SEGMENTS[sequence % len(_SEGMENTS)],
            "Product": _PRODUCTS[sequence % len(_PRODUCTS)],
            "Units": units,
            "Unit Price": round(unit_price, 2),
            "Discount": discount,
            "Net Revenue": round(revenue, 2),
            "Cost": round(cost, 2),
            "Margin": round(revenue - cost, 2),
            "Approved": sequence % 11 != 0,
            "Forecast": f"=L{excel_row}*1.08",
        }

    def __iter__(self) -> Iterator[Mapping[str, object]]:
        for index in range(self._row_count):
            yield self[index]


@dataclass(frozen=True)
class Result:
    engine: str
    sheets: int
    rows_per_sheet: int
    data_cells: int
    build_seconds: float
    save_seconds: float
    total_seconds: float
    build_peak_rss_mib: float
    peak_rss_mib: float
    output_mib: float


def _peak_rss_mib() -> float:
    # Linux reports KiB; macOS reports bytes. This repository's benchmark and
    # CI environments are Linux, while the branch keeps local macOS runs sane.
    rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    if sys.platform == "darwin":
        return rss / _MIB
    return rss / 1024


def _build_workbook(rows_per_sheet: int, sheets: int) -> x.Workbook:
    report_sheets = []
    for sheet_index in range(sheets):
        records = GeneratedRecords(rows_per_sheet, sheet_index)
        table = x.table(
            column_order=_COLUMNS,
            header_style=[x.text_sm, x.bold, x.text_center, x.bg_primary],
            style=[x.table_bordered, x.table_banded, x.table_compact],
        )[records]
        report_sheets.append(
            x.sheet(f"Revenue {sheet_index + 1}", show_gridlines=False)[
                x.row()[
                    x.cell(
                        colspan=len(_COLUMNS),
                        style=[x.text_2xl, x.bold, x.text_blue, x.row_height(34)],
                    )[f"Global revenue ledger · partition {sheet_index + 1}"]
                ],
                x.row(style=[x.text_sm, x.text_gray])[
                    f"{rows_per_sheet:,} transactions",
                    "730-day reporting horizon",
                    "Deterministic benchmark data",
                ],
                x.hstack(
                    x.row(style=[x.bold, x.bg_info])["Actuals", rows_per_sheet],
                    x.row(style=[x.bold, x.bg_success])["Status", "Validated"],
                    x.row(style=[x.bold, x.bg_warning])["Currency", "USD"],
                    gap=2,
                ),
                x.space(height=8),
                table,
            ]
        )
    return x.workbook()[report_sheets]


def _validate_xlsx(path: Path, expected_sheets: int) -> None:
    if path.stat().st_size == 0:
        raise RuntimeError("Benchmark produced an empty workbook")
    with zipfile.ZipFile(path) as archive:
        worksheets = [
            name
            for name in archive.namelist()
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
        ]
    if len(worksheets) != expected_sheets:
        raise RuntimeError(
            f"Expected {expected_sheets} worksheets, found {len(worksheets)}"
        )


def _worker(rows_per_sheet: int, sheets: int, engine: EngineName) -> Result:
    started = time.perf_counter()
    workbook = _build_workbook(rows_per_sheet, sheets)
    built = time.perf_counter()
    build_peak_rss = _peak_rss_mib()

    with tempfile.TemporaryDirectory(prefix="xpyxl-rss-") as tmpdir:
        output_path = Path(tmpdir) / "large-complex.xlsx"
        workbook.save(output_path, engine=engine)
        saved = time.perf_counter()
        peak_rss = _peak_rss_mib()
        _validate_xlsx(output_path, sheets)
        output_mib = output_path.stat().st_size / _MIB

    del workbook
    gc.collect()
    return Result(
        engine=engine,
        sheets=sheets,
        rows_per_sheet=rows_per_sheet,
        data_cells=rows_per_sheet * sheets * len(_COLUMNS),
        build_seconds=built - started,
        save_seconds=saved - built,
        total_seconds=saved - started,
        build_peak_rss_mib=build_peak_rss,
        peak_rss_mib=peak_rss,
        output_mib=output_mib,
    )


def _run_child(args: argparse.Namespace) -> Result:
    command = [
        sys.executable,
        str(Path(__file__).resolve()),
        "--worker",
        "--rows-per-sheet",
        str(args.rows_per_sheet),
        "--sheets",
        str(args.sheets),
        "--engine",
        args.engine,
    ]
    completed = subprocess.run(command, check=True, capture_output=True, text=True)
    return Result(**json.loads(completed.stdout.strip().splitlines()[-1]))


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows-per-sheet", type=int, default=50_000)
    parser.add_argument("--sheets", type=int, default=3)
    parser.add_argument(
        "--engine", choices=("hybrid", "xlsxwriter", "openpyxl"), default="hybrid"
    )
    parser.add_argument("--runs", type=int, default=1)
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.rows_per_sheet < 1 or args.sheets < 1 or args.runs < 1:
        parser.error("rows-per-sheet, sheets, and runs must all be positive")
    return args


def main() -> None:
    args = _parse_args()
    engine = cast(EngineName, args.engine)
    if args.worker:
        print(json.dumps(asdict(_worker(args.rows_per_sheet, args.sheets, engine))))
        return

    results = [_run_child(args) for _ in range(args.runs)]
    representative = min(results, key=lambda result: result.total_seconds)
    summary = {
        **asdict(representative),
        "runs": args.runs,
        "median_total_seconds": statistics.median(
            result.total_seconds for result in results
        ),
        "median_peak_rss_mib": statistics.median(
            result.peak_rss_mib for result in results
        ),
    }
    if args.json:
        print(json.dumps(summary, indent=2, sort_keys=True))
        return

    print(
        f"{summary['data_cells']:,} data cells · {summary['sheets']} sheets · "
        f"{summary['engine']} engine"
    )
    print(
        f"build {summary['build_seconds']:.2f}s · save {summary['save_seconds']:.2f}s · "
        f"total {summary['total_seconds']:.2f}s"
    )
    print(
        f"build peak {summary['build_peak_rss_mib']:.1f} MiB · "
        f"process peak {summary['peak_rss_mib']:.1f} MiB · "
        f"output {summary['output_mib']:.1f} MiB"
    )


if __name__ == "__main__":
    main()
