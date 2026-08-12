"""Peak-RSS benchmark dominated by custom composed xpyxl components.

The default workload renders 100,000 reusable report sections containing 5.4
million heterogeneous cells. Unlike ``benchmark_rss.py``, large tables are not
the dominant structure: every section is assembled from nested rows, columns,
vertical/horizontal stacks, merged cells, spacers, formulas, dates, numbers,
booleans, bytes, blanks, and varied styles.

Run the full benchmark with::

    uv run python scripts/benchmark_components_rss.py

Use ``--components 1000`` for a quick iteration-sized workload.
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
from collections.abc import Iterator, Sequence
from dataclasses import asdict, dataclass
from datetime import date, timedelta
from decimal import Decimal
from pathlib import Path
from typing import cast, overload

import xpyxl as x
from xpyxl.engines import EngineName

__all__ = ["main"]

_MIB = 1024 * 1024
_CELLS_PER_COMPONENT = 54
_SHEET_HEADER_CELLS = 11
_REGIONS = ("AMER", "EMEA", "APAC", "LATAM")
_SEGMENTS = ("Enterprise", "Mid-market", "SMB", "Public sector")
_STATUSES = ("Healthy", "Watch", "At risk", "Escalated")


def _metric_card(
    label: str,
    value: object,
    trend: str,
    background: x.Style,
) -> x.Node:
    """A reusable KPI component composed from two styled rows."""
    return x.vstack(
        x.row(style=[x.text_xs, x.text_gray, background])[label],
        x.row(style=[x.text_lg, x.bold, background])[value, trend],
        style=[x.border_all, x.border_muted],
    )


def _report_component(sequence: int) -> x.Node:
    """Build one deliberately varied, nested report section."""
    units = 1 + sequence % 900
    price = Decimal("9.75") + Decimal(sequence % 500) / Decimal(10)
    revenue = Decimal(units) * price
    cost = revenue * Decimal("0.57")
    margin = revenue - cost
    region = _REGIONS[sequence % len(_REGIONS)]
    status = _STATUSES[sequence % len(_STATUSES)]
    booked = date(2023, 1, 1) + timedelta(days=sequence % 1_095)

    title = x.row()[
        x.cell(
            colspan=16,
            style=[x.text_lg, x.bold, x.text_blue, x.bg_muted, x.row_height(28)],
        )[f"Account review · ACCT-{sequence:08d} · {status}"]
    ]
    metrics = x.hstack(
        _metric_card("Revenue", revenue, f"+{sequence % 17}%", x.bg_info),
        _metric_card("Margin", margin, f"{43 + sequence % 8}%", x.bg_success),
        _metric_card("Units", units, f"{sequence % 31} open", x.bg_warning),
        _metric_card("Risk", status, f"P{1 + sequence % 4}", x.bg_muted),
        gap=1,
    )
    detail = x.row(style=[x.align_middle, x.border_bottom, x.border_muted])[
        f"ORD-{sequence:010d}",
        x.cell(style=[x.date_short])[booked],
        region,
        f"Account {sequence % 25_000:05d}",
        f"Owner {sequence % 750:03d}",
        units,
        x.cell(style=[x.currency_usd])[price],
        x.cell(style=[x.currency_usd, x.bold])[revenue],
        x.cell(style=[x.currency_usd])[cost],
        x.cell(style=[x.currency_usd, x.text_green])[margin],
        x.cell(style=[x.percent])[(sequence % 100) / 100],
        sequence % 9 != 0,
        f"={units}*{float(price):.2f}",
        f"batch-{sequence % 256:02x}".encode(),
        None,
        x.cell(style=[x.wrap, x.row_width(32)])[
            f"Deterministic review note for account {sequence}; region={region}."
        ],
    ]
    channel_grid = x.hstack(
        x.col(style=[x.text_sm, x.bold, x.bg_muted])[
            "Channel",
            "Direct",
            "Partner",
            "Marketplace",
        ],
        x.col(style=[x.text_sm, x.number_comma])[
            "Pipeline",
            units * 3,
            units * 2,
            units,
        ],
        x.col(style=[x.text_sm, x.currency_usd])[
            "Weighted",
            revenue * Decimal("0.7"),
            revenue * Decimal("0.5"),
            revenue * Decimal("0.3"),
        ],
        x.col(style=[x.text_sm, x.text_center])[
            "Approved",
            sequence % 2 == 0,
            sequence % 3 == 0,
            sequence % 5 == 0,
        ],
        gap=1,
    )
    footer = x.row(style=[x.text_xs, x.text_gray, x.italic])[
        "Owner review",
        f"FY{2024 + sequence % 3}",
        f"Q{1 + sequence % 4}",
        region,
        _SEGMENTS[sequence % len(_SEGMENTS)],
        status,
        f"rev-{sequence % 100:02d}",
        "generated",
    ]
    note = x.row()[
        x.cell(colspan=16, style=[x.text_xs, x.wrap, x.bg_muted])[
            "Component boundary: title + KPI cards + detail row + channel grid + footer"
        ]
    ]
    return x.vstack(
        title,
        metrics,
        detail,
        channel_grid,
        footer,
        note,
        x.space(height=6),
        gap=1,
        style=[x.text_sm],
    )


def _sheet_header(sheet_number: int, component_count: int) -> x.Node:
    return x.vstack(
        x.row()[
            x.cell(colspan=16, style=[x.text_3xl, x.bold, x.text_primary])[
                f"Portfolio review · shard {sheet_number + 1}"
            ]
        ],
        x.row(style=[x.text_sm, x.text_gray])[
            f"{component_count:,} composed sections",
            f"{component_count * _CELLS_PER_COMPONENT:,} component cells",
            "No dominant large table",
            "Deterministic data",
        ],
        x.hstack(
            x.row(style=[x.bg_success, x.bold])["Healthy", "on plan"],
            x.row(style=[x.bg_warning, x.bold])["Watch", "review"],
            x.row(style=[x.bg_red, x.text_white, x.bold])["At risk", "escalate"],
            gap=2,
        ),
        x.space(height=10),
        gap=1,
    )


class GeneratedComponents(Sequence[x.Node]):
    """A repeatable sequence of custom components for one worksheet."""

    def __init__(self, component_count: int, sequence_offset: int, sheet: int) -> None:
        self._component_count = component_count
        self._sequence_offset = sequence_offset
        self._sheet = sheet

    def __len__(self) -> int:
        return self._component_count + 1

    @overload
    def __getitem__(self, index: int) -> x.Node: ...

    @overload
    def __getitem__(self, index: slice) -> Sequence[x.Node]: ...

    def __getitem__(self, index: int | slice) -> x.Node | Sequence[x.Node]:
        if isinstance(index, slice):
            return tuple(self[item] for item in range(*index.indices(len(self))))
        if index < 0:
            index += len(self)
        if index < 0 or index >= len(self):
            raise IndexError(index)
        if index == 0:
            return _sheet_header(self._sheet, self._component_count)
        return _report_component(self._sequence_offset + index)

    def __iter__(self) -> Iterator[x.Node]:
        for index in range(len(self)):
            yield self[index]


@dataclass(frozen=True)
class Result:
    engine: str
    sheets: int
    components: int
    component_cells: int
    build_seconds: float
    save_seconds: float
    total_seconds: float
    build_peak_rss_mib: float
    peak_rss_mib: float
    output_mib: float


def _peak_rss_mib() -> float:
    rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return rss / _MIB if sys.platform == "darwin" else rss / 1024


def _component_counts(total: int, sheets: int) -> list[int]:
    quotient, remainder = divmod(total, sheets)
    return [quotient + (index < remainder) for index in range(sheets)]


def _build_workbook(components: int, sheets: int) -> x.Workbook:
    report_sheets = []
    sequence_offset = 0
    for sheet_index, count in enumerate(_component_counts(components, sheets)):
        generated = GeneratedComponents(count, sequence_offset, sheet_index)
        report_sheets.append(
            x.sheet(f"Components {sheet_index + 1}", show_gridlines=False)[generated]
        )
        sequence_offset += count
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


def _worker(components: int, sheets: int, engine: EngineName) -> Result:
    started = time.perf_counter()
    workbook = _build_workbook(components, sheets)
    built = time.perf_counter()
    build_peak_rss = _peak_rss_mib()

    with tempfile.TemporaryDirectory(prefix="xpyxl-components-rss-") as tmpdir:
        output_path = Path(tmpdir) / "large-components.xlsx"
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
        components=components,
        component_cells=components * _CELLS_PER_COMPONENT
        + sheets * _SHEET_HEADER_CELLS,
        build_seconds=built - started,
        save_seconds=saved - built,
        total_seconds=saved - started,
        build_peak_rss_mib=build_peak_rss,
        peak_rss_mib=peak_rss,
        output_mib=output_mib,
    )


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--components", type=int, default=100_000)
    parser.add_argument("--sheets", type=int, default=4)
    parser.add_argument(
        "--engine", choices=("hybrid", "xlsxwriter", "openpyxl"), default="hybrid"
    )
    parser.add_argument("--runs", type=int, default=1)
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.components < 1 or args.sheets < 1 or args.runs < 1:
        parser.error("components, sheets, and runs must all be positive")
    if args.components < args.sheets:
        parser.error("components must be greater than or equal to sheets")
    return args


def _run_child(args: argparse.Namespace) -> Result:
    command = [
        sys.executable,
        str(Path(__file__).resolve()),
        "--worker",
        "--components",
        str(args.components),
        "--sheets",
        str(args.sheets),
        "--engine",
        args.engine,
    ]
    completed = subprocess.run(command, check=True, capture_output=True, text=True)
    return Result(**json.loads(completed.stdout.strip().splitlines()[-1]))


def main() -> None:
    args = _parse_args()
    engine = cast(EngineName, args.engine)
    if args.worker:
        print(json.dumps(asdict(_worker(args.components, args.sheets, engine))))
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
        f"{summary['components']:,} custom components · "
        f"{summary['component_cells']:,} cells · {summary['sheets']} sheets · "
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
