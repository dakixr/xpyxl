"""Benchmark script comparing openpyxl and xlsxwriter rendering engines."""

from __future__ import annotations

import json
import sys
import tempfile
import time
import tracemalloc
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, cast

# Add src directory to Python path before importing xpyxl
_project_root = Path(__file__).resolve().parent.parent
_src_dir = _project_root / "src"
if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

import xpyxl as x  # noqa: E402
from xpyxl.engines import EngineName  # noqa: E402

__all__ = ["main"]

# Number of runs per benchmark for averaging
NUM_RUNS = 3

# Table sizes to test
TABLE_SIZES = [100, 1_000, 10_000, 50_000]


@dataclass
class BenchmarkResult:
    """Result of a single benchmark run."""

    scenario: str
    engine: str
    table_size: int | None
    execution_time: float  # seconds
    memory_peak: float  # MB
    memory_current: float  # MB
    success: bool
    error: str | None = None


def run_benchmark(
    engine_name: EngineName,
    scenario_name: str,
    func: Callable[..., None],
    *args: Any,
    table_size: int | None = None,
) -> BenchmarkResult:
    """Run a benchmark function and collect metrics.

    Args:
        engine_name: Name of the engine ("openpyxl" or "xlsxwriter")
        scenario_name: Name of the benchmark scenario
        func: Function to benchmark (should accept engine_name as first arg)
        *args: Additional arguments to pass to func
        table_size: Optional table size for big table benchmarks

    Returns:
        BenchmarkResult with collected metrics
    """
    times = []
    memory_peaks = []
    memory_currents = []
    last_error = None

    for run in range(NUM_RUNS):
        try:
            # Start memory tracking
            tracemalloc.start()

            # Measure execution time
            start_time = time.perf_counter()
            func(engine_name, *args)
            end_time = time.perf_counter()

            # Get memory statistics
            current, peak = tracemalloc.get_traced_memory()
            tracemalloc.stop()

            execution_time = end_time - start_time
            memory_peak_mb = peak / (1024 * 1024)
            memory_current_mb = current / (1024 * 1024)

            times.append(execution_time)
            memory_peaks.append(memory_peak_mb)
            memory_currents.append(memory_current_mb)

        except Exception as e:
            tracemalloc.stop()
            last_error = str(e)
            print(f"  Run {run + 1} failed: {e}")

    if not times:
        return BenchmarkResult(
            scenario=scenario_name,
            engine=engine_name,
            table_size=table_size,
            execution_time=0.0,
            memory_peak=0.0,
            memory_current=0.0,
            success=False,
            error=last_error,
        )

    # Average the results
    avg_time = sum(times) / len(times)
    avg_peak = sum(memory_peaks) / len(memory_peaks)
    avg_current = sum(memory_currents) / len(memory_currents)

    return BenchmarkResult(
        scenario=scenario_name,
        engine=engine_name,
        table_size=table_size,
        execution_time=avg_time,
        memory_peak=avg_peak,
        memory_current=avg_current,
        success=True,
        error=None,
    )


def benchmark_big_tables(engine_name: EngineName, num_rows: int) -> None:
    """Benchmark rendering a large table.

    Args:
        engine_name: Engine to use ("openpyxl" or "xlsxwriter")
        num_rows: Number of rows in the table
    """
    # Generate table data
    rows = []
    for idx in range(num_rows):
        rows.append(
            {
                "Row": idx,
                "Name": f"Item {idx}",
                "Category": "Even" if idx % 2 == 0 else "Odd",
                "Value": idx * 1.5,
                "Flag": "✔" if idx % 10 == 0 else "",
            }
        )

    table = x.table()[rows]

    sheet = x.sheet("Big Table")[
        x.row(style=[x.text_lg, x.bold])[f"{num_rows:,}-row table"],
        x.row(style=[x.text_sm, x.text_gray])["Performance benchmark test."],
        x.space(),
        table,
    ]

    workbook = x.workbook()[sheet]

    # Save to temporary file
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        workbook.save(tmp_path, engine=engine_name)
    finally:
        # Clean up
        Path(tmp_path).unlink(missing_ok=True)


def benchmark_simple_layouts(engine_name: EngineName) -> None:
    """Benchmark simple layout operations (vstack/hstack).

    Args:
        engine_name: Engine to use ("openpyxl" or "xlsxwriter")
    """
    # Simple vstack
    vstack_layout = x.vstack(
        x.row(style=[x.text_xl, x.bold])["Header"],
        x.row()["Row 1"],
        x.row()["Row 2"],
        x.row()["Row 3"],
        x.space(),
        x.row(style=[x.text_sm, x.text_gray])["Footer"],
        gap=1,
    )

    # Simple hstack
    hstack_layout = x.hstack(
        x.col(style=[x.bold])["Col A", "Col B", "Col C"],
        x.col()["Val 1", "Val 2", "Val 3"],
        x.col()["Val 4", "Val 5", "Val 6"],
        gap=2,
    )

    # Combined layout
    combined = x.vstack(
        vstack_layout,
        x.space(),
        hstack_layout,
        gap=2,
    )

    sheet = x.sheet("Simple Layouts")[combined]
    workbook = x.workbook()[sheet]

    # Save to temporary file
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        workbook.save(tmp_path, engine=engine_name)
    finally:
        # Clean up
        Path(tmp_path).unlink(missing_ok=True)


def benchmark_complex_layouts(engine_name: EngineName) -> None:
    """Benchmark complex multi-sheet layouts with styling.

    Args:
        engine_name: Engine to use ("openpyxl" or "xlsxwriter")
    """
    # Summary sheet with multiple components
    summary_table = x.table(
        header_style=[x.text_sm, x.text_gray, x.text_center],
        style=[x.table_bordered, x.table_banded, x.table_compact],
    )[
        [
            {"Region": "EMEA", "Units": 1200, "Revenue": "$1.6M"},
            {"Region": "APAC", "Units": 900, "Revenue": "$1.1M"},
            {"Region": "AMER", "Units": 1500, "Revenue": "$1.5M"},
        ]
    ]

    stats = x.hstack(
        x.table(style=[x.table_bordered, x.table_compact])[
            [{"Metric": "Total", "Value": "$4.2M"}]
        ],
        x.table(style=[x.table_bordered, x.table_compact])[
            [{"Metric": "Growth", "Value": "+14%"}]
        ],
        gap=2,
    )

    summary_sheet = x.sheet("Summary")[
        x.vstack(
            x.row(style=[x.text_2xl, x.bold, x.text_blue])["Q3 Performance"],
            x.space(),
            summary_table,
            x.space(),
            stats,
            gap=1,
        )
    ]

    # Data sheet with larger table
    data_rows = []
    for i in range(100):
        data_rows.append(
            {
                "ID": i,
                "Product": f"Product {i}",
                "Category": ["A", "B", "C"][i % 3],
                "Price": 10.0 + i * 0.5,
                "Stock": 100 - i,
            }
        )

    data_table = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_compact],
    )[data_rows]

    data_sheet = x.sheet("Data")[
        x.vstack(
            x.row(style=[x.text_lg, x.bold])["Product Inventory"],
            x.space(),
            data_table,
            gap=1,
        )
    ]

    # Create multi-sheet workbook
    workbook = x.workbook()[summary_sheet, data_sheet]

    # Save to temporary file
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        workbook.save(tmp_path, engine=engine_name)
    finally:
        # Clean up
        Path(tmp_path).unlink(missing_ok=True)


def print_results(results: list[BenchmarkResult]) -> None:
    """Print formatted benchmark results to console.

    Args:
        results: List of benchmark results
    """
    print("\n" + "=" * 80)
    print("BENCHMARK RESULTS")
    print("=" * 80)

    # Group results by scenario
    scenarios: dict[str, list[BenchmarkResult]] = {}
    for result in results:
        if result.scenario not in scenarios:
            scenarios[result.scenario] = []
        scenarios[result.scenario].append(result)

    for scenario_name, scenario_results in scenarios.items():
        print(f"\n{scenario_name}")
        print("-" * 80)

        # Check if this is a big tables scenario
        if any(r.table_size is not None for r in scenario_results):
            # Print table format for big tables
            print(f"{'Size':<12} {'Engine':<15} {'Time (s)':<12} {'Memory (MB)':<15}")
            print("-" * 80)

            # Group by table size
            by_size: dict[int, list[BenchmarkResult]] = {}
            for result in scenario_results:
                if result.table_size is not None:
                    if result.table_size not in by_size:
                        by_size[result.table_size] = []
                    by_size[result.table_size].append(result)

            for size in sorted(by_size.keys()):
                size_results = by_size[size]
                for result in sorted(size_results, key=lambda r: r.engine):
                    status = "✓" if result.success else "✗"
                    time_str = f"{result.execution_time:.4f}" if result.success else "N/A"
                    mem_str = (
                        f"{result.memory_peak:.2f}" if result.success else "N/A"
                    )
                    print(
                        f"{status} {size:>8,} {result.engine:<15} {time_str:<12} {mem_str:<15}"
                    )
                    if not result.success and result.error:
                        print(f"    Error: {result.error}")

                # Compare engines for this size
                openpyxl_result = next(
                    (r for r in size_results if r.engine == "openpyxl" and r.success),
                    None,
                )
                xlsxwriter_result = next(
                    (
                        r
                        for r in size_results
                        if r.engine == "xlsxwriter" and r.success
                    ),
                    None,
                )

                if openpyxl_result and xlsxwriter_result:
                    time_ratio = (
                        openpyxl_result.execution_time / xlsxwriter_result.execution_time
                    )
                    mem_ratio = (
                        openpyxl_result.memory_peak / xlsxwriter_result.memory_peak
                    )
                    faster = (
                        "openpyxl"
                        if time_ratio < 1
                        else "xlsxwriter"
                        if time_ratio > 1
                        else "equal"
                    )
                    print(
                        f"    → {faster} is {max(time_ratio, 1/time_ratio):.2f}x faster, "
                        f"{'openpyxl' if mem_ratio > 1 else 'xlsxwriter'} uses {max(mem_ratio, 1/mem_ratio):.2f}x more memory"
                    )
        else:
            # Print simple format for other scenarios
            print(f"{'Engine':<15} {'Time (s)':<12} {'Memory (MB)':<15}")
            print("-" * 80)

            for result in sorted(scenario_results, key=lambda r: r.engine):
                status = "✓" if result.success else "✗"
                time_str = f"{result.execution_time:.4f}" if result.success else "N/A"
                mem_str = f"{result.memory_peak:.2f}" if result.success else "N/A"
                print(f"{status} {result.engine:<15} {time_str:<12} {mem_str:<15}")
                if not result.success and result.error:
                    print(f"    Error: {result.error}")

            # Compare engines
            openpyxl_result = next(
                (r for r in scenario_results if r.engine == "openpyxl" and r.success),
                None,
            )
            xlsxwriter_result = next(
                (
                    r
                    for r in scenario_results
                    if r.engine == "xlsxwriter" and r.success
                ),
                None,
            )

            if openpyxl_result and xlsxwriter_result:
                time_ratio = (
                    openpyxl_result.execution_time / xlsxwriter_result.execution_time
                )
                mem_ratio = (
                    openpyxl_result.memory_peak / xlsxwriter_result.memory_peak
                )
                faster = (
                    "openpyxl"
                    if time_ratio < 1
                    else "xlsxwriter"
                    if time_ratio > 1
                    else "equal"
                )
                print(
                    f"    → {faster} is {max(time_ratio, 1/time_ratio):.2f}x faster, "
                    f"{'openpyxl' if mem_ratio > 1 else 'xlsxwriter'} uses {max(mem_ratio, 1/mem_ratio):.2f}x more memory"
                )

    print("\n" + "=" * 80)


def save_results(results: list[BenchmarkResult], path: Path) -> None:
    """Save benchmark results to JSON file.

    Args:
        results: List of benchmark results
        path: Path to save JSON file
    """
    # Convert results to dictionaries
    data = []
    for result in results:
        data.append(
            {
                "scenario": result.scenario,
                "engine": result.engine,
                "table_size": result.table_size,
                "execution_time": result.execution_time,
                "memory_peak_mb": result.memory_peak,
                "memory_current_mb": result.memory_current,
                "success": result.success,
                "error": result.error,
            }
        )

    output = {
        "metadata": {
            "num_runs": NUM_RUNS,
            "table_sizes": TABLE_SIZES,
        },
        "results": data,
    }

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w") as f:
        json.dump(output, f, indent=2)

    print(f"\nResults saved to: {path.resolve()}")


def main() -> None:
    """Run all benchmarks and generate report."""
    print("Starting benchmark comparison of openpyxl vs xlsxwriter engines...")
    print(f"Running {NUM_RUNS} iterations per test for averaging")
    print("-" * 80)

    all_results: list[BenchmarkResult] = []

    # Benchmark big tables
    print("\n[1/3] Benchmarking Big Tables...")
    for size in TABLE_SIZES:
        print(f"  Testing {size:,} rows...")
        for engine_str in ["openpyxl", "xlsxwriter"]:
            engine = cast(EngineName, engine_str)
            result = run_benchmark(
                engine,
                "Big Tables",
                benchmark_big_tables,
                size,
                table_size=size,
            )
            all_results.append(result)
            status = "✓" if result.success else "✗"
            print(
                f"    {status} {engine}: {result.execution_time:.4f}s, "
                f"{result.memory_peak:.2f} MB"
            )

    # Benchmark simple layouts
    print("\n[2/3] Benchmarking Simple Layouts...")
    for engine_str in ["openpyxl", "xlsxwriter"]:
        engine = cast(EngineName, engine_str)
        print(f"  Testing {engine}...")
        result = run_benchmark(engine, "Simple Layouts", benchmark_simple_layouts)
        all_results.append(result)
        status = "✓" if result.success else "✗"
        print(
            f"    {status} {engine}: {result.execution_time:.4f}s, "
            f"{result.memory_peak:.2f} MB"
        )

    # Benchmark complex layouts
    print("\n[3/3] Benchmarking Complex Layouts...")
    for engine_str in ["openpyxl", "xlsxwriter"]:
        engine = cast(EngineName, engine_str)
        print(f"  Testing {engine}...")
        result = run_benchmark(engine, "Complex Layouts", benchmark_complex_layouts)
        all_results.append(result)
        status = "✓" if result.success else "✗"
        print(
            f"    {status} {engine}: {result.execution_time:.4f}s, "
            f"{result.memory_peak:.2f} MB"
        )

    # Print formatted results
    print_results(all_results)

    # Save results to file
    output_dir = _project_root / ".testing"
    output_path = output_dir / "benchmark_results.json"
    save_results(all_results, output_path)

    # Summary
    successful = sum(1 for r in all_results if r.success)
    total = len(all_results)
    print(f"\nCompleted: {successful}/{total} benchmarks succeeded")


if __name__ == "__main__":
    main()

