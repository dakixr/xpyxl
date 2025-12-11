# xpyxl — Excel in Python

Compose polished spreadsheets with pure Python—no manual coordinates. You assemble rows/columns/cells; xpyxl handles layout, rendering, and styling with utility-style classes.

## Core ideas

- **Positionless composition:** Build sheets declaratively from `row`, `col`, `cell`, `table`, `vstack`, and `hstack`.
- **Composable styling:** Tailwind-inspired utilities (typography, colors, alignment, number formats) applied via `style=[...]`.
- **Deterministic rendering:** Pure-data trees compiled into `.xlsx` files with predictable output—ideal for tests and CI diffing.

## Installation

```bash
uv add xpyxl
pip install xpyxl
```

## Getting started

```python
import xpyxl as x

report = (
    x.workbook()[
        x.sheet("Summary")[
            x.row(style=[x.text_2xl, x.bold, x.text_blue])["Q3 Sales Overview"],
            x.row(style=[x.text_sm, x.text_gray])["Region", "Units", "Price"],
            x.row(style=[x.bg_primary, x.text_white, x.bold])["EMEA", 1200, 19.0],
            x.row()["APAC", 900, 21.0],
            x.row()["AMER", 1500, 18.5],
        ]
    ]
)

report.save("report.xlsx")
```

## Rendering Engines

xpyxl supports multiple rendering engines, allowing you to choose the best one for your needs:

- **openpyxl** (default): Full-featured engine with comprehensive Excel support. Best for complex workbooks with advanced formatting.
- **xlsxwriter**: Fast, memory-efficient engine. Ideal for large datasets and performance-critical applications.

### Using Different Engines

Specify the engine when saving:

```python
import xpyxl as x

workbook = x.workbook()[
    x.sheet("Data")[
        x.row(style=[x.bold])["Name", "Value"],
        x.row()["Item A", 100],
        x.row()["Item B", 200],
    ]
]

# Use openpyxl (default)
workbook.save("output-openpyxl.xlsx", engine="openpyxl")

# Use xlsxwriter
workbook.save("output-xlsxwriter.xlsx", engine="xlsxwriter")
```

## Importing existing sheets

You can pull in a static sheet from an existing Excel file without translating it into xpyxl nodes:

```
template = x.import_sheet("template.xlsx", "Cover")
report = x.workbook()[
    template,
    x.sheet("Data")[x.row()["Item", "Value"], x.row()["A", 1]],
]

report.save("with-template.xlsx", engine="openpyxl")
```

Copying sheets from existing workbooks is supported by the **openpyxl** engine. If you call `save(..., engine="xlsxwriter")` with imported sheets present, a `NotImplementedError` is raised.

### Saving without touching disk

`Workbook.save` accepts a filesystem path, any binary buffer (like `io.BytesIO()`), or no target at all to get the raw bytes back:

```python
import io
from pathlib import Path

buffer = io.BytesIO()
workbook.save(buffer, engine="xlsxwriter")  # writes into memory

raw_bytes = workbook.save(engine="openpyxl")  # returns bytes when no target is provided
Path("report.xlsx").write_bytes(raw_bytes)
```

Both engines produce equivalent Excel files, but may have subtle differences in:
- File size and memory usage
- Rendering performance
- Support for advanced Excel features

Choose **openpyxl** for maximum compatibility and feature support, or **xlsxwriter** when performance and memory efficiency are priorities.

### Performance Benchmarks

Benchmark results comparing the two rendering engines across different scenarios (averaged over 3 runs):

#### Big Tables

| Size    | Engine      | Time (s) | Memory (MB) |
|---------|-------------|----------|-------------|
| 100     | openpyxl    | 0.0977   | 0.68        |
| 100     | xlsxwriter  | 0.0325   | 0.50        |
| 1,000   | openpyxl    | 0.9354   | 3.10        |
| 1,000   | xlsxwriter  | 0.2565   | 2.04        |
| 10,000  | openpyxl    | 11.2455  | 31.26       |
| 10,000  | xlsxwriter  | 4.0539   | 17.81       |
| 50,000  | openpyxl    | 55.5080  | 156.33      |
| 50,000  | xlsxwriter  | 20.5686  | 94.43       |

**Summary**: xlsxwriter is **2.7-3.7x faster** and uses **1.4-1.8x less memory** for large tables. The openpyxl engine has been optimized with style object caching, providing **~15% performance improvement** compared to previous versions.

#### Simple Layouts

| Engine      | Time (s) | Memory (MB) |
|-------------|----------|-------------|
| openpyxl    | 0.0084   | 0.38        |
| xlsxwriter  | 0.0058   | 0.34        |

**Summary**: xlsxwriter is **1.5x faster** with similar memory usage.

#### Complex Layouts

| Engine      | Time (s) | Memory (MB) |
|-------------|----------|-------------|
| openpyxl    | 0.1462   | 0.69        |
| xlsxwriter  | 0.0454   | 0.52        |

**Summary**: xlsxwriter is **3.2x faster** and uses **1.3x less memory** for multi-sheet workbooks with styling.

Run benchmarks yourself:

```bash
uv run scripts/benchmark.py
```

## Primitives

```python
x.row(style=[x.bold, x.bg_warning])[1, 2, 3, 4, 5]
x.col(style=[x.italic])["a", "b", "c"]
x.cell(style=[x.text_green, x.number_precision])[42100]
```

- `row[...]` accepts any sequence (numbers, strings, dataclasses…)
- `col[...]` stacks values vertically
- `cell[...]` wraps a single scalar
- All primitives accept `style=[...]`

## Component: `table`

`x.table(...)` renders a header + body with optional style overrides. Combine with `vstack`/`hstack` for dashboards and reports.

```python
sales_table = x.table(
    header_style=[x.text_sm, x.text_gray, x.align_middle],
    style=[x.table_bordered, x.table_compact],
)[
    {"Region": "EMEA", "Units": 1200, "Price": 19.0},
    {"Region": "APAC", "Units": 900, "Price": 21.0},
    {"Region": "AMER", "Units": 1500, "Price": 18.5},
]

layout = x.vstack(
    x.row(style=[x.text_xl, x.bold])["Q3 Sales Overview"],
    x.space(),
    x.hstack(
        sales_table,
        x.cell(style=[x.text_sm, x.text_gray])["Generated with xpyxl"],
        gap=2,
    ),
)
```

Tables also accept pandas-friendly shapes:
- **records:** `table()[[{"region": "EMEA", "units": 1200}, ...]]` derives the header from dict keys (missing keys are filled with `None`).
- **dict of lists:** `table()[{"region": ["EMEA", "APAC"], "units": [1200, 900]}]` zips columns together (lengths must match).
Headers are inferred from your keys and default to bold text on a muted background; override with `header_style=[...]` when needed.

## Utility styles (non-exhaustive)

- **Typography:** `text_xs/_sm/_base/_lg/_xl/_2xl/_3xl`, `bold`, `italic`, `mono`
- **Text colors:** `text_red`, `text_green`, `text_blue`, `text_orange`, `text_purple`, `text_black`, `text_gray`
- **Backgrounds:** `bg_red`, `bg_primary`, `bg_muted`, `bg_success`, `bg_warning`, `bg_info`
- **Layout & alignment:** `text_left`, `text_center`, `text_right`, `align_top/middle/bottom`, `wrap`, `nowrap`, `wrap_shrink`, `allow_overflow`, `row_height(...)`, `row_width(...)`
- Use `allow_overflow` when you want to keep a column narrow and let the text spill into adjacent empty cells, `row_height(32)` to force a specific row height, and `row_width(12)` to pin a column width.
- **Borders:** `border_all`, `border_top`, `border_bottom`, `border_left`, `border_right`, `border_x`, `border_y`, `border_red`, `border_green`, `border_blue`, `border_orange`, `border_purple`, `border_black`, `border_gray`, `border_white`, `border_muted`, `border_primary`, `border_thin`, `border_medium`, `border_thick`, `border_dashed`, `border_dotted`, `border_double`, `border_none`
- **Tables:** `table_bordered`, `table_banded`, `table_compact`
- **Number/date formats:** `number_comma`, `number_precision`, `percent`, `currency_usd`, `currency_eur`, `date_short`, `datetime_short`, `time_short`

Border utilities can sit on individual cells or be applied at the row/column level for fast outlines.

Mix and match utilities freely—what you see is what you get.

## Layout helpers

- `vstack(a, b, c, gap=1, style=[x.border_all])` vertically stacks components with optional blank rows and shared styles (great for card-like borders).
- `hstack(a, b, gap=1, style=[x.border_all])` arranges components side by side with configurable column gaps and shared wrapper styles.
- `space(rows=1, height=None)` inserts empty rows (optionally with a fixed height) in a `vstack` or empty columns when dropped into an `hstack`.
- `sheet(name, background_color="#F8FAFC")` sets a sheet-wide background fill; the first ~200 rows and 80 columns are painted so the grid feels cohesive.

## Examples & Tests

The `tests/` directory contains example modules that demonstrate various features of xpyxl. Each module exports a `build_workbook()` function (or `build_sample_workbook()` for multi-sheet examples) that returns a `SheetNode` or list of `SheetNode` objects.

- **Multi-sheet sales demo**: `tests/multi_sheet_sales_demo.py` - showcases tables, stacks, spacing, and utility styles across multiple sheets.
- **Border styles demo**: `tests/border_styles_demo.py` - demonstrates border utilities at cell, row, and column levels.
- **Wrap styles demo**: `tests/wrap_styles_demo.py` - shows text wrapping and overflow utilities.
- **Row height demo**: `tests/row_height_demo.py` - examples of manual row height and width controls.
- **Big table demo**: `tests/big_table_demo.py` - performance test with a 1k-row table.

### Running Tests

Run all test modules to generate combined Excel files with both rendering engines:

```bash
uv run scripts/run_tests.py
```

This will:
- Collect sheets from all test modules
- Combine them into a single workbook
- Generate two output files in `.testing/`:
  - `combined-output-openpyxl.xlsx` (rendered with openpyxl engine)
  - `combined-output-xlsxwriter.xlsx` (rendered with xlsxwriter engine)

Each test module contributes one or more sheets to the combined workbook, allowing you to compare rendering output between engines.

## Types & ergonomics

- Modern Python with full type hints.
- Pure Python stack traces; easy to debug, script, and test.
- Deterministic rendering for stable diffs in CI.
