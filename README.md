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

xpyxl supports three rendering engines:

- **hybrid** (default): Combines xlsxwriter speed for generated sheets with openpyxl for importing existing sheets. Best balance of speed and features.
- **openpyxl**: Full-featured with comprehensive Excel support. Best for complex workbooks with advanced formatting.
- **xlsxwriter**: Fast, memory-efficient. Ideal for large datasets and performance-critical applications. Does **not** support `import_sheet`.

Specify the engine when saving:

```python
workbook.save("output.xlsx")                       # hybrid (default)
workbook.save("output.xlsx", engine="openpyxl")    # full-featured
workbook.save("output.xlsx", engine="xlsxwriter")  # fast, generation only
```

`Workbook.save` accepts a filesystem path, any binary buffer (like `io.BytesIO()`), or no target to get raw bytes:

```python
import io
from pathlib import Path

buffer = io.BytesIO()
workbook.save(buffer, engine="xlsxwriter")

raw_bytes = workbook.save(engine="openpyxl")
Path("report.xlsx").write_bytes(raw_bytes)
```

## Importing existing sheets

Pull in a static sheet from an existing Excel file:

```python
report = x.workbook()[
    x.import_sheet("template.xlsx", "Cover"),
    x.sheet("Data")[x.row()["Item", "Value"], x.row()["A", 1]],
]
report.save("with-template.xlsx")  # uses hybrid by default (fast + imports)
```

Imported sheets preserve styles, merges, dimensions, freeze panes, filters, and other properties from the source file.

Engine support for `import_sheet`:

- **hybrid** (default): Combines xlsxwriter speed for generated sheets with openpyxl for importing. Best balance of speed and features.
- **openpyxl**: Full support with native fidelity.
- **xlsxwriter**: Does **not** support `import_sheet`. Use `hybrid` or `openpyxl` instead.

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

Tables accept polars/pandas-friendly shapes:
- **records:** `table()[[{"region": "EMEA", "units": 1200}, ...]]` derives header from dict keys
- **dict of lists:** `table()[{"region": ["EMEA", "APAC"], "units": [1200, 900]}]` zips columns together

## Utility styles

- **Typography:** `text_xs/_sm/_base/_lg/_xl/_2xl/_3xl`, `bold`, `italic`, `mono`
- **Text colors:** `text_red`, `text_green`, `text_blue`, `text_orange`, `text_purple`, `text_black`, `text_gray`
- **Backgrounds:** `bg_red`, `bg_primary`, `bg_muted`, `bg_success`, `bg_warning`, `bg_info`
- **Layout & alignment:** `text_left`, `text_center`, `text_right`, `align_top/middle/bottom`, `wrap`, `nowrap`, `wrap_shrink`, `allow_overflow`, `row_height(...)`, `row_width(...)`
- **Borders:** `border_all`, `border_top/bottom/left/right/x/y`, `border_red/green/blue/...`, `border_thin/medium/thick`, `border_dashed/dotted/double/none`
- **Tables:** `table_bordered`, `table_banded`, `table_compact`
- **Number/date formats:** `number_comma`, `number_precision`, `percent`, `currency_usd`, `currency_eur`, `date_short`, `datetime_short`, `time_short`

## Layout helpers

- `vstack(a, b, c, gap=1, style=[x.border_all])` vertically stacks components with optional blank rows
- `hstack(a, b, gap=1, style=[x.border_all])` arranges components side by side with configurable gaps
- `space(rows=1, height=None)` inserts empty rows/columns
- `sheet(name, background_color="#F8FAFC")` sets a sheet-wide background fill

## Types & ergonomics

- Modern Python with full type hints
- Pure Python stack traces; easy to debug, script, and test
- Deterministic rendering for stable diffs in CI
