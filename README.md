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
    header=["Region", "Units", "Price"],
    header_style=[x.text_sm, x.text_gray, x.align_middle],
    style=[x.table_bordered, x.table_compact],
)[
    ["EMEA", 1200, 19.0],
    ["APAC", 900, 21.0],
    ["AMER", 1500, 18.5],
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

- `vstack(a, b, c, gap=1)` vertically stacks components with optional blank rows between them.
- `hstack(a, b, gap=1)` arranges components side by side with configurable column gaps.
- `space(rows=1, height=None)` inserts empty rows (optionally with a fixed height).

## Examples

- **Multi-sheet sales demo**: see `examples/multi_sheet_sales_demo.py`.
  - Run it to generate a showcase workbook:
```bash
python examples/multi_sheet_sales_demo.py
```
  - Output: `multi-sheet-sales-demo-output.xlsx` with sheets `Summary`, `Raw Data`, `Pipeline`, and `Glossary`, demonstrating tables, stacks, spacing, and utility styles.

## Types & ergonomics

- Modern Python with full type hints.
- Pure Python stack traces; easy to debug, script, and test.
- Deterministic rendering for stable diffs in CI.
