"""Generate a single-sheet workbook with a 1k-row table for performance/demo."""

from __future__ import annotations

import xpyxl as x


def build_workbook() -> x.SheetNode:
    rows = []
    for idx in range(1_000):
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

    return x.sheet("Big Table")[
        x.row(style=[x.text_lg, x.bold])["1k-row table"],
        x.row(style=[x.text_sm, x.text_gray])["Useful for sanity/perf checks."],
        x.space(),
        table,
    ]
