"""Generate a single-sheet workbook with a 1k-row table for performance/demo."""

from __future__ import annotations

from pathlib import Path

import xpyxl as x


def build_workbook() -> x.Workbook:
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

    sheet = x.sheet("Big Table")[
        x.row(style=[x.text_lg, x.bold])["1k-row table"],
        x.row(style=[x.text_sm, x.text_gray])["Useful for sanity/perf checks."],
        x.space(),
        table,
    ]

    return x.workbook()[sheet]


def main(output_path: Path | None = None) -> None:
    wb = build_workbook()
    if output_path is None:
        output_path = Path("big-table-demo-output.xlsx")
    wb.save(output_path)
    print(f"Saved {output_path.resolve()}")


if __name__ == "__main__":
    main()
