"""Showcase manual row height/width utilities."""

from __future__ import annotations

from pathlib import Path

import xpyxl as x


def build_workbook() -> x.Workbook:
    sheet = x.sheet("Row Sizes")[
        x.row(style=[x.text_2xl, x.bold, x.row_height(36)])[
            "Manual Row Heights & Widths"
        ],
        x.row(style=[x.text_sm, x.text_gray])[
            "Use x.row_height(value) and x.row_width(value) anywhere styles are accepted."
        ],
        x.space(),
        x.row(style=[x.row_height(40)])[
            "Row-level height (40)",
            x.cell(style=[x.text_gray])["Applies to every cell in the row."],
        ],
        x.row(style=[x.row_height(20)])[
            "Compact row (20)",
            x.cell(style=[x.text_gray])["Good for dense tables."],
        ],
        x.row()[
            x.cell(style=[x.row_height(50), x.wrap])[
                "Cell-only height (50) with wrapping so the text stays inside the allotted space."
            ],
            x.cell()["Neighbor cell"],
        ],
        x.row(style=[x.row_height(28), x.wrap])[
            "Row + wrap",
            x.cell()[
                "Wrapping text respects the manual height when you want consistent card layouts."
            ],
        ],
        x.space(),
        x.row(style=[x.row_width(14)])[
            "Row width (14)",
            x.cell(style=[x.text_gray])["Each column touched by this row keeps width=14."],
        ],
        x.row()[
            x.cell(style=[x.row_width(25), x.wrap])[
                "Cell-only width (25) ensures this column stays wide enough for details."
            ],
            x.cell()["Neighbor"],
            x.cell()["Other"],
        ],
        x.row(style=[x.row_width(10), x.wrap])[
            "Narrow width",
            x.cell(style=[x.text_gray])["Combine with wrap to keep skinny card columns aligned."],
        ],
        x.row(style=[x.text_sm, x.text_gray])["Generated with xpyxl"],
    ]
    return x.workbook()[sheet]


def main() -> None:
    wb = build_workbook()
    output_path = Path("row-height-demo-output.xlsx")
    wb.save(output_path)
    print(f"Saved {output_path.resolve()}")


if __name__ == "__main__":
    main()
