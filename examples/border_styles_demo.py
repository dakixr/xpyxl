"""Showcase workbook for the border utility classes."""

from pathlib import Path

import xpyxl as x


def cell_border_section() -> x.Node:
    title = x.row(style=[x.text_lg, x.bold])["Cell utilities", "", ""]
    header = x.row(style=[x.text_sm, x.text_gray])["Utility", "Preview", "Notes"]

    variants = [
        ("x.border_all", [x.border_all, x.border_primary], "All sides + brand color"),
        ("x.border_top", [x.border_top, x.border_thick], "Top rule"),
        ("x.border_x", [x.border_x, x.border_green], "Left + right"),
        ("x.border_y", [x.border_y, x.border_dashed], "Top + bottom dashed"),
        ("x.border_bottom", [x.border_bottom, x.border_orange], "Underline style"),
    ]

    rows: list[x.Node] = []
    for label, styles, note in variants:
        preview = x.cell(style=[*styles, x.text_center, x.text_sm])["Sample"]
        rows.append(x.row()[label, preview, note])

    return x.vstack(title, header, *rows)


def row_border_section() -> x.Node:
    title = x.row(style=[x.text_lg, x.bold])["Row-level borders", "", ""]
    header = x.row(style=[x.text_sm, x.text_gray])["Utility", "Preview", "Notes"]

    variants = [
        (
            "x.border_y + x.border_muted",
            [x.border_y, x.border_muted],
            "Soft banded rows",
        ),
        (
            "x.border_top + x.border_bottom + x.border_thick",
            [x.border_top, x.border_bottom, x.border_thick],
            "Strong divider",
        ),
        (
            "x.border_y + x.border_dotted",
            [x.border_y, x.border_dotted],
            "Subtle dotted grid",
        ),
    ]

    rows: list[x.Node] = []
    for label, styles, note in variants:
        rows.append(x.row(style=styles)[label, "Row preview", note])

    return x.vstack(title, header, *rows)


def column_border_section() -> x.Node:
    header = x.row(style=[x.text_lg, x.bold])["Column-level borders"]
    details = x.row(style=[x.text_sm, x.text_gray])[
        "Style the column container once and every cell inherits."
    ]

    columns = x.hstack(
        x.col(style=[x.border_x, x.border_blue])[
            x.cell(style=[x.bold])["x.border_x"],
            "keeps",
            "left & right",
            "outlined",
        ],
        x.col(style=[x.border_y, x.border_red])[
            x.cell(style=[x.bold])["x.border_y"],
            "adds",
            "top & bottom",
            "per cell",
        ],
        x.col(style=[x.border_all, x.border_muted, x.border_thin])[
            x.cell(style=[x.bold])["x.border_all"],
            "acts like",
            "inline cards",
            "for stacks",
        ],
        gap=2,
    )

    return x.vstack(header, details, columns)


def build_workbook() -> x.Workbook:
    border_sheet = x.sheet("Borders")[
        x.row(style=[x.text_2xl, x.bold])["Border Styles Demo"],
        x.row(style=[x.text_sm, x.text_gray])[
            "Cell, row, and column outlines using utility classes."
        ],
        x.space(),
        cell_border_section(),
        x.space(),
        row_border_section(),
        x.space(),
        column_border_section(),
        x.space(),
        x.row(style=[x.text_sm, x.text_gray])["Generated with xpyxl"],
    ]
    return x.workbook()[border_sheet]


def main(output_path: Path | None = None) -> None:
    workbook = build_workbook()
    if output_path is None:
        output_path = Path("border-styles-demo-output.xlsx")
    workbook.save(output_path)
    print(f"Saved {output_path.resolve()}")


if __name__ == "__main__":
    main()
