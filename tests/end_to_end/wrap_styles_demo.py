"""Showcase workbook for the wrapping utilities."""

from __future__ import annotations

import xpyxl as x

LONG_TEXT = (
    "This description is intentionally long so you can resize the column and observe "
    "how wrapping and shrink-to-fit behave in Excel."
)


def wrap_variants_gallery() -> x.Node:
    entries = [
        ("Default", [], "Excel auto wrapping based on column width"),
        ("x.wrap", [x.wrap], "Always wrap text to the next line"),
        (
            "x.nowrap",
            [x.nowrap],
            "Keep the text on one line even if parent containers set wrapping",
        ),
        (
            "x.wrap_shrink",
            [x.wrap_shrink],
            "Wrap text but also shrink the font to keep headings tidy",
        ),
        (
            "x.allow_overflow",
            [x.nowrap, x.allow_overflow],
            "Hold the column width and let text overflow",
        ),
    ]

    cards: list[x.Node] = []
    for label, styles, note in entries:
        cards.append(
            x.table(
                header_style=[x.text_sm, x.text_gray],
                style=[x.table_bordered, x.table_compact],
            )[
                [
                    {label: x.cell(style=styles)[LONG_TEXT]},
                    {label: x.cell(style=[x.text_sm, x.text_gray])[note]},
                ]
            ]
        )

    return x.hstack(*cards, gap=2)


def wrap_variants_section() -> x.Node:
    title = x.row(style=[x.text_lg, x.bold])["Cell-level utilities"]
    return x.vstack(
        title,
        wrap_variants_gallery(),
        style=[x.border_all, x.row_width(28)],
    )


def mix_and_match_section() -> x.Node:
    instructions = x.row(style=[x.text_sm, x.text_gray])[
        "Stack wrapping utilities at the row/column level too."
    ]

    wrap_stack = x.col(style=[x.wrap])[
        x.cell(style=[x.bold])["Row wrap"],
        LONG_TEXT,
        x.cell(style=[x.text_sm, x.text_gray])[
            "Row style enforces wrapping on every cell."
        ],
    ]
    nowrap_stack = x.col(style=[x.nowrap])[
        x.cell(style=[x.bold])["Row nowrap"],
        LONG_TEXT,
        x.cell(style=[x.text_sm, x.text_gray])["Keeps rows to a single line."],
    ]
    shrink_stack = x.col(style=[x.wrap_shrink])[
        x.cell(style=[x.bold])["Wrap & shrink"],
        LONG_TEXT,
        x.cell(style=[x.text_sm, x.text_gray])["Great for skinny annotation columns."],
    ]
    overflow_stack = x.col(style=[x.allow_overflow])[
        x.cell(style=[x.bold])["Allow overflow"],
        LONG_TEXT,
        x.cell(style=[x.text_sm, x.text_gray])[
            "Column width stays fixed; Excel shows spillover."
        ],
    ]

    return x.vstack(
        instructions,
        x.hstack(wrap_stack, nowrap_stack, shrink_stack, overflow_stack, gap=2),
        style=[x.border_all, x.row_width(28)],
    )


def build_workbook() -> x.SheetNode:
    return x.sheet("Wrapping")[
        x.row(style=[x.text_2xl, x.bold])["Wrapping Utilities"],
        x.row(style=[x.text_sm, x.text_gray])[
            "Resize the sample columns in Excel to see differences."
        ],
        x.space(),
        wrap_variants_section(),
        x.space(),
        mix_and_match_section(),
        x.space(),
        x.row(style=[x.text_sm, x.text_gray])["Generated with xpyxl"],
    ]
