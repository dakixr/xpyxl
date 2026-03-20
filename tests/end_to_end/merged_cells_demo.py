"""Manual visual demo for generated merged cells."""

from __future__ import annotations

import xpyxl as x


def summary_banner() -> x.Node:
    return x.row()[
        x.cell(
            style=[x.text_2xl, x.bold, x.text_center, x.bg_primary, x.text_white],
            colspan=4,
        )["Merged Cells Demo"]
    ]


def horizontal_examples() -> x.Node:
    return x.vstack(
        x.row(style=[x.text_lg, x.bold])["Horizontal merges"],
        x.row(style=[x.text_sm, x.text_gray])[
            "These should look like full-width section banners and grouped headers."
        ],
        x.row()[
            x.cell(style=[x.bg_warning, x.bold, x.text_center], colspan=4)[
                "Quarterly Summary"
            ]
        ],
        x.row(style=[x.text_sm, x.text_gray])[
            "Region",
            "Q1",
            "Q2",
            "Q3",
            "Q4",
        ],
        x.row()[
            "EMEA",
            1200,
            1280,
            x.cell(style=[x.bg_info, x.text_center], colspan=2)["Forecast merged"],
        ],
        style=[x.row_width(16)],
    )


def vertical_examples() -> x.Node:
    return x.vstack(
        x.row(style=[x.text_lg, x.bold])["Vertical merges"],
        x.row(style=[x.text_sm, x.text_gray])[
            "The left label should span two rows while the data flows around it."
        ],
        x.row()[
            x.cell(style=[x.bg_success, x.bold, x.align_middle], rowspan=2)[
                "North"
            ],
            "Jan",
            240,
            "On target",
        ],
        x.row()["Feb", 260, "Strong pipeline"],
        x.row()[
            x.cell(style=[x.bg_muted, x.bold, x.align_middle], rowspan=3)["South"],
            "Jan",
            180,
            "Needs review",
        ],
        x.row()["Feb", 195, "Stable"],
        x.row()["Mar", 210, "Recovered"],
        style=[x.row_width(18)],
    )


def layout_mix() -> x.Node:
    left = x.vstack(
        x.row(style=[x.text_lg, x.bold])["Stacked cards"],
        x.row()[
            x.cell(style=[x.bg_primary, x.text_white, x.text_center], colspan=2)[
                "Lead Funnel"
            ]
        ],
        x.row()[x.cell(style=[x.bold], rowspan=2)["Inbound"], "42"],
        x.row()["Follow-up"],
        x.row()[x.cell(style=[x.bold], rowspan=2)["Partner"], "18"],
        x.row()["Qualified"],
        style=[x.border_all, x.row_width(14)],
    )

    right = x.vstack(
        x.row(style=[x.text_lg, x.bold])["Matrix sample"],
        x.row()[
            x.cell(style=[x.bg_warning, x.bold, x.text_center], colspan=3)[
                "Coverage Plan"
            ]
        ],
        x.row()["Tier", "Owner", "Notes"],
        x.row()[
            x.cell(style=[x.align_middle, x.bg_info], rowspan=2)["A"],
            "Iris",
            "Strategic accounts",
        ],
        x.row()["Mateo", "Expansion focus"],
        style=[x.border_all, x.row_width(18)],
    )

    return x.hstack(left, x.space(rows=2), right, gap=0)


def build_workbook() -> x.SheetNode:
    return x.sheet("MergedCells", background_color="#F8FAFC")[
        summary_banner(),
        x.space(),
        horizontal_examples(),
        x.space(),
        vertical_examples(),
        x.space(),
        layout_mix(),
        x.space(),
        x.row(style=[x.text_sm, x.text_gray])[
            "Open the generated .xlsx or combined HTML in .testing to inspect merge alignment."
        ],
    ]
