from pathlib import Path

import xpyxl as x


def stat_card(title: str, value: str, delta: str, *, positive: bool = True) -> x.Node:
    delta_style = x.text_green if positive else x.text_red
    return x.table(
        header_style=[x.text_sm, x.text_gray, x.text_left],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            {title: x.cell(style=[x.text_2xl, x.bold, x.text_black])[value]},
            {title: x.cell(style=[x.text_sm, delta_style])[delta]},
        ]
    ]


def summary_section() -> x.Node:
    headline = x.row(style=[x.text_3xl, x.bold, x.text_blue])["Q3 Revenue Performance"]

    cards = x.hstack(
        stat_card("Revenue", "$4.2M", "+14% vs LY", positive=True),
        stat_card("Win Rate", "52%", "+6 pts", positive=True),
        stat_card("Avg. Deal", "$18.9K", "-3% vs LY", positive=False),
        stat_card("Pipeline", "$6.8M", "+$1.1M QoQ", positive=True),
        gap=1,
    )

    regional_performance = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_banded, x.table_compact],
    )[
        [
            {
                "Region": "EMEA",
                "GM": x.cell(style=[x.text_right])["$1.6M"],
                "Units": x.cell(style=[x.text_right])[1200],
                "YoY": x.cell(style=[x.text_right])["+18%"],
            },
            {
                "Region": "APAC",
                "GM": x.cell(style=[x.text_right])["$1.1M"],
                "Units": x.cell(style=[x.text_right])[930],
                "YoY": x.cell(style=[x.text_right])["+9%"],
            },
            {
                "Region": "AMER",
                "GM": x.cell(style=[x.text_right])["$1.5M"],
                "Units": x.cell(style=[x.text_right])[1480],
                "YoY": x.cell(style=[x.text_right])["+6%"],
            },
        ]
    ]

    top_opportunities = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            {
                "Opportunity": "Atlas Renewals",
                "Stage": "Negotiation",
                "Owner": "S. Patel",
                "Value": x.cell(style=[x.text_right])["$420K"],
            },
            {
                "Opportunity": "Aurora Launch",
                "Stage": "Proposal",
                "Owner": "C. Rivers",
                "Value": x.cell(style=[x.text_right])["$310K"],
            },
            {
                "Opportunity": "Nimbus Edge",
                "Stage": "Discovery",
                "Owner": "L. Gomez",
                "Value": x.cell(style=[x.text_right])["$185K"],
            },
        ]
    ]

    key_updates = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered],
    )[
        [
            {
                "Key Updates": x.cell(style=[x.wrap])[
                    "• APAC backlog cleared; normalization expected by Q4."
                ]
            },
            {
                "Key Updates": x.cell(style=[x.wrap])[
                    "• Marketing launch for Nimbus Edge driving 23% lift in leads."
                ]
            },
            {
                "Key Updates": x.cell(style=[x.wrap])[
                    "• Supply constraints eased; lead times back under 4 weeks."
                ]
            },
        ]
    ]

    lower_row = x.hstack(
        regional_performance,
        top_opportunities,
        key_updates,
        gap=1,
    )

    return x.vstack(
        headline,
        cards,
        x.space(),
        lower_row,
        x.space(),
        x.row(style=[x.text_sm, x.text_gray])["Generated with xsxpy"],
        gap=1,
    )


def raw_data_sheet() -> x.SheetNode:
    data_table = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            {
                "Region": "EMEA",
                "Segment": "Enterprise",
                "Owner": "S. Patel",
                "Units": x.cell(style=[x.text_right])[620],
                "GM": x.cell(style=[x.text_right])["$820K"],
            },
            {
                "Region": "EMEA",
                "Segment": "Mid-Market",
                "Owner": "T. Kato",
                "Units": x.cell(style=[x.text_right])[580],
                "GM": x.cell(style=[x.text_right])["$780K"],
            },
            {
                "Region": "APAC",
                "Segment": "Enterprise",
                "Owner": "L. Gomez",
                "Units": x.cell(style=[x.text_right])[410],
                "GM": x.cell(style=[x.text_right])["$610K"],
            },
            {
                "Region": "APAC",
                "Segment": "SMB",
                "Owner": "K. Zhao",
                "Units": x.cell(style=[x.text_right])[520],
                "GM": x.cell(style=[x.text_right])["$490K"],
            },
            {
                "Region": "AMER",
                "Segment": "Enterprise",
                "Owner": "M. Shaw",
                "Units": x.cell(style=[x.text_right])[870],
                "GM": x.cell(style=[x.text_right])["$910K"],
            },
            {
                "Region": "AMER",
                "Segment": "SMB",
                "Owner": "C. Rivers",
                "Units": x.cell(style=[x.text_right])[610],
                "GM": x.cell(style=[x.text_right])["$590K"],
            },
        ]
    ]

    totals = x.row()[
        x.cell(style=[x.bold, x.text_gray])["Total"],
        "",
        "",
        x.cell(style=[x.bold, x.text_right])[3610],
        x.cell(style=[x.bold, x.text_right])["$4.2M"],
    ]

    notes = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered],
    )[
        [
            {
                "Notes": x.cell(style=[x.wrap])[
                    "Conversion benchmarks calculated using trailing 90 days."
                ]
            },
        ]
    ]

    return x.sheet("Raw Data")[
        x.vstack(
            x.row(style=[x.text_lg, x.bold])["Source Transactions"],
            x.space(),
            data_table,
            totals,
            x.space(),
            notes,
            gap=1,
        )
    ]


def pipeline_sheet() -> x.SheetNode:
    funnel = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            {
                "Stage": "Discovery",
                "Deals": x.cell(style=[x.text_right])[42],
                "Value": x.cell(style=[x.text_right])["$1.9M"],
            },
            {
                "Stage": "Qualification",
                "Deals": x.cell(style=[x.text_right])[33],
                "Value": x.cell(style=[x.text_right])["$1.4M"],
            },
            {
                "Stage": "Proposal",
                "Deals": x.cell(style=[x.text_right])[21],
                "Value": x.cell(style=[x.text_right])["$1.1M"],
            },
            {
                "Stage": "Negotiation",
                "Deals": x.cell(style=[x.text_right])[14],
                "Value": x.cell(style=[x.text_right])["$1.0M"],
            },
            {
                "Stage": "Closed Won",
                "Deals": x.cell(style=[x.text_right])[18],
                "Value": x.cell(style=[x.text_right])["$1.5M"],
            },
        ]
    ]

    forecast = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            {
                "Scenario": "Commit",
                "Probability": x.cell(style=[x.text_right])["75%"],
                "Forecast": x.cell(style=[x.text_right])["$3.2M"],
            },
            {
                "Scenario": "Best",
                "Probability": x.cell(style=[x.text_right])["50%"],
                "Forecast": x.cell(style=[x.text_right])["$4.5M"],
            },
            {
                "Scenario": "Upside",
                "Probability": x.cell(style=[x.text_right])["25%"],
                "Forecast": x.cell(style=[x.text_right])["$6.1M"],
            },
        ]
    ]

    team_heatmap = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_banded, x.table_compact],
    )[
        [
            {
                "Owner": "S. Patel",
                "Active Deals": x.cell(style=[x.text_right])[18],
                "Win Rate": x.cell(style=[x.text_right])["64%"],
            },
            {
                "Owner": "C. Rivers",
                "Active Deals": x.cell(style=[x.text_right])[15],
                "Win Rate": x.cell(style=[x.text_right])["58%"],
            },
            {
                "Owner": "L. Gomez",
                "Active Deals": x.cell(style=[x.text_right])[12],
                "Win Rate": x.cell(style=[x.text_right])["54%"],
            },
            {
                "Owner": "T. Kato",
                "Active Deals": x.cell(style=[x.text_right])[11],
                "Win Rate": x.cell(style=[x.text_right])["49%"],
            },
        ]
    ]

    return x.sheet("Pipeline")[
        x.vstack(
            x.row(style=[x.text_lg, x.bold])["Pipeline & Forecast"],
            x.space(),
            x.hstack(funnel, forecast, gap=2),
            x.space(),
            team_heatmap,
            gap=1,
        )
    ]


def glossary_sheet() -> x.SheetNode:
    utility_grid = x.table(
        header_style=[x.text_sm, x.text_gray],
        style=[x.table_bordered, x.table_compact],
    )[
        [
            {
                "Utility": "text_blue",
                "Description": "Accent headline color",
                "Example": "Q3 Revenue",
            },
            {
                "Utility": "bg_success",
                "Description": "Positive badge background",
                "Example": "+14% Revenue",
            },
            {
                "Utility": "wrap",
                "Description": "Wrap long text inside cells",
                "Example": "Supply constraints eased...",
            },
            {
                "Utility": "number_precision",
                "Description": "Two-decimal numeric format",
                "Example": "42100.00",
            },
        ]
    ]

    return x.sheet("Glossary")[
        x.vstack(
            x.row(style=[x.text_lg, x.bold])["Utility Cheatsheet"],
            x.space(),
            utility_grid,
            gap=1,
        )
    ]


def build_sample_workbook() -> x.Workbook:
    summary_sheet = x.sheet("Summary")[summary_section()]

    workbook = x.workbook()[
        summary_sheet,
        raw_data_sheet(),
        pipeline_sheet(),
        glossary_sheet(),
    ]

    return workbook


def main(output_path: Path | None = None) -> None:
    if output_path is None:
        output_path = Path("multi-sheet-sales-demo-output.xlsx")
    workbook = build_sample_workbook()
    workbook.save(output_path)
    print(f"Saved workbook to {output_path.resolve()}")


if __name__ == "__main__":
    main()
