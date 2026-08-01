"""Tests for the in-process ReportLab document exporter."""

from __future__ import annotations

import tempfile
from io import BytesIO
from pathlib import Path

import pytest
from PIL import Image, ImageChops, ImageStat

import xpyxl as x


def _fixture_workbook() -> x.Workbook:
    return x.workbook()[
        x.sheet("Summary", show_gridlines=True)[
            x.row(style=[x.bold, x.bg_primary, x.text_white, x.text_center])[
                "Region",
                "Revenue",
            ],
            x.row()["EMEA", x.cell(style=[x.currency_usd, x.text_right])[125000]],
            x.row()[
                x.cell(style=[x.bg_muted, x.italic], colspan=2)["Quarterly total"]
            ],
        ]
    ]


def test_reportlab_pdf_bytes() -> None:
    result = _fixture_workbook().export(renderer="reportlab")

    assert isinstance(result, bytes)
    assert result.startswith(b"%PDF")
    assert len(result) > 1_000


def test_reportlab_png_bytes_and_exact_canvas_size() -> None:
    result = _fixture_workbook().export(format="png", renderer="reportlab")

    assert isinstance(result, bytes)
    assert result.startswith(b"\x89PNG\r\n\x1a\n")
    with Image.open(BytesIO(result)) as image:
        assert image.size == (112, 72)
        assert image.convert("RGB").getpixel((2, 2)) == (37, 99, 235)


def test_reportlab_png_scale_controls_pixel_density() -> None:
    result = _fixture_workbook().export(
        format="png",
        renderer="reportlab",
        scale=2,
    )

    assert isinstance(result, bytes)
    with Image.open(BytesIO(result)) as image:
        assert image.size == (224, 144)


def test_reportlab_exports_multiple_pdf_sheets() -> None:
    workbook = x.workbook()[
        x.sheet("One")[x.row()["First"]],
        x.sheet("Two")[x.row()["Second"]],
    ]
    result = workbook.export(renderer="reportlab")

    assert isinstance(result, bytes)
    assert result.count(b"/Type /Page") == 3  # two pages plus the Pages node


def test_reportlab_writes_path_and_stream() -> None:
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "report.pdf"
        path_result = _fixture_workbook().export(path, renderer="reportlab")
        stream = BytesIO()
        stream_result = _fixture_workbook().export(stream, renderer="reportlab")

        assert path_result is None
        assert path.read_bytes().startswith(b"%PDF")
        assert stream_result is None
        assert stream.getvalue().startswith(b"%PDF")


def test_reportlab_and_chromium_pngs_have_matching_geometry() -> None:
    try:
        chromium = _fixture_workbook().export(format="png", renderer="chromium")
    except FileNotFoundError:
        pytest.skip("Chromium is not installed")
    reportlab = _fixture_workbook().export(format="png", renderer="reportlab")

    assert isinstance(chromium, bytes)
    assert isinstance(reportlab, bytes)
    with (
        Image.open(BytesIO(chromium)).convert("RGB") as chromium_image,
        Image.open(BytesIO(reportlab)).convert("RGB") as reportlab_image,
    ):
        assert chromium_image.size == reportlab_image.size
        difference = ImageChops.difference(chromium_image, reportlab_image)
        mean_difference = sum(ImageStat.Stat(difference).mean) / 3
        assert mean_difference < 35


def test_reportlab_png_requires_sheet_for_multiple_sheets() -> None:
    workbook = x.workbook()[
        x.sheet("One")[x.row()[1]],
        x.sheet("Two")[x.row()[2]],
    ]

    with pytest.raises(ValueError, match="requires sheet"):
        workbook.export(format="png", renderer="reportlab")


@pytest.mark.parametrize("scale", [0, -1, float("nan"), float("inf")])
def test_reportlab_rejects_invalid_scale(scale: float) -> None:
    with pytest.raises(ValueError, match="finite number greater than zero"):
        _fixture_workbook().export(
            format="png",
            renderer="reportlab",
            scale=scale,
        )
