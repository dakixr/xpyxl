from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, Literal

__all__ = [
    "Style",
    "combine_styles",
    "normalize_hex",
    "to_argb",
    "text_xs",
    "text_sm",
    "text_base",
    "text_lg",
    "text_xl",
    "text_2xl",
    "text_3xl",
    "bold",
    "italic",
    "mono",
    "muted",
    "text_muted",
    "text_primary",
    "text_white",
    "text_red",
    "text_green",
    "text_blue",
    "text_orange",
    "text_purple",
    "text_black",
    "text_gray",
    "text_left",
    "text_center",
    "text_right",
    "align_top",
    "align_middle",
    "align_bottom",
    "wrap",
    "nowrap",
    "wrap_shrink",
    "allow_overflow",
    "row_height",
    "row_width",
    "bg_red",
    "bg_primary",
    "bg_muted",
    "bg_success",
    "bg_warning",
    "bg_info",
    "border_all",
    "border_top",
    "border_bottom",
    "border_left",
    "border_right",
    "border_x",
    "border_y",
    "border_red",
    "border_green",
    "border_blue",
    "border_orange",
    "border_purple",
    "border_black",
    "border_gray",
    "border_white",
    "border_muted",
    "border_primary",
    "border_thin",
    "border_medium",
    "border_thick",
    "border_dashed",
    "border_dotted",
    "border_double",
    "border_none",
    "table_bordered",
    "table_banded",
    "table_compact",
    "number_comma",
    "number_precision",
    "percent",
    "currency_usd",
    "currency_eur",
    "date_short",
    "datetime_short",
    "time_short",
    "BorderStyleName",
    "BorderStyleLiteral",
]


BorderStyleLiteral = Literal[
    "dashDot",
    "dashDotDot",
    "dashed",
    "dotted",
    "double",
    "hair",
    "medium",
    "mediumDashDot",
    "mediumDashDotDot",
    "mediumDashed",
    "slantDashDot",
    "thick",
    "thin",
]
BorderStyleName = BorderStyleLiteral | Literal["none"]

DEFAULT_BORDER_STYLE_NAME: BorderStyleLiteral = "thin"


def normalize_hex(value: str) -> str:
    text = value.strip()
    if not text:
        raise ValueError("Color values cannot be empty")
    if text.startswith("#"):
        text = text[1:]
    if len(text) == 3:
        text = "".join(ch * 2 for ch in text)
    if len(text) != 6:
        raise ValueError(f"Expected 6 hex characters, got '{value}'")
    return "#" + text.upper()


def to_argb(value: str) -> str:
    rgb = normalize_hex(value)[1:]
    return "FF" + rgb


@dataclass(frozen=True)
class Style:
    name: str = ""
    font_name: str | None = None
    font_size: float | None = None
    font_size_delta: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    mono: bool | None = None
    text_color: str | None = None
    fill_color: str | None = None
    horizontal_align: str | None = None
    vertical_align: str | None = None
    indent: int | None = None
    wrap_text: bool | None = None
    shrink_to_fit: bool | None = None
    auto_width: bool | None = None
    row_height: float | None = None
    row_width: float | None = None
    number_format: str | None = None
    border: BorderStyleName | None = None
    border_color: str | None = None
    border_top: bool | None = None
    border_bottom: bool | None = None
    border_left: bool | None = None
    border_right: bool | None = None
    table_banded: bool | None = None
    table_bordered: bool | None = None
    table_compact: bool | None = None

    def merge(self, other: Style) -> Style:
        base_delta = 0.0 if self.font_size_delta is None else self.font_size_delta
        other_delta = 0.0 if other.font_size_delta is None else other.font_size_delta
        merged_delta = base_delta + other_delta
        delta_value = (
            merged_delta
            if (self.font_size_delta is not None or other.font_size_delta is not None)
            else None
        )
        return Style(
            name=other.name or self.name,
            font_name=other.font_name or self.font_name,
            font_size=other.font_size
            if other.font_size is not None
            else self.font_size,
            font_size_delta=delta_value,
            bold=other.bold if other.bold is not None else self.bold,
            italic=other.italic if other.italic is not None else self.italic,
            mono=other.mono if other.mono is not None else self.mono,
            text_color=other.text_color or self.text_color,
            fill_color=other.fill_color or self.fill_color,
            horizontal_align=other.horizontal_align or self.horizontal_align,
            vertical_align=other.vertical_align or self.vertical_align,
            indent=other.indent if other.indent is not None else self.indent,
            wrap_text=other.wrap_text
            if other.wrap_text is not None
            else self.wrap_text,
            shrink_to_fit=other.shrink_to_fit
            if other.shrink_to_fit is not None
            else self.shrink_to_fit,
            auto_width=other.auto_width
            if other.auto_width is not None
            else self.auto_width,
            row_height=other.row_height
            if other.row_height is not None
            else self.row_height,
            row_width=other.row_width
            if other.row_width is not None
            else self.row_width,
            number_format=other.number_format or self.number_format,
            border=other.border if other.border is not None else self.border,
            border_color=other.border_color
            if other.border_color is not None
            else self.border_color,
            border_top=other.border_top
            if other.border_top is not None
            else self.border_top,
            border_bottom=other.border_bottom
            if other.border_bottom is not None
            else self.border_bottom,
            border_left=other.border_left
            if other.border_left is not None
            else self.border_left,
            border_right=other.border_right
            if other.border_right is not None
            else self.border_right,
            table_banded=other.table_banded
            if other.table_banded is not None
            else self.table_banded,
            table_bordered=other.table_bordered
            if other.table_bordered is not None
            else self.table_bordered,
            table_compact=other.table_compact
            if other.table_compact is not None
            else self.table_compact,
        )


def combine_styles(styles: Iterable[Style], *, base: Style | None = None) -> Style:
    combined = base or Style()
    for style in styles:
        combined = combined.merge(style)
    return combined


def _style(name: str, **kwargs: Any) -> Style:
    return Style(name=name, **kwargs)


# fmt: off
text_xs = _style("text_xs", font_size_delta=-2.0)
text_sm = _style("text_sm", font_size_delta=-1.0)
text_base = _style("text_base", font_size_delta=0.0)
text_lg = _style("text_lg", font_size_delta=1.0)
text_xl = _style("text_xl", font_size_delta=4.0)
text_2xl = _style("text_2xl", font_size_delta=6.0)
text_3xl = _style("text_3xl", font_size_delta=8.0)

bold = _style("bold", bold=True)
italic = _style("italic", italic=True)
mono = _style("mono", mono=True)

muted = _style("muted", text_color=normalize_hex("#6B7280"))
text_muted = muted
text_primary = _style("text_primary", text_color=normalize_hex("#2563EB"))
text_white = _style("text_white", text_color=normalize_hex("#FFFFFF"))
text_red = _style("text_red", text_color=normalize_hex("#DC2626"))
text_green = _style("text_green", text_color=normalize_hex("#16A34A"))
text_blue = _style("text_blue", text_color=normalize_hex("#2563EB"))
text_orange = _style("text_orange", text_color=normalize_hex("#EA580C"))
text_purple = _style("text_purple", text_color=normalize_hex("#7C3AED"))
text_black = _style("text_black", text_color=normalize_hex("#111827"))
text_gray = _style("text_gray", text_color=normalize_hex("#4B5563"))

bg_red = _style("bg_red", fill_color=normalize_hex("#F04438"))
bg_primary = _style("bg_primary", fill_color=normalize_hex("#2563EB"))
bg_muted = _style("bg_muted", fill_color=normalize_hex("#6B7280"))
bg_success = _style("bg_success", fill_color=normalize_hex("#047857"))
bg_warning = _style("bg_warning", fill_color=normalize_hex("#B45309"))
bg_info = _style("bg_info", fill_color=normalize_hex("#0EA5E9"))

text_left = _style("text_left", horizontal_align="left")
text_center = _style("text_center", horizontal_align="center")
text_right = _style("text_right", horizontal_align="right")
align_top = _style("align_top", vertical_align="top")
align_middle = _style("align_middle", vertical_align="center")
align_bottom = _style("align_bottom", vertical_align="bottom")
wrap = _style("wrap", wrap_text=True)
nowrap = _style("nowrap", wrap_text=False)
wrap_shrink = _style("wrap_shrink", wrap_text=True, shrink_to_fit=True)
allow_overflow = _style("allow_overflow", auto_width=False)


def row_height(value: float) -> Style:
    if value <= 0:
        raise ValueError("Row height must be positive")
    return Style(name=f"row_height_{value:g}", row_height=value)


def row_width(value: float) -> Style:
    if value <= 0:
        raise ValueError("Row width must be positive")
    return Style(name=f"row_width_{value:g}", row_width=value)

border_all = _style(
    "border_all",
    border=DEFAULT_BORDER_STYLE_NAME,
    border_top=True,
    border_bottom=True,
    border_left=True,
    border_right=True,
)
border_top = _style("border_top", border=DEFAULT_BORDER_STYLE_NAME, border_top=True)
border_bottom = _style(
    "border_bottom", border=DEFAULT_BORDER_STYLE_NAME, border_bottom=True
)
border_left = _style("border_left", border=DEFAULT_BORDER_STYLE_NAME, border_left=True)
border_right = _style(
    "border_right", border=DEFAULT_BORDER_STYLE_NAME, border_right=True
)
border_x = _style(
    "border_x",
    border=DEFAULT_BORDER_STYLE_NAME,
    border_left=True,
    border_right=True,
)
border_y = _style(
    "border_y",
    border=DEFAULT_BORDER_STYLE_NAME,
    border_top=True,
    border_bottom=True,
)

border_thin = _style("border_thin", border="thin")
border_medium = _style("border_medium", border="medium")
border_thick = _style("border_thick", border="thick")
border_dashed = _style("border_dashed", border="dashed")
border_dotted = _style("border_dotted", border="dotted")
border_double = _style("border_double", border="double")
border_none = _style("border_none", border="none")

border_primary = _style("border_primary", border_color=normalize_hex("#2563EB"))
border_muted = _style("border_muted", border_color=normalize_hex("#6B7280"))
border_red = _style("border_red", border_color=normalize_hex("#DC2626"))
border_green = _style("border_green", border_color=normalize_hex("#16A34A"))
border_blue = _style("border_blue", border_color=normalize_hex("#2563EB"))
border_orange = _style("border_orange", border_color=normalize_hex("#EA580C"))
border_purple = _style("border_purple", border_color=normalize_hex("#7C3AED"))
border_black = _style("border_black", border_color=normalize_hex("#111827"))
border_gray = _style("border_gray", border_color=normalize_hex("#4B5563"))
border_white = _style("border_white", border_color=normalize_hex("#FFFFFF"))

table_bordered = _style("table_bordered", table_bordered=True)
table_banded = _style("table_banded", table_banded=True)
table_compact = _style("table_compact", table_compact=True)

number_comma = _style("number_comma", number_format="#,##0")
number_precision = _style("number_precision", number_format="#,##0.00")
percent = _style("percent", number_format="0.00%")
currency_usd = _style("currency_usd", number_format="$#,##0.00")
currency_eur = _style("currency_eur", number_format="€#,##0.00")
date_short = _style("date_short", number_format="yyyy-mm-dd")
datetime_short = _style("datetime_short", number_format="yyyy-mm-dd hh:mm")
time_short = _style("time_short", number_format="hh:mm")
