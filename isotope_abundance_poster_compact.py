#!/usr/bin/env python3
"""
Create a compact, single-page isotope natural-abundance poster from the
CIAAW-style isotope abundance CSV.

This version:
- uses ONLY the right-most CSV column: "Mean Range Abun"
- aims for a single-page 8.5 x 11 inch PORTRAIT layout by default
- keeps one isotope list column inside each element box
- shows only isotope mass numbers inside boxes (e.g., 40, not 40Ar)
- color-codes element text by isotope count
- reduces whitespace with tighter margins, gutters, and padding
- aligns isotope masses and abundance decimal points across element boxes

Usage:
    pip install pandas reportlab
    python isotope_abundance_poster_compact.py IsotopeAbundances_CIAAWdotOrg_2024.csv out.pdf

Optional examples:
    python isotope_abundance_poster_compact.py input.csv out.pdf --page 8.5x11 --columns 7
    python isotope_abundance_poster_compact.py input.csv out.pdf --full-precision
"""

from __future__ import annotations

import argparse
import math
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


# -----------------------------------------------------------------------------
# Defaults / style
# -----------------------------------------------------------------------------

DEFAULT_PAGE = "8.5x11"      # portrait by default
DEFAULT_COLUMNS = 7          # 0 = automatic
MARGIN_IN = 0.23
GUTTER_IN = 0.085
CARD_GAP_IN = 0.02
TITLE_H_IN = 0.44
FOOTER_H_IN = 0.18
TITLE_FONT = "Helvetica-Bold"
BODY_FONT = "Helvetica"
BODY_BOLD = "Helvetica-Bold"

DEFAULT_FOOTER = (
    "Isotopic abundances from: CIAAW Isotopic compositions of the elements, 2024; "
    "Available online at www.ciaaw.org\r"
    "For isotopes with a range reported, the mean of the range limits is shown.\n"
    "Poster layout by Dan Ohnemus (dan@uga.edu); code available at "
    "github.com/capecodnative/ICP-MS_Interferences_Poster"
)

PALETTE = {
    "ink": HexColor("#111111"),
    "muted": HexColor("#4A4A4A"),
    "border": HexColor("#808080"),
    "header_fill": HexColor("#F2F2F2"),
    "card_fill": colors.white,
    "mono": HexColor("#B23A2E"),
    "two": HexColor("#1F5EAA"),
    "three": HexColor("#238A45"),
}


@dataclass
class Isotope:
    mass: int
    mean_frac: Optional[float]
    mean_text: str


@dataclass
class ElementGroup:
    z: int
    symbol: str
    name: str
    isotopes: list[Isotope]


# -----------------------------------------------------------------------------
# Parsing / formatting
# -----------------------------------------------------------------------------


def parse_float(value) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    s = str(value).strip()
    if not s or s == "-" or s.lower() == "nan":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def format_percent_decimal(frac: Optional[float]) -> str:
    """Format source fraction as a percent string, avoiding scientific notation."""
    if frac is None:
        return "-"
    pct = frac * 100.0

    # Use enough decimals for small values but keep output compact.
    if abs(pct) >= 10:
        s = f"{pct:.4f}"
    elif abs(pct) >= 1:
        s = f"{pct:.4f}"
    elif abs(pct) >= 0.1:
        s = f"{pct:.5f}"
    elif abs(pct) >= 0.01:
        s = f"{pct:.6f}"
    else:
        s = f"{pct:.7f}"

    s = s.rstrip("0").rstrip(".")
    if "." not in s:
        s += ".0"
    return s


def format_percent_full_precision(source_frac: str) -> str:
    """Convert the source decimal fraction to percent without float rounding."""
    s = str(source_frac).strip()
    if not s or s == "-" or s.lower() == "nan":
        return "-"
    try:
        pct = format(Decimal(s) * Decimal(100), "f")
    except InvalidOperation:
        return "-"
    if "." in pct:
        pct = pct.rstrip("0").rstrip(".")
    return pct


def format_abundance(iso: Isotope, full_precision: bool) -> str:
    if full_precision:
        return format_percent_full_precision(iso.mean_text)
    return format_percent_decimal(iso.mean_frac)


def split_decimal_parts(text: str) -> tuple[str, str] | None:
    if text == "-":
        return None
    if "." in text:
        a, b = text.split(".", 1)
        return a, b
    return text, ""


def load_groups(csv_path: Path) -> list[ElementGroup]:
    df = pd.read_csv(csv_path, dtype=str).fillna("")
    required = ["Z", "M", "Symbol", "Element", "Mean Range Abun"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required CSV columns: {missing}")

    groups: list[ElementGroup] = []
    for (z, symbol, element), g in df.groupby(["Z", "Symbol", "Element"], sort=False):
        isotopes: list[Isotope] = []
        for _, row in g.iterrows():
            isotopes.append(
                Isotope(
                    mass=int(str(row["M"]).strip()),
                    mean_frac=parse_float(row["Mean Range Abun"]),
                    mean_text=str(row["Mean Range Abun"]).strip(),
                )
            )
        isotopes.sort(key=lambda x: x.mass)
        groups.append(ElementGroup(int(z), str(symbol), str(element).title(), isotopes))

    groups.sort(key=lambda x: x.z)
    return groups


# -----------------------------------------------------------------------------
# Page / layout helpers
# -----------------------------------------------------------------------------


def parse_page_size(page: str) -> tuple[float, float]:
    page = page.lower().strip()
    if "x" not in page:
        raise ValueError("Page must be like '8.5x11', '11x17', or '36x24'.")
    w, h = page.split("x", 1)
    return float(w) * inch, float(h) * inch


def card_height(group: ElementGroup, scale: float) -> float:
    pad_t = 1.7 * scale
    pad_b = 2.5 * scale
    header_h = 11.2 * scale
    line_h = 5.4 * scale
    return pad_t + header_h + len(group.isotopes) * line_h + pad_b


def split_into_ordered_columns(
    groups: list[ElementGroup],
    ncols: int,
    scale: float,
) -> list[list[ElementGroup]]:
    heights = [card_height(g, scale) + CARD_GAP_IN * inch for g in groups]
    n = len(groups)
    ncols = min(ncols, n)
    prefix = [0.0]
    for h in heights:
        prefix.append(prefix[-1] + h)

    def span_height(start: int, end: int) -> float:
        return prefix[end] - prefix[start]

    # Dynamic programming over contiguous element ranges. This preserves
    # increasing Z order while avoiding one very tall trailing column.
    dp: list[list[tuple[float, float] | None]] = [[None] * (n + 1) for _ in range(ncols + 1)]
    cuts: list[list[int | None]] = [[None] * (n + 1) for _ in range(ncols + 1)]
    dp[0][0] = (0.0, 0.0)

    for k in range(1, ncols + 1):
        for end in range(k, n + 1):
            best: tuple[float, float] | None = None
            best_start: int | None = None
            for start in range(k - 1, end):
                prev = dp[k - 1][start]
                if prev is None:
                    continue
                h = span_height(start, end)
                candidate = (max(prev[0], h), prev[1] + h * h)
                if best is None or candidate < best:
                    best = candidate
                    best_start = start
            dp[k][end] = best
            cuts[k][end] = best_start

    columns: list[list[ElementGroup]] = []
    end = n
    for k in range(ncols, 0, -1):
        start = cuts[k][end]
        assert start is not None
        columns.append(groups[start:end])
        end = start
    columns.reverse()
    return columns


@dataclass
class Layout:
    ncols: int
    scale: float
    cols: list[list[ElementGroup]]
    col_width: float
    margin: float
    title_h: float
    content_h: float
    mass_right_rel: float
    dot_x_rel: float


def data_column_positions(
    groups: list[ElementGroup],
    col_width: float,
    scale: float,
    full_precision: bool,
) -> tuple[float, float]:
    left_pad = 4.2 * scale
    right_pad = 5.2 * scale
    mass_font = 4.8 * scale
    value_font = 4.7 * scale
    max_mass_w = max(
        stringWidth(str(iso.mass), BODY_FONT, mass_font)
        for group in groups
        for iso in group.isotopes
    )
    values = [format_abundance(iso, full_precision) for group in groups for iso in group.isotopes]
    max_left_w, max_right_w = abundance_widths(values, value_font)
    dot_x_rel = col_width - right_pad - max_right_w
    mass_right_rel = max(left_pad + max_mass_w, dot_x_rel - max_left_w - 3.0 * scale)
    return mass_right_rel, dot_x_rel



def choose_layout(
    groups: list[ElementGroup],
    page_w: float,
    page_h: float,
    requested_cols: int,
    full_precision: bool,
) -> Layout:
    margin = MARGIN_IN * inch
    gutter = GUTTER_IN * inch
    title_h = TITLE_H_IN * inch
    footer_h = FOOTER_H_IN * inch
    content_h = page_h - 2 * margin - title_h - footer_h

    if requested_cols > 0:
        col_candidates: Iterable[int] = [requested_cols]
    else:
        # Default target is a compact 8.5 x 11 inch portrait poster.
        col_candidates = [7]

    fits: list[Layout] = []
    fallback: Optional[Layout] = None
    scales = [
        1.80, 1.75, 1.70, 1.65, 1.60, 1.55, 1.50, 1.45, 1.40, 1.35,
        1.30, 1.25, 1.20, 1.15, 1.10, 1.05, 1.00,
        0.95, 0.90, 0.86, 0.82, 0.78, 0.74, 0.70,
        0.66, 0.62, 0.58, 0.54, 0.50,
    ]

    for ncols in col_candidates:
        col_width = (page_w - 2 * margin - gutter * (ncols - 1)) / ncols
        for scale in scales:
            cols = split_into_ordered_columns(groups, ncols, scale)
            max_h = max(
                sum(card_height(g, scale) + CARD_GAP_IN * inch for g in col)
                for col in cols
            )
            mass_right_rel, dot_x_rel = data_column_positions(groups, col_width, scale, full_precision)
            layout = Layout(ncols, scale, cols, col_width, margin, title_h, content_h, mass_right_rel, dot_x_rel)
            fallback = layout
            if max_h <= content_h:
                fits.append(layout)
                break

    if not fits:
        assert fallback is not None
        return fallback

    # Highest scale first; if similar, prefer MORE columns (narrower boxes, less whitespace).
    best_scale = max(x.scale for x in fits)
    near_best = [x for x in fits if x.scale >= best_scale - 0.04]
    return max(near_best, key=lambda x: (x.ncols, x.scale))


# -----------------------------------------------------------------------------
# Drawing
# -----------------------------------------------------------------------------


def draw_fit_text(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    font: str,
    size: float,
    min_size: float,
    color,
):
    c.setFillColor(color)
    s = size
    while s > min_size and stringWidth(text, font, s) > max_width:
        s -= 0.2
    c.setFont(font, s)
    c.drawString(x, y, text)



def draw_decimal_aligned_value(
    c: canvas.Canvas,
    text: str,
    dot_x: float,
    y: float,
    size: float,
    color,
):
    c.setFillColor(color)
    c.setFont(BODY_FONT, size)
    parts = split_decimal_parts(text)
    if parts is None:
        c.drawRightString(dot_x + 7 * size / 6, y, "-")
        return

    left, right = parts
    left_w = stringWidth(left, BODY_FONT, size)
    c.drawString(dot_x - left_w, y, left)
    if not right:
        return
    c.drawString(dot_x, y, ".")
    c.drawString(dot_x + stringWidth(".", BODY_FONT, size), y, right)


def draw_right_aligned_segments(
    c: canvas.Canvas,
    segments: list[tuple[str, object]],
    right_x: float,
    y: float,
    font: str,
    size: float,
):
    total_w = sum(stringWidth(text, font, size) for text, _ in segments)
    x = right_x - total_w
    c.setFont(font, size)
    for text, color in segments:
        c.setFillColor(color)
        c.drawString(x, y, text)
        x += stringWidth(text, font, size)


def element_text_color(group: ElementGroup):
    n = len(group.isotopes)
    if n == 1:
        return PALETTE["mono"]
    if n == 2:
        return PALETTE["two"]
    if n == 3:
        return PALETTE["three"]
    return PALETTE["ink"]


def abundance_widths(values: list[str], font_size: float) -> tuple[float, float]:
    """Return widest text widths to the left and right of the decimal point."""
    max_left_w = 0.0
    max_right_w = 0.0
    for value in values:
        parts = split_decimal_parts(value)
        if parts is None:
            max_right_w = max(max_right_w, 7 * font_size / 6)
            continue
        left, right = parts
        max_left_w = max(max_left_w, stringWidth(left, BODY_FONT, font_size))
        max_right_w = max(
            max_right_w,
            stringWidth("." + right, BODY_FONT, font_size),
        )
    return max_left_w, max_right_w


def abundance_decimal_x(values: list[str], right_x: float, font_size: float) -> float:
    """Right-align abundance values by placing the decimal point from the widest suffix."""
    _, max_right_w = abundance_widths(values, font_size)
    return right_x - max_right_w



def draw_element_card(
    c: canvas.Canvas,
    group: ElementGroup,
    x: float,
    y_top: float,
    w: float,
    scale: float,
    mass_right_rel: float,
    dot_x_rel: float,
    full_precision: bool,
) -> float:
    h = card_height(group, scale)
    y = y_top - h

    radius = 2.2 * scale
    left_pad = 4.2 * scale
    right_pad = 5.2 * scale
    pad_t = 1.7 * scale
    header_h = 11.2 * scale
    line_h = 5.4 * scale
    text_color = element_text_color(group)

    # Card fill.
    c.setFillColor(PALETTE["card_fill"])
    c.roundRect(x, y, w, h, radius, fill=1, stroke=0)

    # Header band.
    c.setFillColor(PALETTE["header_fill"])
    c.roundRect(x, y + h - header_h, w, header_h, radius, fill=1, stroke=0)
    c.rect(x, y + h - header_h, w, header_h * 0.45, fill=1, stroke=0)

    # Stroke the full border last so the header fill cannot cover it.
    c.setLineWidth(0.35)
    c.setStrokeColor(PALETTE["border"])
    c.roundRect(x, y, w, h, radius, fill=0, stroke=1)

    # Header text.
    sym_size = 6.3 * scale
    name_size = 4.7 * scale
    z_size = 3.4 * scale
    header_text_y = y + h - header_h + 2.6 * scale
    z_right_x = x + w - right_pad

    c.setFillColor(text_color)
    c.setFont(BODY_FONT, z_size)
    z_text = str(group.z)
    c.drawString(x + left_pad, header_text_y - 0.4 * scale, z_text)
    z_w = stringWidth(z_text, BODY_FONT, z_size)

    c.setFont(BODY_BOLD, sym_size)
    sym_x = x + left_pad + z_w + 0.8 * scale
    c.drawString(sym_x, header_text_y, group.symbol)

    sym_w = stringWidth(group.symbol, BODY_BOLD, sym_size)
    name_x = sym_x + sym_w + 2.0 * scale
    name_w = max(1.0, z_right_x - name_x)
    draw_fit_text(
        c,
        group.name,
        name_x,
        y + h - header_h + 2.8 * scale,
        name_w,
        BODY_FONT,
        name_size,
        min_size=3.2 * scale,
        color=text_color,
    )

    # Row geometry.
    mass_font = 4.8 * scale
    value_font = 4.7 * scale
    row_y0 = y + h - header_h - pad_t - value_font
    values = [format_abundance(iso, full_precision) for iso in group.isotopes]
    mass_right_x = x + mass_right_rel
    dot_x = x + dot_x_rel

    # Draw rows.
    for i, (iso, value) in enumerate(zip(group.isotopes, values)):
        iy = row_y0 - i * line_h

        c.setFillColor(text_color)
        c.setFont(BODY_FONT, mass_font)
        c.drawRightString(mass_right_x, iy, str(iso.mass))

        draw_decimal_aligned_value(c, value, dot_x, iy, value_font, text_color)

    return h


# -----------------------------------------------------------------------------
# Main render
# -----------------------------------------------------------------------------


def make_poster(
    csv_path: Path,
    pdf_path: Path,
    page: str,
    columns: int,
    title: str,
    source_note: str,
    full_precision: bool = False,
) -> Layout:
    groups = load_groups(csv_path)
    page_w, page_h = parse_page_size(page)
    layout = choose_layout(groups, page_w, page_h, columns, full_precision)

    c = canvas.Canvas(str(pdf_path), pagesize=(page_w, page_h))
    c.setTitle(title)

    # Page background.
    c.setFillColor(colors.white)
    c.rect(0, 0, page_w, page_h, fill=1, stroke=0)

    margin = layout.margin
    gutter = GUTTER_IN * inch

    # Title.
    title_size = 24.0
    c.setFillColor(PALETTE["ink"])
    c.setFont(TITLE_FONT, title_size)
    c.drawString(margin, page_h - margin - title_size + 2, title)

    # Color key.
    key_size = 8.5
    key_y2 = page_h - margin - layout.title_h + 7.0
    key_y1 = key_y2 + key_size + 1.4
    key_line1 = [
        ("Mono-isotopic", PALETTE["mono"]),
        ("; ", PALETTE["muted"]),
        ("Two abundant isotopes;", PALETTE["two"]),
    ]
    key_line2 = [
        ("Three abundant isotopes", PALETTE["three"]),
        ("; ", PALETTE["muted"]),
        ("Four or more", PALETTE["ink"]),
    ]
    draw_right_aligned_segments(c, key_line1, page_w - margin, key_y1, BODY_FONT, key_size)
    draw_right_aligned_segments(c, key_line2, page_w - margin, key_y2, BODY_FONT, key_size)

    # Cards.
    content_top = page_h - margin - layout.title_h
    for j, col in enumerate(layout.cols):
        x = margin + j * (layout.col_width + gutter)
        y = content_top
        for group in col:
            used_h = draw_element_card(
                c,
                group,
                x,
                y,
                layout.col_width,
                layout.scale,
                layout.mass_right_rel,
                layout.dot_x_rel,
                full_precision,
            )
            y -= used_h + CARD_GAP_IN * inch

    # Footer.
    footer_lines = source_note.splitlines()
    footer_size = 7
    footer_y = margin * 0.55
    c.setFillColor(PALETTE["muted"])
    c.setFont(BODY_FONT, footer_size)
    for i, line in enumerate(footer_lines):
        c.drawString(margin, footer_y + (len(footer_lines) - 1 - i) * (footer_size + 1.4), line)

    c.showPage()
    c.save()
    return layout



def main():
    parser = argparse.ArgumentParser(description="Create a compact one-page isotope abundance poster PDF.")
    parser.add_argument("csv", type=Path, help="Input isotope abundance CSV")
    parser.add_argument("pdf", type=Path, help="Output PDF")
    parser.add_argument("--page", default=DEFAULT_PAGE, help="Page size in inches, e.g. 8.5x11, 11x17, 36x24")
    parser.add_argument("--columns", type=int, default=DEFAULT_COLUMNS, help="Poster columns; use 0 for automatic")
    parser.add_argument("--title", default="Natural Abundances of the Isotopes", help="Poster title")
    parser.add_argument(
        "--source-note",
        default=DEFAULT_FOOTER,
        help="Footer note",
    )
    parser.add_argument(
        "--full-precision",
        "--all-precision",
        action="store_true",
        help="Display all decimal precision from the source abundance values instead of compact rounded values.",
    )
    args = parser.parse_args()

    layout = make_poster(
        args.csv,
        args.pdf,
        args.page,
        args.columns,
        args.title,
        args.source_note,
        full_precision=args.full_precision,
    )
    precision = "full precision" if args.full_precision else "compact precision"
    print(f"Wrote {args.pdf} ({layout.ncols} columns, scale {layout.scale:g}, {precision})")


if __name__ == "__main__":
    main()
