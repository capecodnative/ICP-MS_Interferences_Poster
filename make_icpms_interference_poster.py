#!/usr/bin/env python3
"""
Generate a large-format ICP-MS interference poster PDF from a spreadsheet.

Expected spreadsheet columns, matching the sample workbook:
    m/z
    Analyte Element Ion
    Element Overlap Ion Contains
    Overlap ion
    Ion type

Usage:
    python make_icpms_interference_poster.py fullSpreadsheet.xlsx interference_poster.pdf --abundances IsotopeAbundances.csv

Dependencies:
    pip install pandas openpyxl reportlab
"""

from __future__ import annotations

import argparse
import math
import re
import textwrap
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A0, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# -------- User-editable settings --------
COL_MZ = "m/z"
COL_ANALYTE = "Analyte Element Ion"
COL_INTERFERING_ELEMENT = "Element Overlap Ion Contains"
COL_OVERLAP_ION = "Overlap ion"
COL_ION_TYPE = "Ion type"

FIRST_ELEMENT = "Li"
LAST_ELEMENT = "Pb"
PAGE_SIZE = landscape(A0)     # base size; page grows automatically if the matrix is larger
AUTO_GROW_PAGE = True
MIN_COL_WIDTH = 0.2 * inch
MIN_ROW_HEIGHT = 0.19 * inch
MARGIN = 0.08 * inch       # small edge margin; page is fit tightly to title/key/table
MIN_CELL_FONT = 2.4
MAX_CELL_FONT = 4.8
HEADER_FONT = 8.5
TITLE_FONT = 20
GRID_LINE_WIDTH = 0.12
SHOW_EMPTY_GRID = False
ADD_LEGEND = True
FIT_PAGE_TO_CONTENT = True    # trim whitespace by sizing PDF exactly around rendered content
PRUNE_EMPTY_ROWS_AND_COLUMNS = True
TWO_COLUMN_ENTRY_THRESHOLD = 6
CELL_TEXT_PAD = 1.5
MAX_AUTO_COL_WIDTH = 0.90 * inch
MAX_AUTO_ROW_HEIGHT = 0.95 * inch
CITATIONS_TEXT = "Citations: add citations here"
CITATIONS_BOX_WIDTH = 2.7 * inch

# Optional isotope-abundance lookup CSV. Values in the CSV are fractions; the poster displays percent.
DEFAULT_ABUNDANCE_CSV = "IsotopeAbundances.csv"
ABUNDANCE_SYMBOL_COL = "Symbol"
ABUNDANCE_MASS_COL = "M"
ABUNDANCE_VALUE_COL = "Mean Range Abun"
ABUNDANCE_HEADER = "Rel.<br/>Abund.<br/>(%)"
ABUNDANCE_FORMAT = ".6g"

# Ion-type colors from colortypes.xlsx. RGB values are 0-1, as provided.
# Keys are normalized by stripping whitespace and converting to lowercase.
ION_TYPE_RGB = {
    "argide": (1.0000, 0.0000, 0.0000),
    "oxide": (0.2539, 0.4102, 0.8789),
    "dioxide": (0.2539, 0.4102, 0.8789),
    "trioxide": (0.2539, 0.4102, 0.8789),
    "hydroxide": (0.1328, 0.5430, 0.1328),
    "hydride": (0.1328, 0.5430, 0.1328),
    "doubly charged": (0.5000, 0.0000, 0.5000),
    "other polyatomic": (0.5000, 0.0000, 0.5000),
    "doubly charged polyatomic": (0.5000, 0.0000, 0.5000),
    "nitride": (1.0000, 0.0781, 0.5742),
    "chloride": (1.0000, 0.0781, 0.5742),
    "elemental": (1.0000, 0.0781, 0.5742),
    "carbide": (1.0000, 0.0781, 0.5742),
    "sulfide": (1.0000, 0.0781, 0.5742),
    "plasma": (1.0000, 0.0781, 0.5742),
    "fluoride": (1.0000, 0.0781, 0.5742),
}
DEFAULT_ION_RGB = (0.0000, 0.0000, 0.0000)
# ----------------------------------------
# Atomic masses are used for Y-axis sorting. Values are standard atomic weights or representative mass numbers
# where conventional weights are interval/uncertain. Sorting Li through Pb is effectively atomic-number order,
# but masses are included because the requested ordering was by atomic mass.
ATOMIC_MASS = {
    "H": 1.008, "He": 4.0026, "Li": 6.94, "Be": 9.0122, "B": 10.81, "C": 12.011, "N": 14.007,
    "O": 15.999, "F": 18.998, "Ne": 20.180, "Na": 22.990, "Mg": 24.305, "Al": 26.982, "Si": 28.085,
    "P": 30.974, "S": 32.06, "Cl": 35.45, "Ar": 39.948, "K": 39.098, "Ca": 40.078, "Sc": 44.956,
    "Ti": 47.867, "V": 50.942, "Cr": 51.996, "Mn": 54.938, "Fe": 55.845, "Co": 58.933, "Ni": 58.693,
    "Cu": 63.546, "Zn": 65.38, "Ga": 69.723, "Ge": 72.630, "As": 74.922, "Se": 78.971, "Br": 79.904,
    "Kr": 83.798, "Rb": 85.468, "Sr": 87.62, "Y": 88.906, "Zr": 91.224, "Nb": 92.906, "Mo": 95.95,
    "Tc": 98.0, "Ru": 101.07, "Rh": 102.91, "Pd": 106.42, "Ag": 107.87, "Cd": 112.41, "In": 114.82,
    "Sn": 118.71, "Sb": 121.76, "Te": 127.60, "I": 126.90, "Xe": 131.29, "Cs": 132.91, "Ba": 137.33,
    "La": 138.91, "Ce": 140.12, "Pr": 140.91, "Nd": 144.24, "Pm": 145.0, "Sm": 150.36, "Eu": 151.96,
    "Gd": 157.25, "Tb": 158.93, "Dy": 162.50, "Ho": 164.93, "Er": 167.26, "Tm": 168.93, "Yb": 173.05,
    "Lu": 174.97, "Hf": 178.49, "Ta": 180.95, "W": 183.84, "Re": 186.21, "Os": 190.23, "Ir": 192.22,
    "Pt": 195.08, "Au": 196.97, "Hg": 200.59, "Tl": 204.38, "Pb": 207.2,
}
ELEMENTS = list(ATOMIC_MASS.keys())
ELEMENT_RE = re.compile(r"[A-Z][a-z]?")
ISO_ELEMENT_RE = re.compile(r"(\d+)([A-Z][a-z]?)")
TRAILING_CHARGE_RE = re.compile(r"(\+\+|--|\+|-|\d+[+-])$")

@dataclass(frozen=True)
class XKey:
    element: str
    mz: float


def element_range(first: str, last: str) -> List[str]:
    i0, i1 = ELEMENTS.index(first), ELEMENTS.index(last)
    return ELEMENTS[i0:i1 + 1]


def clean_text(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def mz_to_float(x) -> float:
    s = clean_text(x)
    try:
        return float(s)
    except ValueError:
        m = re.search(r"\d+(?:\.\d+)?", s)
        return float(m.group(0)) if m else math.nan


def label_mz(x: float) -> str:
    if math.isnan(x):
        return ""
    return str(int(x)) if abs(x - round(x)) < 1e-9 else f"{x:g}"


def mass_number_from_mz(x: float) -> int | None:
    """Return a display mass number from m/z, or None if m/z is not finite."""
    if math.isnan(x):
        return None
    return int(round(x))


def format_isotope_label(element: str, mz: float) -> str:
    """Return ReportLab markup such as superscripted 84Rb."""
    mass = mass_number_from_mz(mz)
    if mass is None:
        return xml_escape(element)
    return f"<super>{mass}</super>{xml_escape(element)}"


def normalize_charge(token: str) -> str:
    if token == "++":
        return "2+"
    if token == "--":
        return "2-"
    return token


def xml_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def normalize_ion_type(typ: str) -> str:
    return clean_text(typ).lower()


def rgb_to_hex(rgb: Tuple[float, float, float]) -> str:
    r, g, b = rgb
    return "#" + "".join(f"{max(0, min(255, round(v * 255))):02X}" for v in (r, g, b))


def ion_type_hex(typ: str) -> str:
    return rgb_to_hex(ION_TYPE_RGB.get(normalize_ion_type(typ), DEFAULT_ION_RGB))


def format_ion_reportlab(ion: str) -> str:
    """Return ReportLab Paragraph markup with isotope masses and charges superscripted."""
    s = clean_text(ion)
    if not s:
        return ""

    # Preserve descriptive suffixes like " wing" while formatting the chemical-looking prefix.
    m = re.match(r"^([0-9A-Za-z+\-]+)(.*)$", s)
    formula, suffix = (m.group(1), m.group(2)) if m else (s, "")

    charge = ""
    cm = TRAILING_CHARGE_RE.search(formula)
    if cm:
        charge = normalize_charge(cm.group(1))
        formula = formula[:cm.start()]

    parts: List[str] = []
    pos = 0
    for im in ISO_ELEMENT_RE.finditer(formula):
        if im.start() > pos:
            parts.append(xml_escape(formula[pos:im.start()]))
        mass, elem = im.groups()
        parts.append(f"<super>{xml_escape(mass)}</super>{xml_escape(elem)}")
        pos = im.end()
    parts.append(xml_escape(formula[pos:]))
    if charge:
        parts.append(f"<super>{xml_escape(charge)}</super>")
    parts.append(xml_escape(suffix))
    return "".join(parts)


def read_data(path: str, sheet_name=None) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    if isinstance(df, dict):
        # If sheet_name=None returns dict, take first sheet.
        df = next(iter(df.values()))
    missing = [c for c in [COL_MZ, COL_ANALYTE, COL_INTERFERING_ELEMENT, COL_OVERLAP_ION] if c not in df.columns]
    if missing:
        raise ValueError(f"Missing expected column(s): {missing}. Found columns: {list(df.columns)}")
    return df


def format_abundance_percent(value) -> str:
    """Format fractional isotope abundance as a percent string for display."""
    if pd.isna(value):
        return ""
    try:
        pct = float(value) * 100
    except (TypeError, ValueError):
        return clean_text(value)
    return f"{format(pct, ABUNDANCE_FORMAT)}%"


def read_abundance_lookup(path: str | None) -> Dict[Tuple[str, int], str]:
    """Read isotope relative abundances keyed by (element symbol, mass number)."""
    if not path:
        return {}

    csv_path = Path(path)
    if not csv_path.exists():
        # Also try next to this script, useful when DEFAULT_ABUNDANCE_CSV is used.
        csv_path = Path(__file__).resolve().parent / path
    if not csv_path.exists():
        raise FileNotFoundError(f"Isotope-abundance CSV not found: {path}")

    df = pd.read_csv(csv_path)
    required = [ABUNDANCE_SYMBOL_COL, ABUNDANCE_MASS_COL, ABUNDANCE_VALUE_COL]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing isotope-abundance CSV column(s): {missing}. Found columns: {list(df.columns)}")

    lookup: Dict[Tuple[str, int], str] = {}
    for _, r in df.iterrows():
        sym = clean_text(r[ABUNDANCE_SYMBOL_COL])
        mass = r[ABUNDANCE_MASS_COL]
        val = r[ABUNDANCE_VALUE_COL]
        if not sym or pd.isna(mass) or pd.isna(val):
            continue
        try:
            mass_i = int(round(float(mass)))
        except ValueError:
            continue
        lookup[(sym, mass_i)] = format_abundance_percent(val)
    return lookup


def abundance_for_xkey(xk: XKey, abundance_lookup: Dict[Tuple[str, int], str]) -> str:
    mass = mass_number_from_mz(xk.mz)
    if mass is None:
        return ""
    return abundance_lookup.get((xk.element, mass), "")


def build_matrix(df: pd.DataFrame) -> Tuple[List[XKey], List[str], Dict[Tuple[str, XKey], List[Tuple[str, str]]]]:
    elems = set(element_range(FIRST_ELEMENT, LAST_ELEMENT))
    df = df.copy()
    df["_mz"] = df[COL_MZ].map(mz_to_float)
    df["_analyte"] = df[COL_ANALYTE].map(clean_text)
    df["_interfering_elem"] = df[COL_INTERFERING_ELEMENT].map(clean_text)
    df["_overlap"] = df[COL_OVERLAP_ION].map(clean_text)
    df["_type"] = df[COL_ION_TYPE].map(clean_text) if COL_ION_TYPE in df.columns else ""

    df = df[df["_analyte"].isin(elems) & df["_interfering_elem"].isin(elems)]
    records = df[["_mz", "_analyte", "_interfering_elem", "_overlap", "_type"]].to_dict("records")
    xkeys = sorted({XKey(r["_analyte"], r["_mz"]) for r in records if not math.isnan(r["_mz"])},
                   key=lambda k: (ELEMENTS.index(k.element), k.mz))
    ykeys = sorted(elems, key=lambda e: ATOMIC_MASS[e])

    matrix: Dict[Tuple[str, XKey], List[Tuple[str, str]]] = defaultdict(list)
    for r in records:
        if math.isnan(r["_mz"]):
            continue
        xk = XKey(r["_analyte"], r["_mz"])
        matrix[(r["_interfering_elem"], xk)].append((r["_overlap"], r["_type"]))

    # Deduplicate while preserving order.
    for key, vals in list(matrix.items()):
        seen = set()
        out = []
        for ion, typ in vals:
            ident = (ion, typ)
            if ident not in seen:
                seen.add(ident)
                out.append((ion, typ))
        matrix[key] = out
    return xkeys, ykeys, matrix


def split_items_for_cell(items: List[Tuple[str, str]]) -> Tuple[List[Tuple[str, str]], List[Tuple[str, str]]]:
    """Return left/right item lists. Right is empty unless the cell is crowded."""
    if len(items) < TWO_COLUMN_ENTRY_THRESHOLD:
        return items, []
    cut = math.ceil(len(items) / 2)
    return items[:cut], items[cut:]


def make_item_paragraph(ion: str, typ: str, style: ParagraphStyle) -> Paragraph:
    ion_markup = format_ion_reportlab(ion)
    return Paragraph(f'<font color="{ion_type_hex(typ)}">{ion_markup}</font>', style)


def make_cell_flowable(items: List[Tuple[str, str]], style: ParagraphStyle, col_width: float) -> object:
    if not items:
        return ""
    left, right = split_items_for_cell(items)
    if not right:
        return Paragraph("<br/>".join(
            f'<font color="{ion_type_hex(typ)}">{format_ion_reportlab(ion)}</font>'
            for ion, typ in left
        ), style)

    # Nested two-column table for crowded cells. No internal grid: only the main matrix grid is drawn.
    left_paras = [make_item_paragraph(ion, typ, style) for ion, typ in left]
    right_paras = [make_item_paragraph(ion, typ, style) for ion, typ in right]
    nested = Table(
        [[left_paras, right_paras]],
        colWidths=[max(6, (col_width - 2 * CELL_TEXT_PAD) / 2)] * 2,
    )
    nested.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0.2),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return nested


def item_plain_len(ion: str) -> int:
    return len(clean_text(ion))


def needed_col_width(xk: XKey, ykeys: List[str], matrix: Dict[Tuple[str, XKey], List[Tuple[str, str]]], font_size: float) -> float:
    header_len = max(len(xk.element), len("m/z " + label_mz(xk.mz)))
    max_len = header_len
    crowded = False
    for y in ykeys:
        items = matrix.get((y, xk), [])
        if not items:
            continue
        left, right = split_items_for_cell(items)
        if right:
            crowded = True
            longest_side = max(
                item_plain_len(max([ion for ion, _ in left], key=len, default="")),
                item_plain_len(max([ion for ion, _ in right], key=len, default="")),
            )
            max_len = max(max_len, 2 * longest_side + 60)
        else:
            max_len = max(max_len, *(item_plain_len(ion) for ion, _ in items))
    est = max(MIN_COL_WIDTH, max_len * font_size * 1 + 2 * CELL_TEXT_PAD + (8 if crowded else 0))
    return min(MAX_AUTO_COL_WIDTH, est)


def needed_row_height(y: str, xkeys: List[XKey], matrix: Dict[Tuple[str, XKey], List[Tuple[str, str]]], font_size: float) -> float:
    max_lines = 1
    for xk in xkeys:
        items = matrix.get((y, xk), [])
        if not items:
            continue
        left, right = split_items_for_cell(items)
        max_lines = max(max_lines, max(len(left), len(right)))
    est = max(MIN_ROW_HEIGHT, max_lines * (font_size + 1.1) + 2 * CELL_TEXT_PAD)
    return min(MAX_AUTO_ROW_HEIGHT, est)


def needed_flipped_col_width(y: str, xkeys: List[XKey], matrix: Dict[Tuple[str, XKey], List[Tuple[str, str]]], font_size: float) -> float:
    """Column width after flipping axes: columns are interfering elements."""
    max_len = len(y)
    crowded = False
    for xk in xkeys:
        items = matrix.get((y, xk), [])
        if not items:
            continue
        left, right = split_items_for_cell(items)
        if right:
            crowded = True
            longest_side = max(
                item_plain_len(max([ion for ion, _ in left], key=len, default="")),
                item_plain_len(max([ion for ion, _ in right], key=len, default="")),
            )
            max_len = max(max_len, 2 * longest_side + 6)
        else:
            max_len = max(max_len, *(item_plain_len(ion) for ion, _ in items))
    est = max(MIN_COL_WIDTH, max_len * font_size * 0.58 + 2 * CELL_TEXT_PAD + (8 if crowded else 0))
    return min(MAX_AUTO_COL_WIDTH, est)


def needed_flipped_row_height(xk: XKey, ykeys: List[str], matrix: Dict[Tuple[str, XKey], List[Tuple[str, str]]], font_size: float) -> float:
    """Row height after flipping axes: rows are analyte element/mz pairs."""
    max_lines = 2  # analyte row label is two lines: element + m/z
    for y in ykeys:
        items = matrix.get((y, xk), [])
        if not items:
            continue
        left, right = split_items_for_cell(items)
        max_lines = max(max_lines, max(len(left), len(right)))
    est = max(MIN_ROW_HEIGHT, max_lines * (font_size + 1.1) + 2 * CELL_TEXT_PAD)
    return min(MAX_AUTO_ROW_HEIGHT, est)


def prune_empty_axes(xkeys: List[XKey], ykeys: List[str], matrix: Dict[Tuple[str, XKey], List[Tuple[str, str]]]) -> Tuple[List[XKey], List[str]]:
    if not PRUNE_EMPTY_ROWS_AND_COLUMNS:
        return xkeys, ykeys
    valid_x = [xk for xk in xkeys if any(matrix.get((y, xk)) for y in ykeys)]
    valid_y = [y for y in ykeys if any(matrix.get((y, xk)) for xk in xkeys)]
    return valid_x, valid_y


def build_pdf(df: pd.DataFrame, out_pdf: str, abundance_csv: str | None = None) -> None:
    xkeys_all, ykeys_all, matrix = build_matrix(df)
    if not xkeys_all:
        raise ValueError("No Li-through-Pb analyte rows found after filtering.")

    xkeys, ykeys = prune_empty_axes(xkeys_all, ykeys_all, matrix)
    abundance_lookup = read_abundance_lookup(abundance_csv)
    if not xkeys or not ykeys:
        raise ValueError("No populated Li-through-Pb interference cells found after pruning empty rows/columns.")

    # Axes are intentionally flipped relative to the original script:
    #   rows    = analyte element + m/z
    #   columns = interfering element
    base_w, base_h = PAGE_SIZE
    isotope_header_w = 0.58 * inch
    abundance_col_w = 0.68 * inch
    row_header_w = isotope_header_w + abundance_col_w
    col_header_h = 0.46 * inch
    citation_lines = max(1, math.ceil(len(clean_text(CITATIONS_TEXT)) / 55))
    key_block_h = max(0.30 * inch, citation_lines * 0.13 * inch)
    title_block_h = 0.50 * inch if not ADD_LEGEND else 0.54 * inch + key_block_h
    n_cols, n_rows = len(ykeys), len(xkeys)

    # Estimate a compact font size from the base page, then resize rows/columns from actual populated content.
    usable_w_base = base_w - 2 * MARGIN - row_header_w
    usable_h_base = base_h - 2 * MARGIN - col_header_h - title_block_h
    nominal_cell_w = max(MIN_COL_WIDTH, usable_w_base / max(n_cols, 1))
    nominal_cell_h = max(MIN_ROW_HEIGHT, usable_h_base / max(n_rows, 1))
    cell_font = max(MIN_CELL_FONT, min(MAX_CELL_FONT, min(nominal_cell_h * 0.25, nominal_cell_w * 0.22)))

    col_widths_data = [needed_flipped_col_width(y, xkeys, matrix, cell_font) for y in ykeys]
    row_heights_data = [needed_flipped_row_height(xk, ykeys, matrix, cell_font) for xk in xkeys]

    matrix_w = row_header_w + sum(col_widths_data)
    matrix_h = col_header_h + sum(row_heights_data)

    if FIT_PAGE_TO_CONTENT:
        # Tight page: no unused A0 whitespace. The PDF page is just large enough for the
        # title, color key, and compacted table plus the small MARGIN above.
        page_w = 2 * MARGIN + matrix_w
        page_h = 2 * MARGIN + title_block_h + matrix_h
    elif AUTO_GROW_PAGE:
        page_w = max(base_w, 2 * MARGIN + matrix_w)
        page_h = max(base_h, 2 * MARGIN + title_block_h + matrix_h)
    else:
        page_w, page_h = base_w, base_h

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Title"], fontSize=TITLE_FONT, leading=TITLE_FONT + 3, alignment=TA_CENTER)
    hdr_style = ParagraphStyle("Header", parent=styles["Normal"], fontSize=HEADER_FONT, leading=HEADER_FONT + 2, alignment=TA_CENTER)
    row_hdr_style = ParagraphStyle("RowHeader", parent=hdr_style, fontSize=HEADER_FONT + 1.5, leading=HEADER_FONT + 3)
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=cell_font, leading=cell_font + 0.9, alignment=TA_LEFT)
    note_style = ParagraphStyle("Note", parent=styles["Normal"], fontSize=7.5, leading=9, alignment=TA_LEFT)
    citation_style = ParagraphStyle("Citation", parent=note_style, alignment=TA_RIGHT)

    title = Paragraph("ICP-MS Interference/Overlap Matrix", title_style)

    key_tbl = None
    if ADD_LEGEND:
        legend_bits = []
        groups: Dict[str, List[str]] = defaultdict(list)
        for typ, rgb in ION_TYPE_RGB.items():
            groups[rgb_to_hex(rgb)].append(typ)
        for hex_color, type_names in groups.items():
            label = ", ".join(type_names)
            legend_bits.append(f'<font color="{hex_color}">■</font> {xml_escape(label)}')
        legend_para = Paragraph("Ion-type colors: " + "; ".join(legend_bits), note_style)
        citation_para = Paragraph(xml_escape(CITATIONS_TEXT), citation_style)
        citation_w = min(CITATIONS_BOX_WIDTH, max(1.2 * inch, matrix_w * 0.40))
        legend_w = max(1.2 * inch, matrix_w - citation_w)
        key_tbl = Table([[legend_para, citation_para]], colWidths=[legend_w, citation_w])
        key_tbl.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))

    table_data: List[List[object]] = []

    # Column headers are now interfering elements.
    col_header = [Paragraph("Analyte<br/>isotope", hdr_style), Paragraph(ABUNDANCE_HEADER, hdr_style)]
    for y in ykeys:
        col_header.append(Paragraph(xml_escape(y), hdr_style))
    table_data.append(col_header)

    # Row headers are now analyte isotope labels plus a relative-abundance column.
    for xk in xkeys:
        row: List[object] = [
            Paragraph(format_isotope_label(xk.element, xk.mz), row_hdr_style),
            Paragraph(xml_escape(abundance_for_xkey(xk, abundance_lookup)), row_hdr_style),
        ]
        for j, y in enumerate(ykeys):
            row.append(make_cell_flowable(matrix.get((y, xk), []), cell_style, col_widths_data[j]))
        table_data.append(row)

    col_widths = [isotope_header_w, abundance_col_w] + col_widths_data
    row_heights = [col_header_h] + row_heights_data

    tbl = Table(table_data, colWidths=col_widths, rowHeights=row_heights, repeatRows=1)
    style_cmds = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (0, 1), (1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF7")),
        ("BACKGROUND", (0, 1), (1, -1), colors.HexColor("#E8EEF7")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#333333")),
        ("LINEBELOW", (0, 0), (-1, 0), 0.35, colors.HexColor("#333333")),
        ("LINEAFTER", (1, 0), (1, -1), 0.35, colors.HexColor("#333333")),
        ("LEFTPADDING", (0, 0), (-1, -1), CELL_TEXT_PAD),
        ("RIGHTPADDING", (0, 0), (-1, -1), CELL_TEXT_PAD),
        ("TOPPADDING", (0, 0), (-1, -1), CELL_TEXT_PAD),
        ("BOTTOMPADDING", (0, 0), (-1, -1), CELL_TEXT_PAD),
    ]

    # Lightly shade alternating interfering-element column headers for readability.
    for j in range(2, len(ykeys) + 2):
        if (j - 2) % 2 == 1:
            style_cmds.append(("BACKGROUND", (j, 0), (j, 0), colors.HexColor("#DCE6F2")))

    # Grid bars only between valid/populated matrix cells, not across empty whitespace.
    for i, xk in enumerate(xkeys, start=1):
        for j, y in enumerate(ykeys, start=2):
            if matrix.get((y, xk)):
                style_cmds.extend([
                    ("BACKGROUND", (j, i), (j, i), colors.HexColor("#FFF7D6")),
                    ("BOX", (j, i), (j, i), GRID_LINE_WIDTH, colors.HexColor("#808891")),
                ])

    tbl.setStyle(TableStyle(style_cmds))

    doc = SimpleDocTemplate(out_pdf, pagesize=(page_w, page_h), leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=MARGIN, bottomMargin=MARGIN)
    story = [title]
    if key_tbl is not None:
        story += [Spacer(1, 0.03 * inch), key_tbl]
    story += [Spacer(1, 0.05 * inch), tbl]
    doc.build(story)

def main() -> None:
    p = argparse.ArgumentParser(description="Generate a poster-sized ICP-MS interference matrix PDF.")
    p.add_argument("xlsx", help="Input Excel workbook")
    p.add_argument("out_pdf", nargs="?", default="interference_poster.pdf", help="Output PDF path")
    p.add_argument("--sheet", default=0, help="Sheet name or index; default first sheet")
    p.add_argument("--abundances", default=DEFAULT_ABUNDANCE_CSV, help="Isotope-abundance CSV path; default IsotopeAbundances.csv next to this script or in the working directory")
    args = p.parse_args()

    sheet = args.sheet
    try:
        sheet = int(sheet)
    except (TypeError, ValueError):
        pass

    df = read_data(args.xlsx, sheet_name=sheet)
    build_pdf(df, args.out_pdf, abundance_csv=args.abundances)
    print(f"Wrote {args.out_pdf}")


if __name__ == "__main__":
    main()
