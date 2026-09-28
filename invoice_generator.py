#!/usr/bin/env python3
"""
Invoice Generator
=================

Reads a master Excel spreadsheet, validates rows client-by-client, and
creates one editable Word invoice (.docx) per valid client. A PDF copy is
also produced so the generated invoice can be viewed without a Word viewer.

Expected spreadsheet columns:
    Date
    Client
    Typist
    Description
    Client Rate
    Typist Rate
    Mins
    Other info
    Proofing
    Typist total
    Sub total
    VAT
    Total

Only these columns are placed on the client invoice:
    Date, Description, Client Rate, Mins, Sub total, VAT, Total

Important:
- The client name is filled automatically from the spreadsheet.
- Client address lines are deliberately left blank for manual completion.
- Invoice number, invoice date, and reference are deliberately left blank.
- The month/year in the "To:" line is generated from the invoice data.
- A client invoice is generated only when ALL rows for that client are valid.
- Every successfully validated client receives a Schedule of Work.xlsx for easy checking.
- Clients with more than 25 jobs receive a short invoice that refers to the attached schedule.
- Output is organised by client and invoice month for easier long-term tracking.
- Each run creates a separate simplified Excel output showing the calculated/validation/status values written to the master spreadsheet.
- The supplied logo.png is inserted automatically when it sits beside this script.
- The generated DOCX is the editable master. The script then creates a PDF
  copy for viewing. LibreOffice is used when available; otherwise a ReportLab
  preview PDF is generated as a fallback.
- This script does NOT connect to My Docs Online. It operates on a local
  Excel file.
"""

from __future__ import annotations

import argparse
import csv
import re
import shutil
import subprocess
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ReportLab is kept only as a PDF fallback when LibreOffice is unavailable.
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image as RLImage,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path("generated_invoices")
DEFAULT_LOGO = SCRIPT_DIR / "logo.png"
LARGE_CLIENT_JOB_THRESHOLD = 25

INVOICE_REGISTER = SCRIPT_DIR / "invoice_register.csv"
INVOICE_START_NUMBER = 18122


VAT_RATE = Decimal("0.20")

AUDIO_SOURCE_EXTENSIONS = {
    ".mp3", ".m4a", ".wav", ".wma", ".aac", ".flac", ".dss", ".ds2"
}

NON_AUDIO_SOURCE_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".jpg", ".jpeg",
    ".png", ".tif", ".tiff", ".ppt", ".pptx", ".rtf", ".txt"
}

SOURCE_FILE_PATTERN = re.compile(
    r"(?:^|;\s*)Source file:\s*([^;]+)",
    re.IGNORECASE,
)


class WarningApprovalRequired(Exception):
    """Raised when suspicious-but-not-blocking invoice data needs user approval."""


VALIDATION_COLUMN = "Validation Issues"
STATUS_COLUMN = "Status"

STATUS_READY = "Ready"
STATUS_REVIEW = "Needs Review"
STATUS_INVOICED = "Invoice Generated"

REQUIRED_COLUMNS = [
    "Date",
    "Client",
    "Typist",
    "Description",
    "Client Rate",
    "Typist Rate",
    "Mins",
    "Proofing",
    "Typist total",
    "Sub total",
    "VAT",
    "Total",
]

# Columns that represent actual source/job data.  Derived columns such as
# Sub total, VAT, Total, Validation Issues and Status are deliberately excluded.
# A spreadsheet row is treated as completely blank when every one of these
# fields is empty.
SOURCE_DATA_COLUMNS = [
    "Date",
    "Client",
    "Typist",
    "Description",
    "Client Rate",
    "Typist Rate",
    "Mins",
    "Email Subject",
    "Other info",
    "Source File",
    "Proofing",
    "Typist total",
]

REVIEW_FILL = PatternFill(fill_type="solid", fgColor="FCE8E6")
INVOICED_FILL = PatternFill(fill_type="solid", fgColor="E6F4EA")
READY_FILL = PatternFill(fill_type="solid", fgColor="FFF4CC")


# ---------------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------------


def money(value) -> Decimal:
    """Convert an Excel value such as 123.45 or '£123.45' to Decimal."""
    if pd.isna(value):
        return Decimal("0.00")

    text = str(value).strip()
    text = text.replace("£", "").replace(",", "").strip()

    if not text:
        return Decimal("0.00")

    try:
        return Decimal(text).quantize(Decimal("0.01"))
    except InvalidOperation as exc:
        raise ValueError(f"Could not interpret money value: {value!r}") from exc


def format_money(value: Decimal) -> str:
    return f"£{value:,.2f}"


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def source_row_is_blank(row) -> bool:
    """Return True when a row contains no actual job/input data.

    This intentionally ignores derived/output columns such as Sub total, VAT,
    Total, Validation Issues and Status.  That prevents formatted/template rows
    or values left by an earlier run from being treated as real jobs.
    """
    for column in SOURCE_DATA_COLUMNS:
        value = row.get(column, None)
        if pd.isna(value):
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return False
    return True


def safe_filename(name: str) -> str:
    """Make a client name safe for use as a filename."""
    name = clean_text(name)
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:150] or "Unknown_Client"


def parse_spreadsheet_date(value):
    """Parse Excel/CSV date values without misreading ISO dates as UK dates."""
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value
    if isinstance(value, datetime):
        return pd.Timestamp(value)
    if isinstance(value, date):
        return pd.Timestamp(value)

    text = str(value).strip()
    if not text:
        return pd.NaT

    # ISO/Y-M-D values are unambiguous and must not be parsed with dayfirst=True.
    if re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
        return pd.to_datetime(text, yearfirst=True, errors="coerce")

    # The source workbook is UK-style, so ambiguous slash dates are day-first.
    return pd.to_datetime(text, dayfirst=True, errors="coerce")


def format_invoice_date(value) -> str:
    """Format a spreadsheet date in the style used by the supplied invoice."""
    if pd.isna(value):
        return ""

    if isinstance(value, datetime):
        d = value.date()
    elif isinstance(value, date):
        d = value
    else:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return clean_text(value)
        d = parsed.date()

    if 10 <= d.day % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(d.day % 10, "th")

    return f"{d.day}{suffix} {d.strftime('%b %y')}"


def month_year_from_rows(group: pd.DataFrame) -> str:
    """Determine the invoice month from the dates in a client's rows."""
    dates = pd.to_datetime(group["Date"], errors="coerce").dropna()

    if dates.empty:
        return "MONTH YEAR"

    periods = dates.dt.to_period("M").unique()

    if len(periods) > 1:
        months = ", ".join(str(p) for p in sorted(periods))
        raise ValueError(
            f"Client {clean_text(group['Client'].iloc[0])!r} has rows from "
            f"multiple months: {months}. Split the spreadsheet/month first."
        )

    return dates.iloc[0].strftime("%B %Y")


def period_folder_from_rows(group: pd.DataFrame) -> str:
    """Return a sortable YYYY-MM folder name for a client's invoice period."""
    dates = pd.to_datetime(group["Date"], errors="coerce").dropna()
    if dates.empty:
        return "Unknown-Month"

    periods = dates.dt.to_period("M").unique()
    if len(periods) != 1:
        # month_year_from_rows will raise the clearer error elsewhere; this is
        # simply a safe fallback for path construction.
        return "Mixed-Months"

    return str(periods[0])


def validate_columns(df: pd.DataFrame) -> None:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(
            "The spreadsheet is missing these required columns:\n"
            + "\n".join(f"  - {col}" for col in missing)
        )


def calculate_totals(rate, mins):
    """Recalculate subtotal/VAT/total from client rate and minutes."""
    try:
        if pd.isna(rate) or pd.isna(mins):
            raise ValueError
        rate_decimal = Decimal(str(rate))
        mins_decimal = Decimal(str(mins))
    except Exception:
        return Decimal("0.00"), Decimal("0.00"), Decimal("0.00")

    subtotal = (rate_decimal * mins_decimal).quantize(Decimal("0.01"))
    vat = (subtotal * VAT_RATE).quantize(Decimal("0.01"))
    total = (subtotal + vat).quantize(Decimal("0.01"))
    return subtotal, vat, total


def validate_row(row, expected_period):
    """
    Validate only fields that must be present/correct to create a client invoice.

    The monthly spreadsheet is the human review boundary. Once the user has
    reviewed it and clicks Generate Invoices, populated inferred rates are
    treated as accepted values. Rate-inference evidence is informational and
    does not block invoice generation.
    """
    issues = []

    if pd.isna(row["Date"]):
        issues.append("Missing Date")
    if not clean_text(row["Client"]):
        issues.append("Missing Client")
    if not clean_text(row["Description"]):
        issues.append("Missing Description")
    if pd.isna(row["Client Rate"]):
        issues.append("Missing Client Rate")
    if pd.isna(row["Mins"]):
        issues.append("Missing Minutes")

    if expected_period is not None and not pd.isna(row["Date"]):
        row_period = pd.Timestamp(row["Date"]).to_period("M")
        if row_period != expected_period:
            issues.append("Unexpected Month/Year")

    return "; ".join(issues)



def extract_source_filename(row) -> str:
    """
    Read the hidden Source File column written by the current spreadsheet
    generator. Fall back to legacy Other info for older workbooks.
    """
    source_file = clean_text(row.get("Source File", ""))
    if source_file:
        return source_file

    other_info = clean_text(row.get("Other info", ""))
    match = SOURCE_FILE_PATTERN.search(other_info)
    return match.group(1).strip() if match else ""


def numeric_value(value):
    """Return a float for a spreadsheet numeric value, or None when unavailable."""
    if pd.isna(value):
        return None

    text = str(value).strip().replace("£", "").replace(",", "")
    if not text:
        return None

    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def suspicious_row_warnings(row) -> list[str]:
    """
    Return advisory warnings only.

    These do NOT make a row invalid. They highlight unusual combinations that
    deserve a human check before invoice generation.
    """
    warnings = []

    source_file = extract_source_filename(row)
    extension = Path(source_file).suffix.lower() if source_file else ""

    mins = numeric_value(row.get("Mins"))
    client_rate = numeric_value(row.get("Client Rate"))
    typist_rate = numeric_value(row.get("Typist Rate"))

    # The strongest rule: document/image sources normally should not be billed
    # as audio minutes. This is a warning rather than a hard block because some
    # unusual legitimate workflows may exist.
    if extension in NON_AUDIO_SOURCE_EXTENSIONS and mins is not None:
        warnings.append(
            f"Non-audio source {source_file!r} has Mins={mins:g}; "
            "confirm the billing basis"
        )

    if client_rate is not None and (client_rate < 0.50 or client_rate > 5.00):
        warnings.append(
            f"Client Rate £{client_rate:.2f} is outside the expected per-minute range"
        )

    if (
        client_rate is not None
        and typist_rate is not None
        and typist_rate > client_rate
    ):
        warnings.append(
            f"Typist Rate £{typist_rate:.2f} is greater than "
            f"Client Rate £{client_rate:.2f}"
        )

    if mins is not None and mins <= 0:
        warnings.append(f"Mins={mins:g}; expected a positive value")

    if mins is not None and mins > 600:
        warnings.append(f"Mins={mins:g}; unusually high duration")

    return warnings


def collect_invoice_warnings(df: pd.DataFrame, active_row_mask) -> list[dict]:
    warnings = []
    header_row = int(df.attrs.get("excel_header_row", 0))

    for idx, row in df.loc[active_row_mask].iterrows():
        row_warnings = suspicious_row_warnings(row)

        if not row_warnings:
            continue

        # header_row is zero-based. DataFrame idx 0 is the first row after it.
        excel_row = int(idx) + header_row + 2

        warnings.append(
            {
                "excel_row": excel_row,
                "client": clean_text(row.get("Client", "")) or "(missing client)",
                "description": clean_text(row.get("Description", "")),
                "warnings": row_warnings,
            }
        )

    return warnings


def print_invoice_warnings(warnings: list[dict]) -> None:
    if not warnings:
        print("No suspicious invoice-data warnings found.")
        return

    print("")
    print("=" * 80)
    print("SUSPICIOUS DATA WARNINGS")
    print("=" * 80)
    print(
        "These are advisory only. They do not mean the invoice data is definitely wrong."
    )

    for item in warnings:
        print(
            f"[WARNING] Row {item['excel_row']} | "
            f"{item['client']} | {item['description']}"
        )
        for warning in item["warnings"]:
            print(f"    - {warning}")

    print("")
    print(f"{len(warnings)} row(s) have suspicious data that should be checked.")
    print("")



INVOICE_REGISTER_COLUMNS = [
    "Invoice Number",
    "Client",
    "Period",
    "Invoice Date",
    "Created At",
]


def format_today_invoice_date() -> str:
    today = date.today()
    return f"{today.day}.{today.month}.{today.strftime('%y')}"


def read_invoice_register(path: Path = INVOICE_REGISTER) -> list[dict]:
    if not path.exists():
        return []

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_invoice_register(rows: list[dict], path: Path = INVOICE_REGISTER) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=INVOICE_REGISTER_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)


def _invoice_register_key(client: str, period: str) -> tuple[str, str]:
    return clean_text(client).casefold(), clean_text(period).casefold()


def _next_invoice_number(rows: list[dict]) -> int:
    numbers = []
    for row in rows:
        try:
            numbers.append(int(clean_text(row.get("Invoice Number", ""))))
        except (TypeError, ValueError):
            pass
    return max(numbers, default=INVOICE_START_NUMBER - 1) + 1


def get_or_create_invoice_metadata(
    client: str,
    period: str,
    register_path: Path = INVOICE_REGISTER,
) -> tuple[str, str]:
    """
    Reuse the same invoice number for the same client/month on reruns.
    New client/month combinations take the next sequential number.
    """
    rows = read_invoice_register(register_path)
    key = _invoice_register_key(client, period)

    for row in rows:
        if _invoice_register_key(row.get("Client", ""), row.get("Period", "")) != key:
            continue

        invoice_number = clean_text(row.get("Invoice Number", ""))
        invoice_date = clean_text(row.get("Invoice Date", ""))

        if not invoice_date:
            invoice_date = format_today_invoice_date()
            row["Invoice Date"] = invoice_date
            if not clean_text(row.get("Created At", "")):
                row["Created At"] = datetime.now().isoformat(timespec="seconds")
            write_invoice_register(rows, register_path)

        return invoice_number, invoice_date

    invoice_number = str(_next_invoice_number(rows))
    invoice_date = format_today_invoice_date()

    rows.append(
        {
            "Invoice Number": invoice_number,
            "Client": client,
            "Period": period,
            "Invoice Date": invoice_date,
            "Created At": datetime.now().isoformat(timespec="seconds"),
        }
    )
    write_invoice_register(rows, register_path)

    return invoice_number, invoice_date


# ---------------------------------------------------------------------------
# Word document helpers
# ---------------------------------------------------------------------------


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")

    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)

    for margin, value in {
        "top": top,
        "start": start,
        "bottom": bottom,
        "end": end,
    }.items():
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")

    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)

    for edge in ("top", "left", "bottom", "right"):
        if edge not in kwargs:
            continue

        edge_data = kwargs[edge]
        tag = f"w:{edge}"
        element = borders.find(qn(tag))

        if element is None:
            element = OxmlElement(tag)
            borders.append(element)

        for key, value in edge_data.items():
            element.set(qn(f"w:{key}"), str(value))


def remove_table_borders(table):
    """Hide all borders on a layout-only Word table."""
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "nil")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def set_fixed_table_layout(table):
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")


def set_run_font(run, size, bold=False, italic=False, underline=False, name="Arial"):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.underline = underline


# ---------------------------------------------------------------------------
# DOCX invoice generation
# ---------------------------------------------------------------------------


def setup_document() -> Document:
    doc = Document()
    section = doc.sections[0]

    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(19)
    section.right_margin = Mm(19)
    section.top_margin = Mm(9)
    section.bottom_margin = Mm(18)
    section.footer_distance = Mm(6)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.0

    return doc


def add_brand_header(doc: Document, logo_path: Path | None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)

    run = p.add_run("O U T S O U R C I N G   S E C R E T A R I A L   S P E C I A L I S T S")
    set_run_font(run, 7.5)

    if logo_path and logo_path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(4)
        try:
            p.add_run().add_picture(str(logo_path), width=Mm(63))
        except Exception as exc:
            print(f"  WARNING: Could not insert logo {logo_path}: {exc}")
            logo_path = None

    if not logo_path or not logo_path.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run("F I N G E R T I P S")
        set_run_font(run, 22, bold=True, name="Courier New")

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run("T Y P I N G   S E R V I C E S")
        set_run_font(run, 10, bold=True, name="Courier New")
        run.font.color.rgb = RGBColor(181, 31, 53)


def add_invoice_header(doc: Document, client: str, logo_path: Path | None, invoice_number: str, invoice_date: str):
    """Add the reference-style header while leaving unknown details editable/blank."""
    add_brand_header(doc, logo_path)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(15)
    p.paragraph_format.space_after = Pt(28)
    run = p.add_run("INVOICE")
    set_run_font(run, 16, bold=True, underline=True)

    outer = doc.add_table(rows=1, cols=2)
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER
    outer.autofit = False
    set_fixed_table_layout(outer)
    remove_table_borders(outer)

    outer.columns[0].width = Mm(86)
    outer.columns[1].width = Mm(86)

    left = outer.cell(0, 0)
    right = outer.cell(0, 1)
    left.width = Mm(86)
    right.width = Mm(86)
    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    set_cell_margins(left, top=0, start=0, bottom=0, end=300)
    set_cell_margins(right, top=0, start=300, bottom=0, end=0)

    p = left.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(client)
    set_run_font(run, 10)

    # Three genuinely blank editable address lines.
    for _ in range(3):
        p = left.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.add_run(" ")

    details = right.add_table(rows=4, cols=2)
    details.alignment = WD_TABLE_ALIGNMENT.RIGHT
    details.autofit = False
    set_fixed_table_layout(details)
    remove_table_borders(details)
    details.columns[0].width = Mm(36)
    details.columns[1].width = Mm(47)

    rows = [
        ("INVOICE NO.:", invoice_number),
        ("", "Please state as reference"),
        ("DATE:", invoice_date),
        ("REF.:", "________________"),
    ]

    for i, (label, value) in enumerate(rows):
        label_cell = details.cell(i, 0)
        value_cell = details.cell(i, 1)
        set_cell_margins(label_cell, top=0, start=0, bottom=20, end=50)
        set_cell_margins(value_cell, top=0, start=0, bottom=20, end=0)

        p = label_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(label)
        set_run_font(run, 10)

        p = value_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(value)
        if i == 1:
            set_run_font(run, 8, italic=True)
        else:
            set_run_font(run, 10)

    # Spacing before the service-description line.
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)


def add_service_description(doc: Document, month_year: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(17)

    run = p.add_run("To:    ")
    set_run_font(run, 10)

    run = p.add_run(f"Typing services carried out during the month of {month_year}.")
    set_run_font(run, 10)


def add_invoice_table(doc: Document, group: pd.DataFrame):
    table = doc.add_table(rows=1, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_fixed_table_layout(table)

    widths = [19, 66, 15, 15, 23, 18, 22]
    headings = ["Date", "Description", "Rate", "Mins", "Sub total", "VAT", "Total"]

    for i, width in enumerate(widths):
        table.columns[i].width = Mm(width)

    header = table.rows[0]
    set_repeat_table_header(header)
    prevent_row_split(header)

    for i, heading in enumerate(headings):
        cell = header.cells[i]
        cell.width = Mm(widths[i])
        set_cell_margins(cell, top=25, start=30, bottom=35, end=30)
        set_cell_border(
            cell,
            bottom={"val": "single", "sz": "6", "color": "000000"},
        )

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(heading)
        set_run_font(run, 8, bold=True)

    total_sub = Decimal("0.00")
    total_vat = Decimal("0.00")
    total_total = Decimal("0.00")

    for _, row in group.iterrows():
        subtotal = money(row["Sub total"])
        vat = money(row["VAT"])
        total = money(row["Total"])
        rate = money(row["Client Rate"])

        total_sub += subtotal
        total_vat += vat
        total_total += total

        values = [
            format_invoice_date(row["Date"]),
            clean_text(row["Description"]),
            f"{rate:.1f}",
            clean_text(row["Mins"]),
            format_money(subtotal),
            format_money(vat),
            format_money(total),
        ]

        cells = table.add_row().cells
        prevent_row_split(table.rows[-1])

        for i, value in enumerate(values):
            cells[i].width = Mm(widths[i])
            set_cell_margins(cells[i], top=22, start=30, bottom=22, end=30)

            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if i >= 2 else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(value))
            set_run_font(run, 7.8)

    totals_row = table.add_row()
    prevent_row_split(totals_row)
    totals = ["", "", "", "", format_money(total_sub), format_money(total_vat), format_money(total_total)]

    for i, value in enumerate(totals):
        cell = totals_row.cells[i]
        cell.width = Mm(widths[i])
        set_cell_margins(cell, top=40, start=30, bottom=20, end=30)

        if i >= 4:
            set_cell_border(
                cell,
                top={"val": "single", "sz": "6", "color": "000000"},
            )

            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(value)
            set_run_font(run, 8, bold=True)

    # Underline the final grand total, matching the reference invoice emphasis.
    set_cell_border(
        totals_row.cells[6],
        top={"val": "single", "sz": "6", "color": "000000"},
        bottom={"val": "double", "sz": "6", "color": "000000"},
    )

    return total_sub, total_vat, total_total


def calculate_group_totals(group: pd.DataFrame):
    total_sub = Decimal("0.00")
    total_vat = Decimal("0.00")
    total_total = Decimal("0.00")

    for _, row in group.iterrows():
        total_sub += money(row["Sub total"])
        total_vat += money(row["VAT"])
        total_total += money(row["Total"])

    return total_sub, total_vat, total_total


def add_large_client_summary(doc: Document, group: pd.DataFrame):
    """Use a compact invoice for high-volume clients and refer to the Excel schedule."""
    total_sub, total_vat, total_total = calculate_group_totals(group)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(18)
    run = p.add_run("See attached Schedule of Work for full job details.")
    set_run_font(run, 10, italic=True)

    table = doc.add_table(rows=2, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.autofit = False
    set_fixed_table_layout(table)
    remove_table_borders(table)
    widths = [28, 24, 28]
    headings = ["Sub total", "VAT", "Total"]
    values = [format_money(total_sub), format_money(total_vat), format_money(total_total)]

    for i, width in enumerate(widths):
        table.columns[i].width = Mm(width)
        for row in table.rows:
            row.cells[i].width = Mm(width)

        header_cell = table.cell(0, i)
        set_cell_border(header_cell, bottom={"val": "single", "sz": "6", "color": "000000"})
        p = header_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(headings[i])
        set_run_font(run, 8, bold=True)

        value_cell = table.cell(1, i)
        p = value_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(values[i])
        set_run_font(run, 9, bold=True)

    set_cell_border(table.cell(1, 2), bottom={"val": "double", "sz": "6", "color": "000000"})
    return total_sub, total_vat, total_total


def generate_schedule_of_work(group: pd.DataFrame, output_path: Path):
    """Create the attachment used for high-volume clients."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule of Work"

    headers = ["Date", "Description", "Rate", "Mins", "Sub total", "VAT", "Total"]
    ws.append(headers)

    header_font = Font(bold=True)
    thin_black = Side(style="thin", color="000000")

    for cell in ws[1]:
        cell.font = header_font
        cell.border = Border(bottom=thin_black)
        cell.alignment = Alignment(vertical="top")

    total_sub = Decimal("0.00")
    total_vat = Decimal("0.00")
    total_total = Decimal("0.00")

    for _, row in group.iterrows():
        subtotal = money(row["Sub total"])
        vat = money(row["VAT"])
        total = money(row["Total"])
        rate = money(row["Client Rate"])

        total_sub += subtotal
        total_vat += vat
        total_total += total

        parsed_date = parse_spreadsheet_date(row["Date"])
        excel_date = parsed_date.to_pydatetime() if hasattr(parsed_date, "to_pydatetime") else parsed_date

        ws.append([
            excel_date,
            clean_text(row["Description"]),
            float(rate),
            row["Mins"],
            float(subtotal),
            float(vat),
            float(total),
        ])

    ws.append([None, None, None, None, float(total_sub), float(total_vat), float(total_total)])
    total_row = ws.max_row

    for cell in ws[total_row][4:7]:
        cell.font = header_font
        cell.border = Border(top=thin_black)

    for cell in ws[2:total_row]:
        if cell[0].value:
            cell[0].number_format = "d mmm yy"
        for col_idx in (3, 5, 6, 7):
            cell[col_idx - 1].number_format = '£#,##0.00' if col_idx != 4 else '0.00'

    # Rate is currency-style data even though the invoice displays it without a pound sign.
    for row in range(2, total_row + 1):
        ws.cell(row=row, column=3).number_format = '0.0'
        ws.cell(row=row, column=5).number_format = '£#,##0.00'
        ws.cell(row=row, column=6).number_format = '£#,##0.00'
        ws.cell(row=row, column=7).number_format = '£#,##0.00'

    widths = {"A": 15, "B": 60, "C": 12, "D": 12, "E": 15, "F": 13, "G": 15}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{total_row - 1}"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def add_footer_content(doc: Document):
    """Add the payment/VAT/bank information to the invoice body."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(34)
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run("Payable within 14 days of date of invoice.")
    set_run_font(run, 10, bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    run = p.add_run("VAT Number 919 8060 03")
    set_run_font(run, 10, bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run("Bank details: ")
    set_run_font(run, 10, bold=True)
    run = p.add_run("A/C Number: 45389702  /  Sort code: 09-01-27")
    set_run_font(run, 10)


def add_page_footer(doc: Document):
    """Put the company/contact line in Word's real page footer."""
    footer_text = (
        "Fingertips Typing Services Limited, 36a Slade Gardens, Erith, Kent DA8 2HT  \n"
        "T: 0845 300 5245  ●  F: 01322 351190  ●  "
        "E: cbennett@fingertipstyping.co.uk  W: www.fingertipstyping.co.uk"
    )

    for section in doc.sections:
        section.footer_distance = Mm(6)
        footer = section.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.clear()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(footer_text)
        set_run_font(run, 6.0)


def generate_docx_invoice(
    client: str,
    group: pd.DataFrame,
    output_path: Path,
    logo_path: Path | None = None,
    large_client: bool = False,
    invoice_number: str = "",
    invoice_date: str = "",
):
    month_year = month_year_from_rows(group)

    doc = setup_document()
    add_invoice_header(doc, client, logo_path, invoice_number, invoice_date)
    add_service_description(doc, month_year)

    if large_client:
        total_sub, total_vat, total_total = add_large_client_summary(doc, group)
    else:
        total_sub, total_vat, total_total = add_invoice_table(doc, group)

    add_footer_content(doc)
    add_page_footer(doc)

    # Keep document metadata generic.
    doc.core_properties.title = f"Invoice - {client}"
    doc.core_properties.author = "Invoice Automation"

    doc.save(str(output_path))

    return {
        "client": client,
        "rows": len(group),
        "subtotal": total_sub,
        "vat": total_vat,
        "total": total_total,
        "docx": output_path,
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
    }


# ---------------------------------------------------------------------------
# PDF creation
# ---------------------------------------------------------------------------


def find_libreoffice() -> str | None:
    """Find LibreOffice/soffice on Linux, macOS, or common Windows installs."""
    for command in ("soffice", "libreoffice"):
        found = shutil.which(command)
        if found:
            return found

    windows_candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]
    for candidate in windows_candidates:
        if candidate.exists():
            return str(candidate)

    mac_candidate = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac_candidate.exists():
        return str(mac_candidate)

    return None


def convert_docx_to_pdf(docx_path: Path) -> Path | None:
    """Convert DOCX to PDF with LibreOffice. Return None if unavailable/fails."""
    office = find_libreoffice()
    if not office:
        return None

    pdf_path = docx_path.with_suffix(".pdf")
    if pdf_path.exists():
        pdf_path.unlink()

    try:
        completed = subprocess.run(
            [
                office,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(docx_path.parent),
                str(docx_path),
            ],
            capture_output=True,
            text=True,
            timeout=90,
        )
    except (OSError, subprocess.SubprocessError):
        return None

    if completed.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0:
        return pdf_path

    return None


def _pdf_styles():
    base = getSampleStyleSheet()
    return {
        "normal": ParagraphStyle(
            "PDFNormal",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=9.5,
            leading=11,
            spaceAfter=0,
        ),
        "small": ParagraphStyle(
            "PDFSmall",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=9,
            spaceAfter=0,
        ),
        "company": ParagraphStyle(
            "PDFCompany",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=7.5,
            leading=9,
            alignment=TA_CENTER,
            spaceAfter=0,
        ),
        "brand": ParagraphStyle(
            "PDFBrand",
            parent=base["Normal"],
            fontName="Courier-Bold",
            fontSize=20,
            leading=22,
            alignment=TA_CENTER,
            spaceAfter=0,
        ),
        "brand_sub": ParagraphStyle(
            "PDFBrandSub",
            parent=base["Normal"],
            fontName="Courier-Bold",
            fontSize=9,
            leading=11,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#B51F35"),
            spaceAfter=0,
        ),
        "title": ParagraphStyle(
            "PDFTitle",
            parent=base["Normal"],
            fontName="Helvetica-Bold",
            fontSize=16,
            leading=18,
            spaceAfter=0,
        ),
        "right": ParagraphStyle(
            "PDFRight",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=9.5,
            leading=11,
            alignment=TA_RIGHT,
            spaceAfter=0,
        ),
        "table": ParagraphStyle(
            "PDFTable",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=7.3,
            leading=8.5,
            spaceAfter=0,
        ),
        "table_right": ParagraphStyle(
            "PDFTableRight",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=7.3,
            leading=8.5,
            alignment=TA_RIGHT,
            spaceAfter=0,
        ),
    }


def _draw_pdf_page_footer(canvas, doc):
    """Draw the company/contact line at the physical bottom of every PDF page."""
    footer_text = (
        "Fingertips Typing Services Limited, 36a Slade Gardens, Erith, Kent DA8 2HT  "
        "T: 0845 300 5245  •  F: 01322 351190  •  "
        "E: cbennett@fingertipstyping.co.uk  W: www.fingertipstyping.co.uk"
    )

    canvas.saveState()
    available_width = A4[0] - (38 * mm)
    font_size = 6.0
    while font_size > 4.5 and canvas.stringWidth(footer_text, "Helvetica", font_size) > available_width:
        font_size -= 0.2

    canvas.setFont("Helvetica", font_size)
    canvas.drawCentredString(A4[0] / 2, 7 * mm, footer_text)
    canvas.restoreState()


def generate_pdf_preview(
    client: str,
    group: pd.DataFrame,
    output_path: Path,
    logo_path: Path | None = None,
    large_client: bool = False,
    invoice_number: str = "",
    invoice_date: str = "",
):
    """Fallback PDF generation when LibreOffice is not available."""
    styles = _pdf_styles()
    month_year = month_year_from_rows(group)

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=19 * mm,
        leftMargin=19 * mm,
        topMargin=9 * mm,
        bottomMargin=14 * mm,
        title=f"Invoice - {client}",
        author="Invoice Automation",
    )

    story = [
        Paragraph(
            "O U T S O U R C I N G &nbsp; S E C R E T A R I A L &nbsp; S P E C I A L I S T S",
            styles["company"],
        ),
        Spacer(1, 4),
    ]

    if logo_path and logo_path.exists():
        try:
            logo = RLImage(str(logo_path), width=63 * mm, height=27 * mm, kind="proportional")
            logo.hAlign = "CENTER"
            story.extend([logo, Spacer(1, 4)])
        except Exception:
            story.extend(
                [
                    Paragraph("F I N G E R T I P S", styles["brand"]),
                    Paragraph("T Y P I N G &nbsp; S E R V I C E S", styles["brand_sub"]),
                ]
            )
    else:
        story.extend(
            [
                Paragraph("F I N G E R T I P S", styles["brand"]),
                Paragraph("T Y P I N G &nbsp; S E R V I C E S", styles["brand_sub"]),
            ]
        )

    story.extend(
        [
            Spacer(1, 15),
            Paragraph("<u>INVOICE</u>", styles["title"]),
            Spacer(1, 28),
        ]
    )

    left = [
        Paragraph(client, styles["normal"]),
        Paragraph("&nbsp;", styles["normal"]),
        Paragraph("&nbsp;", styles["normal"]),
        Paragraph("&nbsp;", styles["normal"]),
    ]

    right = Table(
        [
            [Paragraph("INVOICE NO.:", styles["normal"]), Paragraph(invoice_number, styles["right"])],
            ["", Paragraph("<i>Please state as reference</i>", styles["small"])],
            [Paragraph("DATE:", styles["normal"]), Paragraph(invoice_date, styles["right"])],
            [Paragraph("REF.:", styles["normal"]), Paragraph("________________", styles["right"])],
        ],
        colWidths=[36 * mm, 47 * mm],
    )
    right.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )

    details = Table([[left, right]], colWidths=[86 * mm, 86 * mm])
    details.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.extend(
        [
            details,
            Spacer(1, 18),
            Paragraph(
                f"To:&nbsp;&nbsp;&nbsp;&nbsp;Typing services carried out during the month of {month_year}.",
                styles["normal"],
            ),
            Spacer(1, 17),
        ]
    )

    total_sub, total_vat, total_total = calculate_group_totals(group)

    if large_client:
        story.append(
            Paragraph(
                "<i>See attached Schedule of Work for full job details.</i>",
                styles["normal"],
            )
        )
        story.append(Spacer(1, 18))

        summary_data = [
            [
                Paragraph("<b>Sub total</b>", styles["table_right"]),
                Paragraph("<b>VAT</b>", styles["table_right"]),
                Paragraph("<b>Total</b>", styles["table_right"]),
            ],
            [
                Paragraph(f"<b>{format_money(total_sub)}</b>", styles["table_right"]),
                Paragraph(f"<b>{format_money(total_vat)}</b>", styles["table_right"]),
                Paragraph(f"<b>{format_money(total_total)}</b>", styles["table_right"]),
            ],
        ]
        summary = Table(summary_data, colWidths=[28 * mm, 24 * mm, 28 * mm], hAlign="RIGHT")
        summary.setStyle(
            TableStyle(
                [
                    ("LINEBELOW", (0, 0), (-1, 0), 0.6, colors.black),
                    ("LINEBELOW", (2, 1), (2, 1), 1.0, colors.black),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(summary)
    else:
        data = [[
            Paragraph("<b>Date</b>", styles["table"]),
            Paragraph("<b>Description</b>", styles["table"]),
            Paragraph("<b>Rate</b>", styles["table_right"]),
            Paragraph("<b>Mins</b>", styles["table_right"]),
            Paragraph("<b>Sub total</b>", styles["table_right"]),
            Paragraph("<b>VAT</b>", styles["table_right"]),
            Paragraph("<b>Total</b>", styles["table_right"]),
        ]]

        for _, row in group.iterrows():
            subtotal = money(row["Sub total"])
            vat = money(row["VAT"])
            total = money(row["Total"])
            rate = money(row["Client Rate"])

            data.append([
                Paragraph(format_invoice_date(row["Date"]), styles["table"]),
                Paragraph(clean_text(row["Description"]), styles["table"]),
                Paragraph(f"{rate:.1f}", styles["table_right"]),
                Paragraph(clean_text(row["Mins"]), styles["table_right"]),
                Paragraph(format_money(subtotal), styles["table_right"]),
                Paragraph(format_money(vat), styles["table_right"]),
                Paragraph(format_money(total), styles["table_right"]),
            ])

        data.append([
            "", "", "", "",
            Paragraph(f"<b>{format_money(total_sub)}</b>", styles["table_right"]),
            Paragraph(f"<b>{format_money(total_vat)}</b>", styles["table_right"]),
            Paragraph(f"<b>{format_money(total_total)}</b>", styles["table_right"]),
        ])

        table = Table(
            data,
            colWidths=[19 * mm, 66 * mm, 15 * mm, 15 * mm, 23 * mm, 18 * mm, 22 * mm],
            repeatRows=1,
        )
        table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LINEBELOW", (0, 0), (-1, 0), 0.6, colors.black),
                    ("LINEABOVE", (4, -1), (-1, -1), 0.6, colors.black),
                    ("LINEBELOW", (6, -1), (6, -1), 1.0, colors.black),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(table)

    story.extend(
        [
            Spacer(1, 34),
            Paragraph("<b>Payable within 14 days of date of invoice.</b>", styles["normal"]),
            Spacer(1, 10),
            Paragraph("<b>VAT Number 919 8060 03</b>", styles["normal"]),
            Spacer(1, 10),
            Paragraph(
                "<b>Bank details:</b>&nbsp;&nbsp;A/C Number: 45389702&nbsp;&nbsp;/&nbsp;&nbsp;Sort code: 09-01-27",
                styles["normal"],
            ),
        ]
    )

    doc.build(
        story,
        onFirstPage=_draw_pdf_page_footer,
        onLaterPages=_draw_pdf_page_footer,
    )
    return output_path


def generate_invoice_files(
    client: str,
    group: pd.DataFrame,
    output_dir: Path,
    logo_path: Path | None = None,
):
    safe_client = safe_filename(client)
    month_year = month_year_from_rows(group)
    period_folder = period_folder_from_rows(group)

    # Keep every client's history together, split into sortable monthly folders.
    client_month_dir = output_dir / safe_client / period_folder
    client_month_dir.mkdir(parents=True, exist_ok=True)

    docx_path = client_month_dir / f"Invoice - {month_year}.docx"
    pdf_path = client_month_dir / f"Invoice - {month_year}.pdf"

    large_client = len(group) > LARGE_CLIENT_JOB_THRESHOLD
    invoice_number, invoice_date = get_or_create_invoice_metadata(
        client,
        period_folder,
    )

    result = generate_docx_invoice(
        client,
        group,
        docx_path,
        logo_path,
        large_client=large_client,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
    )

    # Always create a Schedule of Work for every successfully validated client.
    # For large clients the invoice refers to it; for smaller clients it is still
    # useful as a simple per-client/month audit view of the underlying job data.
    schedule_path = client_month_dir / f"Schedule of Work - {month_year}.xlsx"
    generate_schedule_of_work(group, schedule_path)

    converted_pdf = convert_docx_to_pdf(docx_path)
    pdf_method = "LibreOffice conversion"

    if converted_pdf is None:
        print("  LibreOffice conversion unavailable; creating ReportLab PDF preview instead.")
        generate_pdf_preview(
            client,
            group,
            pdf_path,
            logo_path,
            large_client=large_client,
        )
        converted_pdf = pdf_path
        pdf_method = "ReportLab fallback preview"

    result["pdf"] = converted_pdf
    result["pdf_method"] = pdf_method
    result["schedule"] = schedule_path
    result["large_client"] = large_client
    result["output_dir"] = client_month_dir
    return result


# ---------------------------------------------------------------------------
# Spreadsheet write-back
# ---------------------------------------------------------------------------


def write_status_back_to_excel(excel_path: Path, df: pd.DataFrame):
    """Update calculated/status columns without replacing the whole workbook."""
    workbook = load_workbook(excel_path)
    sheet = workbook.active

    header_map = {}
    for cell in sheet[1]:
        if cell.value is not None:
            header_map[str(cell.value).strip()] = cell.column

    update_columns = ["Sub total", "VAT", "Total", VALIDATION_COLUMN, STATUS_COLUMN]

    for column_name in update_columns:
        if column_name not in header_map:
            new_col = sheet.max_column + 1
            sheet.cell(row=1, column=new_col, value=column_name)
            header_map[column_name] = new_col

    for df_index, row in df.iterrows():
        excel_row = int(df_index) + 2

        # Completely blank source rows are not jobs.  Clear any derived values
        # that may have been written by an older version of the script and do
        # not give them validation/status information.
        if source_row_is_blank(row):
            for column_name in update_columns:
                cell = sheet.cell(excel_row, header_map[column_name])
                cell.value = None
                if column_name == STATUS_COLUMN:
                    cell.fill = PatternFill(fill_type=None)
            continue

        sheet.cell(excel_row, header_map["Sub total"], value=float(row["Sub total"]))
        sheet.cell(excel_row, header_map["VAT"], value=float(row["VAT"]))
        sheet.cell(excel_row, header_map["Total"], value=float(row["Total"]))
        sheet.cell(excel_row, header_map[VALIDATION_COLUMN], value=clean_text(row[VALIDATION_COLUMN]))
        status_cell = sheet.cell(excel_row, header_map[STATUS_COLUMN], value=clean_text(row[STATUS_COLUMN]))

        if row[STATUS_COLUMN] == STATUS_REVIEW:
            status_cell.fill = REVIEW_FILL
        elif row[STATUS_COLUMN] == STATUS_INVOICED:
            status_cell.fill = INVOICED_FILL
        elif row[STATUS_COLUMN] == STATUS_READY:
            status_cell.fill = READY_FILL

    workbook.save(excel_path)


def create_spreadsheet_output_report(df: pd.DataFrame, output_dir: Path) -> Path:
    """Create a simple view of the values the program writes to the master sheet.

    This is intentionally not a before/after audit. It shows the final calculated
    totals, validation result and status for each processed spreadsheet row.
    """
    report_columns = [
        "Excel Row",
        "Client",
        "Date",
        "Description",
        "Sub total",
        "VAT",
        "Total",
        VALIDATION_COLUMN,
        STATUS_COLUMN,
    ]

    records = []
    for idx, row in df.iterrows():
        # Ignore completely blank source/job rows in the viewing report.
        if source_row_is_blank(row):
            continue

        records.append({
            "Excel Row": int(idx) + 2,
            "Client": clean_text(row.get("Client", "")),
            "Date": row.get("Date", ""),
            "Description": clean_text(row.get("Description", "")),
            "Sub total": row.get("Sub total", 0),
            "VAT": row.get("VAT", 0),
            "Total": row.get("Total", 0),
            VALIDATION_COLUMN: clean_text(row.get(VALIDATION_COLUMN, "")),
            STATUS_COLUMN: clean_text(row.get(STATUS_COLUMN, "")),
        })

    report_df = pd.DataFrame(records, columns=report_columns)
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    report_path = output_dir / f"Spreadsheet Output - {timestamp}.xlsx"

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        report_df.to_excel(writer, index=False, sheet_name="Updated Values")
        ws = writer.book["Updated Values"]
        ws.freeze_panes = "A2"
        if ws.max_row >= 1:
            ws.auto_filter.ref = ws.dimensions

        widths = {
            "A": 11, "B": 28, "C": 14, "D": 55,
            "E": 14, "F": 14, "G": 14, "H": 35, "I": 18,
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

            for col_idx in (5, 6, 7):
                row[col_idx - 1].number_format = '£#,##0.00'

            status_cell = row[8]
            if status_cell.value == STATUS_REVIEW:
                status_cell.fill = REVIEW_FILL
            elif status_cell.value == STATUS_INVOICED:
                status_cell.fill = INVOICED_FILL
            elif status_cell.value == STATUS_READY:
                status_cell.fill = READY_FILL

    return report_path


# ---------------------------------------------------------------------------
# Main workflow
# ---------------------------------------------------------------------------


def read_master_spreadsheet(excel_path: Path) -> pd.DataFrame:
    """
    Locate the real header row automatically.

    The generated workbook contains title/information rows above the business
    table, so relying on Excel row 1 (or permanently hard-coding row 4) is
    brittle. We identify the row containing the key business columns.
    """
    preview = pd.read_excel(excel_path, header=None, nrows=15)

    required_markers = {
        "Date",
        "Client",
        "Description",
        "Client Rate",
        "Mins",
    }

    for header_row, row in preview.iterrows():
        values = {
            str(value).strip()
            for value in row.tolist()
            if not pd.isna(value)
        }

        if required_markers.issubset(values):
            df = pd.read_excel(excel_path, header=header_row)
            df.columns = df.columns.map(lambda value: str(value).strip())
            df.attrs["excel_header_row"] = int(header_row)
            return df

    raise ValueError(
        "Could not locate the spreadsheet header row. "
        "Expected Date, Client, Description, Client Rate and Mins columns."
    )


def generate_invoices(excel_path: Path, output_dir: Path, logo_path: Path | None = None, allow_warnings: bool = False, check_warnings_only: bool = False):
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Reading: {excel_path}")
    print("Checking data...\n")

    df = read_master_spreadsheet(excel_path)
    validate_columns(df)

    df["Date"] = df["Date"].apply(parse_spreadsheet_date)

    # These columns may arrive from Excel as integers when they initially contain
    # only zeros. Cast them before writing decimal values back into the DataFrame.
    for money_column in ("Sub total", "VAT", "Total"):
        df[money_column] = pd.to_numeric(df[money_column], errors="coerce").astype(float)

    if VALIDATION_COLUMN not in df.columns:
        df[VALIDATION_COLUMN] = ""
    if STATUS_COLUMN not in df.columns:
        df[STATUS_COLUMN] = ""

    # These are text columns. Excel sheets with entirely blank columns can be
    # inferred by pandas as float columns, which produces warnings when text is
    # assigned later. Force them to object/string-friendly dtype first.
    df[VALIDATION_COLUMN] = df[VALIDATION_COLUMN].astype(object)
    df[STATUS_COLUMN] = df[STATUS_COLUMN].astype(object)

    # The sheet is expected to represent a single invoice month. Use the most
    # common valid month as the expected period and flag outliers for review.
    # Establish which spreadsheet rows are real jobs BEFORE calculating or
    # writing any derived fields.  This is the key blank-row guard.
    active_row_mask = ~df.apply(source_row_is_blank, axis=1)

    expected_period = None
    valid_dates = df.loc[active_row_mask, "Date"].dropna()
    if not valid_dates.empty:
        expected_period = valid_dates.dt.to_period("M").value_counts().idxmax()

    invoice_warnings = collect_invoice_warnings(df, active_row_mask)
    print_invoice_warnings(invoice_warnings)

    if check_warnings_only:
        return {
            "warnings": invoice_warnings,
            "warning_count": len(invoice_warnings),
        }

    if invoice_warnings and not allow_warnings:
        raise WarningApprovalRequired(
            f"{len(invoice_warnings)} row(s) contain suspicious data. "
            "Review the warnings and rerun with --allow-warnings to continue."
        )

    # Recalculate money columns and validate only genuine job rows.
    for idx, row in df.iterrows():
        if not active_row_mask.loc[idx]:
            # Keep the DataFrame clean as well, so the separate spreadsheet
            # output cannot accidentally include blank template rows.
            df.at[idx, "Sub total"] = float("nan")
            df.at[idx, "VAT"] = float("nan")
            df.at[idx, "Total"] = float("nan")
            df.at[idx, VALIDATION_COLUMN] = ""
            df.at[idx, STATUS_COLUMN] = ""
            continue

        sub, vat, total = calculate_totals(row["Client Rate"], row["Mins"])
        df.at[idx, "Sub total"] = float(sub)
        df.at[idx, "VAT"] = float(vat)
        df.at[idx, "Total"] = float(total)

        issues = validate_row(df.loc[idx], expected_period)
        df.at[idx, VALIDATION_COLUMN] = issues

        if issues:
            df.at[idx, STATUS_COLUMN] = STATUS_REVIEW
        elif clean_text(df.at[idx, STATUS_COLUMN]) != STATUS_INVOICED:
            df.at[idx, STATUS_COLUMN] = STATUS_READY

    # Work only with genuine job rows that have a client name.  A genuinely
    # blank spreadsheet/template row is ignored completely; a non-blank row
    # with a missing client still receives a validation issue in the workbook.
    client_rows = df[active_row_mask & (df["Client"].apply(clean_text) != "")].copy()
    client_rows["Client"] = client_rows["Client"].apply(clean_text)

    if client_rows.empty:
        write_status_back_to_excel(excel_path, df)
        raise ValueError("No rows containing a Client were found.")

    clients = sorted(client_rows["Client"].unique(), key=str.casefold)
    print(f"Found {len(clients)} unique client(s).\n")

    results = []
    review_clients = []
    skipped_existing = []

    for client in clients:
        group = client_rows[client_rows["Client"] == client].copy()
        group_indices = group.index

        issue_mask = group[VALIDATION_COLUMN].fillna("").astype(str).str.strip() != ""

        # Client-level rule: one invalid row blocks the entire invoice.
        if issue_mask.any():
            # The problematic rows are already marked Needs Review by validate_row.
            # Valid rows stay Ready, but the existence of any issue blocks the
            # entire client's invoice until the spreadsheet is corrected.
            review_clients.append(client)

            print(f"[REVIEW] {client}")
            print(f"    {int(issue_mask.sum())} row(s) require attention")

            for idx, row in group[issue_mask].iterrows():
                excel_row = int(idx) + int(df.attrs.get("excel_header_row", 0)) + 2
                print(f"        Row {excel_row} - {row[VALIDATION_COLUMN]}")

            print("    Invoice not generated until the missing/invalid data above is corrected.\n")
            continue

        # Do not silently regenerate rows already marked as invoiced.
        if (group[STATUS_COLUMN] == STATUS_INVOICED).all():
            skipped_existing.append(client)
            print(f"[SKIP] {client} - already marked as invoiced; skipped")
            continue

        df.loc[group_indices, STATUS_COLUMN] = STATUS_READY

        print(f"[OK] {client}")
        print(f"    Generating editable invoice from {len(group)} row(s)...")

        result = generate_invoice_files(client, group, output_dir, logo_path)
        results.append(result)

        df.loc[group_indices, STATUS_COLUMN] = STATUS_INVOICED
        print(
            f"    Invoice no. {result['invoice_number']} | "
            f"Date {result['invoice_date']}"
        )

        print(f"    DOCX: {result['docx']}")
        print(f"    PDF:  {result['pdf']} ({result['pdf_method']})")
        if result.get("schedule"):
            print(f"    Schedule: {result['schedule']}")
        print(
            f"    Subtotal {format_money(result['subtotal'])} | "
            f"VAT {format_money(result['vat'])} | "
            f"Total {format_money(result['total'])}\n"
        )

    try:
        write_status_back_to_excel(excel_path, df)
        print("Master spreadsheet status updated.")
    except PermissionError:
        print(
            "[WARNING] Could not write status changes back to the master spreadsheet. "
            "It is probably still open in Excel."
        )
        print(
            "[WARNING] The invoices were still generated successfully. "
            "Close the spreadsheet before the next invoice run if you want "
            "'Invoice Generated' statuses written back."
        )

    spreadsheet_output = create_spreadsheet_output_report(df, output_dir)

    print("Finished")
    print(f"{len(results)} invoice(s) created")
    print(f"{len(review_clients)} client(s) require review")
    if skipped_existing:
        print(f"{len(skipped_existing)} client(s) were already invoiced and skipped")
    print(f"Output folder: {output_dir.resolve()}")
    print(f"Spreadsheet output view: {spreadsheet_output}")

    return results


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Generate editable DOCX invoices and PDF viewing copies from an Excel spreadsheet."
        )
    )

    parser.add_argument(
        "excel_file",
        type=Path,
        help="Path to the master Excel spreadsheet",
    )

    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_DIR,
        help="Directory in which to create invoices (default: generated_invoices)",
    )

    parser.add_argument(
        "--logo",
        type=Path,
        default=DEFAULT_LOGO,
        help=(
            "Optional logo image. If it is not found, a text-based Fingertips "
            "header is used instead (default: logo.png)."
        ),
    )

    parser.add_argument(
        "--check-warnings",
        action="store_true",
        help=(
            "Check the spreadsheet for suspicious-but-nonblocking data and exit "
            "without generating invoices. Exit code 2 means warnings were found."
        ),
    )

    parser.add_argument(
        "--allow-warnings",
        action="store_true",
        help="Generate invoices even when suspicious-data warnings are present.",
    )

    args = parser.parse_args()

    logo_path = args.logo
    if logo_path and not logo_path.is_absolute():
        # The default logo lives beside this script, not necessarily in the
        # directory PowerShell happens to be running from.
        script_relative = SCRIPT_DIR / logo_path
        if script_relative.exists():
            logo_path = script_relative

    if not logo_path or not logo_path.exists():
        print(f"WARNING: Logo not found at {logo_path}; using text header instead.")
        logo_path = None

    try:
        result = generate_invoices(
            args.excel_file,
            args.output,
            logo_path,
            allow_warnings=args.allow_warnings,
            check_warnings_only=args.check_warnings,
        )

        if args.check_warnings and result.get("warning_count", 0):
            raise SystemExit(2)

    except WarningApprovalRequired as exc:
        print(f"WARNING APPROVAL REQUIRED: {exc}")
        raise SystemExit(2)
    except Exception as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
