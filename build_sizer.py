#!/usr/bin/env python3
"""
A&S Capital Sizer — Excel Template Builder
Generates a professional loan sizing workbook with branding, formulas,
data validation dropdowns, and preserved Zillow Market Data.

Output: assets/AS_Capital_Sizer.xlsx
"""

import os
import copy
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, NamedStyle, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

# ---------------------------------------------------------------------------
# PATH CONFIGURATION
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(SCRIPT_DIR, "assets")
OUTPUT_PATH = os.path.join(ASSETS_DIR, "AS_Capital_Sizer.xlsx")
EXISTING_PATH = OUTPUT_PATH  # we read from the same file before overwriting

# ---------------------------------------------------------------------------
# COLOUR PALETTE
# ---------------------------------------------------------------------------
DEEP_BLUE   = "0B5394"
POWDER_BLUE = "A3D5E0"
LIGHT_BLUE  = "E0F0F8"
DARK_TEXT    = "2C3E50"
WHITE        = "FFFFFF"
PASS_GREEN   = "E8F8F0"
FAIL_RED     = "FDEDEC"
LIGHT_GRAY   = "F2F2F2"
MED_GRAY     = "999999"
DARK_GRAY    = "444444"
BLACK        = "000000"

# ---------------------------------------------------------------------------
# REUSABLE STYLE OBJECTS
# ---------------------------------------------------------------------------
THIN_BORDER = Border(
    left=Side(style="thin", color=BLACK),
    right=Side(style="thin", color=BLACK),
    top=Side(style="thin", color=BLACK),
    bottom=Side(style="thin", color=BLACK),
)
BOTTOM_BORDER = Border(bottom=Side(style="thin", color=MED_GRAY))

FONT_TITLE = Font(name="Calibri", size=16, bold=True, color=WHITE)
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
FONT_SUBSECTION = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
FONT_LABEL = Font(name="Calibri", size=10, color=DARK_TEXT)
FONT_INPUT = Font(name="Calibri", size=10, color=BLACK)
FONT_COMPUTED = Font(name="Calibri", size=10, color=DARK_TEXT)
FONT_COMPUTED_BOLD = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
FONT_BIG_RESULT = Font(name="Calibri", size=12, bold=True, color=WHITE)
FONT_COLHEAD = Font(name="Calibri", size=11, bold=True, color=WHITE)
FONT_NOTE = Font(name="Calibri", size=9, italic=True, color=MED_GRAY)
FONT_PASS = Font(name="Calibri", size=10, bold=True, color="27AE60")
FONT_FAIL = Font(name="Calibri", size=10, bold=True, color="E74C3C")
FONT_REF_HEADER = Font(name="Calibri", size=13, bold=True, color=WHITE)
FONT_REF_SECTION = Font(name="Calibri", size=11, bold=True, color=DEEP_BLUE)
FONT_REF = Font(name="Calibri", size=9, color=DARK_TEXT)
FONT_REF_BOLD = Font(name="Calibri", size=9, bold=True, color=DARK_TEXT)

FILL_DEEP_BLUE  = PatternFill(start_color=DEEP_BLUE, end_color=DEEP_BLUE, fill_type="solid")
FILL_POWDER     = PatternFill(start_color=POWDER_BLUE, end_color=POWDER_BLUE, fill_type="solid")
FILL_LIGHT_BLUE = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
FILL_WHITE      = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
FILL_LIGHT_GRAY = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
FILL_PASS       = PatternFill(start_color=PASS_GREEN, end_color=PASS_GREEN, fill_type="solid")
FILL_FAIL       = PatternFill(start_color=FAIL_RED, end_color=FAIL_RED, fill_type="solid")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT   = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT  = Alignment(horizontal="right", vertical="center")

FMT_CURRENCY = '$#,##0'
FMT_CURRENCY_DEC = '$#,##0.00'
FMT_PCT = '0.0%'
FMT_RATE = '0.000%'
FMT_INT = '#,##0'
FMT_DATE = 'MM/DD/YYYY'

# ---------------------------------------------------------------------------
# DROPDOWN LISTS
# ---------------------------------------------------------------------------
STATES = "AL,AK,AZ,AR,CA,CO,CT,DC,DE,FL,GA,HI,ID,IL,IN,IA,KS,KY,LA,ME,MD,MA,MI,MN,MS,MO,MT,NE,NV,NH,NJ,NM,NY,NC,ND,OH,OK,OR,PA,RI,SC,SD,TN,TX,UT,VT,VA,WA,WV,WI,WY"

DROPDOWNS = {
    "deal_type": "Fix & Flip,Bridge,Fix & Hold,Ground Up Construction",
    "transaction": "Purchase,Refinance (Rate & Term),Refinance (Cash Out)",
    "loan_term": "6 Months,12 Months,13-18 Months,19-24 Months",
    "deal_product": "Light Rehab,Heavy Rehab,Bridge,Construction",
    "property_type": "SFR,Townhome,Condo,PUD,2-4 Unit,5-10 MFR,11-20 MFR,21-50 MFR",
    "state": STATES,
    "condition": "Excellent,Good,Fair,Poor",
    "experience": "Yes,No,Limited",
    "guarantors": "1,2,3,4",
}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def _apply_input_style(cell, fmt=None):
    """White bg, thin border — an input cell."""
    cell.fill = FILL_WHITE
    cell.border = THIN_BORDER
    cell.font = FONT_INPUT
    cell.alignment = ALIGN_LEFT
    if fmt:
        cell.number_format = fmt


def _apply_computed_style(cell, fmt=None, bold=False):
    """Light blue bg — a computed/formula cell."""
    cell.fill = FILL_LIGHT_BLUE
    cell.border = THIN_BORDER
    cell.font = FONT_COMPUTED_BOLD if bold else FONT_COMPUTED
    cell.alignment = ALIGN_RIGHT
    if fmt:
        cell.number_format = fmt


def _apply_label_style(cell):
    cell.font = FONT_LABEL
    cell.alignment = ALIGN_LEFT


def _section_header(ws, row, col_start, col_end, text):
    """Powder-blue section header spanning col_start:col_end."""
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row, end_column=col_end
    )
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = FONT_SECTION
    cell.fill = FILL_POWDER
    cell.alignment = ALIGN_LEFT
    cell.border = THIN_BORDER
    # fill merged cells with same border/fill
    for c in range(col_start + 1, col_end + 1):
        mc = ws.cell(row=row, column=c)
        mc.fill = FILL_POWDER
        mc.border = THIN_BORDER


def _sub_header(ws, row, col_start, col_end, text):
    """Lighter sub-header for guarantor blocks etc."""
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row, end_column=col_end
    )
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = FONT_SUBSECTION
    cell.fill = FILL_LIGHT_GRAY
    cell.alignment = ALIGN_LEFT
    cell.border = THIN_BORDER
    for c in range(col_start + 1, col_end + 1):
        mc = ws.cell(row=row, column=c)
        mc.fill = FILL_LIGHT_GRAY
        mc.border = THIN_BORDER


def _title_row(ws, row, col_start, col_end, text):
    """Deep blue full-width title."""
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row, end_column=col_end
    )
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = FONT_TITLE
    cell.fill = FILL_DEEP_BLUE
    cell.alignment = ALIGN_CENTER
    cell.border = THIN_BORDER
    for c in range(col_start + 1, col_end + 1):
        mc = ws.cell(row=row, column=c)
        mc.fill = FILL_DEEP_BLUE
        mc.border = THIN_BORDER


def _add_dropdown(ws, cell_range, formula_list, prompt_title="", prompt_body=""):
    dv = DataValidation(
        type="list",
        formula1=f'"{formula_list}"',
        allow_blank=True,
        showDropDown=False,
    )
    if prompt_title:
        dv.prompt = prompt_body
        dv.promptTitle = prompt_title
        dv.showInputMessage = True
    dv.add(cell_range)
    ws.add_data_validation(dv)


def _label(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    _apply_label_style(c)
    return c


def _input_cell(ws, row, col, fmt=None):
    c = ws.cell(row=row, column=col)
    _apply_input_style(c, fmt)
    return c


def _formula_cell(ws, row, col, formula, fmt=None, bold=False):
    c = ws.cell(row=row, column=col, value=formula)
    _apply_computed_style(c, fmt, bold)
    return c


def _merged_input(ws, row, col_start, col_end, fmt=None):
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row, end_column=col_end
    )
    c = ws.cell(row=row, column=col_start)
    _apply_input_style(c, fmt)
    for cc in range(col_start + 1, col_end + 1):
        ws.cell(row=row, column=cc).border = THIN_BORDER
    return c


# ============================================================================
# SHEET 1: SIZER (INPUT SHEET)
# ============================================================================

def build_sizer_sheet(wb):
    ws = wb.active
    ws.title = "Sizer"

    # Column widths  A=3, B=22, C=18, D=18, E=22, F=18, G=18, H=3
    widths = {1: 3, 2: 24, 3: 20, 4: 20, 5: 24, 6: 20, 7: 20, 8: 3}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    # Row heights
    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 36

    # ---- Title ----
    _title_row(ws, 2, 1, 8, "A&S CAPITAL SIZER")

    # ==================================================================
    # DEAL INFORMATION (rows 4-8)
    # ==================================================================
    _section_header(ws, 4, 2, 3, "DEAL INFORMATION")
    _label(ws, 5, 2, "Deal Type")
    _input_cell(ws, 5, 3)
    _add_dropdown(ws, "C5", DROPDOWNS["deal_type"], "Deal Type", "Select deal type")

    _label(ws, 6, 2, "Transaction Type")
    _input_cell(ws, 6, 3)
    _add_dropdown(ws, "C6", DROPDOWNS["transaction"], "Transaction", "Select transaction type")

    _label(ws, 7, 2, "Loan Term")
    _input_cell(ws, 7, 3)
    _add_dropdown(ws, "C7", DROPDOWNS["loan_term"], "Loan Term", "Select loan term")

    _label(ws, 8, 2, "Deal Product")
    _input_cell(ws, 8, 3)
    _add_dropdown(ws, "C8", DROPDOWNS["deal_product"], "Product", "Select deal product")

    # ==================================================================
    # PROPERTY INFORMATION (rows 10-18)
    # ==================================================================
    _section_header(ws, 10, 2, 3, "PROPERTY INFORMATION")
    _label(ws, 11, 2, "Property Address")
    _merged_input(ws, 11, 3, 4)

    _label(ws, 12, 2, "City")
    _input_cell(ws, 12, 3)
    _label(ws, 12, 4, "State")
    _input_cell(ws, 12, 5)
    _add_dropdown(ws, "E12", DROPDOWNS["state"], "State", "Select state")

    _label(ws, 13, 2, "ZIP Code")
    _input_cell(ws, 13, 3, FMT_INT)
    _label(ws, 13, 4, "County")
    _input_cell(ws, 13, 5)

    _label(ws, 14, 2, "Property Type")
    _input_cell(ws, 14, 3)
    _add_dropdown(ws, "C14", DROPDOWNS["property_type"], "Property Type", "Select property type")

    _label(ws, 15, 2, "# Units")
    _input_cell(ws, 15, 3, FMT_INT)

    _label(ws, 16, 2, "Square Footage")
    _input_cell(ws, 16, 3, FMT_INT)
    _label(ws, 16, 4, "Lot Size (SF)")
    _input_cell(ws, 16, 5, FMT_INT)

    _label(ws, 17, 2, "Year Built")
    _input_cell(ws, 17, 3, "0")

    _label(ws, 18, 2, "Condition")
    _input_cell(ws, 18, 3)
    _add_dropdown(ws, "C18", DROPDOWNS["condition"], "Condition", "Select property condition")

    # ==================================================================
    # VALUATION (rows 20-25)
    # ==================================================================
    _section_header(ws, 20, 2, 3, "VALUATION")
    _label(ws, 21, 2, "Purchase Price")
    _input_cell(ws, 21, 3, FMT_CURRENCY)
    _label(ws, 22, 2, "As-Is Value")
    _input_cell(ws, 22, 3, FMT_CURRENCY)
    _label(ws, 23, 2, "After Repair Value (ARV)")
    _input_cell(ws, 23, 3, FMT_CURRENCY)
    _label(ws, 24, 2, "Rehab Budget")
    _input_cell(ws, 24, 3, FMT_CURRENCY)
    _label(ws, 25, 2, "Total Project Cost")
    _formula_cell(ws, 25, 3, "=C21+C24", FMT_CURRENCY, bold=True)

    # ---- ZHVI MARKET DATA (right side, rows 20-23) ----
    _section_header(ws, 20, 5, 6, "ZHVI MARKET DATA")
    _label(ws, 21, 5, "ZHVI (Zillow)")
    # VLOOKUP: ZIP in C13 -> Zillow Market Data C:AN column 38 (last ZHVI month)
    _formula_cell(
        ws, 21, 6,
        '=IFERROR(VLOOKUP(C13,\'Zillow Market Data\'!C:AN,38,FALSE),"")',
        FMT_CURRENCY
    )
    _label(ws, 22, 5, "Value / ZHVI Ratio")
    _formula_cell(ws, 22, 6, '=IFERROR(C22/F21,"")', '0.00x')

    _label(ws, 23, 5, "Deal vs Market")
    # Conditional text formula
    _formula_cell(
        ws, 23, 6,
        '=IF(F22="","",IF(F22>3,"HIGH RISK",IF(F22>2,"ELEVATED","NORMAL")))'
    )

    # ==================================================================
    # LOAN REQUEST (rows 28-32)
    # ==================================================================
    _section_header(ws, 28, 2, 3, "LOAN REQUEST")
    _label(ws, 29, 2, "Initial Loan Amount")
    _input_cell(ws, 29, 3, FMT_CURRENCY)
    _label(ws, 30, 2, "Rehab Holdback")
    _input_cell(ws, 30, 3, FMT_CURRENCY)
    _label(ws, 31, 2, "Interest Reserve")
    _input_cell(ws, 31, 3, FMT_CURRENCY)
    _label(ws, 32, 2, "Total Loan Amount")
    _formula_cell(ws, 32, 3, "=C29+C30+C31", FMT_CURRENCY, bold=True)

    # ---- LEVERAGE RATIOS (right side) ----
    _section_header(ws, 28, 5, 6, "LEVERAGE RATIOS")
    _label(ws, 29, 5, "LTV (Loan / As-Is Value)")
    _formula_cell(ws, 29, 6, '=IFERROR(C29/C22,"")', FMT_PCT)
    _label(ws, 30, 5, "LTC (Loan / Cost)")
    _formula_cell(ws, 30, 6, '=IFERROR(C32/C25,"")', FMT_PCT)
    _label(ws, 31, 5, "LTARV")
    _formula_cell(ws, 31, 6, '=IFERROR(C32/C23,"")', FMT_PCT)

    # ==================================================================
    # BORROWER INFORMATION (rows 36-48)
    # ==================================================================
    _section_header(ws, 36, 2, 3, "BORROWER INFORMATION")
    _label(ws, 37, 2, "Borrowing Entity")
    _merged_input(ws, 37, 3, 4)

    _label(ws, 38, 2, "# Guarantors")
    _input_cell(ws, 38, 3)
    _add_dropdown(ws, "C38", DROPDOWNS["guarantors"], "Guarantors", "Number of guarantors")

    # Guarantor 1
    _sub_header(ws, 40, 2, 3, "GUARANTOR 1")
    _label(ws, 41, 2, "Full Name")
    _merged_input(ws, 41, 3, 4)
    _label(ws, 42, 2, "FICO Score")
    _input_cell(ws, 42, 3, FMT_INT)
    _label(ws, 43, 2, "FICO Date")
    _input_cell(ws, 43, 3, FMT_DATE)

    # Guarantor 2
    _sub_header(ws, 45, 2, 3, "GUARANTOR 2")
    _label(ws, 46, 2, "Full Name")
    _merged_input(ws, 46, 3, 4)
    _label(ws, 47, 2, "FICO Score")
    _input_cell(ws, 47, 3, FMT_INT)
    _label(ws, 48, 2, "FICO Date")
    _input_cell(ws, 48, 3, FMT_DATE)

    # ==================================================================
    # EXPERIENCE & LIQUIDITY (rows 50-54)
    # ==================================================================
    _section_header(ws, 50, 2, 3, "EXPERIENCE & LIQUIDITY")
    _label(ws, 51, 2, "# Completed Projects")
    _input_cell(ws, 51, 3, FMT_INT)
    _label(ws, 52, 2, "Similar Experience")
    _input_cell(ws, 52, 3)
    _add_dropdown(ws, "C52", DROPDOWNS["experience"], "Experience", "Similar project experience?")
    _label(ws, 53, 2, "Verified Liquidity ($)")
    _input_cell(ws, 53, 3, FMT_CURRENCY)
    _label(ws, 54, 2, "Monthly PITIA")
    _input_cell(ws, 54, 3, FMT_CURRENCY)

    # ---- Note row ----
    ws.merge_cells("B57:G57")
    note = ws.cell(row=57, column=2,
        value="Complete all fields above, then review the Sizing sheet for auto-calculated leverage and pricing.")
    note.font = FONT_NOTE
    note.alignment = ALIGN_CENTER

    # ---- Print setup ----
    ws.sheet_properties.pageSetUpPr = None
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.print_area = "A1:H57"
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)

    return ws


# ============================================================================
# SHEET 2: SIZING (AUTO-CALCULATING SUMMARY)
# ============================================================================

def build_sizing_sheet(wb):
    ws = wb.create_sheet("Sizing")

    # Column widths: A=3, B=28, C=20, D=20, E=3, F=28, G=20
    widths = {1: 3, 2: 30, 3: 22, 4: 22, 5: 3, 6: 30, 7: 22}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 36

    # ---- Title ----
    _title_row(ws, 2, 1, 7, "A&S CAPITAL \u2014 DEAL SIZING")

    # ==================================================================
    # DEAL OVERVIEW (rows 4-12)
    # ==================================================================
    _section_header(ws, 4, 2, 4, "DEAL OVERVIEW")

    overview_rows = [
        (5, "Deal Type",     '=Sizer!C5'),
        (6, "Transaction",   '=Sizer!C6'),
        (7, "Property",      '=Sizer!C11&", "&Sizer!C12&", "&Sizer!E12&" "&Sizer!C13'),
        (8, "Property Type",  '=Sizer!C14&" ("&Sizer!C15&" units)"'),
        (9, "FICO Score",    '=Sizer!C42'),
        (10, "Experience",   '=Sizer!C51&" projects"'),
        (11, "ZHVI Ratio",   '=Sizer!F22'),
        (12, "Product",      '=Sizer!C8'),
    ]
    for r, lbl, fml in overview_rows:
        _label(ws, r, 2, lbl)
        c = ws.cell(row=r, column=3, value=fml)
        c.font = FONT_COMPUTED
        c.fill = FILL_LIGHT_BLUE
        c.border = THIN_BORDER
        c.alignment = ALIGN_LEFT
        if r == 9:
            c.number_format = FMT_INT
        if r == 11:
            c.number_format = '0.00x'

    # ==================================================================
    # VALUATION (rows 14-20)
    # ==================================================================
    _section_header(ws, 14, 2, 4, "VALUATION")

    val_rows = [
        (15, "Purchase Price",         '=Sizer!C21'),
        (16, "As-Is Value",            '=Sizer!C22'),
        (17, "ARV",                    '=Sizer!C23'),
        (18, "Rehab Budget",           '=Sizer!C24'),
        (19, "Total Project Cost",     '=Sizer!C25'),
        (20, "Borrower Request (Total)", '=Sizer!C32'),
    ]
    for r, lbl, fml in val_rows:
        _label(ws, r, 2, lbl)
        _formula_cell(ws, r, 3, fml, FMT_CURRENCY)

    # ==================================================================
    # SIDE-BY-SIDE: COLCHIS vs FIDELIS (rows 22-35)
    # ==================================================================
    # Column headers
    for col, text in [(3, "COLCHIS CAPITAL"), (6, "FIDELIS INVESTORS")]:
        c = ws.cell(row=22, column=col, value=text)
        c.font = FONT_COLHEAD
        c.fill = FILL_DEEP_BLUE
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # Sub-headers row 23
    sub_headers = [
        (3, "Max Leverage"), (4, "Max $ Amount"),
        (6, "Max Leverage"), (7, "Max $ Amount"),
    ]
    for col, text in sub_headers:
        c = ws.cell(row=23, column=col, value=text)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # --- Colchis leverage rows (25-31) ---
    leverage_labels = [
        (25, "Max LTV (As-Is)"),
        (26, "Max LTC"),
        (27, "Max LTARV"),
    ]
    # Colchis value refs: As-Is=C16, TPC=C19, ARV=C17 on this sheet
    colchis_dollar_formulas = [
        '=IFERROR(C25*C16,"")',  # LTV * As-Is
        '=IFERROR(C26*C19,"")',  # LTC * TPC
        '=IFERROR(C27*C17,"")',  # LTARV * ARV
    ]
    fidelis_dollar_formulas = [
        '=IFERROR(F25*C16,"")',
        '=IFERROR(F26*C19,"")',
        '=IFERROR(F27*C17,"")',
    ]

    for i, (r, lbl) in enumerate(leverage_labels):
        _label(ws, r, 2, lbl)
        # Colchis leverage input
        _input_cell(ws, r, 3, FMT_PCT)
        # Colchis dollar amount (computed)
        _formula_cell(ws, r, 4, colchis_dollar_formulas[i], FMT_CURRENCY)
        # Fidelis leverage input
        _input_cell(ws, r, 6, FMT_PCT)
        # Fidelis dollar amount (computed)
        _formula_cell(ws, r, 7, fidelis_dollar_formulas[i], FMT_CURRENCY)

    # Guidelines Max Loan
    _label(ws, 29, 2, "Guidelines Max Loan")
    ws.cell(row=29, column=2).font = FONT_COMPUTED_BOLD
    _formula_cell(ws, 29, 4, '=IFERROR(MIN(D25,D26,D27),"")', FMT_CURRENCY, bold=True)
    _formula_cell(ws, 29, 7, '=IFERROR(MIN(G25,G26,G27),"")', FMT_CURRENCY, bold=True)

    # Max Loan Cap
    _label(ws, 30, 2, "Max Loan Amount Cap")
    cap_col = ws.cell(row=30, column=4, value=3500000)
    cap_col.number_format = FMT_CURRENCY
    cap_col.font = FONT_COMPUTED
    cap_col.fill = FILL_LIGHT_BLUE
    cap_col.border = THIN_BORDER
    cap_col.alignment = ALIGN_RIGHT
    cap_fid = ws.cell(row=30, column=7, value=5000000)
    cap_fid.number_format = FMT_CURRENCY
    cap_fid.font = FONT_COMPUTED
    cap_fid.fill = FILL_LIGHT_BLUE
    cap_fid.border = THIN_BORDER
    cap_fid.alignment = ALIGN_RIGHT

    # FINAL MAX LOAN (big, bold, deep blue)
    _label(ws, 31, 2, "FINAL MAX LOAN")
    ws.cell(row=31, column=2).font = Font(name="Calibri", size=12, bold=True, color=DARK_TEXT)
    for col, fml in [(4, '=IFERROR(MIN(D29,D30),"")'), (7, '=IFERROR(MIN(G29,G30),"")' )]:
        c = ws.cell(row=col, column=col)
        c = ws.cell(row=31, column=col, value=fml)
        c.font = FONT_BIG_RESULT
        c.fill = FILL_DEEP_BLUE
        c.number_format = FMT_CURRENCY
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # Estimated Rate / Points
    _label(ws, 33, 2, "Estimated Rate")
    _input_cell(ws, 33, 3, FMT_RATE)
    _input_cell(ws, 33, 6, FMT_RATE)
    _label(ws, 34, 2, "Points (Origination)")
    _input_cell(ws, 34, 3, FMT_PCT)
    _input_cell(ws, 34, 6, FMT_PCT)

    # ==================================================================
    # LOAN PROCEEDS COMPARISON (rows 37-45)
    # ==================================================================
    _section_header(ws, 37, 2, 7, "LOAN PROCEEDS COMPARISON")

    # Sub-header row
    proceed_heads = [
        (3, "Borrower Request"), (4, "Colchis Max"),
        (6, "Fidelis Max"), (7, "Col. vs Fid."),
    ]
    for col, txt in proceed_heads:
        c = ws.cell(row=38, column=col, value=txt)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    proceed_items = [
        (39, "Initial Loan Amount",  '=Sizer!C29', '=IFERROR(MIN(Sizer!C29,D31),"")',
             '=IFERROR(MIN(Sizer!C29,G31),"")', '=IFERROR(F39-D39,"")'),
        (40, "Rehab Holdback",       '=Sizer!C30', '=IFERROR(MIN(Sizer!C30,D31-D39),"")',
             '=IFERROR(MIN(Sizer!C30,G31-F39),"")', '=IFERROR(F40-D40,"")'),
        (41, "Interest Reserve",     '=Sizer!C31', '=IFERROR(MIN(Sizer!C31,D31-D39-D40),"")',
             '=IFERROR(MIN(Sizer!C31,G31-F39-F40),"")', '=IFERROR(F41-D41,"")'),
        (42, "Total Loan Amount",    '=SUM(C39:C41)', '=SUM(D39:D41)',
             '=SUM(F39:F41)', '=IFERROR(F42-D42,"")'),
    ]
    for r, lbl, c_fml, d_fml, f_fml, g_fml in proceed_items:
        _label(ws, r, 2, lbl)
        _formula_cell(ws, r, 3, c_fml, FMT_CURRENCY)
        _formula_cell(ws, r, 4, d_fml, FMT_CURRENCY)
        _formula_cell(ws, r, 6, f_fml, FMT_CURRENCY)
        _formula_cell(ws, r, 7, g_fml, FMT_CURRENCY)
        if r == 42:
            for cc in [2, 3, 4, 6, 7]:
                ws.cell(row=r, column=cc).font = FONT_COMPUTED_BOLD

    # Actual ratios
    _label(ws, 44, 2, "Actual LTV")
    _formula_cell(ws, 44, 3, '=IFERROR(C39/C16,"")', FMT_PCT)
    _formula_cell(ws, 44, 4, '=IFERROR(D42/C16,"")', FMT_PCT)
    _formula_cell(ws, 44, 6, '=IFERROR(F42/C16,"")', FMT_PCT)

    _label(ws, 45, 2, "Actual LTARV")
    _formula_cell(ws, 45, 3, '=IFERROR(C42/C17,"")', FMT_PCT)
    _formula_cell(ws, 45, 4, '=IFERROR(D42/C17,"")', FMT_PCT)
    _formula_cell(ws, 45, 6, '=IFERROR(F42/C17,"")', FMT_PCT)

    # ==================================================================
    # GUIDELINE CHECKS (rows 47-55)
    # ==================================================================
    _section_header(ws, 47, 2, 7, "GUIDELINE CHECKS")

    # Sub-header
    check_heads = [(3, "Actual Value"), (4, "Colchis"), (6, "Fidelis")]
    for col, txt in check_heads:
        c = ws.cell(row=48, column=col, value=txt)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # --- Row 49: Min FICO ---
    _label(ws, 49, 2, "Min FICO")
    _formula_cell(ws, 49, 3, '=C9', FMT_INT)
    # Colchis min FICO = 680
    c49d = ws.cell(row=49, column=4, value='=IF(C9="","",IF(C9>=680,"PASS","FAIL"))')
    c49d.border = THIN_BORDER
    c49d.alignment = ALIGN_CENTER
    c49d.font = FONT_COMPUTED
    # Fidelis min FICO = 660
    c49f = ws.cell(row=49, column=6, value='=IF(C9="","",IF(C9>=660,"PASS","FAIL"))')
    c49f.border = THIN_BORDER
    c49f.alignment = ALIGN_CENTER
    c49f.font = FONT_COMPUTED

    # --- Row 50: Min Loan Amount ---
    _label(ws, 50, 2, "Min Loan Amount")
    _formula_cell(ws, 50, 3, '=C20', FMT_CURRENCY)
    c50d = ws.cell(row=50, column=4, value='=IF(C20="","",IF(C20>=100000,"PASS","FAIL"))')
    c50d.border = THIN_BORDER; c50d.alignment = ALIGN_CENTER; c50d.font = FONT_COMPUTED
    c50f = ws.cell(row=50, column=6, value='=IF(C20="","",IF(C20>=75000,"PASS","FAIL"))')
    c50f.border = THIN_BORDER; c50f.alignment = ALIGN_CENTER; c50f.font = FONT_COMPUTED

    # --- Row 51: Max Loan Amount ---
    _label(ws, 51, 2, "Max Loan Amount")
    _formula_cell(ws, 51, 3, '=C20', FMT_CURRENCY)
    c51d = ws.cell(row=51, column=4, value='=IF(C20="","",IF(C20<=3500000,"PASS","FAIL"))')
    c51d.border = THIN_BORDER; c51d.alignment = ALIGN_CENTER; c51d.font = FONT_COMPUTED
    c51f = ws.cell(row=51, column=6, value='=IF(C20="","",IF(C20<=5000000,"PASS","FAIL"))')
    c51f.border = THIN_BORDER; c51f.alignment = ALIGN_CENTER; c51f.font = FONT_COMPUTED

    # --- Row 52: State Eligible ---
    _label(ws, 52, 2, "State Eligible?")
    _formula_cell(ws, 52, 3, '=Sizer!E12')
    # Colchis excludes IL
    c52d = ws.cell(row=52, column=4,
        value='=IF(Sizer!E12="","",IF(OR(Sizer!E12="IL",Sizer!E12="NV",Sizer!E12="ND",Sizer!E12="SD",Sizer!E12="VT"),"FAIL","PASS"))')
    c52d.border = THIN_BORDER; c52d.alignment = ALIGN_CENTER; c52d.font = FONT_COMPUTED
    c52f = ws.cell(row=52, column=6, value='=IF(Sizer!E12="","","PASS")')
    c52f.border = THIN_BORDER; c52f.alignment = ALIGN_CENTER; c52f.font = FONT_COMPUTED

    # --- Row 53: Overall ---
    _label(ws, 53, 2, "Overall Eligibility")
    ws.cell(row=53, column=2).font = FONT_COMPUTED_BOLD
    c53d = ws.cell(row=53, column=4,
        value='=IF(OR(D49="FAIL",D50="FAIL",D51="FAIL",D52="FAIL"),"FAIL","PASS")')
    c53d.border = THIN_BORDER; c53d.alignment = ALIGN_CENTER
    c53d.font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
    c53f = ws.cell(row=53, column=6,
        value='=IF(OR(F49="FAIL",F50="FAIL",F51="FAIL",F52="FAIL"),"FAIL","PASS")')
    c53f.border = THIN_BORDER; c53f.alignment = ALIGN_CENTER
    c53f.font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)

    # ---- Print setup ----
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = "A1:G55"
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)

    return ws


# ============================================================================
# SHEET 3: COLCHIS LEVERAGE (Reference Table)
# ============================================================================

def build_colchis_leverage_sheet(wb):
    ws = wb.create_sheet("Colchis Leverage")

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 3
    ws.column_dimensions["G"].width = 22
    ws.column_dimensions["H"].width = 22
    ws.column_dimensions["I"].width = 22

    ws.row_dimensions[1].height = 30

    # Title
    ws.merge_cells("A1:I1")
    c = ws.cell(row=1, column=1, value="COLCHIS CAPITAL \u2014 LEVERAGE GUIDELINES")
    c.font = FONT_REF_HEADER
    c.fill = FILL_DEEP_BLUE
    c.alignment = ALIGN_CENTER
    for cc in range(2, 10):
        ws.cell(row=1, column=cc).fill = FILL_DEEP_BLUE

    # ---- Helper to build a grid ----
    def _grid(start_row, title, col_headers, data_rows, note=""):
        r = start_row
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c = ws.cell(row=r, column=2, value=title)
        c.font = FONT_REF_SECTION
        c.fill = FILL_LIGHT_GRAY
        c.alignment = ALIGN_LEFT
        for cc in range(3, 6):
            ws.cell(row=r, column=cc).fill = FILL_LIGHT_GRAY
        r += 1

        if note:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            ws.cell(row=r, column=2, value=note).font = FONT_NOTE
            r += 1

        # Column headers
        for i, h in enumerate(col_headers):
            c = ws.cell(row=r, column=2 + i, value=h)
            c.font = FONT_REF_BOLD
            c.fill = FILL_POWDER
            c.border = THIN_BORDER
            c.alignment = ALIGN_CENTER
        r += 1

        # Data rows
        for row_data in data_rows:
            for i, val in enumerate(row_data):
                c = ws.cell(row=r, column=2 + i, value=val)
                c.font = FONT_REF
                c.border = THIN_BORDER
                c.alignment = ALIGN_CENTER if i > 0 else ALIGN_LEFT
            r += 1

        return r + 1  # next available row

    # =====================================================================
    # SINGLE FAMILY GRIDS
    # =====================================================================
    row = 3
    row = _grid(row,
        "SINGLE FAMILY \u2014 LIGHT REHAB (Purchase)",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",   "90% / 92.5% / 75%", "90% / 92.5% / 75%", "90% / 90% / 75%"],
            ["720-739", "90% / 90% / 75%",   "87.5% / 90% / 70%", "85% / 90% / 70%"],
            ["700-719", "85% / 90% / 70%",   "85% / 87.5% / 70%", "80% / 85% / 65%"],
            ["680-699", "80% / 85% / 65%",   "80% / 85% / 65%",   "75% / 80% / 65%"],
            ["660-679", "75% / 80% / 65%",   "N/A",               "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "SINGLE FAMILY \u2014 HEAVY REHAB (Purchase)",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",   "90% / 92.5% / 75%", "87.5% / 92.5% / 75%", "85% / 90% / 70%"],
            ["720-739", "87.5% / 90% / 70%", "85% / 90% / 70%",     "82.5% / 87.5% / 70%"],
            ["700-719", "85% / 87.5% / 70%", "82.5% / 87.5% / 70%", "80% / 85% / 65%"],
            ["680-699", "80% / 85% / 65%",   "77.5% / 82.5% / 65%", "75% / 80% / 65%"],
            ["660-679", "75% / 80% / 65%",   "N/A",                 "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "SINGLE FAMILY \u2014 BRIDGE (Purchase)",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",   "80% / 75%", "80% / 75%", "75% / 70%"],
            ["720-739", "80% / 75%", "77.5% / 70%", "75% / 70%"],
            ["700-719", "75% / 70%", "75% / 70%", "72.5% / 65%"],
            ["680-699", "72.5% / 65%", "72.5% / 65%", "70% / 65%"],
        ],
        note="Format: Max LTV / Max LTARV"
    )

    row = _grid(row,
        "SINGLE FAMILY \u2014 GROUND UP CONSTRUCTION",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",   "90% LTC / 75% LTARV", "90% LTC / 75% LTARV", "85% LTC / 70% LTARV"],
            ["720-739", "90% LTC / 70% LTARV", "87.5% LTC / 70% LTARV", "85% LTC / 70% LTARV"],
            ["700-719", "87.5% LTC / 70% LTARV", "85% LTC / 70% LTARV", "82.5% LTC / 65% LTARV"],
            ["680-699", "85% LTC / 65% LTARV", "82.5% LTC / 65% LTARV", "80% LTC / 65% LTARV"],
        ],
    )

    # =====================================================================
    # MULTIFAMILY GRIDS
    # =====================================================================
    row = _grid(row,
        "MULTIFAMILY (5+) \u2014 LIGHT REHAB / BRIDGE",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",   "80% / 85% / 70%", "77.5% / 82.5% / 70%", "75% / 80% / 65%"],
            ["720-739", "77.5% / 82.5% / 70%", "75% / 80% / 65%", "72.5% / 77.5% / 65%"],
            ["700-719", "75% / 80% / 65%", "72.5% / 77.5% / 65%", "70% / 75% / 60%"],
            ["680-699", "72.5% / 77.5% / 65%", "70% / 75% / 60%", "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    # =====================================================================
    # PRICING GRID
    # =====================================================================
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    c = ws.cell(row=row, column=2, value="COLCHIS PRICING GRID")
    c.font = FONT_REF_SECTION
    c.fill = FILL_LIGHT_GRAY
    for cc in range(3, 6):
        ws.cell(row=row, column=cc).fill = FILL_LIGHT_GRAY
    row += 1

    pricing_headers = ["FICO", "Tier 1 (8+ Proj)", "Tier 2 (4-7 Proj)", "Tier 3 (0-3 Proj)"]
    for i, h in enumerate(pricing_headers):
        c = ws.cell(row=row, column=2 + i, value=h)
        c.font = FONT_REF_BOLD
        c.fill = FILL_POWDER
        c.border = THIN_BORDER
        c.alignment = ALIGN_CENTER
    row += 1

    pricing_data = [
        ["740+",    "9.50% + 1.0pt", "10.00% + 1.5pt", "10.50% + 2.0pt"],
        ["720-739", "10.00% + 1.5pt", "10.50% + 1.5pt", "11.00% + 2.0pt"],
        ["700-719", "10.50% + 1.5pt", "11.00% + 2.0pt", "11.50% + 2.0pt"],
        ["680-699", "11.00% + 2.0pt", "11.50% + 2.0pt", "12.00% + 2.5pt"],
        ["660-679", "11.50% + 2.0pt", "N/A", "N/A"],
    ]
    for rd in pricing_data:
        for i, val in enumerate(rd):
            c = ws.cell(row=row, column=2 + i, value=val)
            c.font = FONT_REF
            c.border = THIN_BORDER
            c.alignment = ALIGN_CENTER if i > 0 else ALIGN_LEFT
        row += 1

    # General notes
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    c = ws.cell(row=row, column=2,
        value="Loan Range: $100K - $3.5M  |  Terms: 12-24 months  |  Min FICO: 680 (660 Tier 1 only)")
    c.font = FONT_NOTE
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    c = ws.cell(row=row, column=2,
        value="Excluded states: IL, NV, ND, SD, VT  |  Prepay: None  |  Extension: 3-6 mo at 1pt")
    c.font = FONT_NOTE

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    return ws


# ============================================================================
# SHEET 4: FIDELIS LEVERAGE (Reference Table)
# ============================================================================

def build_fidelis_leverage_sheet(wb):
    ws = wb.create_sheet("Fidelis Leverage")

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 3
    ws.column_dimensions["G"].width = 22
    ws.column_dimensions["H"].width = 22
    ws.column_dimensions["I"].width = 22

    ws.row_dimensions[1].height = 30

    # Title
    ws.merge_cells("A1:I1")
    c = ws.cell(row=1, column=1, value="FIDELIS INVESTORS \u2014 LEVERAGE GUIDELINES")
    c.font = FONT_REF_HEADER
    c.fill = FILL_DEEP_BLUE
    c.alignment = ALIGN_CENTER
    for cc in range(2, 10):
        ws.cell(row=1, column=cc).fill = FILL_DEEP_BLUE

    def _grid(start_row, title, col_headers, data_rows, note=""):
        r = start_row
        span = len(col_headers)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2 + span - 1)
        c = ws.cell(row=r, column=2, value=title)
        c.font = FONT_REF_SECTION
        c.fill = FILL_LIGHT_GRAY
        c.alignment = ALIGN_LEFT
        for cc in range(3, 2 + span):
            ws.cell(row=r, column=cc).fill = FILL_LIGHT_GRAY
        r += 1

        if note:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2 + span - 1)
            ws.cell(row=r, column=2, value=note).font = FONT_NOTE
            r += 1

        for i, h in enumerate(col_headers):
            c = ws.cell(row=r, column=2 + i, value=h)
            c.font = FONT_REF_BOLD
            c.fill = FILL_POWDER
            c.border = THIN_BORDER
            c.alignment = ALIGN_CENTER
        r += 1

        for row_data in data_rows:
            for i, val in enumerate(row_data):
                c = ws.cell(row=r, column=2 + i, value=val)
                c.font = FONT_REF
                c.border = THIN_BORDER
                c.alignment = ALIGN_CENTER if i > 0 else ALIGN_LEFT
            r += 1

        return r + 1

    # =====================================================================
    # NATIONAL PROGRAM (excluding FL, CA, NY)
    # =====================================================================
    row = 3
    row = _grid(row,
        "NATIONAL PROGRAM \u2014 FIX & FLIP / BRIDGE (excl. FL, CA, NY)",
        ["FICO", "Exp 5+", "Exp 3-4", "Exp 1-2", "Exp 0"],
        [
            ["760+",   "90% / 90% / 75%", "87.5% / 87.5% / 72.5%", "85% / 85% / 70%", "82.5% / 82.5% / 67.5%"],
            ["740-759", "87.5% / 87.5% / 72.5%", "85% / 85% / 70%", "82.5% / 82.5% / 67.5%", "80% / 80% / 65%"],
            ["720-739", "85% / 85% / 70%", "82.5% / 82.5% / 67.5%", "80% / 80% / 65%", "77.5% / 77.5% / 62.5%"],
            ["700-719", "82.5% / 82.5% / 67.5%", "80% / 80% / 65%", "77.5% / 77.5% / 62.5%", "75% / 75% / 60%"],
            ["680-699", "80% / 80% / 65%", "77.5% / 77.5% / 62.5%", "75% / 75% / 60%", "72.5% / 72.5% / 57.5%"],
            ["660-679", "77.5% / 77.5% / 62.5%", "75% / 75% / 60%", "N/A", "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "FLORIDA PROGRAM \u2014 FIX & FLIP / BRIDGE",
        ["FICO", "Exp 5+", "Exp 3-4", "Exp 1-2", "Exp 0"],
        [
            ["760+",   "85% / 87.5% / 72.5%", "82.5% / 85% / 70%", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%"],
            ["740-759", "82.5% / 85% / 70%", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%"],
            ["720-739", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%"],
            ["700-719", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%"],
            ["680-699", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%", "N/A"],
            ["660-679", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%", "N/A", "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "CA / NY PROGRAM \u2014 FIX & FLIP / BRIDGE",
        ["FICO", "Exp 5+", "Exp 3-4", "Exp 1-2", "Exp 0"],
        [
            ["760+",   "85% / 87.5% / 72.5%", "82.5% / 85% / 70%", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%"],
            ["740-759", "82.5% / 85% / 70%", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%"],
            ["720-739", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%"],
            ["700-719", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%"],
            ["680-699", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%", "N/A"],
            ["660-679", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%", "N/A", "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    # MULTIFAMILY
    row = _grid(row,
        "MULTIFAMILY (5-20 UNITS) \u2014 NATIONAL",
        ["FICO", "Exp 5+", "Exp 3-4", "Exp 1-2"],
        [
            ["740+",   "80% / 82.5% / 70%", "77.5% / 80% / 67.5%", "75% / 77.5% / 65%"],
            ["720-739", "77.5% / 80% / 67.5%", "75% / 77.5% / 65%", "72.5% / 75% / 62.5%"],
            ["700-719", "75% / 77.5% / 65%", "72.5% / 75% / 62.5%", "70% / 72.5% / 60%"],
            ["680-699", "72.5% / 75% / 62.5%", "70% / 72.5% / 60%", "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    # Ground Up Construction
    row = _grid(row,
        "GROUND UP CONSTRUCTION \u2014 NATIONAL",
        ["FICO", "Exp 5+", "Exp 3-4", "Exp 1-2"],
        [
            ["740+",   "87.5% LTC / 72.5% LTARV", "85% LTC / 70% LTARV", "82.5% LTC / 67.5% LTARV"],
            ["720-739", "85% LTC / 70% LTARV", "82.5% LTC / 67.5% LTARV", "80% LTC / 65% LTARV"],
            ["700-719", "82.5% LTC / 67.5% LTARV", "80% LTC / 65% LTARV", "77.5% LTC / 62.5% LTARV"],
            ["680-699", "80% LTC / 65% LTARV", "77.5% LTC / 62.5% LTARV", "75% LTC / 60% LTARV"],
        ],
    )

    # Pricing
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    c = ws.cell(row=row, column=2, value="FIDELIS PRICING GRID")
    c.font = FONT_REF_SECTION
    c.fill = FILL_LIGHT_GRAY
    for cc in range(3, 7):
        ws.cell(row=row, column=cc).fill = FILL_LIGHT_GRAY
    row += 1

    pricing_headers = ["FICO", "Tier 1 (5+ Proj)", "Tier 2 (3-4 Proj)", "Tier 3 (1-2 Proj)", "Tier 4 (0 Proj)"]
    for i, h in enumerate(pricing_headers):
        c = ws.cell(row=row, column=2 + i, value=h)
        c.font = FONT_REF_BOLD
        c.fill = FILL_POWDER
        c.border = THIN_BORDER
        c.alignment = ALIGN_CENTER
    row += 1

    pricing_data = [
        ["760+",    "9.25% + 1.0pt", "9.75% + 1.0pt", "10.25% + 1.5pt", "10.75% + 2.0pt"],
        ["740-759", "9.75% + 1.0pt", "10.25% + 1.5pt", "10.75% + 1.5pt", "11.25% + 2.0pt"],
        ["720-739", "10.25% + 1.5pt", "10.75% + 1.5pt", "11.25% + 2.0pt", "11.75% + 2.0pt"],
        ["700-719", "10.75% + 1.5pt", "11.25% + 2.0pt", "11.75% + 2.0pt", "12.25% + 2.5pt"],
        ["680-699", "11.25% + 2.0pt", "11.75% + 2.0pt", "12.25% + 2.5pt", "12.75% + 2.5pt"],
        ["660-679", "11.75% + 2.0pt", "12.25% + 2.5pt", "N/A", "N/A"],
    ]
    for rd in pricing_data:
        for i, val in enumerate(rd):
            c = ws.cell(row=row, column=2 + i, value=val)
            c.font = FONT_REF
            c.border = THIN_BORDER
            c.alignment = ALIGN_CENTER if i > 0 else ALIGN_LEFT
        row += 1

    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    c = ws.cell(row=row, column=2,
        value="Loan Range: $75K - $5M  |  Terms: 6-24 months  |  Min FICO: 660 (Tier 1-2 only)")
    c.font = FONT_NOTE
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    c = ws.cell(row=row, column=2,
        value="All states eligible  |  Prepay: None  |  Extension: Available at 0.5-1.0pt")
    c.font = FONT_NOTE

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    return ws


# ============================================================================
# SHEET 5: ZILLOW MARKET DATA (copy from existing file)
# ============================================================================

def copy_zillow_data(wb, existing_path):
    """
    Read the Zillow Market Data sheet from the existing workbook and
    write every cell (value + basic number format) into the new workbook.
    """
    print("  Reading existing Zillow Market Data ...")
    src = load_workbook(existing_path, read_only=True, data_only=True)
    src_ws = src["Zillow Market Data"]

    ws = wb.create_sheet("Zillow Market Data")

    total_rows = src_ws.max_row
    total_cols = src_ws.max_column

    print(f"  Copying {total_rows:,} rows x {total_cols} columns ...")

    for row_idx, row in enumerate(src_ws.iter_rows(min_row=1, max_row=total_rows,
                                                    max_col=total_cols,
                                                    values_only=True), start=1):
        for col_idx, val in enumerate(row, start=1):
            if val is None:
                continue
            new_cell = ws.cell(row=row_idx, column=col_idx, value=val)

            # Style the header row
            if row_idx == 1:
                new_cell.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
                new_cell.fill = FILL_DEEP_BLUE
                new_cell.alignment = ALIGN_CENTER
                new_cell.border = THIN_BORDER
            else:
                new_cell.font = Font(name="Calibri", size=9, color=DARK_TEXT)
                # Format date header columns' values as currency
                if col_idx >= 10:  # Column J onward = ZHVI values
                    new_cell.number_format = '$#,##0'

        if row_idx % 5000 == 0:
            print(f"    ... {row_idx:,} / {total_rows:,} rows")

    # Freeze header row
    ws.freeze_panes = "A2"

    # Auto-filter
    ws.auto_filter.ref = f"A1:{get_column_letter(total_cols)}{total_rows}"

    # Column widths for key columns
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 6
    ws.column_dimensions["G"].width = 16
    ws.column_dimensions["H"].width = 36
    ws.column_dimensions["I"].width = 20

    src.close()
    print("  Zillow Market Data copied successfully.")
    return ws


# ============================================================================
# MAIN BUILD FUNCTION
# ============================================================================

def build_sizer_workbook():
    print("=" * 60)
    print("A&S Capital Sizer \u2014 Excel Template Builder")
    print("=" * 60)

    # Ensure output directory exists
    os.makedirs(ASSETS_DIR, exist_ok=True)

    # ------------------------------------------------------------------
    # Step 1: Read Zillow data from existing file (if it exists)
    # ------------------------------------------------------------------
    has_zillow = False
    if os.path.exists(EXISTING_PATH):
        try:
            test_wb = load_workbook(EXISTING_PATH, read_only=True)
            if "Zillow Market Data" in test_wb.sheetnames:
                has_zillow = True
                print(f"[OK] Found existing Zillow Market Data in {EXISTING_PATH}")
            test_wb.close()
        except Exception as e:
            print(f"[WARN] Could not read existing file: {e}")

    # ------------------------------------------------------------------
    # Step 2: Create new workbook
    # ------------------------------------------------------------------
    wb = Workbook()

    print("\n[1/5] Building Sizer (Input) sheet ...")
    build_sizer_sheet(wb)

    print("[2/5] Building Sizing (Calculation) sheet ...")
    build_sizing_sheet(wb)

    print("[3/5] Building Colchis Leverage sheet ...")
    build_colchis_leverage_sheet(wb)

    print("[4/5] Building Fidelis Leverage sheet ...")
    build_fidelis_leverage_sheet(wb)

    print("[5/5] Copying Zillow Market Data sheet ...")
    if has_zillow:
        copy_zillow_data(wb, EXISTING_PATH)
    else:
        # Create placeholder sheet
        ws = wb.create_sheet("Zillow Market Data")
        ws.cell(row=1, column=1, value="RegionName")
        ws.cell(row=1, column=2, value="ZHVI")
        ws.cell(row=1, column=1).font = Font(bold=True)
        ws.cell(row=1, column=2).font = Font(bold=True)
        print("  [WARN] No existing Zillow data found. Created placeholder sheet.")

    # ------------------------------------------------------------------
    # Step 3: Set active sheet to Sizer
    # ------------------------------------------------------------------
    wb.active = wb.sheetnames.index("Sizer")

    # ------------------------------------------------------------------
    # Step 4: Save
    # ------------------------------------------------------------------
    print(f"\nSaving workbook to: {OUTPUT_PATH}")
    wb.save(OUTPUT_PATH)
    file_size = os.path.getsize(OUTPUT_PATH)
    print(f"[DONE] File saved successfully ({file_size / 1024 / 1024:.1f} MB)")
    print("=" * 60)


# ============================================================================
if __name__ == "__main__":
    build_sizer_workbook()
