#!/usr/bin/env python3
"""
A&S Capital Sizer -- Excel Template Builder  (v2 - Auto-Calculating)
Generates a professional loan sizing workbook with:
  - Auto-calculating leverage maximums via VLOOKUP against hidden lookup tables
  - Auto-calculated pricing via embedded Colchis pricing grid
  - ZHVI market data integration
  - Full guideline pass/fail checks

Output: assets/AS_Capital_Sizer.xlsx
"""

import os
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
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
EXISTING_PATH = OUTPUT_PATH

# ---------------------------------------------------------------------------
# COLOUR PALETTE  (NO YELLOW ANYWHERE)
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

FONT_TITLE         = Font(name="Calibri", size=16, bold=True, color=WHITE)
FONT_SECTION        = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
FONT_SUBSECTION     = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
FONT_LABEL          = Font(name="Calibri", size=10, color=DARK_TEXT)
FONT_INPUT          = Font(name="Calibri", size=10, color=BLACK)
FONT_COMPUTED       = Font(name="Calibri", size=10, color=DARK_TEXT)
FONT_COMPUTED_BOLD  = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
FONT_BIG_RESULT     = Font(name="Calibri", size=12, bold=True, color=WHITE)
FONT_COLHEAD        = Font(name="Calibri", size=11, bold=True, color=WHITE)
FONT_NOTE           = Font(name="Calibri", size=9, italic=True, color=MED_GRAY)
FONT_PASS           = Font(name="Calibri", size=10, bold=True, color="27AE60")
FONT_FAIL           = Font(name="Calibri", size=10, bold=True, color="E74C3C")
FONT_REF_HEADER     = Font(name="Calibri", size=13, bold=True, color=WHITE)
FONT_REF_SECTION    = Font(name="Calibri", size=11, bold=True, color=DEEP_BLUE)
FONT_REF            = Font(name="Calibri", size=9, color=DARK_TEXT)
FONT_REF_BOLD       = Font(name="Calibri", size=9, bold=True, color=DARK_TEXT)

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

FMT_CURRENCY     = '$#,##0'
FMT_CURRENCY_DEC = '$#,##0.00'
FMT_PCT          = '0.0%'
FMT_RATE         = '0.000%'
FMT_INT          = '#,##0'
FMT_DATE         = 'MM/DD/YYYY'
FMT_TEXT         = '@'

# ---------------------------------------------------------------------------
# DROPDOWN LISTS
# ---------------------------------------------------------------------------
STATES = "AL,AK,AZ,AR,CA,CO,CT,DC,DE,FL,GA,HI,ID,IL,IN,IA,KS,KY,LA,ME,MD,MA,MI,MN,MS,MO,MT,NE,NV,NH,NJ,NM,NY,NC,ND,OH,OK,OR,PA,RI,SC,SD,TN,TX,UT,VT,VA,WA,WV,WI,WY"

DROPDOWNS = {
    "deal_type":     "Fix & Flip,Bridge,Fix & Hold,Ground Up Construction",
    "transaction":   "Purchase,Refinance (Rate & Term),Refinance (Cash Out)",
    "loan_term":     "12 Months,18 Months,24 Months",
    "deal_product":  "Light Rehab,Heavy Rehab,Bridge,Construction",
    "property_type": "SFR,Townhome,Condo,PUD,2-4 Unit,5-10 MFR,11-20 MFR,21-50 MFR",
    "state":         STATES,
    "experience":    "Yes,No,Limited",
    "guarantors":    "1,2,3,4",
}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def _apply_input_style(cell, fmt=None):
    """White bg, thin border -- an input cell."""
    cell.fill = FILL_WHITE
    cell.border = THIN_BORDER
    cell.font = FONT_INPUT
    cell.alignment = ALIGN_LEFT
    if fmt:
        cell.number_format = fmt


def _apply_computed_style(cell, fmt=None, bold=False):
    """Light blue bg -- a computed/formula cell."""
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
    ws.sheet_properties.tabColor = None  # No tab color

    # Column widths  A=3, B=22, C=20, D=20, E=22, F=20, G=20, H=3
    widths = {1: 3, 2: 22, 3: 20, 4: 20, 5: 22, 6: 20, 7: 20, 8: 3}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

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
    # PROPERTY INFORMATION (rows 10-17)
    # No Condition field, no County field
    # ==================================================================
    _section_header(ws, 10, 2, 3, "PROPERTY INFORMATION")

    _label(ws, 11, 2, "Property Address")
    _merged_input(ws, 11, 3, 4)

    _label(ws, 12, 2, "City")
    _input_cell(ws, 12, 3)
    _label(ws, 12, 4, "State")
    _input_cell(ws, 12, 5)
    _add_dropdown(ws, "E12", DROPDOWNS["state"], "State", "Select state")

    # ZIP Code -- TEXT format, no commas
    _label(ws, 13, 2, "ZIP Code")
    c_zip = _input_cell(ws, 13, 3, FMT_TEXT)

    _label(ws, 14, 2, "Property Type")
    _input_cell(ws, 14, 3)
    _add_dropdown(ws, "C14", DROPDOWNS["property_type"], "Property Type", "Select property type")

    _label(ws, 15, 2, "# Units")
    _input_cell(ws, 15, 3, FMT_INT)

    _label(ws, 16, 2, "Square Footage")
    _input_cell(ws, 16, 3, FMT_INT)
    _label(ws, 16, 4, "Lot Size (SF)")
    _input_cell(ws, 16, 5, FMT_INT)

    # Year Built -- auto-fills current year for Ground Up Construction
    _label(ws, 17, 2, "Year Built")
    _formula_cell(
        ws, 17, 3,
        '=IF(C5="Ground Up Construction",YEAR(TODAY()),"")',
        "0"
    )
    # Note next to Year Built
    note_yb = ws.cell(row=17, column=4, value="Auto-fills for Ground Up Construction")
    note_yb.font = FONT_NOTE
    note_yb.alignment = ALIGN_LEFT

    # ==================================================================
    # VALUATION (rows 19-25)   with ZHVI on right side
    # No "Condition" or "County" fields anywhere
    # ==================================================================
    _section_header(ws, 19, 2, 3, "VALUATION")

    _label(ws, 20, 2, "Purchase Price")
    _input_cell(ws, 20, 3, FMT_CURRENCY)

    _label(ws, 21, 2, "Purchase Date")
    _input_cell(ws, 21, 3, FMT_DATE)

    _label(ws, 22, 2, "As-Is Value")
    _input_cell(ws, 22, 3, FMT_CURRENCY)

    _label(ws, 23, 2, "After Repair Value (ARV)")
    _input_cell(ws, 23, 3, FMT_CURRENCY)

    _label(ws, 24, 2, "Rehab Budget")
    _input_cell(ws, 24, 3, FMT_CURRENCY)

    _label(ws, 25, 2, "Total Project Cost")
    _formula_cell(ws, 25, 3, "=C20+C24", FMT_CURRENCY, bold=True)

    # ---- ZHVI MARKET DATA (right side, rows 19-22) ----
    _section_header(ws, 19, 5, 6, "ZHVI MARKET DATA")

    _label(ws, 20, 5, "ZHVI (Zillow)")
    _formula_cell(
        ws, 20, 6,
        '=IFERROR(VLOOKUP(C13,\'Zillow Market Data\'!C:AN,38,FALSE),"")',
        FMT_CURRENCY
    )

    _label(ws, 21, 5, "Value / ZHVI Ratio")
    _formula_cell(ws, 21, 6, '=IFERROR(C22/F20,"")', '0.00"x"')

    _label(ws, 22, 5, "Deal vs Market")
    _formula_cell(
        ws, 22, 6,
        '=IF(F21="","",IF(F21>3,"HIGH RISK",IF(F21>2,"ELEVATED","NORMAL")))'
    )

    # ==================================================================
    # LOAN REQUEST (rows 27-31)  with leverage ratios on right
    # ==================================================================
    _section_header(ws, 27, 2, 3, "LOAN REQUEST")

    _label(ws, 28, 2, "Initial Loan Amount")
    _input_cell(ws, 28, 3, FMT_CURRENCY)

    _label(ws, 29, 2, "Rehab Holdback")
    _input_cell(ws, 29, 3, FMT_CURRENCY)

    _label(ws, 30, 2, "Interest Reserve")
    _input_cell(ws, 30, 3, FMT_CURRENCY)

    _label(ws, 31, 2, "Total Loan Amount")
    _formula_cell(ws, 31, 3, "=C28+C29+C30", FMT_CURRENCY, bold=True)

    # ---- LEVERAGE RATIOS (right side) ----
    _section_header(ws, 27, 5, 6, "LEVERAGE RATIOS")

    _label(ws, 28, 5, "LTV (Loan / As-Is Value)")
    _formula_cell(ws, 28, 6, '=IFERROR(C28/C22,"")', FMT_PCT)

    _label(ws, 29, 5, "LTC (Loan / Cost)")
    _formula_cell(ws, 29, 6, '=IFERROR(C31/C25,"")', FMT_PCT)

    _label(ws, 30, 5, "LTARV")
    _formula_cell(ws, 30, 6, '=IFERROR(C31/C23,"")', FMT_PCT)

    # ==================================================================
    # BORROWER INFORMATION (rows 33-46)
    # No "Verified Liquidity" or "Monthly PITIA"
    # ==================================================================
    _section_header(ws, 33, 2, 3, "BORROWER INFORMATION")

    _label(ws, 34, 2, "Borrowing Entity")
    _merged_input(ws, 34, 3, 4)

    _label(ws, 35, 2, "# Guarantors")
    _input_cell(ws, 35, 3)
    _add_dropdown(ws, "C35", DROPDOWNS["guarantors"], "Guarantors", "Number of guarantors")

    # Guarantor 1
    _sub_header(ws, 37, 2, 3, "GUARANTOR 1")
    _label(ws, 38, 2, "Full Name")
    _merged_input(ws, 38, 3, 4)
    _label(ws, 39, 2, "FICO Score")
    _input_cell(ws, 39, 3, FMT_INT)

    # Guarantor 2
    _sub_header(ws, 41, 2, 3, "GUARANTOR 2")
    _label(ws, 42, 2, "Full Name")
    _merged_input(ws, 42, 3, 4)
    _label(ws, 43, 2, "FICO Score")
    _input_cell(ws, 43, 3, FMT_INT)

    # ==================================================================
    # EXPERIENCE (rows 45-48)
    # No "Verified Liquidity", no "Monthly PITIA"
    # ==================================================================
    _section_header(ws, 45, 2, 3, "EXPERIENCE")

    _label(ws, 46, 2, "# Completed Projects")
    _input_cell(ws, 46, 3, FMT_INT)

    _label(ws, 47, 2, "Similar Experience")
    _input_cell(ws, 47, 3)
    _add_dropdown(ws, "C47", DROPDOWNS["experience"], "Experience", "Similar project experience?")

    # ---- Note row ----
    ws.merge_cells("B50:G50")
    note = ws.cell(
        row=50, column=2,
        value="Complete all fields above, then review the Sizing sheet for auto-calculated leverage and pricing."
    )
    note.font = FONT_NOTE
    note.alignment = ALIGN_CENTER

    # ---- Print setup ----
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.print_area = "A1:H50"
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)

    return ws


# ============================================================================
# SHEET 2: SIZING (AUTO-CALCULATING SUMMARY)
# ============================================================================

def build_sizing_sheet(wb):
    ws = wb.create_sheet("Sizing")
    ws.sheet_properties.tabColor = None

    # Column widths: A=3, B=30, C=22, D=22, E=5, F=22, G=22
    widths = {1: 3, 2: 30, 3: 22, 4: 22, 5: 5, 6: 22, 7: 22}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 36

    # ---- Title ----
    _title_row(ws, 2, 1, 7, "A&S CAPITAL \u2014 DEAL SIZING & PRICING")

    # ==================================================================
    # DEAL OVERVIEW (rows 4-13)
    # ==================================================================
    _section_header(ws, 4, 2, 4, "DEAL OVERVIEW")

    overview_rows = [
        (5,  "Deal Type",              '=Sizer!C5'),
        (6,  "Transaction",            '=Sizer!C6'),
        (7,  "Product",                '=Sizer!C8'),
        (8,  "Property",               '=Sizer!C11&", "&Sizer!C12&", "&Sizer!E12&" "&Sizer!C13'),
        (9,  "Property Type",          '=Sizer!C14'),
        (10, "Units",                  '=Sizer!C15'),
        (11, "FICO Score",             '=Sizer!C39'),
        (12, "Experience (# Projects)",'=Sizer!C46'),
        (13, "ZHVI Ratio",             '=Sizer!F21'),
    ]
    for r, lbl, fml in overview_rows:
        _label(ws, r, 2, lbl)
        c = ws.cell(row=r, column=3, value=fml)
        c.font = FONT_COMPUTED
        c.fill = FILL_LIGHT_BLUE
        c.border = THIN_BORDER
        c.alignment = ALIGN_LEFT
        if r == 11:
            c.number_format = FMT_INT
        if r == 10:
            c.number_format = FMT_INT
        if r == 12:
            c.number_format = FMT_INT
        if r == 13:
            c.number_format = '0.00"x"'

    # ==================================================================
    # VALUATION (rows 15-21)
    # ==================================================================
    _section_header(ws, 15, 2, 4, "VALUATION")

    val_rows = [
        (16, "Purchase Price",          '=Sizer!C20'),
        (17, "As-Is Value",             '=Sizer!C22'),
        (18, "ARV",                     '=Sizer!C23'),
        (19, "Rehab Budget",            '=Sizer!C24'),
        (20, "Total Project Cost",      '=Sizer!C25'),
        (21, "Borrower Request (Total)",'=Sizer!C31'),
    ]
    for r, lbl, fml in val_rows:
        _label(ws, r, 2, lbl)
        _formula_cell(ws, r, 3, fml, FMT_CURRENCY)

    # ==================================================================
    # COLCHIS CLASSIFICATION (rows 23-27)
    # Auto-determine experience tier, FICO bucket, lookup key
    # ==================================================================
    _section_header(ws, 23, 2, 4, "COLCHIS CLASSIFICATION")

    _label(ws, 24, 2, "Experience Tier")
    _formula_cell(
        ws, 24, 3,
        '=IF(C12="","",IF(C12>=8,"8+",IF(C12>=4,"4-7","0-3")))',
    )

    _label(ws, 25, 2, "FICO Bucket")
    _formula_cell(
        ws, 25, 3,
        '=IF(C11="","",IF(C11>=740,"740+",IF(C11>=700,"700-739",IF(C11>=680,"680-699","<680 (Ineligible)"))))',
    )

    _label(ws, 26, 2, "Product Category")
    _formula_cell(
        ws, 26, 3,
        '=IF(C7="","",C7)',
    )

    _label(ws, 27, 2, "Leverage Lookup Key")
    _formula_cell(
        ws, 27, 3,
        '=IF(OR(C26="",C25="",C24=""),"",C26&"|"&C25&"|"&C24)',
    )

    # ==================================================================
    # COLCHIS LEVERAGE LIMITS (rows 29-37)
    # Auto-looked-up from hidden "Colchis Leverage Data" sheet
    # ==================================================================
    _section_header(ws, 29, 2, 4, "COLCHIS LEVERAGE LIMITS")

    # Sub-header row
    for col, txt in [(3, "Max Leverage"), (4, "Max $ Amount")]:
        c = ws.cell(row=30, column=col, value=txt)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # Max LTV (As-Is)
    _label(ws, 31, 2, "Max LTV (As-Is)")
    _formula_cell(
        ws, 31, 3,
        '=IFERROR(VLOOKUP(C27,\'Colchis Leverage Data\'!A:D,2,FALSE),"")',
        FMT_PCT
    )
    _formula_cell(
        ws, 31, 4,
        '=IFERROR(C31*C17,"")',
        FMT_CURRENCY
    )

    # Max LTC
    _label(ws, 32, 2, "Max LTC")
    _formula_cell(
        ws, 32, 3,
        '=IFERROR(VLOOKUP(C27,\'Colchis Leverage Data\'!A:D,3,FALSE),"")',
        FMT_PCT
    )
    _formula_cell(
        ws, 32, 4,
        '=IFERROR(C32*C20,"")',
        FMT_CURRENCY
    )

    # Max LTARV
    _label(ws, 33, 2, "Max LTARV")
    _formula_cell(
        ws, 33, 3,
        '=IFERROR(VLOOKUP(C27,\'Colchis Leverage Data\'!A:D,4,FALSE),"")',
        FMT_PCT
    )
    _formula_cell(
        ws, 33, 4,
        '=IFERROR(C33*C18,"")',
        FMT_CURRENCY
    )

    # Guidelines Max Loan
    _label(ws, 35, 2, "Guidelines Max Loan")
    ws.cell(row=35, column=2).font = FONT_COMPUTED_BOLD
    _formula_cell(ws, 35, 4, '=IFERROR(MIN(D31,D32,D33),"")', FMT_CURRENCY, bold=True)

    # Max Loan Amount Cap
    _label(ws, 36, 2, "Max Loan Amount Cap")
    cap_cell = ws.cell(row=36, column=4, value=3500000)
    cap_cell.number_format = FMT_CURRENCY
    cap_cell.font = FONT_COMPUTED
    cap_cell.fill = FILL_LIGHT_BLUE
    cap_cell.border = THIN_BORDER
    cap_cell.alignment = ALIGN_RIGHT

    # FINAL MAX LOAN -- big, deep blue bg, white text
    _label(ws, 37, 2, "FINAL MAX LOAN")
    ws.cell(row=37, column=2).font = Font(name="Calibri", size=12, bold=True, color=DARK_TEXT)
    c_final = ws.cell(row=37, column=4, value='=IFERROR(MIN(D35,D36),"")')
    c_final.font = FONT_BIG_RESULT
    c_final.fill = FILL_DEEP_BLUE
    c_final.number_format = FMT_CURRENCY
    c_final.alignment = ALIGN_CENTER
    c_final.border = THIN_BORDER

    # ==================================================================
    # LOAN PROCEEDS CALCULATION (rows 39-47)
    # ==================================================================
    _section_header(ws, 39, 2, 4, "LOAN PROCEEDS CALCULATION")

    # Sub-header row 40
    for col, txt in [(3, "Borrower Requested"), (4, "Guidelines Max")]:
        c = ws.cell(row=40, column=col, value=txt)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # Initial Loan Amount
    _label(ws, 41, 2, "Initial Loan Amount")
    _formula_cell(ws, 41, 3, '=Sizer!C28', FMT_CURRENCY)
    _formula_cell(ws, 41, 4, '=IFERROR(MIN(Sizer!C28,D37),"")', FMT_CURRENCY)

    # Financed Rehab Budget
    _label(ws, 42, 2, "Financed Rehab Budget")
    _formula_cell(ws, 42, 3, '=Sizer!C29', FMT_CURRENCY)
    _formula_cell(ws, 42, 4, '=IFERROR(MIN(Sizer!C29,D37-D41),"")', FMT_CURRENCY)

    # Interest Reserve
    _label(ws, 43, 2, "Interest Reserve")
    _formula_cell(ws, 43, 3, '=Sizer!C30', FMT_CURRENCY)
    _formula_cell(ws, 43, 4, '=IFERROR(MIN(Sizer!C30,D37-D41-D42),"")', FMT_CURRENCY)

    # Loan Amount (total row, bold)
    _label(ws, 44, 2, "Loan Amount")
    ws.cell(row=44, column=2).font = FONT_COMPUTED_BOLD
    _formula_cell(ws, 44, 3, '=SUM(C41:C43)', FMT_CURRENCY, bold=True)
    _formula_cell(ws, 44, 4, '=SUM(D41:D43)', FMT_CURRENCY, bold=True)

    # Actual ratios row 46-47
    _label(ws, 46, 2, "Actual LTV (Req / Max)")
    _formula_cell(ws, 46, 3, '=IFERROR(C41/C17,"")', FMT_PCT)
    _formula_cell(ws, 46, 4, '=IFERROR(D44/C17,"")', FMT_PCT)

    _label(ws, 47, 2, "Actual LTARV (Req / Max)")
    _formula_cell(ws, 47, 3, '=IFERROR(C44/C18,"")', FMT_PCT)
    _formula_cell(ws, 47, 4, '=IFERROR(D44/C18,"")', FMT_PCT)

    # ==================================================================
    # COLCHIS PRICING (rows 49-55)
    # ==================================================================
    _section_header(ws, 49, 2, 4, "COLCHIS BUY RATE")

    # Build a pricing lookup key: "Product|FICO Bucket|LTC Bucket"
    # LTC bucket based on actual LTC from the deal
    _label(ws, 50, 2, "Actual LTC %")
    _formula_cell(
        ws, 50, 3,
        '=IFERROR(C44/C20,"")',
        FMT_PCT
    )

    _label(ws, 51, 2, "LTC Bucket")
    _formula_cell(
        ws, 51, 3,
        '=IF(C50="","",IF(C50<=0.7,"<=70%",IF(C50<=0.75,"<=75%",IF(C50<=0.8,"<=80%",IF(C50<=0.85,"<=85%",IF(C50<=0.9,"<=90%","<=95%"))))))',
    )

    _label(ws, 52, 2, "Pricing Key")
    _formula_cell(
        ws, 52, 3,
        '=IF(OR(C26="",C25="",C51=""),"",C26&"|"&C25&"|"&C51)',
    )

    _label(ws, 53, 2, "Base Rate (Buy Rate)")
    _formula_cell(
        ws, 53, 3,
        '=IFERROR(VLOOKUP(C52,\'Colchis Leverage Data\'!F:G,2,FALSE),"")',
        FMT_RATE
    )

    _label(ws, 54, 2, "Loan Interest Rate")
    _formula_cell(
        ws, 54, 3,
        '=IFERROR(C53+0.005,"")',
        FMT_RATE
    )
    # Note about spread
    note_rate = ws.cell(row=54, column=4, value="(Buy rate + 50bps spread)")
    note_rate.font = FONT_NOTE
    note_rate.alignment = ALIGN_LEFT

    # ==================================================================
    # FINAL CREDIT CHECK (rows 56-68)
    # ==================================================================
    _section_header(ws, 56, 2, 4, "FINAL CREDIT CHECK")

    # Sub-header
    for col, txt in [(3, "Actual Value"), (4, "Result")]:
        c = ws.cell(row=57, column=col, value=txt)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # --- Check rows ---
    # Min FICO (680)
    _label(ws, 58, 2, "Min FICO (680)")
    _formula_cell(ws, 58, 3, '=C11', FMT_INT)
    c58 = ws.cell(row=58, column=4, value='=IF(C11="","",IF(C11>=680,"PASS","FAIL"))')
    c58.border = THIN_BORDER; c58.alignment = ALIGN_CENTER; c58.font = FONT_COMPUTED

    # Max Loan ($3.5M)
    _label(ws, 59, 2, "Max Loan ($3.5M)")
    _formula_cell(ws, 59, 3, '=D44', FMT_CURRENCY)
    c59 = ws.cell(row=59, column=4, value='=IF(D44="","",IF(D44<=3500000,"PASS","FAIL"))')
    c59.border = THIN_BORDER; c59.alignment = ALIGN_CENTER; c59.font = FONT_COMPUTED

    # Min Loan ($100K)
    _label(ws, 60, 2, "Min Loan ($100K)")
    _formula_cell(ws, 60, 3, '=D44', FMT_CURRENCY)
    c60 = ws.cell(row=60, column=4, value='=IF(D44="","",IF(D44>=100000,"PASS","FAIL"))')
    c60.border = THIN_BORDER; c60.alignment = ALIGN_CENTER; c60.font = FONT_COMPUTED

    # State Eligible (Colchis excludes IL, NV, ND, SD, VT)
    _label(ws, 61, 2, "State Eligible")
    _formula_cell(ws, 61, 3, '=Sizer!E12')
    c61 = ws.cell(
        row=61, column=4,
        value='=IF(Sizer!E12="","",IF(OR(Sizer!E12="IL",Sizer!E12="NV",Sizer!E12="ND",Sizer!E12="SD",Sizer!E12="VT"),"FAIL","PASS"))'
    )
    c61.border = THIN_BORDER; c61.alignment = ALIGN_CENTER; c61.font = FONT_COMPUTED

    # Leverage Eligible (lookup returned a value)
    _label(ws, 62, 2, "Leverage Eligible")
    _formula_cell(ws, 62, 3, '=C27')
    c62 = ws.cell(
        row=62, column=4,
        value='=IF(C27="","",IF(AND(C31<>"",C32<>"",C33<>""),"PASS","FAIL - Ineligible Combo"))'
    )
    c62.border = THIN_BORDER; c62.alignment = ALIGN_CENTER; c62.font = FONT_COMPUTED

    # LTV Check
    _label(ws, 63, 2, "LTV Within Limits")
    _formula_cell(ws, 63, 3, '=IFERROR(C41/C17,"")', FMT_PCT)
    c63 = ws.cell(
        row=63, column=4,
        value='=IF(OR(C63="",C31=""),"",IF(C63<=C31,"PASS","FAIL"))'
    )
    c63.border = THIN_BORDER; c63.alignment = ALIGN_CENTER; c63.font = FONT_COMPUTED

    # LTC Check
    _label(ws, 64, 2, "LTC Within Limits")
    _formula_cell(ws, 64, 3, '=IFERROR(C44/C20,"")', FMT_PCT)
    c64 = ws.cell(
        row=64, column=4,
        value='=IF(OR(C64="",C32=""),"",IF(C64<=C32,"PASS","FAIL"))'
    )
    c64.border = THIN_BORDER; c64.alignment = ALIGN_CENTER; c64.font = FONT_COMPUTED

    # LTARV Check
    _label(ws, 65, 2, "LTARV Within Limits")
    _formula_cell(ws, 65, 3, '=IFERROR(C44/C18,"")', FMT_PCT)
    c65 = ws.cell(
        row=65, column=4,
        value='=IF(OR(C65="",C33=""),"",IF(C65<=C33,"PASS","FAIL"))'
    )
    c65.border = THIN_BORDER; c65.alignment = ALIGN_CENTER; c65.font = FONT_COMPUTED

    # MASTER CHECK
    _label(ws, 67, 2, "MASTER CHECK")
    ws.cell(row=67, column=2).font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
    c_master = ws.cell(
        row=67, column=4,
        value='=IF(COUNTBLANK(D58:D65)=8,"",IF(COUNTIF(D58:D65,"FAIL")+COUNTIF(D58:D65,"FAIL*")>0,"FAIL","PASS"))'
    )
    c_master.border = THIN_BORDER
    c_master.alignment = ALIGN_CENTER
    c_master.font = Font(name="Calibri", size=12, bold=True, color=DARK_TEXT)

    # ---- Print setup ----
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = "A1:G68"
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)

    return ws


# ============================================================================
# SHEET 3: COLCHIS LEVERAGE DATA  (HIDDEN lookup table)
# ============================================================================

def build_colchis_leverage_data_sheet(wb):
    """
    Hidden sheet with two lookup tables:
      Columns A-D: Leverage lookup  (Key | MaxLTV | MaxLTC | MaxLTARV)
      Columns F-G: Pricing lookup   (Key | BaseRate)
    """
    ws = wb.create_sheet("Colchis Leverage Data")

    # Column widths
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 3
    ws.column_dimensions["F"].width = 40
    ws.column_dimensions["G"].width = 12

    # ---- LEVERAGE TABLE HEADER ----
    headers_lev = [("A", "Lookup Key"), ("B", "Max LTV"), ("C", "Max LTC"), ("D", "Max LTARV")]
    for col_letter, hdr in headers_lev:
        c = ws.cell(row=1, column=ord(col_letter) - 64, value=hdr)
        c.font = FONT_REF_BOLD
        c.fill = FILL_POWDER
        c.border = THIN_BORDER
        c.alignment = ALIGN_CENTER

    # ---- PRICING TABLE HEADER ----
    headers_pr = [("F", "Pricing Key"), ("G", "Base Rate")]
    for col_letter, hdr in headers_pr:
        c = ws.cell(row=1, column=ord(col_letter) - 64, value=hdr)
        c.font = FONT_REF_BOLD
        c.fill = FILL_POWDER
        c.border = THIN_BORDER
        c.alignment = ALIGN_CENTER

    # ================================================================
    # LEVERAGE DATA
    # Format: "Product|FICO Bucket|Experience Tier" -> LTV, LTC, LTARV
    # ================================================================
    leverage_data = [
        # ---- SF Light Rehab ----
        ("Light Rehab|740+|8+",     0.900, 0.925, 0.750),
        ("Light Rehab|740+|4-7",    0.900, 0.925, 0.750),
        ("Light Rehab|740+|0-3",    0.900, 0.900, 0.750),
        ("Light Rehab|700-739|8+",  0.900, 0.925, 0.750),
        ("Light Rehab|700-739|4-7", 0.900, 0.925, 0.750),
        ("Light Rehab|700-739|0-3", 0.875, 0.900, 0.750),
        ("Light Rehab|680-699|8+",  0.875, 0.900, 0.750),
        ("Light Rehab|680-699|4-7", 0.850, 0.875, 0.750),
        ("Light Rehab|680-699|0-3", 0.850, 0.850, 0.700),

        # ---- SF Heavy Rehab ----
        ("Heavy Rehab|740+|8+",     0.800, 0.850, 0.700),
        ("Heavy Rehab|740+|4-7",    0.800, 0.850, 0.700),
        # 740+ / 0-3: ineligible -- omitted so VLOOKUP returns #N/A -> ""
        ("Heavy Rehab|700-739|8+",  0.800, 0.850, 0.700),
        ("Heavy Rehab|700-739|4-7", 0.800, 0.850, 0.700),
        # 700-739 / 0-3: ineligible
        ("Heavy Rehab|680-699|8+",  0.750, 0.825, 0.650),
        ("Heavy Rehab|680-699|4-7", 0.750, 0.800, 0.650),
        # 680-699 / 0-3: ineligible

        # ---- SF Bridge ----
        ("Bridge|740+|8+",     0.750, 0.750, 0.750),
        ("Bridge|740+|4-7",    0.750, 0.750, 0.750),
        ("Bridge|740+|0-3",    0.750, 0.750, 0.750),
        ("Bridge|700-739|8+",  0.750, 0.750, 0.750),
        ("Bridge|700-739|4-7", 0.750, 0.750, 0.750),
        ("Bridge|700-739|0-3", 0.700, 0.700, 0.700),
        ("Bridge|680-699|8+",  0.700, 0.700, 0.700),
        ("Bridge|680-699|4-7", 0.700, 0.700, 0.700),
        ("Bridge|680-699|0-3", 0.650, 0.650, 0.650),

        # ---- SF Construction ----
        # Construction uses different experience tiers: 6+, 4-5, 0-3
        # But our FICO/Experience tier formulas map to 8+, 4-7, 0-3.
        # For construction, 8+ maps to "6+" tier, 4-7 maps to "4-5" tier.
        ("Construction|740+|8+",     0.600, 0.900, 0.700),
        ("Construction|740+|4-7",    0.600, 0.850, 0.700),
        # 740+ / 0-3: ineligible
        ("Construction|700-739|8+",  0.600, 0.900, 0.700),
        ("Construction|700-739|4-7", 0.600, 0.850, 0.700),
        # 700-739 / 0-3: ineligible
        ("Construction|680-699|8+",  0.600, 0.850, 0.700),
        ("Construction|680-699|4-7", 0.600, 0.825, 0.650),
        # 680-699 / 0-3: ineligible
    ]

    row = 2
    for key, ltv, ltc, ltarv in leverage_data:
        ws.cell(row=row, column=1, value=key).font = FONT_REF
        ws.cell(row=row, column=1).border = THIN_BORDER

        c_ltv = ws.cell(row=row, column=2, value=ltv)
        c_ltv.number_format = FMT_PCT
        c_ltv.font = FONT_REF
        c_ltv.border = THIN_BORDER
        c_ltv.alignment = ALIGN_CENTER

        c_ltc = ws.cell(row=row, column=3, value=ltc)
        c_ltc.number_format = FMT_PCT
        c_ltc.font = FONT_REF
        c_ltc.border = THIN_BORDER
        c_ltc.alignment = ALIGN_CENTER

        c_ltarv = ws.cell(row=row, column=4, value=ltarv)
        c_ltarv.number_format = FMT_PCT
        c_ltarv.font = FONT_REF
        c_ltarv.border = THIN_BORDER
        c_ltarv.alignment = ALIGN_CENTER

        row += 1

    # ================================================================
    # PRICING DATA
    # Format: "Product|FICO Bucket|LTC Bucket" -> Base Rate
    # ================================================================
    pricing_data = [
        # Light Rehab pricing
        ("Light Rehab|740+|<=70%",     0.07750),
        ("Light Rehab|740+|<=75%",     0.07750),
        ("Light Rehab|740+|<=80%",     0.07750),
        ("Light Rehab|740+|<=85%",     0.07875),
        ("Light Rehab|740+|<=90%",     0.08000),
        ("Light Rehab|740+|<=95%",     0.08250),
        ("Light Rehab|700-739|<=70%",  0.07750),
        ("Light Rehab|700-739|<=75%",  0.07750),
        ("Light Rehab|700-739|<=80%",  0.07875),
        ("Light Rehab|700-739|<=85%",  0.08000),
        ("Light Rehab|700-739|<=90%",  0.08125),
        ("Light Rehab|700-739|<=95%",  0.08375),
        ("Light Rehab|680-699|<=70%",  0.07875),
        ("Light Rehab|680-699|<=75%",  0.08000),
        ("Light Rehab|680-699|<=80%",  0.08125),
        ("Light Rehab|680-699|<=85%",  0.08250),
        ("Light Rehab|680-699|<=90%",  0.08375),
        # 680-699 / <=95%: N/A -- omitted

        # Heavy Rehab pricing (use same grid shifted up by 25bps)
        ("Heavy Rehab|740+|<=70%",     0.08000),
        ("Heavy Rehab|740+|<=75%",     0.08000),
        ("Heavy Rehab|740+|<=80%",     0.08125),
        ("Heavy Rehab|740+|<=85%",     0.08250),
        ("Heavy Rehab|740+|<=90%",     0.08500),
        ("Heavy Rehab|700-739|<=70%",  0.08000),
        ("Heavy Rehab|700-739|<=75%",  0.08125),
        ("Heavy Rehab|700-739|<=80%",  0.08250),
        ("Heavy Rehab|700-739|<=85%",  0.08375),
        ("Heavy Rehab|700-739|<=90%",  0.08500),
        ("Heavy Rehab|680-699|<=70%",  0.08250),
        ("Heavy Rehab|680-699|<=75%",  0.08375),
        ("Heavy Rehab|680-699|<=80%",  0.08500),
        ("Heavy Rehab|680-699|<=85%",  0.08625),

        # Bridge pricing
        ("Bridge|740+|<=70%",     0.07500),
        ("Bridge|740+|<=75%",     0.07750),
        ("Bridge|700-739|<=70%",  0.07750),
        ("Bridge|700-739|<=75%",  0.07750),
        ("Bridge|680-699|<=70%",  0.08000),

        # Construction pricing
        ("Construction|740+|<=70%",     0.08250),
        ("Construction|740+|<=75%",     0.08500),
        ("Construction|740+|<=80%",     0.08750),
        ("Construction|740+|<=85%",     0.09000),
        ("Construction|740+|<=90%",     0.09250),
        ("Construction|700-739|<=70%",  0.08500),
        ("Construction|700-739|<=75%",  0.08750),
        ("Construction|700-739|<=80%",  0.09000),
        ("Construction|700-739|<=85%",  0.09250),
        ("Construction|700-739|<=90%",  0.09500),
        ("Construction|680-699|<=70%",  0.08750),
        ("Construction|680-699|<=75%",  0.09000),
        ("Construction|680-699|<=80%",  0.09250),
        ("Construction|680-699|<=85%",  0.09500),
    ]

    pr_row = 2
    for key, rate in pricing_data:
        ws.cell(row=pr_row, column=6, value=key).font = FONT_REF
        ws.cell(row=pr_row, column=6).border = THIN_BORDER

        c_rate = ws.cell(row=pr_row, column=7, value=rate)
        c_rate.number_format = FMT_RATE
        c_rate.font = FONT_REF
        c_rate.border = THIN_BORDER
        c_rate.alignment = ALIGN_CENTER

        pr_row += 1

    # HIDE this sheet
    ws.sheet_state = 'hidden'

    return ws


# ============================================================================
# SHEET 4: COLCHIS LEVERAGE (Human-Readable Reference)
# ============================================================================

def build_colchis_leverage_sheet(wb):
    ws = wb.create_sheet("Colchis Leverage")
    ws.sheet_properties.tabColor = None

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

    row = 3
    row = _grid(row,
        "SINGLE FAMILY \u2014 LIGHT REHAB (Purchase)",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",    "90% / 92.5% / 75%", "90% / 92.5% / 75%", "90% / 90% / 75%"],
            ["700-739", "90% / 92.5% / 75%", "90% / 92.5% / 75%", "87.5% / 90% / 75%"],
            ["680-699", "87.5% / 90% / 75%", "85% / 87.5% / 75%", "85% / 85% / 70%"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "SINGLE FAMILY \u2014 HEAVY REHAB (Purchase)",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",    "80% / 85% / 70%",   "80% / 85% / 70%",   "N/A"],
            ["700-739", "80% / 85% / 70%",   "80% / 85% / 70%",   "N/A"],
            ["680-699", "75% / 82.5% / 65%", "75% / 80% / 65%",   "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "SINGLE FAMILY \u2014 BRIDGE (Purchase)",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",    "75%", "75%", "75%"],
            ["700-739", "75%", "75%", "70%"],
            ["680-699", "70%", "70%", "65%"],
        ],
        note="Format: Max LTV"
    )

    row = _grid(row,
        "SINGLE FAMILY \u2014 GROUND UP CONSTRUCTION",
        ["FICO", "Exp 8+ (6+)", "Exp 4-7 (4-5)", "Exp 0-3"],
        [
            ["740+",    "60% LTV / 90% LTC / 70% LTARV", "60% LTV / 85% LTC / 70% LTARV", "N/A"],
            ["700-739", "60% LTV / 90% LTC / 70% LTARV", "60% LTV / 85% LTC / 70% LTARV", "N/A"],
            ["680-699", "60% LTV / 85% LTC / 70% LTARV", "60% LTV / 82.5% LTC / 65% LTARV", "N/A"],
        ],
    )

    # Pricing Grid
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    c = ws.cell(row=row, column=2, value="COLCHIS PRICING GRID (Light Rehab Buy Rate)")
    c.font = FONT_REF_SECTION
    c.fill = FILL_LIGHT_GRAY
    for cc in range(3, 6):
        ws.cell(row=row, column=cc).fill = FILL_LIGHT_GRAY
    row += 1

    pricing_headers = ["FICO", "<=70% LTC", "<=80% LTC", "<=90% LTC"]
    for i, h in enumerate(pricing_headers):
        c = ws.cell(row=row, column=2 + i, value=h)
        c.font = FONT_REF_BOLD
        c.fill = FILL_POWDER
        c.border = THIN_BORDER
        c.alignment = ALIGN_CENTER
    row += 1

    pricing_data = [
        ["740+",    "7.750%", "7.750%", "8.000%"],
        ["700-739", "7.750%", "7.875%", "8.125%"],
        ["680-699", "7.875%", "8.125%", "8.375%"],
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
        value="Loan Range: $100K - $3.5M  |  Terms: 12-24 months  |  Min FICO: 680")
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
# SHEET 5: FIDELIS LEVERAGE (Human-Readable Reference)
# ============================================================================

def build_fidelis_leverage_sheet(wb):
    ws = wb.create_sheet("Fidelis Leverage")
    ws.sheet_properties.tabColor = None

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 22
    ws.column_dimensions["G"].width = 3

    ws.row_dimensions[1].height = 30

    ws.merge_cells("A1:G1")
    c = ws.cell(row=1, column=1, value="FIDELIS INVESTORS \u2014 LEVERAGE GUIDELINES")
    c.font = FONT_REF_HEADER
    c.fill = FILL_DEEP_BLUE
    c.alignment = ALIGN_CENTER
    for cc in range(2, 8):
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

    row = 3
    row = _grid(row,
        "NATIONAL PROGRAM \u2014 FIX & FLIP / BRIDGE (excl. FL, CA, NY)",
        ["FICO", "Exp 5+", "Exp 3-4", "Exp 1-2", "Exp 0"],
        [
            ["760+",    "90% / 90% / 75%", "87.5% / 87.5% / 72.5%", "85% / 85% / 70%", "82.5% / 82.5% / 67.5%"],
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
            ["760+",    "85% / 87.5% / 72.5%", "82.5% / 85% / 70%", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%"],
            ["740-759", "82.5% / 85% / 70%", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%"],
            ["720-739", "80% / 82.5% / 67.5%", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%"],
            ["700-719", "77.5% / 80% / 65%", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%"],
            ["680-699", "75% / 77.5% / 62.5%", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%", "N/A"],
            ["660-679", "72.5% / 75% / 60%", "70% / 72.5% / 57.5%", "N/A", "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "GROUND UP CONSTRUCTION \u2014 NATIONAL",
        ["FICO", "Exp 5+", "Exp 3-4", "Exp 1-2"],
        [
            ["740+",    "87.5% LTC / 72.5% LTARV", "85% LTC / 70% LTARV", "82.5% LTC / 67.5% LTARV"],
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
# SHEET 6: ZILLOW MARKET DATA (copy from existing file)
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
    ws.sheet_properties.tabColor = None

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

            if row_idx == 1:
                new_cell.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
                new_cell.fill = FILL_DEEP_BLUE
                new_cell.alignment = ALIGN_CENTER
                new_cell.border = THIN_BORDER
            else:
                new_cell.font = Font(name="Calibri", size=9, color=DARK_TEXT)
                if col_idx >= 10:
                    new_cell.number_format = '$#,##0'

        if row_idx % 5000 == 0:
            print(f"    ... {row_idx:,} / {total_rows:,} rows")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(total_cols)}{total_rows}"

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
    print("A&S Capital Sizer -- Excel Template Builder (v2)")
    print("Auto-Calculating Leverage Architecture")
    print("=" * 60)

    os.makedirs(ASSETS_DIR, exist_ok=True)

    # ------------------------------------------------------------------
    # Step 1: Check for existing Zillow data
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

    print("\n[1/6] Building Sizer (Input) sheet ...")
    build_sizer_sheet(wb)

    print("[2/6] Building Sizing (Auto-Calculating) sheet ...")
    build_sizing_sheet(wb)

    print("[3/6] Building Colchis Leverage Data (Hidden Lookup) sheet ...")
    build_colchis_leverage_data_sheet(wb)

    print("[4/6] Building Colchis Leverage (Reference) sheet ...")
    build_colchis_leverage_sheet(wb)

    print("[5/6] Building Fidelis Leverage (Reference) sheet ...")
    build_fidelis_leverage_sheet(wb)

    print("[6/6] Copying Zillow Market Data sheet ...")
    if has_zillow:
        copy_zillow_data(wb, EXISTING_PATH)
    else:
        ws = wb.create_sheet("Zillow Market Data")
        ws.sheet_properties.tabColor = None
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
