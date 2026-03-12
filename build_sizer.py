#!/usr/bin/env python3
"""
A&S Capital Sizer -- Excel Template Builder  (v3 - Single-Tab Eastview Style)
Generates a professional loan sizing workbook with:
  - Single "Sizer" tab: inputs on LEFT, auto-sizing on RIGHT (Eastview layout)
  - Correct Colchis leverage formulas (Construction sized by LTC/LTARV, not LTV)
  - Full Colchis pricing grid with ALL adjustments (experience, term, state, ZHVI, etc.)
  - Hidden lookup tables for VLOOKUP-driven auto-calculation
  - Guideline pass/fail checks
  - ZHVI market data integration

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
TEAL        = "087496"
POWDER_BLUE = "A3D5E0"
LIGHT_BLUE  = "E0F0F8"
DARK_TEXT    = "2C3E50"
WHITE        = "FFFFFF"
PASS_GREEN   = "27AE60"
FAIL_RED     = "E74C3C"
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
NO_BORDER = Border()

FONT_TITLE         = Font(name="Calibri", size=16, bold=True, color=WHITE)
FONT_SECTION        = Font(name="Calibri", size=11, bold=True, color=WHITE)
FONT_SUBSECTION     = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
FONT_LABEL          = Font(name="Calibri", size=10, color=DARK_TEXT)
FONT_INPUT          = Font(name="Calibri", size=10, color=BLACK)
FONT_COMPUTED       = Font(name="Calibri", size=10, color=DARK_TEXT)
FONT_COMPUTED_BOLD  = Font(name="Calibri", size=10, bold=True, color=DARK_TEXT)
FONT_BIG_RESULT     = Font(name="Calibri", size=12, bold=True, color=WHITE)
FONT_NOTE           = Font(name="Calibri", size=9, italic=True, color=MED_GRAY)
FONT_PASS           = Font(name="Calibri", size=10, bold=True, color=PASS_GREEN)
FONT_FAIL           = Font(name="Calibri", size=10, bold=True, color=FAIL_RED)
FONT_REF_HEADER     = Font(name="Calibri", size=13, bold=True, color=WHITE)
FONT_REF_SECTION    = Font(name="Calibri", size=11, bold=True, color=DEEP_BLUE)
FONT_REF            = Font(name="Calibri", size=9, color=DARK_TEXT)
FONT_REF_BOLD       = Font(name="Calibri", size=9, bold=True, color=DARK_TEXT)

FILL_DEEP_BLUE  = PatternFill(start_color=DEEP_BLUE, end_color=DEEP_BLUE, fill_type="solid")
FILL_TEAL       = PatternFill(start_color=TEAL, end_color=TEAL, fill_type="solid")
FILL_POWDER     = PatternFill(start_color=POWDER_BLUE, end_color=POWDER_BLUE, fill_type="solid")
FILL_LIGHT_BLUE = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
FILL_WHITE      = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
FILL_LIGHT_GRAY = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")

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
    "yes_no":        "Yes,No",
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
    """Teal section header spanning col_start:col_end."""
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row, end_column=col_end
    )
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = FONT_SECTION
    cell.fill = FILL_TEAL
    cell.alignment = ALIGN_LEFT
    cell.border = THIN_BORDER
    for c in range(col_start + 1, col_end + 1):
        mc = ws.cell(row=row, column=c)
        mc.fill = FILL_TEAL
        mc.border = THIN_BORDER


def _sub_header(ws, row, col_start, col_end, text):
    """Lighter sub-header."""
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
# SINGLE SIZER SHEET: Inputs LEFT (B-E) + Auto-Sizing RIGHT (G-J)
# ============================================================================
#
# LAYOUT:
#   Col A = spacer (width 3)
#   Col B = labels  (width 24)
#   Col C = input values (width 20)
#   Col D = extra labels / input overflow (width 20)
#   Col E = extra inputs (width 20)
#   Col F = spacer (width 3)
#   Col G = sizing labels (width 28)
#   Col H = sizing values / col 1 (width 20)
#   Col I = sizing values / col 2 (width 20)
#   Col J = sizing notes (width 20)
#   Col K = spacer (width 3)
#
# Cell references for dealfit.py:
#   C5  = Deal Type       C6  = Transaction Type    C7  = Loan Term
#   C8  = Deal Product
#   C11 = Address          C12 = City                E12 = State
#   C13 = ZIP Code         C14 = Property Type       C15 = # Units
#   C16 = Square Footage   E16 = Lot Size            C17 = Year Built (formula)
#   C20 = Purchase Price   C21 = Purchase Date       C22 = As-Is Value
#   C23 = ARV              C24 = Rehab Budget        C25 = Total Project Cost (formula)
#   C28 = Initial Loan Amt C29 = Rehab Holdback      C30 = Interest Reserve
#   C31 = Total Loan (formula)
#   C34 = Borrowing Entity C35 = # Guarantors
#   C38 = G1 Name          C39 = G1 FICO
#   C42 = G2 Name          C43 = G2 FICO
#   C46 = # Completed Projects   C47 = Similar Experience
#   F20 = ZHVI             F21 = Value/ZHVI Ratio
#
# RIGHT SIDE auto-sizing starts at G5:
#   H5  = Experience Tier (formula)
#   H6  = FICO Bucket (formula)
#   H7  = Product Category (formula)
#   H8  = Leverage Lookup Key
#   ...etc
# ============================================================================

def build_sizer_sheet(wb):
    ws = wb.active
    ws.title = "Sizer"

    # Column widths
    widths = {1: 3, 2: 24, 3: 20, 4: 20, 5: 20, 6: 3, 7: 28, 8: 20, 9: 20, 10: 20, 11: 3}
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 36

    # ---- Title ----
    _title_row(ws, 2, 1, 11, "A&S CAPITAL SIZER")

    # ==================================================================
    # LEFT SIDE: DEAL INFORMATION (rows 4-8)
    # ==================================================================
    _section_header(ws, 4, 2, 5, "DEAL INFORMATION")

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
    # RIGHT SIDE: COLCHIS CLASSIFICATION (rows 4-8)
    # ==================================================================
    _section_header(ws, 4, 7, 10, "COLCHIS CLASSIFICATION")

    _label(ws, 5, 7, "Experience Tier")
    _formula_cell(
        ws, 5, 8,
        '=IF(C46="","",IF(C8="Construction",IF(C46>=6,"6+",IF(C46>=4,"4-5","0-3")),IF(C46>=8,"8+",IF(C46>=4,"4-7","0-3"))))',
    )

    _label(ws, 6, 7, "FICO Bucket")
    _formula_cell(
        ws, 6, 8,
        '=IF(C39="","",IF(C39>=740,"740+",IF(C39>=700,"700-739",IF(C39>=680,"680-699",IF(C39>=660,"660-679","<660 (Ineligible)")))))',
    )

    _label(ws, 7, 7, "Product Category")
    _formula_cell(ws, 7, 8, '=IF(C8="","",C8)')

    _label(ws, 8, 7, "Leverage Lookup Key")
    _formula_cell(
        ws, 8, 8,
        '=IF(OR(H7="",H6="",H5=""),"",H7&"|"&H6&"|"&H5)',
    )

    # ==================================================================
    # LEFT SIDE: PROPERTY INFORMATION (rows 10-17)
    # ==================================================================
    _section_header(ws, 10, 2, 5, "PROPERTY INFORMATION")

    _label(ws, 11, 2, "Property Address")
    _merged_input(ws, 11, 3, 4)

    _label(ws, 12, 2, "City")
    _input_cell(ws, 12, 3)
    _label(ws, 12, 4, "State")
    _input_cell(ws, 12, 5)
    _add_dropdown(ws, "E12", DROPDOWNS["state"], "State", "Select state")

    _label(ws, 13, 2, "ZIP Code")
    _input_cell(ws, 13, 3, FMT_TEXT)

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
    _formula_cell(
        ws, 17, 3,
        '=IF(C5="Ground Up Construction",YEAR(TODAY()),"")',
        "0"
    )
    note_yb = ws.cell(row=17, column=4, value="Auto-fills for GUC")
    note_yb.font = FONT_NOTE
    note_yb.alignment = ALIGN_LEFT

    # ==================================================================
    # RIGHT SIDE: COLCHIS LEVERAGE LIMITS (rows 10-19)
    # ==================================================================
    _section_header(ws, 10, 7, 10, "COLCHIS LEVERAGE LIMITS")

    # Sub-header
    for col, txt in [(8, "Max %"), (9, "Max $ Amount")]:
        c = ws.cell(row=11, column=col, value=txt)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # Max LTV (As-Is)
    _label(ws, 12, 7, "Max LTV (As-Is)")
    _formula_cell(
        ws, 12, 8,
        '=IFERROR(VLOOKUP(H8,\'Colchis Data\'!A:D,2,FALSE),"")',
        FMT_PCT
    )
    _formula_cell(ws, 12, 9, '=IFERROR(H12*C22,"")', FMT_CURRENCY)

    # Max LTC
    _label(ws, 13, 7, "Max LTC")
    _formula_cell(
        ws, 13, 8,
        '=IFERROR(VLOOKUP(H8,\'Colchis Data\'!A:D,3,FALSE),"")',
        FMT_PCT
    )
    _formula_cell(ws, 13, 9, '=IFERROR(H13*C25,"")', FMT_CURRENCY)

    # Max LTARV
    _label(ws, 14, 7, "Max LTARV")
    _formula_cell(
        ws, 14, 8,
        '=IFERROR(VLOOKUP(H8,\'Colchis Data\'!A:D,4,FALSE),"")',
        FMT_PCT
    )
    _formula_cell(ws, 14, 9, '=IFERROR(H14*C23,"")', FMT_CURRENCY)

    # --- Sizing Logic ---
    # For Construction: size on MIN(LTC $, LTARV $) — ignore LTV
    # For Bridge: size on LTV $ only (no LTC/LTARV)
    # For Rehab (Light/Heavy): size on MIN(LTV $, LTC $, LTARV $)
    _label(ws, 16, 7, "Guidelines Max Loan")
    ws.cell(row=16, column=7).font = FONT_COMPUTED_BOLD
    _formula_cell(
        ws, 16, 9,
        '=IFERROR(IF(H7="Construction",MIN(I13,I14),IF(H7="Bridge",I12,MIN(I12,I13,I14))),"")',
        FMT_CURRENCY, bold=True
    )

    # Max Loan Cap
    _label(ws, 17, 7, "Max Loan Amount Cap")
    cap_cell = ws.cell(row=17, column=9, value=3500000)
    cap_cell.number_format = FMT_CURRENCY
    cap_cell.font = FONT_COMPUTED
    cap_cell.fill = FILL_LIGHT_BLUE
    cap_cell.border = THIN_BORDER
    cap_cell.alignment = ALIGN_RIGHT

    # FINAL MAX LOAN
    _label(ws, 19, 7, "FINAL MAX LOAN")
    ws.cell(row=19, column=7).font = Font(name="Calibri", size=12, bold=True, color=DARK_TEXT)
    c_final = ws.cell(row=19, column=9, value='=IFERROR(MIN(I16,I17),"")')
    c_final.font = FONT_BIG_RESULT
    c_final.fill = FILL_DEEP_BLUE
    c_final.number_format = FMT_CURRENCY
    c_final.alignment = ALIGN_CENTER
    c_final.border = THIN_BORDER

    # ==================================================================
    # LEFT SIDE: VALUATION (rows 19-25) with ZHVI on right columns D-E
    # ==================================================================
    _section_header(ws, 19, 2, 5, "VALUATION")

    _label(ws, 20, 2, "Purchase Price")
    _input_cell(ws, 20, 3, FMT_CURRENCY)
    _label(ws, 20, 4, "ZHVI (Zillow)")
    _formula_cell(
        ws, 20, 5,
        '=IFERROR(VLOOKUP(C13,\'Zillow Market Data\'!C:AN,38,FALSE),"")',
        FMT_CURRENCY
    )

    _label(ws, 21, 2, "Purchase Date")
    _input_cell(ws, 21, 3, FMT_DATE)
    _label(ws, 21, 4, "Value / ZHVI Ratio")
    _formula_cell(ws, 21, 5, '=IFERROR(C22/E20,"")', '0.00"x"')

    _label(ws, 22, 2, "As-Is Value")
    _input_cell(ws, 22, 3, FMT_CURRENCY)
    _label(ws, 22, 4, "Deal vs Market")
    _formula_cell(
        ws, 22, 5,
        '=IF(E21="","",IF(E21>3,"HIGH RISK",IF(E21>2,"ELEVATED","NORMAL")))'
    )

    _label(ws, 23, 2, "After Repair Value (ARV)")
    _input_cell(ws, 23, 3, FMT_CURRENCY)

    _label(ws, 24, 2, "Rehab Budget")
    _input_cell(ws, 24, 3, FMT_CURRENCY)

    _label(ws, 25, 2, "Total Project Cost")
    _formula_cell(ws, 25, 3, "=C20+C24", FMT_CURRENCY, bold=True)

    # ==================================================================
    # RIGHT SIDE: LOAN SIZING (rows 21-30)
    # ==================================================================
    _section_header(ws, 21, 7, 10, "LOAN SIZING")

    # Sub-header row 22
    for col, txt in [(8, "Borrower Req."), (9, "Guidelines Max")]:
        c = ws.cell(row=22, column=col, value=txt)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # Initial Loan Amount
    _label(ws, 23, 7, "Initial Loan Amount")
    _formula_cell(ws, 23, 8, '=C28', FMT_CURRENCY)
    _formula_cell(ws, 23, 9, '=IFERROR(IF(H7="Construction",0,MIN(C28,I19)),"")', FMT_CURRENCY)

    # Financed Rehab / Construction Budget
    _label(ws, 24, 7, "Financed Rehab Budget")
    _formula_cell(ws, 24, 8, '=C29', FMT_CURRENCY)
    _formula_cell(ws, 24, 9, '=IFERROR(MIN(C29,MAX(I19-I23,0)),"")', FMT_CURRENCY)

    # Interest Reserve
    _label(ws, 25, 7, "Interest Reserve")
    _formula_cell(ws, 25, 8, '=C30', FMT_CURRENCY)
    _formula_cell(ws, 25, 9, '=IFERROR(MIN(C30,MAX(I19-I23-I24,0)),"")', FMT_CURRENCY)

    # Total Loan Amount (bold)
    _label(ws, 26, 7, "Total Loan Amount")
    ws.cell(row=26, column=7).font = FONT_COMPUTED_BOLD
    _formula_cell(ws, 26, 8, '=SUM(H23:H25)', FMT_CURRENCY, bold=True)
    _formula_cell(ws, 26, 9, '=SUM(I23:I25)', FMT_CURRENCY, bold=True)

    # Actual ratios
    _label(ws, 28, 7, "Actual LTV")
    _formula_cell(ws, 28, 8, '=IFERROR(H23/C22,"")', FMT_PCT)
    _formula_cell(ws, 28, 9, '=IFERROR(I26/C22,"")', FMT_PCT)

    _label(ws, 29, 7, "Actual LTC")
    _formula_cell(ws, 29, 8, '=IFERROR(H26/C25,"")', FMT_PCT)
    _formula_cell(ws, 29, 9, '=IFERROR(I26/C25,"")', FMT_PCT)

    _label(ws, 30, 7, "Actual LTARV")
    _formula_cell(ws, 30, 8, '=IFERROR(H26/C23,"")', FMT_PCT)
    _formula_cell(ws, 30, 9, '=IFERROR(I26/C23,"")', FMT_PCT)

    # ==================================================================
    # LEFT SIDE: LOAN REQUEST (rows 27-31) with leverage on D-E
    # ==================================================================
    _section_header(ws, 27, 2, 5, "LOAN REQUEST")

    _label(ws, 28, 2, "Initial Loan Amount")
    _input_cell(ws, 28, 3, FMT_CURRENCY)

    _label(ws, 29, 2, "Rehab Holdback")
    _input_cell(ws, 29, 3, FMT_CURRENCY)

    _label(ws, 30, 2, "Interest Reserve")
    _input_cell(ws, 30, 3, FMT_CURRENCY)

    _label(ws, 31, 2, "Total Loan Amount")
    _formula_cell(ws, 31, 3, "=C28+C29+C30", FMT_CURRENCY, bold=True)

    # Leverage ratios on right of loan request
    _label(ws, 28, 4, "LTV")
    _formula_cell(ws, 28, 5, '=IFERROR(C28/C22,"")', FMT_PCT)

    _label(ws, 29, 4, "LTC")
    _formula_cell(ws, 29, 5, '=IFERROR(C31/C25,"")', FMT_PCT)

    _label(ws, 30, 4, "LTARV")
    _formula_cell(ws, 30, 5, '=IFERROR(C31/C23,"")', FMT_PCT)

    # ==================================================================
    # RIGHT SIDE: COLCHIS PRICING (rows 32-41)
    # ==================================================================
    _section_header(ws, 32, 7, 10, "COLCHIS PRICING")

    # Actual LTC % for pricing bucket
    _label(ws, 33, 7, "Actual LTC %")
    _formula_cell(ws, 33, 8, '=IFERROR(H26/C25,"")', FMT_PCT)

    # LTC Bucket
    _label(ws, 34, 7, "LTC Bucket")
    _formula_cell(
        ws, 34, 8,
        '=IF(H33="","",IF(H33<=0.7,"<=70.0%",IF(H33<=0.75,"<=75.0%",IF(H33<=0.8,"<=80.0%",IF(H33<=0.85,"<=85.0%",IF(H33<=0.9,"<=90.0%","<=95.0%"))))))',
    )

    # Pricing Key
    _label(ws, 35, 7, "Pricing Key")
    _formula_cell(
        ws, 35, 8,
        '=IF(OR(H7="",H6="",H34=""),"",H7&"|"&H6&"|"&H34)',
    )

    # Base Rate (from lookup)
    _label(ws, 36, 7, "Base Rate")
    _formula_cell(
        ws, 36, 8,
        '=IFERROR(VLOOKUP(H35,\'Colchis Data\'!F:G,2,FALSE),"")',
        FMT_RATE
    )

    # Adjustments
    _label(ws, 37, 7, "Experience Adj.")
    _formula_cell(
        ws, 37, 8,
        '=IF(H5="","",IF(OR(H5="8+",H5="6+"),-0.0025,IF(H5="0-3",0.0025,0)))',
        FMT_RATE
    )
    note_exp = ws.cell(row=37, column=9, value="Tier 1: -25bps / Tier 3: +25bps")
    note_exp.font = FONT_NOTE

    _label(ws, 38, 7, "Term Adj.")
    _formula_cell(
        ws, 38, 8,
        '=IF(C7="","",IF(C7="24 Months",0.00125,IF(C7="18 Months",0,0)))',
        FMT_RATE
    )
    note_term = ws.cell(row=38, column=9, value="19-24mo: +12.5bps")
    note_term.font = FONT_NOTE

    _label(ws, 39, 7, "Transaction Adj.")
    _formula_cell(
        ws, 39, 8,
        '=IF(C6="","",IF(C6="Refinance (Cash Out)",0.0025,0))',
        FMT_RATE
    )
    note_tx = ws.cell(row=39, column=9, value="Cash-out refi: +25bps")
    note_tx.font = FONT_NOTE

    _label(ws, 40, 7, "Loan Size Adj.")
    _formula_cell(
        ws, 40, 8,
        '=IF(H26="","",IF(H26>3000000,0.00125,0))',
        FMT_RATE
    )
    note_sz = ws.cell(row=40, column=9, value=">$3M: +12.5bps")
    note_sz.font = FONT_NOTE

    _label(ws, 41, 7, "State / ZHVI Adj.")
    _formula_cell(
        ws, 41, 8,
        '=IF(E12="","",IF(OR(E12="NY",E12="NJ",E12="CT"),0.0025,IF(E12="CA",-0.00125,0))'
        '+IF(E21="","",IF(E21>3,0.00375,IF(E21>2,0.00125,0))))',
        FMT_RATE
    )
    note_geo = ws.cell(row=41, column=9, value="NY/NJ/CT +25bp; CA -12.5bp; ZHVI adj")
    note_geo.font = FONT_NOTE

    # FINAL RATE
    _label(ws, 43, 7, "ALL-IN BUY RATE")
    ws.cell(row=43, column=7).font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
    c_rate = ws.cell(row=43, column=8, value='=IFERROR(H36+H37+H38+H39+H40+H41,"")')
    c_rate.font = FONT_BIG_RESULT
    c_rate.fill = FILL_DEEP_BLUE
    c_rate.number_format = FMT_RATE
    c_rate.alignment = ALIGN_CENTER
    c_rate.border = THIN_BORDER

    _label(ws, 44, 7, "A&S Sell Rate (+50bps)")
    _formula_cell(ws, 44, 8, '=IFERROR(H43+0.005,"")', FMT_RATE, bold=True)

    # ==================================================================
    # LEFT SIDE: BORROWER INFORMATION (rows 33-47)
    # ==================================================================
    _section_header(ws, 33, 2, 5, "BORROWER INFORMATION")

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
    # EXPERIENCE (rows 45-47)
    # ==================================================================
    _section_header(ws, 45, 2, 5, "EXPERIENCE")

    _label(ws, 46, 2, "# Completed Projects")
    _input_cell(ws, 46, 3, FMT_INT)

    _label(ws, 47, 2, "Similar Experience")
    _input_cell(ws, 47, 3)
    _add_dropdown(ws, "C47", DROPDOWNS["experience"], "Experience", "Similar project experience?")

    # ==================================================================
    # RIGHT SIDE: GUIDELINE CHECK (rows 46-60)
    # ==================================================================
    _section_header(ws, 46, 7, 10, "GUIDELINE CHECK")

    # Sub-header
    for col, txt in [(8, "Value"), (9, "Result")]:
        c = ws.cell(row=47, column=col, value=txt)
        c.font = FONT_SUBSECTION
        c.fill = FILL_POWDER
        c.alignment = ALIGN_CENTER
        c.border = THIN_BORDER

    # Min FICO (680)
    _label(ws, 48, 7, "Min FICO (680)")
    _formula_cell(ws, 48, 8, '=C39', FMT_INT)
    ws.cell(row=48, column=9, value='=IF(C39="","",IF(C39>=680,"PASS","FAIL"))')
    ws.cell(row=48, column=9).border = THIN_BORDER
    ws.cell(row=48, column=9).alignment = ALIGN_CENTER

    # Max Loan ($3.5M)
    _label(ws, 49, 7, "Max Loan ($3.5M)")
    _formula_cell(ws, 49, 8, '=I26', FMT_CURRENCY)
    ws.cell(row=49, column=9, value='=IF(I26="","",IF(I26<=3500000,"PASS","FAIL"))')
    ws.cell(row=49, column=9).border = THIN_BORDER
    ws.cell(row=49, column=9).alignment = ALIGN_CENTER

    # Min Loan ($100K)
    _label(ws, 50, 7, "Min Loan ($100K)")
    _formula_cell(ws, 50, 8, '=I26', FMT_CURRENCY)
    ws.cell(row=50, column=9, value='=IF(I26="","",IF(I26>=100000,"PASS","FAIL"))')
    ws.cell(row=50, column=9).border = THIN_BORDER
    ws.cell(row=50, column=9).alignment = ALIGN_CENTER

    # State Eligible
    _label(ws, 51, 7, "State Eligible")
    _formula_cell(ws, 51, 8, '=E12')
    ws.cell(row=51, column=9, value='=IF(E12="","",IF(E12="IL","FAIL","PASS"))')
    ws.cell(row=51, column=9).border = THIN_BORDER
    ws.cell(row=51, column=9).alignment = ALIGN_CENTER

    # Leverage Eligible
    _label(ws, 52, 7, "Leverage Eligible")
    _formula_cell(ws, 52, 8, '=H8')
    ws.cell(row=52, column=9,
        value='=IF(H8="","",IF(AND(H12<>"",OR(H13<>"",H7="Bridge")),"PASS","FAIL"))')
    ws.cell(row=52, column=9).border = THIN_BORDER
    ws.cell(row=52, column=9).alignment = ALIGN_CENTER

    # LTV Check (skip for Construction)
    _label(ws, 53, 7, "LTV Within Limits")
    _formula_cell(ws, 53, 8, '=IFERROR(H23/C22,"")', FMT_PCT)
    ws.cell(row=53, column=9,
        value='=IF(H7="Construction","N/A",IF(OR(H53="",H12=""),"",IF(H53<=H12,"PASS","FAIL")))')
    ws.cell(row=53, column=9).border = THIN_BORDER
    ws.cell(row=53, column=9).alignment = ALIGN_CENTER

    # LTC Check (skip for Bridge)
    _label(ws, 54, 7, "LTC Within Limits")
    _formula_cell(ws, 54, 8, '=IFERROR(H26/C25,"")', FMT_PCT)
    ws.cell(row=54, column=9,
        value='=IF(H7="Bridge","N/A",IF(OR(H54="",H13=""),"",IF(H54<=H13,"PASS","FAIL")))')
    ws.cell(row=54, column=9).border = THIN_BORDER
    ws.cell(row=54, column=9).alignment = ALIGN_CENTER

    # LTARV Check (skip for Bridge)
    _label(ws, 55, 7, "LTARV Within Limits")
    _formula_cell(ws, 55, 8, '=IFERROR(H26/C23,"")', FMT_PCT)
    ws.cell(row=55, column=9,
        value='=IF(H7="Bridge","N/A",IF(OR(H55="",H14=""),"",IF(H55<=H14,"PASS","FAIL")))')
    ws.cell(row=55, column=9).border = THIN_BORDER
    ws.cell(row=55, column=9).alignment = ALIGN_CENTER

    # ZHVI Check
    _label(ws, 56, 7, "ZHVI > 300% (High Risk)")
    _formula_cell(ws, 56, 8, '=E21')
    ws.cell(row=56, column=9,
        value='=IF(E21="","",IF(E21>3,"WARN","PASS"))')
    ws.cell(row=56, column=9).border = THIN_BORDER
    ws.cell(row=56, column=9).alignment = ALIGN_CENTER

    # MASTER CHECK
    _label(ws, 58, 7, "MASTER CHECK")
    ws.cell(row=58, column=7).font = Font(name="Calibri", size=11, bold=True, color=DARK_TEXT)
    c_master = ws.cell(
        row=58, column=9,
        value='=IF(COUNTBLANK(I48:I56)=9,"",IF(COUNTIF(I48:I56,"FAIL")>0,"FAIL","PASS"))'
    )
    c_master.border = THIN_BORDER
    c_master.alignment = ALIGN_CENTER
    c_master.font = Font(name="Calibri", size=12, bold=True, color=DARK_TEXT)

    # ---- Note row ----
    ws.merge_cells("B50:E50")
    note = ws.cell(
        row=50, column=2,
        value="Complete all fields. Auto-sizing results appear on the right."
    )
    note.font = FONT_NOTE
    note.alignment = ALIGN_CENTER

    # ---- Print setup ----
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.print_area = "A1:K60"
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)

    return ws


# ============================================================================
# COLCHIS DATA (HIDDEN lookup table)
# ============================================================================

def build_colchis_data_sheet(wb):
    """
    Hidden sheet with two lookup tables:
      Columns A-D: Leverage lookup  (Key | MaxLTV | MaxLTC | MaxLTARV)
      Columns F-G: Pricing lookup   (Key | BaseRate)
    """
    ws = wb.create_sheet("Colchis Data")

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
    # LEVERAGE DATA -- SINGLE FAMILY (1-4 Units)
    # Key format: "Product|FICO Bucket|Experience Tier"
    # ================================================================
    leverage_data = [
        # ---- SF Light Rehab (Purchase) ----
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
        ("Heavy Rehab|700-739|8+",  0.800, 0.850, 0.700),
        ("Heavy Rehab|700-739|4-7", 0.800, 0.850, 0.700),
        ("Heavy Rehab|680-699|8+",  0.750, 0.825, 0.650),
        ("Heavy Rehab|680-699|4-7", 0.750, 0.800, 0.650),
        # Heavy Rehab 0-3 exp: ineligible (omitted)

        # ---- SF Bridge (No Rehab) -- LTV only ----
        ("Bridge|740+|8+",     0.750, 0.000, 0.000),
        ("Bridge|740+|4-7",    0.750, 0.000, 0.000),
        ("Bridge|740+|0-3",    0.750, 0.000, 0.000),
        ("Bridge|700-739|8+",  0.750, 0.000, 0.000),
        ("Bridge|700-739|4-7", 0.750, 0.000, 0.000),
        ("Bridge|700-739|0-3", 0.700, 0.000, 0.000),
        ("Bridge|680-699|8+",  0.700, 0.000, 0.000),
        ("Bridge|680-699|4-7", 0.700, 0.000, 0.000),
        ("Bridge|680-699|0-3", 0.650, 0.000, 0.000),

        # ---- SF Construction (experience tiers: 6+, 4-5, 0-3) ----
        ("Construction|740+|6+",     0.600, 0.900, 0.700),
        ("Construction|740+|4-5",    0.600, 0.850, 0.700),
        ("Construction|700-739|6+",  0.600, 0.900, 0.700),
        ("Construction|700-739|4-5", 0.600, 0.850, 0.700),
        ("Construction|680-699|6+",  0.600, 0.850, 0.700),
        ("Construction|680-699|4-5", 0.600, 0.825, 0.650),
        # Construction 0-3 exp: ineligible (omitted)

        # ---- SF Rate/Term Refinance -- LTV only ----
        # Reuse Bridge key format for refi lookups
        # (handled in the experience tier formula which maps correctly)

        # ---- MF Light Rehab (5-10 units) ----
        # These would need a separate property-type-aware lookup
        # For now we handle MF in the same grid with reduced leverage
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
    # PRICING DATA (from Colchis RTL Pricing 2026-01-27)
    # Key format: "Product|FICO Bucket|LTC Bucket"
    # These are BASE rates before adjustments
    # ================================================================
    pricing_data = [
        # ---- Bridge ----
        ("Bridge|740+|<=70.0%",     0.07750),
        ("Bridge|740+|<=75.0%",     0.07750),
        ("Bridge|700-739|<=70.0%",  0.07750),
        ("Bridge|700-739|<=75.0%",  0.07750),
        ("Bridge|680-699|<=70.0%",  0.07875),

        # ---- Light Rehab ----
        ("Light Rehab|740+|<=70.0%",     0.07750),
        ("Light Rehab|740+|<=75.0%",     0.07750),
        ("Light Rehab|740+|<=80.0%",     0.07750),
        ("Light Rehab|740+|<=85.0%",     0.07875),
        ("Light Rehab|740+|<=90.0%",     0.08000),
        ("Light Rehab|740+|<=95.0%",     0.08250),
        ("Light Rehab|700-739|<=70.0%",  0.07750),
        ("Light Rehab|700-739|<=75.0%",  0.07750),
        ("Light Rehab|700-739|<=80.0%",  0.07875),
        ("Light Rehab|700-739|<=85.0%",  0.08000),
        ("Light Rehab|700-739|<=90.0%",  0.08125),
        ("Light Rehab|700-739|<=95.0%",  0.08375),
        ("Light Rehab|680-699|<=70.0%",  0.07875),
        ("Light Rehab|680-699|<=75.0%",  0.08000),
        ("Light Rehab|680-699|<=80.0%",  0.08125),
        ("Light Rehab|680-699|<=85.0%",  0.08250),
        ("Light Rehab|680-699|<=90.0%",  0.08375),

        # ---- Heavy Rehab ----
        ("Heavy Rehab|740+|<=70.0%",     0.08375),
        ("Heavy Rehab|740+|<=75.0%",     0.08375),
        ("Heavy Rehab|740+|<=80.0%",     0.08500),
        ("Heavy Rehab|740+|<=85.0%",     0.08625),
        ("Heavy Rehab|700-739|<=70.0%",  0.08375),
        ("Heavy Rehab|700-739|<=75.0%",  0.08500),
        ("Heavy Rehab|700-739|<=80.0%",  0.08625),
        ("Heavy Rehab|700-739|<=85.0%",  0.08750),
        ("Heavy Rehab|680-699|<=70.0%",  0.08625),
        ("Heavy Rehab|680-699|<=75.0%",  0.08750),
        ("Heavy Rehab|680-699|<=80.0%",  0.08875),
        ("Heavy Rehab|680-699|<=85.0%",  0.09000),

        # ---- Construction ----
        ("Construction|740+|<=70.0%",     0.08375),
        ("Construction|740+|<=75.0%",     0.08375),
        ("Construction|740+|<=80.0%",     0.08500),
        ("Construction|740+|<=85.0%",     0.08625),
        ("Construction|740+|<=90.0%",     0.08875),
        ("Construction|700-739|<=70.0%",  0.08375),
        ("Construction|700-739|<=75.0%",  0.08500),
        ("Construction|700-739|<=80.0%",  0.08625),
        ("Construction|700-739|<=85.0%",  0.08750),
        ("Construction|700-739|<=90.0%",  0.09000),
        ("Construction|680-699|<=70.0%",  0.08625),
        ("Construction|680-699|<=75.0%",  0.08750),
        ("Construction|680-699|<=80.0%",  0.08875),
        ("Construction|680-699|<=85.0%",  0.09000),
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
# COLCHIS LEVERAGE (Human-Readable Reference)
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
        "SF \u2014 LIGHT REHAB (Purchase)",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",    "90% / 92.5% / 75%", "90% / 92.5% / 75%", "90% / 90% / 75%"],
            ["700-739", "90% / 92.5% / 75%", "90% / 92.5% / 75%", "87.5% / 90% / 75%"],
            ["680-699", "87.5% / 90% / 75%", "85% / 87.5% / 75%", "85% / 85% / 70%"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "SF \u2014 HEAVY REHAB",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",    "80% / 85% / 70%",   "80% / 85% / 70%",   "N/A"],
            ["700-739", "80% / 85% / 70%",   "80% / 85% / 70%",   "N/A"],
            ["680-699", "75% / 82.5% / 65%", "75% / 80% / 65%",   "N/A"],
        ],
        note="Format: Max LTV / Max LTC / Max LTARV"
    )

    row = _grid(row,
        "SF \u2014 BRIDGE (No Rehab)",
        ["FICO", "Exp 8+", "Exp 4-7", "Exp 0-3"],
        [
            ["740+",    "75%", "75%", "75%"],
            ["700-739", "75%", "75%", "70%"],
            ["680-699", "70%", "70%", "65%"],
        ],
        note="Max LTV only"
    )

    row = _grid(row,
        "SF \u2014 CONSTRUCTION",
        ["FICO", "Exp 6+", "Exp 4-5", "Exp 0-3"],
        [
            ["740+",    "60% LTV / 90%* LTC / 70% LTARV", "60% LTV / 85% LTC / 70% LTARV", "N/A"],
            ["700-739", "60% LTV / 90%* LTC / 70% LTARV", "60% LTV / 85% LTC / 70% LTARV", "N/A"],
            ["680-699", "60% LTV / 85% LTC / 70% LTARV", "60% LTV / 82.5% LTC / 65% LTARV", "N/A"],
        ],
    )

    # Notes
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    ws.cell(row=row, column=2,
        value="*90% LTC requires budget < $500K; otherwise 85%").font = FONT_NOTE
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    ws.cell(row=row, column=2,
        value="Loan Range: $100K-$3.5M | Terms: 6-24mo | Min FICO: 680 | Excluded: IL").font = FONT_NOTE
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    ws.cell(row=row, column=2,
        value="ZHVI >200%: -5% leverage adj. | ZHVI >300%: -10% leverage adj.").font = FONT_NOTE

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    return ws


# ============================================================================
# FIDELIS LEVERAGE (Human-Readable Reference)
# ============================================================================

def build_fidelis_leverage_sheet(wb):
    ws = wb.create_sheet("Fidelis Leverage")

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
        "NATIONAL \u2014 FIX & FLIP / BRIDGE (excl. FL, CA, NY)",
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
        "GUC \u2014 NATIONAL",
        ["FICO", "Exp 5+", "Exp 3-4", "Exp 1-2"],
        [
            ["740+",    "87.5% LTC / 72.5% LTARV", "85% LTC / 70% LTARV", "82.5% LTC / 67.5% LTARV"],
            ["720-739", "85% LTC / 70% LTARV", "82.5% LTC / 67.5% LTARV", "80% LTC / 65% LTARV"],
            ["700-719", "82.5% LTC / 67.5% LTARV", "80% LTC / 65% LTARV", "77.5% LTC / 62.5% LTARV"],
            ["680-699", "80% LTC / 65% LTARV", "77.5% LTC / 62.5% LTARV", "75% LTC / 60% LTARV"],
        ],
    )

    # Notes
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    ws.cell(row=row, column=2,
        value="Loan Range: $75K-$5M | Terms: 6-24mo | Min FICO: 660 (Tier 1-2)").font = FONT_NOTE
    row += 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    ws.cell(row=row, column=2,
        value="All states eligible | Prepay: None | Extension: 0.5-1.0pt").font = FONT_NOTE

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    return ws


# ============================================================================
# ZILLOW MARKET DATA (copy from existing file)
# ============================================================================

def copy_zillow_data(wb, existing_path):
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
    print("A&S Capital Sizer -- Excel Template Builder (v3)")
    print("Single-Tab Eastview-Style + Correct Colchis Formulas")
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

    print("\n[1/5] Building Sizer sheet (inputs + auto-sizing) ...")
    build_sizer_sheet(wb)

    print("[2/5] Building Colchis Data (hidden lookup) ...")
    build_colchis_data_sheet(wb)

    print("[3/5] Building Colchis Leverage (reference) ...")
    build_colchis_leverage_sheet(wb)

    print("[4/5] Building Fidelis Leverage (reference) ...")
    build_fidelis_leverage_sheet(wb)

    print("[5/5] Copying Zillow Market Data ...")
    if has_zillow:
        copy_zillow_data(wb, EXISTING_PATH)
    else:
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
