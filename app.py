"""
app.py — A&S Capital AI Agent
Main Streamlit UI with four core capabilities:
  1. Excel Sizer Filler (RTL / DSCR / MF / GUC)
  2. Underwriting Conditions Generator
  3. Committee Presentation Builder
  4. Borrower Presentation Builder

Run with:  streamlit run app.py
"""

import os
from datetime import datetime, date
import streamlit as st
from dotenv import load_dotenv

from modules.sizer import fill_sizer, get_template_path, SIZER_TEMPLATES
from modules.auto_sizer import auto_fill_sizer
from modules.underwriting import load_guidelines_pdf, get_guidelines_path, generate_conditions, GUIDELINES_FILES
from modules.committee_deck import build_committee_deck
from modules.borrower_deck import build_borrower_deck

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
LENDER_ORIGINATION_FEE = 0.02   # 2 points — always fixed for the lender

# ---------------------------------------------------------------------------
# Environment & paths
# ---------------------------------------------------------------------------
# Priority: 1) Streamlit Cloud secrets  2) .env file  3) manual .env fallback
ANTHROPIC_API_KEY = ""
TAVILY_API_KEY = ""

# Try Streamlit Cloud secrets first (used when deployed)
try:
    ANTHROPIC_API_KEY = st.secrets.get("ANTHROPIC_API_KEY", "")
    TAVILY_API_KEY = st.secrets.get("TAVILY_API_KEY", "")
except Exception:
    pass

# Fall back to .env file (used when running locally)
if not ANTHROPIC_API_KEY or not TAVILY_API_KEY:
    ENV_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    load_dotenv(ENV_PATH, override=True)
    if not ANTHROPIC_API_KEY:
        ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
    if not TAVILY_API_KEY:
        TAVILY_API_KEY = os.getenv("TAVILY_API_KEY", "")

    # Manual fallback: read .env line by line
    if not ANTHROPIC_API_KEY or not TAVILY_API_KEY:
        try:
            with open(ENV_PATH, "r") as _ef:
                for _line in _ef:
                    _line = _line.strip()
                    if _line and not _line.startswith("#") and "=" in _line:
                        _k, _v = _line.split("=", 1)
                        _k, _v = _k.strip(), _v.strip().strip("\"'")
                        if _k == "ANTHROPIC_API_KEY" and not ANTHROPIC_API_KEY:
                            ANTHROPIC_API_KEY = _v
                        elif _k == "TAVILY_API_KEY" and not TAVILY_API_KEY:
                            TAVILY_API_KEY = _v
        except FileNotFoundError:
            pass

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
IC_TEMPLATE = os.path.join(ASSETS_DIR, "AS_Capital_IC_Template (1).pptx")

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
LOGO_PATH = os.path.join(ASSETS_DIR, "as_logo.png")

st.set_page_config(
    page_title="A&S Capital Agent",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    /* ---- Global: White background + Powder Blue / Light Cyan accents ---- */
    .stApp {
        background-color: #FFFFFF;
    }
    .block-container {
        padding-top: 1.5rem;
        background-color: #FFFFFF;
    }

    /* Sidebar: light powder blue */
    section[data-testid="stSidebar"] {
        background-color: #E0F0F8 !important;
    }
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] .stMarkdown h1,
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #2C3E50 !important;
    }

    /* Headings: dark slate for contrast */
    h1, h2, h3 { color: #2C3E50; }

    /* Tabs: powder blue underline on active */
    .stTabs [data-baseweb="tab"] {
        font-weight: 600;
        color: #2C3E50;
    }
    .stTabs [aria-selected="true"] {
        border-bottom-color: #A3D5E0 !important;
        color: #2C3E50 !important;
    }

    /* Primary buttons: powder blue / light cyan */
    .stButton > button[kind="primary"],
    .stButton > button[data-testid="stBaseButton-primary"] {
        background-color: #A3D5E0 !important;
        border-color: #8DC8D6 !important;
        color: #2C3E50 !important;
        font-weight: 600;
    }
    .stButton > button[kind="primary"]:hover,
    .stButton > button[data-testid="stBaseButton-primary"]:hover {
        background-color: #8DC8D6 !important;
        border-color: #78BFCD !important;
    }

    /* Download buttons */
    .stDownloadButton > button {
        background-color: #A3D5E0 !important;
        border-color: #8DC8D6 !important;
        color: #2C3E50 !important;
        font-weight: 600;
    }
    .stDownloadButton > button:hover {
        background-color: #8DC8D6 !important;
    }

    /* Input fields: subtle powder blue border on focus */
    .stTextInput input:focus,
    .stNumberInput input:focus,
    .stTextArea textarea:focus {
        border-color: #A3D5E0 !important;
        box-shadow: 0 0 0 1px #A3D5E0 !important;
    }

    /* Selectbox */
    .stSelectbox [data-baseweb="select"] {
        border-color: #D0E8F0;
    }

    /* Dividers: light cyan */
    hr { border-color: #D0E8F0 !important; }

    /* Success messages */
    .stSuccess {
        background-color: #E8F8F0 !important;
    }

    /* Sidebar dividers */
    section[data-testid="stSidebar"] hr {
        border-color: #B8D8E8 !important;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)
    else:
        st.title("A&S Capital")
    st.caption("AI-Powered Deal Agent")
    st.divider()

    api_ok = bool(ANTHROPIC_API_KEY and ANTHROPIC_API_KEY != "your-anthropic-api-key-here")
    tavily_ok = bool(TAVILY_API_KEY and TAVILY_API_KEY != "your-tavily-api-key-here")

    st.markdown("**API Status**")
    st.markdown(f"{'✅' if api_ok else '❌'}  Anthropic (Claude)")
    st.markdown(f"{'✅' if tavily_ok else '⚠️'}  Tavily {'(needed for comps)' if not tavily_ok else ''}")

    if not api_ok:
        st.warning("Add your ANTHROPIC_API_KEY to the .env file.")

    st.divider()

    # Check which asset files are present
    st.markdown("**Template Files**")
    for label, filename in [
        ("RTL Sizer", SIZER_TEMPLATES.get("RTL")),
        ("DSCR Sizer", SIZER_TEMPLATES.get("DSCR")),
        ("MF Sizer", SIZER_TEMPLATES.get("MF")),
        ("GUC Sizer", SIZER_TEMPLATES.get("GUC")),
        ("IC Template", "AS_Capital_IC_Template (1).pptx"),
    ]:
        path = os.path.join(ASSETS_DIR, filename) if filename else ""
        exists = os.path.exists(path) if path else False
        st.markdown(f"{'✅' if exists else '❌'}  {label}")

    st.divider()
    st.markdown("**Guidelines PDFs**")
    for label, filename in [
        ("RTL Guidelines", GUIDELINES_FILES.get("RTL")),
        ("DSCR Guidelines", GUIDELINES_FILES.get("DSCR")),
        ("MF Guidelines", GUIDELINES_FILES.get("MF")),
        ("GUC Guidelines", GUIDELINES_FILES.get("GUC")),
    ]:
        path = os.path.join(ASSETS_DIR, filename) if filename else ""
        exists = os.path.exists(path) if path else False
        st.markdown(f"{'✅' if exists else '❌'}  {label}")



# ---------------------------------------------------------------------------
# Helper: date input that returns datetime
# ---------------------------------------------------------------------------
def date_to_datetime(d):
    """Convert a date object to datetime for Excel compatibility."""
    if isinstance(d, date) and not isinstance(d, datetime):
        return datetime(d.year, d.month, d.day)
    return d


# ---------------------------------------------------------------------------
# Main content
# ---------------------------------------------------------------------------
st.title("A&S Capital — Deal Agent")

tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "⚡ Automatic Sizer",
    "📊 Manual Sizer",
    "📋 Underwriting Conditions",
    "🏛️ Committee Deck",
    "📄 Borrower Deck",
])


# ===========================================================================
# TAB 0: AUTOMATIC SIZER FILLER
# ===========================================================================
with tab0:
    st.header("Automatic Sizer Filler")
    st.caption("Drop any deal documents below — the AI will read everything, fill the correct sizer, and highlight missing info in red.")

    if not api_ok:
        st.warning("Set your ANTHROPIC_API_KEY in .env to use this feature.")
    else:
        st.markdown(
            """
            <div style="background-color:#E0F0F8; border:2px dashed #A3D5E0; border-radius:10px;
                        padding:15px; margin-bottom:15px; text-align:center;">
                <p style="margin:0; color:#2C3E50; font-size:16px;">
                    📂 <strong>Document Dropbox</strong> — Upload appraisals, loan apps, credit reports,
                    broker sheets, term sheets, or any deal documents
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        uploaded_files = st.file_uploader(
            "Upload Deal Documents",
            type=["xlsx", "xls", "pdf"],
            accept_multiple_files=True,
            key="auto_sizer_upload",
            help="Drag & drop or click to upload. Supports PDF and Excel files.",
        )

        auto_lt_override = st.selectbox(
            "Loan Type",
            ["Auto-Detect", "RTL", "DSCR", "MF", "GUC"],
            format_func=lambda x: {
                "Auto-Detect": "Auto-Detect (let AI determine)",
                "RTL": "RTL (Fix & Flip / Bridge)",
                "DSCR": "DSCR (Rental)",
                "MF": "Multifamily (5+ Units)",
                "GUC": "Ground Up Construction",
            }.get(x, x),
            key="auto_lt",
        )

        if uploaded_files:
            st.info(f"📄 **{len(uploaded_files)}** document(s) uploaded: " +
                    ", ".join(f.name for f in uploaded_files))

            if st.button("⚡ Auto-Fill Sizer", type="primary", key="auto_generate"):
                with st.spinner("Reading all documents and extracting deal data with AI..."):
                    override = auto_lt_override if auto_lt_override != "Auto-Detect" else None
                    files = [{"bytes": f.getvalue(), "name": f.name} for f in uploaded_files]
                    sizer_bytes, detected_type, filled_count, missing_fields, extracted = auto_fill_sizer(
                        api_key=ANTHROPIC_API_KEY,
                        assets_dir=ASSETS_DIR,
                        files=files,
                        loan_type_override=override,
                    )

                st.success(f"Sizer filled — **{detected_type}** loan detected, **{filled_count}** cells written.")

                # Show missing fields highlighted in red
                if missing_fields:
                    st.warning(
                        f"⚠️ **{len(missing_fields)}** required field(s) could not be found and are "
                        f"**highlighted in red** in the sizer:"
                    )
                    cols = st.columns(3)
                    for i, field in enumerate(missing_fields):
                        label = field.replace("_", " ").title()
                        cols[i % 3].markdown(f"- 🔴 {label}")
                else:
                    st.success("✅ All required fields were filled — no red highlights!")

                # Show what was extracted
                with st.expander("📋 Extracted Fields (click to review)", expanded=False):
                    for k, v in extracted.items():
                        label = k.replace("_", " ").title()
                        st.markdown(f"**{label}:** {v}")

                st.download_button(
                    "⬇️ Download Completed Sizer",
                    data=sizer_bytes,
                    file_name=f"AS_Capital_{detected_type}_Sizer_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


# ===========================================================================
# TAB 1: MANUAL SIZER FILLER
# ===========================================================================
with tab1:
    st.header("Manual Sizer Filler")
    st.caption("Select a loan type, fill in the deal inputs, and download the completed sizer.")

    loan_type = st.selectbox("Loan Type", ["RTL", "DSCR", "MF", "GUC"],
        format_func=lambda x: {"RTL": "RTL (Fix & Flip / Bridge)",
                                "DSCR": "DSCR (Rental)",
                                "MF": "Multifamily (5+ Units)",
                                "GUC": "Ground Up Construction"}.get(x, x),
        key="sz_loan_type")

    # Check template exists
    try:
        template_path = get_template_path(ASSETS_DIR, loan_type)
        if not os.path.exists(template_path):
            st.error(f"Template not found: {template_path}\nCopy it to the assets/ folder.")
            st.stop()
    except ValueError as e:
        st.error(str(e))
        st.stop()

    inputs = {}

    # ----- RTL Form -----
    if loan_type == "RTL":
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Loan Purpose")
            inputs["closing_date"] = date_to_datetime(st.date_input("Expected Closing Date", key="rtl_cd"))

            st.subheader("Borrower / Entity")
            inputs["entity_name"] = st.text_input("Entity Name", key="rtl_ent")
            inputs["num_owners"] = st.number_input("Number of Owners", min_value=1, max_value=4, value=1, key="rtl_no")
            inputs["guarantor_1_name"] = st.text_input("Guarantor 1 Full Name", key="rtl_g1n")
            inputs["guarantor_1_fico"] = st.number_input("Guarantor 1 FICO", min_value=0, max_value=850, value=700, key="rtl_g1f")
            inputs["guarantor_1_credit_date"] = date_to_datetime(st.date_input("Guarantor 1 Credit Report Date", key="rtl_g1d"))
            inputs["guarantor_1_is_guarantor"] = st.selectbox("Guarantor 1 Is Guarantor?", ["Yes", "No"], key="rtl_g1g")
            inputs["guarantor_1_ownership"] = st.number_input("Guarantor 1 % Ownership", 0.0, 1.0, 1.0, 0.05, key="rtl_g1o")

            st.subheader("Experience (Guarantor 1)")
            inputs["g1_rehab_sold"] = st.number_input("Rehab completed & sold", min_value=0, value=0, key="rtl_exp1")
            inputs["g1_rehab_refinanced"] = st.number_input("Rehab completed & refinanced as rental", min_value=0, value=0, key="rtl_exp2")
            inputs["g1_acquired_rental"] = st.number_input("Acquired as rental", min_value=0, value=0, key="rtl_exp3")

        with col2:
            st.subheader("Property 1")
            inputs["prop1_address"] = st.text_input("Street Address", key="rtl_addr")
            inputs["prop1_city"] = st.text_input("City", key="rtl_city")
            inputs["prop1_state"] = st.text_input("State (2-letter)", key="rtl_st")
            inputs["prop1_zip"] = st.number_input("ZIP Code", min_value=0, max_value=99999, value=0, key="rtl_zip", format="%d")
            inputs["prop1_type"] = st.selectbox("Property Type", ["SFR", "Townhome", "Condo", "PUD", "2 Unit", "3 Unit", "4 Unit"], key="rtl_pt")
            inputs["prop1_appraisal_date"] = date_to_datetime(st.date_input("Appraisal Date", key="rtl_apd"))
            inputs["prop1_as_is_value"] = st.number_input("As-Is Value ($)", min_value=0, value=0, step=10000, key="rtl_aiv")
            inputs["prop1_secondary_aiv"] = st.number_input("Secondary As-Is Value ($)", min_value=0, value=0, step=10000, key="rtl_saiv")
            inputs["prop1_arv"] = st.number_input("After-Repair Value ($)", min_value=0, value=0, step=10000, key="rtl_arv")
            inputs["prop1_secondary_arv"] = st.number_input("Secondary ARV ($)", min_value=0, value=0, step=10000, key="rtl_sarv")
            inputs["prop1_rehab_budget"] = st.number_input("Rehab Budget ($)", min_value=0, value=0, step=10000, key="rtl_rhb")
            inputs["prop1_pre_rehab_sqft"] = st.number_input("Pre-Rehab Sq Ft", min_value=0, value=0, key="rtl_prsf")
            inputs["prop1_post_rehab_sqft"] = st.number_input("Post-Rehab Sq Ft", min_value=0, value=0, key="rtl_posf")
            inputs["prop1_purchase_date"] = date_to_datetime(st.date_input("Purchase Date", key="rtl_pd"))
            inputs["prop1_purchase_price"] = st.number_input("Purchase Price ($)", min_value=0, value=0, step=10000, key="rtl_pp")
            inputs["prop1_change_of_use"] = st.selectbox("Change of Use?", ["No", "Yes"], key="rtl_cou")

        st.subheader("Summary Sheet — Loan Structure")
        c1, c2, c3 = st.columns(3)
        with c1:
            inputs["loan_program"] = st.selectbox("Loan Program", ["Fix & Flip", "Bridge", "Bridge Plus"], key="rtl_lp")
            inputs["loan_term"] = st.selectbox("Loan Term", ["12 Months", "18 Months", "24 Months", "36 Months"], key="rtl_lt")
        with c2:
            inputs["initial_loan_amount"] = st.number_input("Initial Loan Amount ($)", min_value=0, value=0, step=10000, key="rtl_ila")
            inputs["interest_reserves"] = st.number_input("Financed Interest Reserves ($)", min_value=0, value=0, step=1000, key="rtl_ir")
        with c3:
            inputs["financed_rehab"] = st.number_input("Financed Rehab Budget ($)", min_value=0, value=0, step=10000, key="rtl_fr")
            inputs["loan_id"] = st.text_input("Loan ID", key="rtl_lid")

    # ----- DSCR Form -----
    elif loan_type == "DSCR":
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Loan Purpose")
            inputs["closing_date"] = date_to_datetime(st.date_input("Closing Date", key="dscr_cd"))
            inputs["prop_purchase_date"] = date_to_datetime(st.date_input("Property Purchase Date", key="dscr_ppd"))

            st.subheader("Borrower / Entity")
            inputs["entity_name"] = st.text_input("Entity Name", key="dscr_ent")
            inputs["num_guarantors"] = st.number_input("Number of Guarantors", 1, 4, 1, key="dscr_ng")
            inputs["guarantor_1_first"] = st.text_input("Guarantor 1 First Name", key="dscr_g1f")
            inputs["guarantor_1_last"] = st.text_input("Guarantor 1 Last Name", key="dscr_g1l")
            inputs["guarantor_1_fico"] = st.number_input("Guarantor 1 FICO", 0, 850, 700, key="dscr_g1fi")
            inputs["guarantor_1_credit_date"] = date_to_datetime(st.date_input("Guarantor 1 Credit Date", key="dscr_g1d"))
            inputs["guarantor_1_ownership"] = st.number_input("Guarantor 1 % Ownership", 0.0, 1.0, 1.0, 0.05, key="dscr_g1o")

            st.subheader("Loan Structure")
            inputs["property_type"] = st.selectbox("Predominant Property Type",
                ["SFR", "Townhome", "Condo", "PUD", "2 Unit", "3 Unit", "4 Unit", "5 Unit", "6 Unit", "7 Unit", "8 Unit", "9 Unit"], key="dscr_pt")
            inputs["amortization"] = st.selectbox("Amortization", ["Fully Amortizing", "Interest Only"], key="dscr_am")
            inputs["rate_type"] = st.selectbox("Rate Type", ["FIXED 30", "5/1 ARM", "7/1 ARM"], key="dscr_rt")
            inputs["verified_liquidity"] = st.number_input("Verified Liquidity ($)", 0, step=10000, key="dscr_liq")
            inputs["loan_id"] = st.text_input("Loan ID", key="dscr_lid")

        with col2:
            st.subheader("Property Details")
            inputs["prop_address"] = st.text_input("Property Address", key="dscr_addr")
            inputs["prop_city"] = st.text_input("City", key="dscr_city")
            inputs["prop_state"] = st.text_input("State (2-letter)", key="dscr_state")
            inputs["prop_zip"] = st.number_input("ZIP Code", min_value=0, max_value=99999, value=0, key="dscr_zip", format="%d")
            inputs["prop_type"] = st.selectbox("Property Type",
                ["SFR", "Townhome", "Condo", "PUD", "2 Unit", "3 Unit", "4 Unit"], key="dscr_prop_type")
            inputs["prop_sqft"] = st.number_input("Square Footage", min_value=0, value=0, key="dscr_sqft")
            inputs["prop_num_units"] = st.number_input("Number of Units", min_value=1, max_value=9, value=1, key="dscr_units")
            inputs["prop_appraisal_date"] = date_to_datetime(st.date_input("Appraisal Date", key="dscr_apd"))
            inputs["prop_appraisal_value"] = st.number_input("Appraisal As-Is Value ($)", 0, step=10000, key="dscr_apv")
            inputs["prop_purchase_price"] = st.number_input("Purchase Price ($)", 0, step=10000, key="dscr_pp")

            st.subheader("Property Rent & Expenses")
            inputs["prop_monthly_rent"] = st.number_input("Monthly Rent in Place ($)", 0, step=100, key="dscr_rent")
            inputs["prop_market_rent"] = st.number_input("Monthly Market Rent ($)", 0, step=100, key="dscr_mktrent")
            inputs["prop_annual_taxes"] = st.number_input("Annual Taxes ($)", 0, step=500, key="dscr_taxes")
            inputs["prop_annual_hazard_ins"] = st.number_input("Annual Hazard Insurance ($)", 0, step=500, key="dscr_hazins")
            inputs["prop_annual_flood_ins"] = st.number_input("Annual Flood Insurance ($)", 0, step=500, key="dscr_floodins")
            inputs["prop_annual_hoa"] = st.number_input("Annual HOA Fees ($)", 0, step=500, key="dscr_hoa")

    # ----- MF Form -----
    elif loan_type == "MF":
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Property")
            inputs["address"] = st.text_input("Address", key="mf_addr")
            inputs["city"] = st.text_input("City", key="mf_city")
            inputs["state"] = st.text_input("State (2-letter)", key="mf_st")
            inputs["zip_code"] = st.number_input("ZIP Code", 0, 99999, 0, key="mf_zip", format="%d")
            inputs["num_units"] = st.number_input("Number of Units", 5, value=5, key="mf_units")

            st.subheader("Loan")
            inputs["closing_date"] = date_to_datetime(st.date_input("Expected Closing Date", key="mf_cd"))
            inputs["loan_program"] = st.selectbox("Loan Program", ["Bridge", "CAPEX"], key="mf_lp")
            inputs["loan_term"] = st.selectbox("Loan Term", ["12 Months", "18 Months", "24 Months", "36 Months"], key="mf_lt")

            st.subheader("Borrower")
            inputs["entity_name"] = st.text_input("Entity Name", key="mf_ent")
            inputs["guarantor_1_name"] = st.text_input("Guarantor 1 Name", key="mf_g1n")
            inputs["guarantor_1_fico"] = st.number_input("Guarantor 1 FICO", 0, 850, 700, key="mf_g1f")
            inputs["guarantor_1_credit_date"] = date_to_datetime(st.date_input("Credit Report Date", key="mf_g1d"))
            inputs["guarantor_1_ownership"] = st.number_input("% Ownership", 0.0, 1.0, 1.0, 0.05, key="mf_g1o")

        with col2:
            st.subheader("Valuation & Rehab")
            inputs["purchase_price"] = st.number_input("Purchase Price ($)", 0, step=10000, key="mf_pp")
            inputs["purchase_date"] = date_to_datetime(st.date_input("Purchase Date", key="mf_pd"))
            inputs["appraisal_date"] = date_to_datetime(st.date_input("Appraisal Date", key="mf_apd"))
            inputs["as_is_value"] = st.number_input("As-Is Value ($)", 0, step=10000, key="mf_aiv")
            inputs["arv"] = st.number_input("ARV ($)", 0, step=10000, key="mf_arv")
            inputs["rehab_budget"] = st.number_input("Rehab Budget ($)", 0, step=10000, key="mf_rb")
            inputs["pre_rehab_sqft"] = st.number_input("Pre-Rehab Sq Ft", 0, key="mf_prsf")
            inputs["post_rehab_sqft"] = st.number_input("Post-Rehab Sq Ft", 0, key="mf_posf")

            st.subheader("Property Economics")
            inputs["gross_potential_rev"] = st.number_input("Annual Gross Potential Revenue ($)", 0, step=1000, key="mf_gpr")
            inputs["opex_vacancy"] = st.number_input("Annual Opex & Vacancy ($)", 0, step=1000, key="mf_opex")
            inputs["annual_taxes"] = st.number_input("Annual Taxes ($)", 0, step=1000, key="mf_tax")
            inputs["annual_insurance"] = st.number_input("Annual Insurance ($)", 0, step=1000, key="mf_ins")

            st.subheader("Loan Proceeds")
            inputs["initial_loan_amount"] = st.number_input("Initial Loan Amount ($)", 0, step=10000, key="mf_ila")
            inputs["interest_reserves"] = st.number_input("Interest Reserves ($)", 0, step=1000, key="mf_ir")
            inputs["verified_liquidity"] = st.number_input("Verified Liquidity ($)", 0, step=10000, key="mf_liq")
            inputs["loan_id"] = st.text_input("Loan ID", key="mf_lid")

    # ----- GUC Form -----
    elif loan_type == "GUC":
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Property")
            inputs["address"] = st.text_input("Address", key="guc_addr")
            inputs["city"] = st.text_input("City", key="guc_city")
            inputs["state"] = st.text_input("State (2-letter)", key="guc_st")
            inputs["zip_code"] = st.number_input("ZIP Code", 0, 99999, 0, key="guc_zip", format="%d")
            inputs["num_units"] = st.number_input("Number of Units", 1, value=1, key="guc_units")

            st.subheader("Loan")
            inputs["closing_date"] = date_to_datetime(st.date_input("Expected Closing Date", key="guc_cd"))
            inputs["loan_term"] = st.selectbox("Loan Term", ["12 Months", "18 Months", "24 Months"], key="guc_lt")

            st.subheader("Borrower")
            inputs["entity_name"] = st.text_input("Entity Name", key="guc_ent")
            inputs["guarantor_1_name"] = st.text_input("Guarantor 1 Name", key="guc_g1n")
            inputs["guarantor_1_fico"] = st.number_input("Guarantor 1 FICO", 0, 850, 700, key="guc_g1f")
            inputs["guarantor_1_credit_date"] = date_to_datetime(st.date_input("Credit Report Date", key="guc_g1d"))
            inputs["guarantor_1_ownership"] = st.number_input("% Ownership", 0.0, 1.0, 1.0, 0.05, key="guc_g1o")
            inputs["g1_construction_sold"] = st.number_input("GUC Completed & Sold", 0, key="guc_exp_sold")
            inputs["g1_construction_rented"] = st.number_input("GUC Completed & Rented", 0, key="guc_exp_rented")

        with col2:
            st.subheader("Valuation & Rehab")
            inputs["purchase_price"] = st.number_input("Purchase Price ($)", 0, step=10000, key="guc_pp")
            inputs["purchase_date"] = date_to_datetime(st.date_input("Purchase Date", key="guc_pd"))
            inputs["appraisal_date"] = date_to_datetime(st.date_input("Appraisal Date", key="guc_apd"))
            inputs["as_is_value"] = st.number_input("As-Is Value ($)", 0, step=10000, key="guc_aiv")
            inputs["arv"] = st.number_input("ARV ($)", 0, step=10000, key="guc_arv")
            inputs["rehab_budget"] = st.number_input("Construction Budget ($)", 0, step=10000, key="guc_rb")
            inputs["post_completion_sqft"] = st.number_input("Post-Completion Sq Ft", 0, key="guc_posf")

            st.subheader("Leverage Deductions")
            inputs["entitled_land"] = st.selectbox("Entitled Land?", ["Yes", "No"], key="guc_el")
            inputs["approved_permits"] = st.selectbox("Approved Permits & Plans?", ["Yes", "No"], key="guc_ap")
            inputs["interest_reserves_flag"] = st.selectbox("Interest Reserves?", ["Yes", "No"], key="guc_irf")

            inputs["verified_liquidity"] = st.number_input("Verified Liquidity ($)", 0, step=10000, key="guc_liq")
            inputs["loan_id"] = st.text_input("Loan ID", key="guc_lid")

    # Pre-processing before fill
    if loan_type == "GUC":
        # Sum sold + rented into the single template cell
        sold = inputs.pop("g1_construction_sold", 0) or 0
        rented = inputs.pop("g1_construction_rented", 0) or 0
        inputs["g1_construction_completed"] = sold + rented

    if loan_type == "DSCR":
        # Lender origination always 2%
        inputs["lender_orig_pct"] = LENDER_ORIGINATION_FEE

    # Generate button
    if st.button("Generate Sizer", type="primary", key="sz_generate"):
        with st.spinner("Filling sizer template..."):
            result, count = fill_sizer(template_path, loan_type, inputs)

        st.success(f"Sizer filled — {count} cells written.")
        st.download_button(
            "⬇️ Download Completed Sizer",
            data=result,
            file_name=f"AS_Capital_{loan_type}_Sizer_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ===========================================================================
# TAB 2: UNDERWRITING CONDITIONS
# ===========================================================================
with tab2:
    st.header("Underwriting Conditions Generator")
    st.caption("Select loan type, enter deal details, and generate conditions using AI + Eastview guidelines.")

    if not api_ok:
        st.warning("Set your ANTHROPIC_API_KEY in .env to use this feature.")
    else:
        uw_type = st.selectbox("Loan Type", ["RTL", "DSCR", "MF", "GUC"],
            format_func=lambda x: {"RTL": "RTL (Fix & Flip / Bridge)",
                                    "DSCR": "DSCR (Rental)",
                                    "MF": "Multifamily (5+ Units)",
                                    "GUC": "Ground Up Construction"}.get(x, x),
            key="uw_type")

        # Check guidelines exist
        try:
            guidelines_path = get_guidelines_path(ASSETS_DIR, uw_type)
            if not os.path.exists(guidelines_path):
                st.error(f"Guidelines PDF not found: {guidelines_path}\nCopy it to the assets/ folder.")
                st.stop()
        except ValueError:
            st.error("Unknown loan type")
            st.stop()

        col1, col2 = st.columns(2)
        deal = {}
        with col1:
            st.subheader("Property")
            deal["property_address"] = st.text_input("Property Address", key="uw_addr")
            deal["city_state_zip"] = st.text_input("City, State ZIP", key="uw_csz")
            deal["property_type"] = st.text_input("Property Type", key="uw_pt")
            deal["num_units"] = st.number_input("Units", 1, key="uw_units")
            deal["as_is_value"] = st.number_input("As-Is Value ($)", 0, step=10000, key="uw_aiv")
            deal["arv"] = st.number_input("ARV ($)", 0, step=10000, key="uw_arv")
            deal["loan_amount"] = st.number_input("Loan Amount ($)", 0, step=10000, key="uw_la")
            deal["loan_purpose"] = st.selectbox("Loan Purpose", ["Purchase", "Refinance (Cash Out)", "Refinance (Rate & Term)"], key="uw_lp")
            deal["rehab_budget"] = st.number_input("Rehab Budget ($)", 0, step=10000, key="uw_rb")

        with col2:
            st.subheader("Borrower")
            deal["borrower_name"] = st.text_input("Borrower / Guarantor Name", key="uw_bn")
            deal["entity_name"] = st.text_input("Entity Name", key="uw_en")
            deal["fico_score"] = st.number_input("FICO Score", 0, 850, 700, key="uw_fico")
            deal["experience"] = st.number_input("Experience (# deals)", 0, key="uw_exp")
            deal["liquidity"] = st.number_input("Liquidity / Reserves ($)", 0, step=10000, key="uw_liq")
            deal["additional_notes"] = st.text_area("Additional Notes", key="uw_notes", height=150)

        if st.button("Generate Conditions", type="primary", key="uw_generate"):
            with st.spinner("Loading guidelines and generating conditions..."):
                guidelines_text = load_guidelines_pdf(guidelines_path)
                conditions = generate_conditions(
                    api_key=ANTHROPIC_API_KEY,
                    guidelines_text=guidelines_text,
                    loan_type=uw_type,
                    deal_details=deal,
                )
            st.subheader("Underwriting Conditions")
            st.markdown(conditions)
            st.download_button("⬇️ Download as Text", conditions,
                file_name=f"UW_Conditions_{uw_type}_{datetime.now().strftime('%Y%m%d')}.txt",
                mime="text/plain")


# ===========================================================================
# TAB 3: COMMITTEE DECK
# ===========================================================================
with tab3:
    st.header("Committee Presentation Builder")
    st.caption("Fill in deal details to generate a completed IC presentation.")

    if not os.path.exists(IC_TEMPLATE):
        st.error(f"IC template not found: {IC_TEMPLATE}\nCopy it to the assets/ folder.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Cover / Property")
            cd_addr = st.text_input("Property Address", key="cd_addr")
            cd_csz = st.text_input("City, State ZIP", key="cd_csz")
            cd_pt = st.text_input("Property Type", value="SFR", key="cd_pt")
            cd_txn = st.selectbox("Transaction Type", ["Purchase", "Refinance (Cash Out)", "Refinance (Rate & Term)"], key="cd_txn")
            cd_lt = st.selectbox("Loan Type", ["Fix & Flip", "Bridge", "Bridge Plus", "DSCR", "GUC"], key="cd_lt")
            cd_ln = st.text_input("Loan Number", key="cd_ln")
            cd_class = st.text_input("Classification (A+, A, B, C)", key="cd_class")

            st.subheader("Property Details")
            cd_yb = st.text_input("Year Built", key="cd_yb")
            cd_sf = st.text_input("Square Footage", key="cd_sf")
            cd_bldg = st.text_input("# Buildings", key="cd_bldg")
            cd_lotsf = st.text_input("Lot SF", key="cd_lotsf")
            cd_acres = st.text_input("Lot Acres", key="cd_acres")
            cd_sub = st.text_input("Subdivision", key="cd_sub")
            cd_term = st.text_input("Loan Term", value="12 Month", key="cd_term")

        with col2:
            st.subheader("Financial Metrics")
            cd_tla = st.number_input("Total Loan Amount ($)", 0, step=10000, key="cd_tla")
            cd_ltarv = st.number_input("LTV to ARV (%)", 0.0, 1.0, 0.0, 0.01, key="cd_ltarv")
            cd_rate = st.number_input("Interest Rate", value=0.0, step=0.005, format="%.3f", key="cd_rate")
            cd_pp = st.number_input("Purchase Price ($)", 0, step=10000, key="cd_pp")
            cd_ppd = st.text_input("Purchase Date", key="cd_ppd")
            cd_aiv = st.number_input("As-Is Value ($)", 0, step=10000, key="cd_aiv")
            cd_arv = st.number_input("After Repair Value ($)", 0, step=10000, key="cd_arv")
            cd_rhb = st.number_input("Rehab Budget ($)", 0, step=10000, key="cd_rhb")

            st.subheader("Loan Breakdown")
            cd_il = st.number_input("Initial Loan ($)", 0, step=10000, key="cd_il")
            cd_iltc = st.number_input("Initial LTC (%)", 0.0, 1.0, 0.0, 0.01, key="cd_iltc")
            cd_ir = st.number_input("Interest Reserve ($)", 0, step=1000, key="cd_ir")
            cd_hb = st.number_input("Holdback / Rehab ($)", 0, step=10000, key="cd_hb")
            cd_hbltc = st.number_input("Holdback LTC (%)", 0.0, 1.0, 0.0, 0.01, key="cd_hbltc")
            cd_tltc = st.number_input("Total LTC (%)", 0.0, 1.0, 0.0, 0.01, key="cd_tltc")

        st.subheader("Ownership")
        cd_ao = st.text_input("Assessment Owner", key="cd_ao")
        cd_be = st.text_input("Buyer Entity", key="cd_be")
        cd_ef = st.text_input("Existing Financing", key="cd_ef")

        cd_highlights = st.text_area("Investment Highlights (leave blank to auto-generate with AI)", key="cd_hl", height=120)

        if st.button("Build Committee Deck", type="primary", key="cd_generate"):
            with st.spinner("Building committee deck..."):
                result = build_committee_deck(
                    template_path=IC_TEMPLATE,
                    anthropic_api_key=ANTHROPIC_API_KEY if api_ok else "",
                    tavily_api_key=TAVILY_API_KEY if tavily_ok else "",
                    property_address=cd_addr,
                    city_state_zip=cd_csz,
                    property_type=cd_pt,
                    transaction_type=cd_txn,
                    loan_type=cd_lt,
                    loan_number=cd_ln,
                    total_loan_amount=cd_tla,
                    ltv_to_arv=cd_ltarv,
                    interest_rate=cd_rate,
                    purchase_price=cd_pp,
                    purchase_date=cd_ppd,
                    as_is_value=cd_aiv,
                    after_repair_value=cd_arv,
                    rehab_budget=cd_rhb,
                    initial_loan=cd_il,
                    initial_ltc=cd_iltc,
                    interest_reserve=cd_ir,
                    holdback_rehab=cd_hb,
                    holdback_ltc=cd_hbltc,
                    total_ltc=cd_tltc,
                    year_built=cd_yb,
                    square_footage=cd_sf,
                    num_buildings=cd_bldg,
                    lot_sf=cd_lotsf,
                    lot_acres=cd_acres,
                    subdivision=cd_sub,
                    loan_term=cd_term,
                    assessment_owner=cd_ao,
                    buyer_entity=cd_be,
                    existing_financing=cd_ef,
                    investment_highlights=cd_highlights,
                    classification=cd_class,
                )
            st.success("Committee deck generated!")
            st.download_button("⬇️ Download Committee Deck", data=result,
                file_name=f"AS_Capital_IC_{cd_addr.replace(' ','_')[:25]}_{datetime.now().strftime('%Y%m%d')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")


# ===========================================================================
# TAB 4: BORROWER DECK
# ===========================================================================
with tab4:
    st.header("Borrower Presentation Builder")
    st.caption("Generate a professional loan proposal for the borrower.")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Property")
        bp_addr = st.text_input("Property Address", key="bp_addr")
        bp_csz = st.text_input("City, State ZIP", key="bp_csz")
        bp_pt = st.text_input("Property Type", value="SFR", key="bp_pt")
        bp_units = st.number_input("Units", 1, key="bp_units")
        bp_sf = st.number_input("Square Footage", 0, key="bp_sf")
        bp_aiv = st.number_input("As-Is Value ($)", 0, step=10000, key="bp_aiv")
        bp_arv = st.number_input("ARV ($)", 0, step=10000, key="bp_arv")

    with col2:
        st.subheader("Loan")
        bp_borr = st.text_input("Borrower Name", key="bp_borr")
        bp_ent = st.text_input("Entity Name", key="bp_ent")
        bp_purpose = st.selectbox("Loan Purpose", ["Purchase", "Refinance", "Construction"], key="bp_purpose")
        bp_ltype = st.selectbox("Loan Type", ["RTL", "DSCR", "MF", "GUC"], key="bp_ltype")
        bp_amt = st.number_input("Loan Amount ($)", 0, step=10000, key="bp_amt")
        bp_term = st.number_input("Term (months)", 1, value=12, key="bp_term")
        bp_rate = st.number_input("Interest Rate", value=0.0, step=0.005, format="%.3f", key="bp_rate")
        bp_rhb = st.number_input("Rehab Budget ($)", 0, step=10000, key="bp_rhb")
        bp_irmo = st.number_input("Interest Reserve (months)", 0, key="bp_irmo")

    bp_reqs = st.text_area("Additional Requirements (one per line)", key="bp_reqs", height=80)

    c1, c2, c3 = st.columns(3)
    with c1:
        bp_cn = st.text_input("Contact Name", value="A&S Capital Originations", key="bp_cn")
    with c2:
        bp_ce = st.text_input("Contact Email", value="originations@ascapital.com", key="bp_ce")
    with c3:
        bp_cp = st.text_input("Contact Phone", value="305.749.0848", key="bp_cp")

    if st.button("Build Borrower Deck", type="primary", key="bp_generate"):
        if bp_amt == 0 or bp_aiv == 0:
            st.error("Enter at least a Loan Amount and As-Is Value.")
        else:
            with st.spinner("Building borrower deck..."):
                result = build_borrower_deck(
                    property_address=bp_addr, city_state_zip=bp_csz, property_type=bp_pt,
                    num_units=bp_units, square_footage=bp_sf, as_is_value=bp_aiv, arv=bp_arv,
                    borrower_name=bp_borr, entity_name=bp_ent, loan_purpose=bp_purpose,
                    loan_type=bp_ltype, loan_amount=bp_amt, loan_term_months=bp_term,
                    interest_rate=bp_rate, origination_fee_pct=LENDER_ORIGINATION_FEE, rehab_budget=bp_rhb,
                    interest_reserve_months=bp_irmo, additional_requirements=bp_reqs,
                    contact_name=bp_cn, contact_email=bp_ce, contact_phone=bp_cp,
                )
            st.success("Borrower deck generated!")
            st.download_button("⬇️ Download Borrower Deck", data=result,
                file_name=f"AS_Capital_Proposal_{bp_addr.replace(' ','_')[:25]}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.divider()
st.caption("A&S Capital AI Agent — Built with Streamlit, Claude, and Tavily")
