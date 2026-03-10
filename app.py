"""
app.py — A&S Capital AI Agent
Main Streamlit UI with three core capabilities:
  1. The Sizernator — Automatic Sizer Filler (RTL / DSCR / MF / GUC)
  2. Dealfit — Capital Partner Comparison (Colchis, Fidelis, Eastview)
  3. Committee Presentation Builder

Run with:  streamlit run app.py
"""

import os
from datetime import datetime
import streamlit as st
from dotenv import load_dotenv

from modules.auto_sizer import auto_fill_sizer
from modules.committee_deck import build_committee_deck
from modules.dealfit import run_dealfit, read_sizer

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
    page_title="Roberto Jr. — A&S Capital",
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
api_ok = bool(ANTHROPIC_API_KEY and ANTHROPIC_API_KEY != "your-anthropic-api-key-here")
tavily_ok = bool(TAVILY_API_KEY and TAVILY_API_KEY != "your-tavily-api-key-here")

with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)
    else:
        st.title("A&S Capital")
    st.markdown("### Roberto Jr.")
    st.caption("AI-Powered Deal Agent")




# ---------------------------------------------------------------------------
# Main content
# ---------------------------------------------------------------------------
st.title("A&S Capital — Roberto Jr.")

tab0, tab_dealfit, tab3 = st.tabs([
    "⚡ The Sizernator",
    "🎯 Dealfit",
    "🏛️ Committee Deck",
])


# ===========================================================================
# TAB 0: AUTOMATIC SIZER FILLER
# ===========================================================================
with tab0:
    st.header("The Sizernator")
    st.caption("Drop any deal documents below — the AI will read everything, fill the correct sizer, and list anything it couldn't find.")

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

                # Fields exempt from the "Not Found" list (unimportant / user-fills-later)
                EXEMPT_FIELDS = {
                    "closing_date", "entity_name", "num_guarantors", "num_owners",
                    "guarantor_1_first", "guarantor_1_last",
                    "guarantor_1_name", "guarantor_1_credit_date",
                    "guarantor_2_first", "guarantor_2_last",
                    "guarantor_2_name", "guarantor_2_credit_date",
                    "guarantor_3_first", "guarantor_3_last",
                    "guarantor_3_name", "guarantor_3_credit_date",
                    "guarantor_4_first", "guarantor_4_last",
                    "guarantor_4_name", "guarantor_4_credit_date",
                    "amortization", "rate_type",
                    "prop_sqft", "prop_appraisal_date", "prop1_appraisal_date",
                    "appraisal_date", "prop1_pre_rehab_sqft",
                    "post_completion_sqft", "verified_liquidity",
                }

                # Filter missing fields: remove optional and exempt ones
                display_missing = [
                    f for f in missing_fields
                    if f not in EXEMPT_FIELDS
                ]

                # Persist results in session_state so they survive reruns (e.g. download click)
                st.session_state["sizer_result"] = {
                    "sizer_bytes": sizer_bytes,
                    "detected_type": detected_type,
                    "filled_count": filled_count,
                    "display_missing": display_missing,
                    "extracted": extracted,
                }

        # ---- Display persisted results (survives download-click reruns) ----
        if "sizer_result" in st.session_state:
            res = st.session_state["sizer_result"]

            st.success(f"Sizer filled — **{res['detected_type']}** loan detected, **{res['filled_count']}** cells written.")

            # Layout: download button + not-found list side by side
            dl_col, nf_col = st.columns([1, 2])

            with dl_col:
                st.download_button(
                    "⬇️ Download Completed Sizer",
                    data=res["sizer_bytes"],
                    file_name=f"AS_Capital_{res['detected_type']}_Sizer_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with nf_col:
                if res["display_missing"]:
                    st.markdown(f"**📝 Not Found ({len(res['display_missing'])})** — fill these in Excel:")
                    for field in res["display_missing"]:
                        label = field.replace("_", " ").title()
                        st.markdown(f"&nbsp;&nbsp;• {label}")
                else:
                    st.success("✅ All important fields were found and filled!")

            # Show what was extracted
            with st.expander("📋 Extracted Fields (click to review)", expanded=False):
                for k, v in res["extracted"].items():
                    label = k.replace("_", " ").title()
                    st.markdown(f"**{label}:** {v}")




# ===========================================================================
# TAB: DEALFIT — Capital Partner Comparison
# ===========================================================================
AS_SIZER_TEMPLATE = os.path.join(ASSETS_DIR, "AS_Capital_Sizer.xlsx")

with tab_dealfit:
    st.header("Dealfit")
    st.caption("Upload your filled A&S Capital Sizer — Roberto Jr. will evaluate the deal against Colchis, Fidelis, and Eastview to find the best capital partner fit.")

    if not api_ok:
        st.warning("Set your ANTHROPIC_API_KEY in .env to use this feature.")
    else:
        # Download blank sizer template
        dl_col, upload_col = st.columns([1, 3])
        with dl_col:
            if os.path.exists(AS_SIZER_TEMPLATE):
                with open(AS_SIZER_TEMPLATE, "rb") as tf:
                    st.download_button(
                        "📄 Download Blank Sizer",
                        data=tf.read(),
                        file_name="AS_Capital_Sizer.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Download the blank A&S Capital Sizer template to fill out",
                    )

        st.markdown(
            """
            <div style="background-color:#E0F0F8; border:2px dashed #A3D5E0; border-radius:10px;
                        padding:15px; margin-bottom:15px; text-align:center;">
                <p style="margin:0; color:#2C3E50; font-size:16px;">
                    📊 <strong>Upload Filled Sizer</strong> — Drop your completed A&S Capital Sizer here
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        sizer_file = st.file_uploader(
            "Upload Filled A&S Capital Sizer",
            type=["xlsx"],
            key="dealfit_upload",
            help="Upload the A&S Capital Sizer Excel file you filled out",
        )

        if sizer_file:
            st.info(f"📄 Uploaded: **{sizer_file.name}**")

            if st.button("🎯 Analyze Deal", type="primary", key="dealfit_run"):
                with st.spinner("Reading sizer and evaluating against all capital partners..."):
                    file_bytes = sizer_file.getvalue()
                    deal, results, recommendation = run_dealfit(
                        api_key=ANTHROPIC_API_KEY,
                        file_bytes=file_bytes,
                    )

                # Store results in session state
                st.session_state["dealfit_result"] = {
                    "deal": deal,
                    "results": results,
                    "recommendation": recommendation,
                }

        # Display persisted results
        if "dealfit_result" in st.session_state:
            dfr = st.session_state["dealfit_result"]
            deal = dfr["deal"]
            results = dfr["results"]
            recommendation = dfr["recommendation"]

            # --- Deal Summary ---
            st.markdown("---")
            st.subheader("Deal Summary")
            sc1, sc2, sc3, sc4 = st.columns(4)
            with sc1:
                st.metric("Deal Type", deal.deal_type or "N/A")
                st.metric("Transaction", deal.transaction_type or "N/A")
            with sc2:
                st.metric("Total Loan", f"${deal.total_loan:,.0f}" if deal.total_loan else "N/A")
                st.metric("FICO", str(deal.primary_fico) if deal.primary_fico else "N/A")
            with sc3:
                st.metric("LTV", f"{deal.ltv:.1%}" if deal.ltv else "N/A")
                st.metric("LTARV", f"{deal.ltarv:.1%}" if deal.ltarv else "N/A")
            with sc4:
                st.metric("LTC", f"{deal.ltc:.1%}" if deal.ltc else "N/A")
                st.metric("Experience", f"{deal.completed_projects} projects")

            # --- AI Recommendation ---
            st.markdown("---")
            st.subheader("🎯 Recommendation")
            st.markdown(
                f'<div style="background-color:#E0F0F8; border-left:4px solid #A3D5E0; '
                f'padding:15px; border-radius:5px; color:#2C3E50;">{recommendation}</div>',
                unsafe_allow_html=True,
            )

            # --- Partner Comparison Table ---
            st.markdown("---")
            st.subheader("Capital Partner Comparison")

            for r in results:
                if r.eligible:
                    color = "#E8F8F0"   # green tint
                    icon = "✅"
                    border = "#2ECC71"
                else:
                    color = "#FDEDEC"   # red tint
                    icon = "❌"
                    border = "#E74C3C"

                with st.container():
                    st.markdown(
                        f'<div style="background-color:{color}; border-left:4px solid {border}; '
                        f'padding:12px; border-radius:5px; margin-bottom:10px;">'
                        f'<h4 style="margin:0; color:#2C3E50;">{icon} {r.partner_name}</h4></div>',
                        unsafe_allow_html=True,
                    )

                    if r.eligible:
                        mc1, mc2, mc3, mc4 = st.columns(4)
                        with mc1:
                            st.markdown(f"**Max LTV:** {r.max_ltv:.1%}" if r.max_ltv else "**Max LTV:** N/A")
                            st.markdown(f"**Max LTARV:** {r.max_ltarv:.1%}" if r.max_ltarv else "**Max LTARV:** N/A")
                        with mc2:
                            st.markdown(f"**Max LTC:** {r.max_ltc:.1%}" if r.max_ltc else "**Max LTC:** N/A")
                            if r.max_loan_amount > 0:
                                st.markdown(f"**Max Loan:** ${r.max_loan_amount:,.0f}")
                        with mc3:
                            if r.estimated_rate is not None:
                                st.markdown(f"**Est. Rate:** {r.estimated_rate:.3%}")
                            else:
                                st.markdown(f"**Est. Rate:** {r.rate_notes}")
                        with mc4:
                            st.markdown(f"**Tier:** {r.experience_tier}")
                            st.markdown(f"**FICO:** {r.fico_tier}")

                        if r.warnings:
                            with st.expander(f"⚠️ Warnings ({len(r.warnings)})", expanded=False):
                                for w in r.warnings:
                                    st.markdown(f"- {w}")
                    else:
                        st.markdown("**Ineligible:**")
                        for reason in r.ineligible_reasons:
                            st.markdown(f"- {reason}")

                    st.markdown("")  # spacer

            # --- Full Deal Details Expander ---
            with st.expander("📋 Full Deal Details", expanded=False):
                dc1, dc2 = st.columns(2)
                with dc1:
                    st.markdown("**Property**")
                    st.markdown(f"Address: {deal.address}")
                    st.markdown(f"Location: {deal.city}, {deal.state} {deal.zip_code}")
                    st.markdown(f"Type: {deal.property_type} ({deal.num_units} units)")
                    st.markdown(f"Sq Ft: {deal.sqft:,}" if deal.sqft else "Sq Ft: N/A")
                    st.markdown("**Valuation**")
                    st.markdown(f"Purchase Price: ${deal.purchase_price:,.0f}")
                    st.markdown(f"As-Is Value: ${deal.as_is_value:,.0f}")
                    st.markdown(f"ARV: ${deal.arv:,.0f}")
                    st.markdown(f"Rehab Budget: ${deal.rehab_budget:,.0f}")
                with dc2:
                    st.markdown("**Loan Request**")
                    st.markdown(f"Loan Amount: ${deal.loan_amount:,.0f}")
                    st.markdown(f"Rehab Holdback: ${deal.rehab_holdback:,.0f}")
                    st.markdown(f"Interest Reserve: ${deal.interest_reserve:,.0f}")
                    st.markdown(f"**Total Loan: ${deal.total_loan:,.0f}**")
                    st.markdown("**Borrower**")
                    st.markdown(f"Entity: {deal.entity_name}")
                    st.markdown(f"FICO: {deal.primary_fico}")
                    st.markdown(f"Experience: {deal.completed_projects} projects")
                    st.markdown(f"Liquidity: ${deal.verified_liquidity:,.0f}")
                    if deal.zhvi > 0:
                        st.markdown(f"ZHVI: ${deal.zhvi:,.0f} | Ratio: {deal.value_zhvi_ratio:.1%}")


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


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.divider()
st.caption("Roberto Jr. — A&S Capital AI Deal Agent")
