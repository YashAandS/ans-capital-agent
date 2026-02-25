"""
modules/underwriting.py
Underwriting Conditions Generator — loads the correct guidelines PDF based on
loan type, accepts deal details, sends both to Claude, returns conditions.

Supports 4 guideline types: RTL, DSCR, MF, GUC
"""

import os
import pdfplumber
from anthropic import Anthropic


# Map loan types to their guidelines PDF filenames
GUIDELINES_FILES = {
    "RTL":  "Eastview RTL Guidelines_v4.2 (1) (1).pdf",
    "DSCR": "Eastview DSCR Guidelines_v7.2 (1) (1).pdf",
    "MF":   "Eastview MF Guidelines_v1.0 (1) (1).pdf",
    "GUC":  "Eastview GUC Guidelines_v1.1 (2) (1) (1).pdf",
}


def load_guidelines_pdf(pdf_path: str) -> str:
    """Extract all text from a guidelines PDF. Returns a single string."""
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if text:
                pages.append(f"--- Page {i} ---\n{text}")
    return "\n\n".join(pages)


def get_guidelines_path(assets_dir: str, loan_type: str) -> str:
    """Get the path to the guidelines PDF for a given loan type."""
    filename = GUIDELINES_FILES.get(loan_type)
    if not filename:
        raise ValueError(f"Unknown loan type: {loan_type}")
    return os.path.join(assets_dir, filename)


def generate_conditions(
    api_key: str,
    guidelines_text: str,
    loan_type: str,
    deal_details: dict,
) -> str:
    """
    Send deal details + underwriting guidelines to Claude and get back
    a structured list of underwriting conditions.

    Args:
        api_key: Anthropic API key
        guidelines_text: Full text of the relevant guidelines PDF
        loan_type: One of RTL, DSCR, MF, GUC
        deal_details: Dict with deal info (address, amounts, borrower info, etc.)

    Returns:
        Formatted string of underwriting conditions
    """
    client = Anthropic(api_key=api_key)

    # Build a deal summary from the dict
    summary_lines = []
    for key, value in deal_details.items():
        if value is not None and value != "":
            label = key.replace("_", " ").title()
            summary_lines.append(f"{label}: {value}")
    deal_summary = "\n".join(summary_lines)

    loan_type_names = {
        "RTL": "Residential Transition Loan (Fix & Flip / Bridge)",
        "DSCR": "DSCR (Debt Service Coverage Ratio) Rental Loan",
        "MF": "Multifamily 5+ Unit Bridge Loan",
        "GUC": "Ground Up Construction Loan",
    }
    loan_name = loan_type_names.get(loan_type, loan_type)

    system_prompt = f"""You are a senior underwriter at A&S Capital, a private lending company
that originates {loan_name} loans through the Eastview platform.

Your job is to produce a clear, actionable list of underwriting conditions for each deal
based on the Eastview underwriting guidelines and the specific deal details provided.

For each condition, provide:
- A condition number
- The condition category (e.g., Property, Borrower, Title, Insurance, Legal, Appraisal,
  Construction, Environmental, Financial, Closing, Third-Party Reports)
- A clear description of what is required
- Whether it is "Prior to Funding", "Prior to Closing", or "Prior to First Draw" (for GUC)

Be specific to THIS deal — reference actual guideline sections, LTV/DSCR thresholds,
FICO requirements, experience tiers, and property type requirements from the guidelines.
Flag any areas where the deal may not meet standard guidelines and note exceptions
that would need committee approval.

Format the output as a clean, numbered list grouped by category. End with a summary of
any flags or exceptions noted."""

    user_message = f"""Here are the Eastview {loan_name} underwriting guidelines:

{guidelines_text}

---

Here is the deal to underwrite:

LOAN TYPE: {loan_name}

{deal_summary}

---

Please produce the full list of underwriting conditions for this deal. Group by
category and note timing (Prior to Funding / Prior to Closing / Prior to First Draw).
Reference specific guideline requirements where applicable."""

    response = client.messages.create(
        model="claude-sonnet-4-5-20250514",
        max_tokens=4096,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
    )

    return response.content[0].text
