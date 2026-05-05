"""
Case-Specific Risk Assessment Generator
Giant Compliance Ltd | AML/CFT Compliance Tools

Streamlit web app — enter a company number and case details,
download a pre-filled Word document in seconds.
"""

import os
import sys
import re
import io
import urllib.parse
from datetime import date, datetime

import streamlit as st
import requests
from dateutil import relativedelta

# Add parent directory so we can import generate_case_ra helpers
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_case_ra import (
    fetch_company, fetch_officers, fetch_pscs, fetch_filing_count,
    build_document, clean_company_number, sic_desc
)

# ─── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Case Risk Assessment | Giant Compliance",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─── Styling ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main > div { padding-top: 1.5rem; }
    .stButton > button {
        width: 100%;
        background-color: #1F3864;
        color: white;
        border: none;
        padding: 0.6rem 1rem;
        font-size: 1rem;
        border-radius: 6px;
    }
    .stButton > button:hover { background-color: #2E75B6; }
    .stDownloadButton > button {
        width: 100%;
        background-color: #375623;
        color: white;
        border: none;
        padding: 0.75rem 1rem;
        font-size: 1.05rem;
        border-radius: 6px;
    }
    .stDownloadButton > button:hover { background-color: #4d7a30; }
    h1 { color: #1F3864 !important; font-size: 1.5rem !important; }
    h3 { color: #2E75B6 !important; font-size: 1.05rem !important; margin-top: 1.25rem !important; }
    .info-box {
        background: #EEF4FB;
        border-left: 4px solid #2E75B6;
        padding: 0.75rem 1rem;
        border-radius: 0 6px 6px 0;
        font-size: 0.9rem;
        margin: 0.75rem 0;
    }
    .success-box {
        background: #C6EFCE;
        border-left: 4px solid #375623;
        padding: 0.75rem 1rem;
        border-radius: 0 6px 6px 0;
        font-size: 0.9rem;
        margin: 0.75rem 0;
    }
    .warn-box {
        background: #FFEB9C;
        border-left: 4px solid #7D4C00;
        padding: 0.75rem 1rem;
        border-radius: 0 6px 6px 0;
        font-size: 0.9rem;
        margin: 0.75rem 0;
    }
    .ch-tag {
        display: inline-block;
        background: #2E75B6;
        color: white;
        font-size: 0.7rem;
        padding: 1px 6px;
        border-radius: 3px;
        margin-left: 4px;
        vertical-align: middle;
    }
    .app-tag {
        display: inline-block;
        background: #375623;
        color: white;
        font-size: 0.7rem;
        padding: 1px 6px;
        border-radius: 3px;
        margin-left: 4px;
        vertical-align: middle;
    }
</style>
""", unsafe_allow_html=True)

# ─── Password gate ─────────────────────────────────────────────────────────────
def check_password() -> bool:
    """Simple password gate. Password stored in st.secrets or env var."""
    correct = st.secrets.get("APP_PASSWORD", os.environ.get("APP_PASSWORD", ""))

    if not correct:
        return True

    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False

    if st.session_state["authenticated"]:
        return True

    st.markdown("### Sign in")
    pwd = st.text_input("Password", type="password", placeholder="Enter access password")
    if st.button("Continue"):
        if pwd == correct:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False

# ─── CH API key ────────────────────────────────────────────────────────────────
def get_api_key() -> str:
    return st.secrets.get("CH_API_KEY", os.environ.get("CH_API_KEY", ""))

# ─── Helpers ──────────────────────────────────────────────────────────────────
APPT_TYPES = [
    "Creditors' Voluntary Liquidation (CVL)",
    "Administration",
    "Members' Voluntary Liquidation (MVL)",
    "Administrative Receivership",
    "Company Voluntary Arrangement (CVA)",
    "Individual Voluntary Arrangement (IVA)",
    "Bankruptcy",
    "Compulsory Liquidation / Winding-Up",
    "Other",
]

def fmt_date_display(raw: str) -> str:
    """Format YYYY-MM-DD to DD/MM/YYYY for display."""
    try:
        return datetime.strptime(raw, "%Y-%m-%d").strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        return raw or "—"

# Source catalogue for adverse media screening
_SOURCES = [
    ("OFSI",              "OFSI (UK Sanctions)",
     "https://www.gov.uk/government/publications/financial-sanctions-consolidated-list-of-targets/consolidated-list-of-targets#search:{q}"),
    ("OFAC",              "OFAC (US Sanctions)",
     "https://sanctionssearch.ofac.treas.gov/Results.aspx?txtLastName={q}"),
    ("UN",                "UN Sanctions",
     "https://www.un.org/securitycouncil/sanctions/search?keywords={q}"),
    ("EU",                "EU Sanctions",
     "https://eam.esma.europa.eu/search?term={q}"),
    ("CH Disqualified",   "CH Disqualified Directors",
     "https://find-and-update.company-information.service.gov.uk/search/disqualified-officers?q={q}"),
    ("Insolvency Service","Insolvency Service Register",
     "https://www.insolvencydirect.bis.gov.uk/IIR/index.asp?search={q}"),
    ("FCA Register",      "FCA Register",
     "https://register.fca.org.uk/s/search#search={q}"),
    ("Find Case Law",     "Find Case Law",
     "https://caselaw.nationalarchives.gov.uk/search?query={q}"),
    ("Google News",       "Google News",
     "https://news.google.com/search?q={q}"),
    ("FT",                "Financial Times",
     "https://www.ft.com/search?q={q}"),
    ("BIJ",               "Bureau of Investigative Journalism",
     "https://www.thebureauinvestigates.com/?s={q}"),
    ("ICIJ",              "ICIJ Offshore Leaks / Panama Papers",
     "https://offshoreleaks.icij.org/search?q={q}"),
    ("Wikipedia",         "Wikipedia",
     "https://en.wikipedia.org/w/index.php?search={q}"),
    ("Wikileaks",         "Wikileaks",
     "https://search.wikileaks.org/?q={q}"),
]

_CATEGORIES = [
    ("Sanctions Lists",               ["OFSI", "OFAC", "UN", "EU"]),
    ("Official Registers",            ["CH Disqualified", "Insolvency Service", "FCA Register", "Find Case Law"]),
    ("Media & Investigative Sources", ["Google News", "FT", "BIJ", "ICIJ", "Wikipedia", "Wikileaks"]),
]

_MEDIA_KEYS = {"Google News", "FT", "BIJ", "ICIJ", "Wikipedia", "Wikileaks"}

_source_map = {k: (lbl, tpl) for k, lbl, tpl in _SOURCES}

# ─── Main app ──────────────────────────────────────────────────────────────────
def main():
    # Header
    st.title("Case Risk Assessment Generator")
    st.caption("AML/CFT Compliance Tools  ·  Giant Compliance Ltd")

    api_key = get_api_key()
    if not api_key:
        st.error("Companies House API key not configured. Contact Giant Compliance.")
        st.stop()

    # ── Step 1: Company lookup ──────────────────────────────────────────────
    st.markdown("### 1  Company number")

    col1, col2 = st.columns([3, 1])
    with col1:
        company_input = st.text_input(
            "Companies House registration number",
            placeholder="e.g. 12345678 or SC123456",
            label_visibility="collapsed",
        )
    with col2:
        lookup_btn = st.button("Look up", use_container_width=True)

    company_data = {}
    officers     = {"current": [], "resigned": []}
    pscs         = []
    filing       = {}

    if lookup_btn or ("company_data" in st.session_state and company_input):
        if not company_input.strip():
            st.warning("Please enter a company number.")
        else:
            cn = clean_company_number(company_input)
            with st.spinner(f"Fetching {cn} from Companies House…"):
                company_data = fetch_company(cn, api_key)
                if not company_data.get("name"):
                    st.error(f"Company {cn} not found on Companies House. Please check the number.")
                    st.stop()
                officers = fetch_officers(cn, api_key)
                pscs     = fetch_pscs(cn, api_key)
                filing   = fetch_filing_count(cn, api_key)

            # Cache in session
            st.session_state["company_data"] = company_data
            st.session_state["officers"]     = officers
            st.session_state["pscs"]         = pscs
            st.session_state["filing"]       = filing

    # Restore from session if navigated back
    if not company_data and "company_data" in st.session_state:
        company_data = st.session_state["company_data"]
        officers     = st.session_state["officers"]
        pscs         = st.session_state["pscs"]
        filing       = st.session_state["filing"]

    # Show CH summary card
    if company_data.get("name"):
        overdue = company_data.get("accounts_overdue", False)
        sic_list = company_data.get("sic_codes", [])
        sic_text = "; ".join(f"{c} — {sic_desc(c)}" for c in sic_list) if sic_list else "Not registered"

        cur_dirs = [o["name"] for o in officers.get("current", [])]
        psc_names = [p["name"] for p in pscs]

        box_class = "warn-box" if overdue else "success-box"
        overdue_note = " <strong>⚠ Accounts overdue at CH</strong>" if overdue else ""

        st.markdown(f"""
        <div class="{box_class}">
            <strong>{company_data['name']}</strong> &nbsp;·&nbsp; {company_data['number']}<br>
            Status: {company_data.get('status', '').replace('-',' ').title()}&nbsp;&nbsp;
            Incorporated: {company_data.get('inc_date', '—')}&nbsp;&nbsp;
            Trading: {company_data.get('trading_years', '—')}{overdue_note}<br>
            SIC: {sic_text}<br>
            Directors: {', '.join(cur_dirs) if cur_dirs else '—'}<br>
            PSCs: {', '.join(psc_names) if psc_names else 'None registered'}
        </div>
        """, unsafe_allow_html=True)

        # ── Step 2: Case details ────────────────────────────────────────────
        st.markdown("### 2  Case details")
        st.caption("These are added to the document. All fields optional except appointment type.")

        case_ref    = st.text_input("Case reference", placeholder="e.g. CVL001")
        appt_type   = st.selectbox("Appointment type", APPT_TYPES)
        appt_date   = st.date_input(
            "Date of appointment",
            value=None,
            min_value=date(2000, 1, 1),
            max_value=date.today(),
            format="DD/MM/YYYY",
        )
        ip_name     = st.text_input("IP / Officeholder name", placeholder="Full name")
        ip_licence  = st.text_input("IP licence number", placeholder="e.g. IPA12345")
        assessed_by = st.text_input("Assessed by", placeholder="Name of person completing this form")

        # ── Step 3: Adverse Media & Sanctions Screening ─────────────────────
        st.markdown("### 3  Adverse media &amp; sanctions screening")
        st.markdown("""
        <div class="info-box">
            Click a principal's name to open that search in a new tab. Record any findings in the
            boxes below &mdash; these pre-fill Section&nbsp;5 of the document. All fields are optional;
            leave blank if nothing to report. Set <strong>Concern?</strong> to
            <strong>Y</strong> for anything that needs to appear in the Evidence Log.
        </div>
        """, unsafe_allow_html=True)

        # Build deduplicated principal list
        _seen = set()
        principals_to_screen = []
        for _o in officers.get("current", []):
            if _o["name"] not in _seen:
                _seen.add(_o["name"])
                principals_to_screen.append(_o["name"])
        for _p in pscs:
            if _p["name"] not in _seen:
                _seen.add(_p["name"])
                principals_to_screen.append(_p["name"])

        if principals_to_screen:
            st.markdown(
                "**Principals to screen:** "
                + " · ".join(f"`{n}`" for n in principals_to_screen)
            )
        else:
            st.caption("No principals found from Companies House.")

        adverse_findings = []

        for _cat_emoji, _cat_label, _cat_keys in [
            ("\U0001f512", "Sanctions Lists",               ["OFSI", "OFAC", "UN", "EU"]),
            ("\U0001f3db️", "Official Registers",      ["CH Disqualified", "Insolvency Service", "FCA Register", "Find Case Law"]),
            ("\U0001f4f0", "Media & Investigative Sources", ["Google News", "FT", "BIJ", "ICIJ", "Wikipedia", "Wikileaks"]),
        ]:
            with st.expander(f"{_cat_emoji}  {_cat_label}", expanded=False):
                for _src_key in _cat_keys:
                    _src_label, _url_tpl = _source_map[_src_key]
                    st.markdown(f"**{_src_label}**")

                    # Per-principal search links
                    _link_parts = []
                    for _name in principals_to_screen:
                        _q = urllib.parse.quote_plus(_name)
                        _u = _url_tpl.replace("{q}", _q)
                        _link_parts.append(f"[{_name}]({_u})")
                    # Entity search for media sources
                    if _src_key in _MEDIA_KEYS and company_data.get("name"):
                        _qco = urllib.parse.quote_plus(company_data["name"])
                        _uco = _url_tpl.replace("{q}", _qco)
                        _link_parts.append(f"[{company_data['name']} *(entity)*]({_uco})")

                    if _link_parts:
                        st.markdown("Search: " + " &nbsp;·&nbsp; ".join(_link_parts))
                    else:
                        st.markdown(f"[Open {_src_label}]({_url_tpl.replace('{q}', '')})")

                    _col_f, _col_u, _col_c = st.columns([4, 3, 1])
                    with _col_f:
                        _finding = st.text_area(
                            "Finding",
                            key=f"af_finding_{_src_key}",
                            placeholder="Describe result — or NIL if clear",
                            height=70,
                            label_visibility="collapsed",
                        )
                    with _col_u:
                        _url_in = st.text_input(
                            "URL / Reference",
                            key=f"af_url_{_src_key}",
                            placeholder="https://… or document title",
                            label_visibility="collapsed",
                        )
                    with _col_c:
                        _concern = st.selectbox(
                            "Concern?",
                            ["—", "Y", "N"],
                            key=f"af_concern_{_src_key}",
                            label_visibility="collapsed",
                        )

                    if _finding.strip() or _url_in.strip() or _concern in ("Y", "N"):
                        adverse_findings.append({
                            "source":  _src_key,
                            "persons": ", ".join(principals_to_screen),
                            "date":    date.today().strftime("%d/%m/%Y"),
                            "finding": _finding.strip(),
                            "url":     _url_in.strip(),
                            "concern": _concern if _concern in ("Y", "N") else "",
                        })
                    st.divider()

        # Summary badge if any concerns flagged
        _flagged = [f for f in adverse_findings if f.get("concern") == "Y"]
        if _flagged:
            st.warning(
                f"⚠️ **{len(_flagged)} concern(s) flagged** — "
                + ", ".join(f["source"] for f in _flagged)
                + ". These will appear in the Evidence Log (Section 5, Part E)."
            )

        # ── Step 4: Generate ────────────────────────────────────────────────
        st.markdown("### 4  Generate document")
        st.markdown("""
        <div class="info-box">
            The document will be pre-filled with Companies House data (marked
            <span class="ch-tag">CH</span>) and any screening results entered above
            (marked <span class="app-tag">APP</span>). Sections for PEP screening,
            source of wealth, insolvency risk flags and the overall risk rating are
            left blank for completion on file.
        </div>
        """, unsafe_allow_html=True)

        if st.button("Generate risk assessment"):
            appt_date_str = appt_date.strftime("%d/%m/%Y") if appt_date else ""

            case_inputs = {
                "case_ref":    case_ref,
                "appt_type":   appt_type,
                "appt_date":   appt_date_str,
                "ip_name":     ip_name,
                "ip_licence":  ip_licence,
                "assessed_by": assessed_by,
            }

            with st.spinner("Building document…"):
                doc = build_document(
                    company_data, officers, pscs, filing, case_inputs,
                    adverse_findings=adverse_findings,
                )

                # Save to bytes buffer (no temp file needed)
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)

                # Build filename
                safe_name = re.sub(r'[^\w\s-]', '', company_data["name"]).strip().replace(" ", "_")[:40]
                filename  = f"Case_RA_{company_data['number']}_{safe_name}.docx"

                st.session_state["doc_bytes"] = buf.getvalue()
                st.session_state["doc_name"]  = filename

        # Show download button if document is ready
        if "doc_bytes" in st.session_state:
            st.success("Document ready.")
            st.download_button(
                label="Download risk assessment (.docx)",
                data=st.session_state["doc_bytes"],
                file_name=st.session_state["doc_name"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
            st.markdown("""
            <div class="info-box">
                <strong>Next steps:</strong> Open the document and verify the Companies House data.
                Section&nbsp;5 (Adverse Media) is pre-filled from your screening above &mdash; add any
                further narrative and complete Sections&nbsp;6&ndash;9 on file. Save to the case file
                and retain for a minimum of five years from the end of the engagement (MLR 2017).
            </div>
            """, unsafe_allow_html=True)

    # ── Footer ──────────────────────────────────────────────────────────────
    st.divider()
    st.caption(
        "Giant Compliance Ltd  ·  AML/CFT Compliance Tools  ·  "
        "Data sourced from Companies House Open Data. "
        "Always verify auto-populated information before relying on it."
    )


# ─── Entry point ──────────────────────────────────────────────────────────────
if check_password():
    main()
