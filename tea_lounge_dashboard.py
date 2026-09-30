# Run locally with:  streamlit run tea_lounge_dashboard.py
import os
import re
import io
import json
import hmac
import html
import shutil
import calendar
from datetime import date, datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ============================================================
# TEA LOUNGE INVENTORY TRACKER — Monthly History Edition
# Run:
#     pip install streamlit pandas numpy openpyxl plotly
#     python tea_lounge_dashboard.py
#
# HOW MONTHLY HISTORY WORKS
# --------------------------------------------------------------
# Upload one workbook at a time (or many, over time). Any sheet whose
# name contains a month + year — e.g. "January26", "Feb 2026",
# "March-2026" — is read as an Ingredient Summary & Inventory report
# for that month and saved to disk in ./tea_lounge_store/. Once saved,
# that month stays available forever (across restarts), and you browse
# history with the Year / Month selectors in the sidebar — you do not
# need to re-upload anything to see a past month again.
#
# A sheet whose name contains "item sales" (e.g. "Monthly Item Sales
# Report") is read separately for its "Total Quantity" column and
# matched to a month either via a Month/Period column inside the sheet,
# via a month name in the sheet's own name, or — if the workbook only
# contains one month of inventory data — attributed to that month.
#
# Expected inventory columns (same as before):
#     Ingredient / Item, Total Units Sold, Total Required Qty,
#     Projected Use, Pre. Month Stock, Received 1, Received 2,
#     Total Stock, Month End Stock, Total Use
# Expected item-sales columns:
#     Item, Total Quantity  (optionally a Month / Period column)
# ============================================================

st.set_page_config(
    page_title="Tea Lounge Inventory Tracker",
    page_icon="🍃",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# OPTIONAL PASSWORD GATE
# When the dashboard is on a public URL, anyone with the link could see the
# sales figures. Set a secret / environment variable called APP_PASSWORD on
# the host and the app asks for it first. If it isn't set (e.g. on your own
# PC) nothing changes and no password is asked.
# -----------------------------

def require_password():
    try:
        expected = st.secrets.get("APP_PASSWORD")
    except Exception:
        expected = None
    expected = expected or os.environ.get("APP_PASSWORD")
    if not expected or st.session_state.get("_authed"):
        return

    st.markdown("### 🍃 Tea Lounge")
    typed = st.text_input("Password", type="password", key="_pw_input")
    if typed:
        if hmac.compare_digest(str(typed), str(expected)):
            st.session_state["_authed"] = True
            st.rerun()
        st.error("Wrong password.")
    st.stop()


require_password()

# -----------------------------
# THEME DETECTION
# Streamlit's Settings menu lets a person flip between light/dark; we read
# that live choice (falling back to the app's configured default, then to
# "light") and pick a matching color palette below, so the dashboard's own
# background/cards genuinely switch instead of always forcing light mint.
# -----------------------------

def detect_theme_mode():
    try:
        t = st.context.theme.type
        if t in ("light", "dark"):
            return t
    except Exception:
        pass
    try:
        base = st.get_option("theme.base")
        if base in ("light", "dark"):
            return base
    except Exception:
        pass
    return "light"


THEME_MODE = "light"  # the whole app is light-only now (detect_theme_mode() is no longer used)

PALETTES = {
    "light": {
        "main_bg": "#f2fbf6",
        "section_title": "#0b3d2e",
        "small_note": "#4f8a72",
        "card_bg": "#e6f9f0",
        "card_border": "#bfe9d3",
        "card_label": "#4b6358",
        "card_value": "#0b3d2e",
        "card_sub": "#7a8f86",
        "warn_bg": "#fff7ea",
        "warn_border": "#f3d79a",
        "warn_text": "#7a4b06",
        "warn_sub": "#6b4a10",
        "alert_bg": "#fff7ea",
        "alert_border": "#f3d79a",
        "alert_text": "#6b4a10",
        "alert_strong": "#7a4b06",
        "chart_font": "#0b3d2e",
    },
    "dark": {
        "main_bg": "#0f1712",
        "section_title": "#eafff5",
        "small_note": "#8fd6b3",
        "card_bg": "#16241d",
        "card_border": "#234a37",
        "card_label": "#9fd9bd",
        "card_value": "#eafff5",
        "card_sub": "#7fae95",
        "warn_bg": "#3a2c12",
        "warn_border": "#6b4a10",
        "warn_text": "#f3d79a",
        "warn_sub": "#e0c187",
        "alert_bg": "#3a2c12",
        "alert_border": "#6b4a10",
        "alert_text": "#e0c187",
        "alert_strong": "#f3d79a",
        "chart_font": "#eafff5",
    },
}

T = PALETTES[THEME_MODE]

# -----------------------------
# CSS
# The sidebar and header banner keep their own deep-green brand colors in
# both modes (they're already dark, so they stay legible either way). The
# main canvas, section titles, KPI cards and alert card use __TOKEN__
# placeholders below, swapped for the palette picked above — plain string
# replacement (not an f-string), since CSS's { } would collide with
# Python's f-string formatting syntax.
# -----------------------------
CSS_TEMPLATE = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    /* Main canvas — switches between light mint and dark forest based on
       the person's Streamlit theme choice (see THEME_MODE above). */
    html, body,
    [data-testid="stApp"],
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewContainer"] > .main,
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"],
    .main,
    .block-container {
        background-color: __MAIN_BG__ !important;
    }

    [data-testid="stHeader"] {
        background-color: __MAIN_BG__ !important;
    }

    [data-testid="stToolbar"] {
        background-color: transparent !important;
    }

    [data-testid="stMain"] *,
    [data-testid="stAppViewContainer"] .main * {
        --background-color: __MAIN_BG__;
    }
    section[data-testid="stSidebar"] {
        background-color: #164a3b !important;
    }
    section[data-testid="stSidebar"] * {
        color: #eafff5 !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div,
    section[data-testid="stSidebar"] input {
        background-color: #1f5a48 !important;
        color: #eafff5 !important;
        border-color: #2c7460 !important;
    }
    /* File uploader — deliberately brute-force. "stFileUploader" is the one
       wrapper testid Streamlit has kept stable across versions; everything
       inside it (whatever internal testid/markup that version uses) gets
       painted deep green here so this survives Streamlit UI redesigns. */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
        background-color: #124f3b !important;
        border: 1.5px dashed #2fbf87 !important;
        border-radius: 10px !important;
        padding: 6px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] * {
        background-color: #124f3b !important;
        color: #eafff5 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] [role="button"] {
        background-color: #1c6b4f !important;
        color: #eafff5 !important;
        border: 1px solid #2fbf87 !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] svg,
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] svg path {
        fill: #eafff5 !important;
    }
    /* Hide the default "Limit 200MB per file • XLSX, XLS" helper line */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] small {
        display: none !important;
    }

    .block-container { padding-top: 2.6rem; padding-bottom: 3rem; max-width: 1250px; }

    /* ---------- Header banner (stays deep green + light text in both modes) ---------- */
    .app-header {
        background: linear-gradient(135deg, #064e3b 0%, #0f9d68 100%);
        border-radius: 20px;
        padding: 30px 36px 26px 36px;
        margin-bottom: 26px;
        box-shadow: 0 10px 28px rgba(6, 78, 59, 0.22);
        overflow: visible;
    }
    .app-header .eyebrow {
        color: #b9f5da;
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        margin: 0 0 6px 0;
        line-height: 1.4;
    }
    .app-header .dashboard-title {
        color: #eafff5;
        font-size: clamp(26px, 3.4vw, 38px);
        font-weight: 800;
        line-height: 1.35;
        margin: 0;
        padding: 4px 0;
        overflow: visible;
        white-space: normal;
    }
    .app-header .dashboard-subtitle {
        color: rgba(234, 255, 245, 0.85);
        font-size: 14.5px;
        margin-top: 10px;
        line-height: 1.6;
    }
    .app-header .dashboard-subtitle b { color: #eafff5; }

    /* ---------- Section titles ---------- */
    .section-title {
        font-size: 20px;
        font-weight: 750;
        margin-top: 26px;
        margin-bottom: 12px;
        color: __SECTION_TITLE__;
        border-left: 5px solid #2fbf87;
        padding-left: 12px;
    }
    .small-note { color: __SMALL_NOTE__; font-size: 12.5px; margin-top: -4px; margin-bottom: 10px; }

    /* ---------- Callout note (used by the Upload monthly files page) ---------- */
    .recipe-note {
        background: __CARD_BG__;
        border-left: 4px solid #2fbf87;
        border-radius: 8px;
        padding: 10px 14px;
        margin: 0 0 10px 0;
        color: __CARD_VALUE__;
        font-size: 13px;
        line-height: 1.5;
    }
    .recipe-note b { color: __CARD_VALUE__; }

    /* ---------- Quad KPI cards (same green gradient as the header card) ---------- */
    .quad-card {
        background: linear-gradient(135deg, #064e3b 0%, #0f9d68 100%);
        border: none;
        border-radius: 14px;
        padding: 16px 16px 14px 16px;
        box-shadow: 0 2px 8px rgba(16, 40, 30, 0.05);
        height: 100%;
        min-height: 152px;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
    }
    .quad-card.quad-warn {
        background: linear-gradient(135deg, #064e3b 0%, #0f9d68 100%);
        border: none;
    }
    .quad-icon { font-size: 20px; margin-bottom: 6px; }
    .quad-label {
        color: #b9f5da !important;
        font-size: 11.5px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin: 0 0 4px 0;
    }
    .quad-card.quad-warn .quad-label { color: #b9f5da !important; }
    .quad-value {
        color: #eafff5 !important;
        font-size: 25px;
        font-weight: 800;
        line-height: 1.2;
        margin: 0;
    }
    .quad-card.quad-warn .quad-value { color: #eafff5 !important; }
    .quad-sub { color: __CARD_SUB__ !important; font-size: 11.5px; margin-top: 5px; line-height: 1.4; }
    .quad-card.quad-warn .quad-sub { color: __WARN_SUB__ !important; }

    /* ---------- Alert card ---------- */
    .alert-card {
        background: __ALERT_BG__;
        border: 1px solid __ALERT_BORDER__;
        border-radius: 14px;
        padding: 14px 18px;
        margin-bottom: 10px;
    }
    .alert-card b { color: __ALERT_STRONG__ !important; }
    .alert-card, .alert-card * { color: __ALERT_TEXT__ !important; }
</style>
"""

CSS = CSS_TEMPLATE
for _key, _val in T.items():
    CSS = CSS.replace(f"__{_key.upper()}__", _val)

st.markdown(CSS, unsafe_allow_html=True)

# -----------------------------
# GLOBAL CHROME
# Streamlit paints "primary" buttons in its own accent red, which clashes
# with the Tea Lounge greens. Every primary button (sidebar nav, Save,
# Create menu) is repainted with the same deep-green gradient used by the
# header banner and KPI cards. Kept separate from CSS_TEMPLATE above so the
# palette tokens there stay untouched.
# -----------------------------
st.markdown(
    """
    <style>
        button[kind="primary"],
        button[data-testid="baseButton-primary"],
        [data-testid="stSidebar"] button[kind="primary"] {
            background: linear-gradient(135deg, #064e3b 0%, #0f9d68 100%) !important;
            border: 1px solid #0f9d68 !important;
            color: #eafff5 !important;
            box-shadow: none !important;
        }
        button[kind="primary"]:hover,
        button[data-testid="baseButton-primary"]:hover {
            background: linear-gradient(135deg, #053f30 0%, #0c8558 100%) !important;
            border-color: #12b67a !important;
            color: #ffffff !important;
        }
        button[kind="primary"]:focus,
        button[data-testid="baseButton-primary"]:focus {
            box-shadow: 0 0 0 2px rgba(15, 157, 104, 0.35) !important;
            color: #ffffff !important;
        }

        /* ============ SIDEBAR DESIGN ============ */

        /* Layout: the sidebar becomes a flex column so the footer card can
           sit at the very bottom whatever else the page adds above it. */
        section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
            display: flex; flex-direction: column;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
            flex: 1 1 auto; display: flex; flex-direction: column;
            padding: 1.5rem 1.1rem 1.3rem 1.1rem !important;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] > [data-testid="stVerticalBlock"] {
            flex: 1 1 auto;
        }
        section[data-testid="stSidebar"] hr { border-color: rgba(234, 255, 245, 0.14) !important; }

        /* Brand block: mint rounded avatar + name / subtitle */
        .tl-brand { display: flex; align-items: center; gap: 12px; padding: 2px 0 10px 0; }
        .tl-brand .tl-avatar {
            width: 44px; height: 44px; flex: 0 0 44px; border-radius: 12px;
            background: #e3f5ec; font-weight: 800; font-size: 16px; letter-spacing: 0.02em;
            display: flex; align-items: center; justify-content: center;
        }
        .tl-brand .tl-name { font-size: 17px; font-weight: 800; line-height: 1.15; }
        .tl-brand .tl-sub { font-size: 12px; line-height: 1.3; }
        section[data-testid="stSidebar"] .tl-brand .tl-avatar { color: #0f3d2e !important; }
        section[data-testid="stSidebar"] .tl-brand .tl-name { color: #ffffff !important; }
        section[data-testid="stSidebar"] .tl-brand .tl-sub { color: #b4d4c5 !important; }

        /* Any other sidebar button keeps a simple left-aligned row */
        [data-testid="stSidebar"] .stButton button {
            justify-content: flex-start !important;
            text-align: left !important;
            font-weight: 600 !important;
            border-radius: 10px !important;
            padding: 9px 14px !important;
        }

        /* Navigation rows: flat, icon + label; the active page gets a soft
           translucent pill with a hairline border (buttons carry the
           key prefix "tlnav_", which Streamlit exposes as a css class). */
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button {
            width: 100%;
            min-height: 46px;
            justify-content: flex-start !important;
            text-align: left !important;
            gap: 10px;
            padding: 11px 14px !important;
            border-radius: 12px !important;
            background: transparent !important;
            border: 1px solid transparent !important;
            box-shadow: none !important;
            transition: background 0.15s ease, border-color 0.15s ease;
        }
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button p {
            font-size: 15px !important;
            font-weight: 600 !important;
            color: #e6f4ec !important;
        }
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button:hover {
            background: rgba(234, 255, 245, 0.07) !important;
            border-color: transparent !important;
        }
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button[kind="primary"],
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button[data-testid="stBaseButton-primary"],
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button[kind="primary"]:hover,
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button[data-testid="stBaseButton-primary"]:hover {
            background: rgba(234, 255, 245, 0.11) !important;
            border: 1px solid rgba(234, 255, 245, 0.28) !important;
        }
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button[kind="primary"] p,
        section[data-testid="stSidebar"] [class*="st-key-tlnav_"] button[data-testid="stBaseButton-primary"] p {
            color: #ffffff !important;
            font-weight: 700 !important;
        }

        /* Collapse button: round white chip with a dark-green chevron */
        section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] button {
            background: #ffffff !important;
            border-radius: 50% !important;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.25) !important;
        }
        section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] button * {
            color: #164a3b !important;
        }

        /* Footer card, pinned to the bottom of the sidebar */
        section[data-testid="stSidebar"] :is([data-testid="stElementContainer"], .element-container):has(.tl-note) {
            order: 99;
            margin-top: auto;
        }
        .tl-note {
            background: rgba(234, 255, 245, 0.08);
            border: 1px solid rgba(234, 255, 245, 0.10);
            border-radius: 14px;
            padding: 14px 16px;
        }
        section[data-testid="stSidebar"] .tl-note b { color: #ffffff !important; font-size: 13.5px; }
        section[data-testid="stSidebar"] .tl-note span {
            color: #a9cbbb !important; font-size: 12px; line-height: 1.5; display: block; margin-top: 4px;
        }
    </style>
    """.replace("__CARD_BG__", T["card_bg"])
       .replace("__CARD_VALUE__", T["card_value"]),
    unsafe_allow_html=True,
)


# -----------------------------
# LIGHT UI — every page
# Streamlit's own widgets (dropdowns, uploaders, buttons, metrics, header
# icons) are repainted light so the Dashboard and Upload pages match the
# Gross Profit Tracker. Only the main area is touched — the sidebar keeps
# its deep-green design. (Streamlit is also launched with
# --theme.base=light above, which lights up tables and other widgets.)
# -----------------------------
_LT = PALETTES["light"]
LIGHT_UI_CSS = """
<style>
    html { color-scheme: light; }
    [data-testid="stHeader"] *, [data-testid="stToolbar"] * { color: __TXT__ !important; }

    [data-testid="stMain"] label,
    [data-testid="stMain"] [data-testid="stWidgetLabel"] * { color: __TXT__ !important; font-weight: 600; }
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] h3,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] h4 { color: __TXT__ !important; }
    [data-testid="stMain"] [data-testid="stMetric"] * { color: __TXT__ !important; }

    [data-testid="stMain"] div[data-baseweb="select"] > div {
        background-color: #ffffff !important; border: 1px solid __LINE__ !important; border-radius: 10px !important;
    }
    [data-testid="stMain"] div[data-baseweb="select"] * { color: __TXT__ !important; }
    [data-testid="stMain"] div[data-baseweb="select"] svg { fill: __TXT__ !important; }
    div[data-baseweb="popover"] ul, div[data-baseweb="popover"] [role="listbox"] { background-color: #ffffff !important; }
    div[data-baseweb="popover"] li, div[data-baseweb="popover"] [role="option"] {
        background-color: #ffffff !important; color: __TXT__ !important;
    }
    div[data-baseweb="popover"] li:hover, div[data-baseweb="popover"] [role="option"]:hover {
        background-color: __SOFT__ !important;
    }

    [data-testid="stMain"] [data-testid="stFileUploader"] section,
    [data-testid="stMain"] [data-testid="stFileUploaderDropzone"] {
        background-color: #ffffff !important; border: 1.5px dashed __GREEN__ !important; border-radius: 12px !important;
    }
    [data-testid="stMain"] [data-testid="stFileUploader"] * { color: __TXT__ !important; }
    [data-testid="stMain"] [data-testid="stFileUploader"] svg { fill: __TXT__ !important; }

    [data-testid="stMain"] button[kind="secondary"],
    [data-testid="stMain"] button[data-testid="stBaseButton-secondary"],
    [data-testid="stMain"] [data-testid="stDownloadButton"] button,
    [data-testid="stMain"] [data-testid="stFileUploader"] button {
        background-color: #ffffff !important; color: __TXT__ !important;
        border: 1px solid __LINE__ !important; border-radius: 10px !important;
    }
    [data-testid="stMain"] button[kind="secondary"] *,
    [data-testid="stMain"] [data-testid="stDownloadButton"] button * { color: __TXT__ !important; }
</style>
"""
for _token, _value in {
    "__TXT__": _LT["section_title"], "__SOFT__": _LT["card_bg"],
    "__LINE__": _LT["card_border"], "__GREEN__": "#2fbf87",
}.items():
    LIGHT_UI_CSS = LIGHT_UI_CSS.replace(_token, _value)
st.markdown(LIGHT_UI_CSS, unsafe_allow_html=True)


# ============================================================
# GENERIC HELPERS
# ============================================================

def clean_number(x):
    """Convert numbers such as '5,335.86' or '—' to numeric."""
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)

    s = str(x).strip().replace(",", "")
    if s in {"", "-", "—", "–", "nan", "None"}:
        return np.nan

    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else np.nan


def normalize_text(x):
    if pd.isna(x):
        return ""
    return re.sub(r"\s+", " ", str(x).strip())


def html_escape(value):
    """str() + HTML-escape, so any user-typed text (an ingredient name, a
    file name) can never break out of the markup it's inserted into."""
    return html.escape(str(value))


def fmt_num(x, decimals=1):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "—"
    x = float(x)
    if x.is_integer():
        return f"{x:,.0f}"
    return f"{x:,.{decimals}f}"


# ============================================================
# MONTH / YEAR PARSING
# Turns free-form text (sheet names, "Month" column values, etc.)
# into a (year, month) pair. Handles things like "January26",
# "Jan 2026", "March-2026", "September 2022".
# ============================================================

MONTH_NAME_TO_NUM = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sept": 9, "sep": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}


def parse_month_year(text):
    """Best-effort extraction of (year, month) from a free-form string.
    Returns None if no month name + year could be found."""
    if text is None:
        return None
    s = re.sub(r"[_\-]", " ", str(text).strip().lower())
    if not s:
        return None

    month_num, month_match = None, None
    for name in sorted(MONTH_NAME_TO_NUM, key=len, reverse=True):
        m = re.search(r"\b" + re.escape(name) + r"[a-z]*", s)
        if m:
            month_num = MONTH_NAME_TO_NUM[name]
            month_match = m
            break
    if month_num is None:
        return None

    rest = s[month_match.end():] + " " + s[:month_match.start()]
    year_match = re.search(r"(20\d{2}|\d{2})\b", rest)
    if not year_match:
        return None

    year = int(year_match.group(1))
    if year < 100:
        year += 2000
    return (year, month_num)


def month_key(year, month):
    return f"{year:04d}-{month:02d}"


def month_label(year, month):
    return f"{calendar.month_name[month]} {year}"


# ============================================================
# EXCEL READERS
# ============================================================

def find_header_row(raw):
    """Finds the header row containing 'Ingredient / Item' etc.
    Returns None (not 0) if no such row is found, so callers can tell
    'this isn't an inventory sheet' apart from 'header is on row 0'."""
    for i in range(min(15, len(raw))):
        values = [normalize_text(v).lower() for v in raw.iloc[i].tolist()]
        if "ingredient / item" in values and (
            "total units sold" in values or "projected use" in values
        ):
            return i
    return None


def read_inventory_sheet(raw, sheet_name):
    """
    Reads an 'Ingredient Summary & Inventory' sheet: Ingredient / Item,
    Total Units Sold, Total Required Qty, Projected Use, Pre. Month Stock,
    Received 1, Received 2, Total Stock, Month End Stock, Total Use.
    Section-header rows (e.g. 'Miscellaneous') are detected automatically
    and used to tag each ingredient with a Category. `raw` is the sheet
    already loaded with header=None (read once and shared by the caller,
    since each sheet also needs to be tested against the sales-report
    header pattern).
    """
    title_cell = normalize_text(raw.iloc[0, 0]) if len(raw) else ""

    header_row = find_header_row(raw)
    if header_row is None:
        raise ValueError("Could not identify the 'Ingredient / Item' column in this sheet.")
    df = raw.iloc[header_row + 1:].copy()
    headers = raw.iloc[header_row].tolist()

    clean_headers = []
    for i, h in enumerate(headers):
        h = normalize_text(h)
        clean_headers.append(h if h else f"Column_{i+1}")
    df.columns = clean_headers
    df = df.reset_index(drop=True)

    aliases = {
        "Ingredient": ["ingredient / item", "ingredient/item", "ingredient"],
        "Total_Units_Sold": ["total units sold", "units sold"],
        "Total_Required_Qty": ["total required qty", "required qty"],
        "Projected_Use": ["projected use"],
        "Pre_Month_Stock": ["pre. month stock", "pre month stock", "previous month stock"],
        "Received_1": ["received 1"],
        "Received_2": ["received 2"],
        "Total_Stock": ["total stock"],
        "Month_End_Stock": ["month end stock"],
        "Total_Use": ["total use"],
    }

    lower_map = {normalize_text(c).lower(): c for c in df.columns}
    rename_map = {}
    for standard, possible_names in aliases.items():
        for name in possible_names:
            if name in lower_map:
                rename_map[lower_map[name]] = standard
                break
    df = df.rename(columns=rename_map)

    if "Ingredient" not in df.columns:
        raise ValueError(
            "Could not identify the 'Ingredient / Item' column in this sheet."
        )

    df["Ingredient"] = df["Ingredient"].apply(normalize_text)

    # Detect section-header rows (e.g. "Miscellaneous"): rows where the
    # Ingredient cell has text but every other column is blank.
    other_cols = [c for c in df.columns if c != "Ingredient"]

    def is_category_row(row):
        if row["Ingredient"] == "":
            return False
        vals = [row[c] for c in other_cols]
        return all(pd.isna(v) or normalize_text(v) == "" for v in vals)

    categories = []
    current_category = "Beverage Ingredients"
    keep_mask = []
    for _, row in df.iterrows():
        if row["Ingredient"] == "":
            categories.append(current_category)
            keep_mask.append(False)
            continue
        if is_category_row(row):
            current_category = row["Ingredient"]
            categories.append(current_category)
            keep_mask.append(False)
            continue
        categories.append(current_category)
        keep_mask.append(True)

    df["Category"] = categories
    df = df[pd.Series(keep_mask, index=df.index)].copy()

    numeric_cols = [
        "Total_Units_Sold", "Total_Required_Qty", "Projected_Use",
        "Pre_Month_Stock", "Received_1", "Received_2",
        "Total_Stock", "Month_End_Stock", "Total_Use",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)
        else:
            df[col] = np.nan

    df = df[df["Ingredient"].notna() & (df["Ingredient"] != "")].reset_index(drop=True)
    df["Received_Total"] = df["Received_1"].fillna(0) + df["Received_2"].fillna(0)

    return df, title_cell


SALES_ITEM_ALIASES = [
    "item", "menu item", "item name", "product", "product name",
    "item description", "menu item name",
]
SALES_QTY_ALIASES = [
    "total quantity", "quantity", "qty", "units sold",
    "qty sold", "sold qty", "sales qty", "quantity sold",
]


def find_sales_header_row(raw):
    """Scans further (25 rows) and matches a wider set of header aliases
    than the inventory sheet, since item-sales report headers vary more
    from report to report. Returns None (not 0) if nothing is found, so
    callers can tell 'no header found' apart from 'header is on row 0'."""
    for i in range(min(25, len(raw))):
        values = [normalize_text(v).lower() for v in raw.iloc[i].tolist()]
        has_qty = any(v in SALES_QTY_ALIASES for v in values)
        has_item = any(v in SALES_ITEM_ALIASES for v in values)
        if has_qty and has_item:
            return i
    return None


def read_item_sales_sheet(raw, sheet_name):
    """
    Reads a 'Monthly Item Sales Report' sheet, e.g.:
        SL | Item Name | Selling Price | Quantity | Total Amount | Ranking
    ...followed by a summary row such as "Total :" with the grand total in
    the Quantity column. Only the Item name and Quantity columns are
    required; a Month/Period column is used (if present) to split one
    sheet's totals across several months. `raw` is the sheet already
    loaded with header=None (read once and shared with the inventory
    reader, since each sheet has to be tested against both patterns).
    Returns (df, None) on success, or (None, reason) on failure so the
    caller can surface *why* the sheet wasn't recognized. The returned
    df carries a "_Is_Total_Row" column marking the summary row.
    """
    header_row = find_sales_header_row(raw)
    if header_row is None:
        return None, (
            "no row with both an item-name header "
            f"({', '.join(SALES_ITEM_ALIASES[:3])}, ...) and a quantity header "
            f"({', '.join(SALES_QTY_ALIASES[:3])}, ...) was found in the first 25 rows"
        )

    body = raw.iloc[header_row + 1:].copy()
    headers = raw.iloc[header_row].tolist()

    clean_headers = []
    for i, h in enumerate(headers):
        h = normalize_text(h)
        clean_headers.append(h if h else f"Column_{i+1}")
    body.columns = clean_headers
    body = body.reset_index(drop=True)

    # Find the "Total" summary row by scanning every original column, since
    # the label (e.g. "Total :", merged across a few cells) can land in any
    # column, not necessarily the Item Name column.
    def row_has_total_marker(row):
        for v in row.tolist():
            text = normalize_text(v).lower().rstrip(":").strip()
            if text == "total" or text.startswith("total "):
                return True
        return False

    is_total_row = body.apply(row_has_total_marker, axis=1)

    aliases = {
        "Item": SALES_ITEM_ALIASES,
        "Total_Quantity": SALES_QTY_ALIASES,
        "Month": ["month", "period"],
    }
    lower_map = {normalize_text(c).lower(): c for c in body.columns}
    rename_map = {}
    for standard, names in aliases.items():
        for name in names:
            if name in lower_map:
                rename_map[lower_map[name]] = standard
                break
    body = body.rename(columns=rename_map)

    if "Item" not in body.columns or "Total_Quantity" not in body.columns:
        found_cols = ", ".join(body.columns.astype(str))
        return None, f"found a header row, but couldn't match Item/Quantity columns (columns seen: {found_cols})"

    body["Item"] = body["Item"].apply(normalize_text)
    body["Total_Quantity"] = body["Total_Quantity"].apply(clean_number)
    body["_Is_Total_Row"] = is_total_row.values

    # Keep genuine item rows (non-blank Item, not the Total row) plus the
    # Total row itself — its own Item cell is often blank because "Total :"
    # sits in a different, merged column.
    keep = body["_Is_Total_Row"] | (body["Item"].notna() & (body["Item"] != ""))
    body = body[keep].reset_index(drop=True)

    return body, None


def summarize_item_sales(sales_df):
    """Turns a raw item-sales table into (cups_sold, menu_item_count).

    The report has one row per menu item (Item Name column) followed by a
    summary row such as "Total :" holding the grand total Quantity. Cups
    sold = that Total row's Quantity, read directly rather than re-summed.
    Menu items = count of distinct item rows, excluding the Total row.
    Falls back to summing item rows if no Total row was detected.
    """
    if "_Is_Total_Row" in sales_df.columns:
        is_total_row = sales_df["_Is_Total_Row"].fillna(False)
    else:
        is_total_row = sales_df["Item"].str.strip().str.lower().str.rstrip(":").str.strip() == "total"

    total_row = sales_df[is_total_row]
    item_rows = sales_df[~is_total_row]
    item_rows = item_rows[item_rows["Item"].notna() & (item_rows["Item"] != "")]

    menu_item_count = item_rows["Item"].nunique()
    if not total_row.empty and pd.notna(total_row["Total_Quantity"].iloc[0]):
        cups_sold = total_row["Total_Quantity"].iloc[0]
    else:
        cups_sold = item_rows["Total_Quantity"].sum()

    return cups_sold, menu_item_count


# ============================================================
# PERSISTENT MONTHLY STORAGE
# Saved next to this script so the data survives restarts and keeps
# every month you've ever uploaded, not just the current file.
# ============================================================

DATA_DIR = os.environ.get("TEA_LOUNGE_DATA_DIR") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "tea_lounge_store"
)
INVENTORY_DIR = os.path.join(DATA_DIR, "inventory")
META_PATH = os.path.join(DATA_DIR, "meta.json")


def ensure_store():
    os.makedirs(INVENTORY_DIR, exist_ok=True)
    if not os.path.exists(META_PATH):
        with open(META_PATH, "w") as f:
            json.dump({}, f)


def inventory_path(year, month):
    return os.path.join(INVENTORY_DIR, f"{month_key(year, month)}.csv")


def load_meta():
    ensure_store()
    try:
        with open(META_PATH) as f:
            return json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        return {}


def save_meta(meta):
    ensure_store()
    with open(META_PATH, "w") as f:
        json.dump(meta, f, indent=2)


def save_month_inventory(year, month, df, source_name):
    ensure_store()
    df.to_csv(inventory_path(year, month), index=False)
    meta = load_meta()
    meta.setdefault(month_key(year, month), {})
    meta[month_key(year, month)]["source_file"] = source_name
    save_meta(meta)


def load_month_inventory(year, month):
    path = inventory_path(year, month)
    if not os.path.exists(path):
        return None
    return pd.read_csv(path)


def list_stored_months():
    ensure_store()
    months = []
    for fname in os.listdir(INVENTORY_DIR):
        if fname.endswith(".csv"):
            key = fname[:-4]
            try:
                y, m = key.split("-")
                months.append((int(y), int(m)))
            except ValueError:
                continue
    return sorted(months)


def set_month_sales(year, month, total_qty, menu_items):
    meta = load_meta()
    key = month_key(year, month)
    meta.setdefault(key, {})
    meta[key]["total_quantity"] = None if total_qty is None or pd.isna(total_qty) else float(total_qty)
    meta[key]["menu_items"] = None if menu_items is None else int(menu_items)
    save_meta(meta)


def get_month_meta(year, month):
    return load_meta().get(month_key(year, month), {})


# ============================================================
# GROSS PROFIT TRACKER — DATA ENGINE
# One small JSON file per month lives in ./tea_lounge_store/gp_tracker/.
# It holds only what was *typed* into the workbook (ingredient costs,
# selling price, cups sold). Every figure the tracker shows — cost per
# cup, gross profit, margin, totals — is recalculated from those inputs
# when the page is drawn, so nothing depends on Excel having cached its
# formula results, and a saved month can never drift out of step with
# the maths. It is a separate folder from the inventory history, so the
# dashboard's months are never affected by it.
# ============================================================

GP_DIR = os.path.join(DATA_DIR, "gp_tracker")

# Header text -> standard field. Headings are matched by name (after
# trimming stray spaces and lower-casing), never by column position.
GP_COL_ALIASES = {
    "Ingredient": ["ingredient / item", "ingredient/item", "ingredient", "item"],
    "Qty_Raw": ["qty raw (as given)", "qty raw"],
    "Qty_Value": ["qty value"],
    "Qty_Unit": ["qty unit", "unit"],
    "Cost": ["cost", "cost (tk)", "ingredient cost"],
    "Cost_Per_Cup": ["cost per cup (tk)", "cost per cup"],
    "Selling_Price": ["selling price per cup (tk)", "selling price per cup", "selling price"],
    "Qty_Sold": ["qty sold", "quantity sold"],
}
GP_MAP_NAME_ALIASES = [
    "item name", "menu/item name", "menu / item name", "menu item name", "menu name", "menu",
]
GP_MAP_ID_ALIASES = ["menu id", "menu code", "menu_id"]

# A typed "Cost Per Cup" only counts as an override when it differs from
# the sum of the ingredient lines by more than this many Tk.
GP_COST_TOLERANCE = 0.005

# The colour code the dashboard's own charts already use.
GP_GREEN_PALE = "#a7e8c8"
GP_GREEN_LIGHT = "#38d996"
GP_GREEN = "#2fbf87"
GP_GREEN_DARK = "#0b6e4f"
GP_LOSS = "#e5484d"


def _gp_num(x):
    """Number or None (never NaN), so results serialise cleanly to JSON."""
    value = clean_number(x)
    return None if pd.isna(value) else float(value)


def _gp_column_map(values, alias_map):
    """{standard field: column index} for one header row."""
    found = {}
    for idx, val in enumerate(values):
        key = normalize_text(val).lower()
        if not key:
            continue
        for standard, names in alias_map.items():
            if key in names and standard not in found:
                found[standard] = idx
                break
    return found


def _gp_find_header_row(raw, alias_map, required, max_scan=12):
    for i in range(min(max_scan, len(raw))):
        found = _gp_column_map(raw.iloc[i].tolist(), alias_map)
        if all(name in found for name in required):
            return i
    return None


def read_gp_mapping(raw):
    """The 'Menu Mapping' sheet -> {MENU ID (upper-case): menu name}."""
    alias_map = {"Name": GP_MAP_NAME_ALIASES, "ID": GP_MAP_ID_ALIASES}
    header_row = _gp_find_header_row(raw, alias_map, ("Name", "ID"))
    if header_row is None:
        return None
    cols = _gp_column_map(raw.iloc[header_row].tolist(), alias_map)
    mapping = {}
    for _, row in raw.iloc[header_row + 1:].iterrows():
        menu_id = normalize_text(row.iloc[cols["ID"]])
        name = normalize_text(row.iloc[cols["Name"]])
        if menu_id and name:
            mapping[menu_id.upper()] = name
    return mapping or None


def read_gp_sheet(raw):
    """One menu's sheet. Selling price, cups sold and an optional typed
    cost-per-cup sit on the first data row; every row with an ingredient
    name is one ingredient line. Returns None if it isn't a GP sheet."""
    header_row = _gp_find_header_row(
        raw, GP_COL_ALIASES, ("Ingredient", "Selling_Price", "Qty_Sold")
    )
    if header_row is None:
        return None
    cols = _gp_column_map(raw.iloc[header_row].tolist(), GP_COL_ALIASES)

    def cell(row, field):
        return row.iloc[cols[field]] if field in cols else None

    selling_price = qty_sold = typed_cost = None
    ingredients = []
    for _, row in raw.iloc[header_row + 1:].iterrows():
        if selling_price is None:
            selling_price = _gp_num(cell(row, "Selling_Price"))
        if qty_sold is None:
            qty_sold = _gp_num(cell(row, "Qty_Sold"))
        if typed_cost is None:
            typed_cost = _gp_num(cell(row, "Cost_Per_Cup"))
        name = normalize_text(cell(row, "Ingredient"))
        if not name:
            continue
        ingredients.append({
            "name": name,
            "qty_raw": normalize_text(cell(row, "Qty_Raw")),
            "qty_value": _gp_num(cell(row, "Qty_Value")),
            "qty_unit": normalize_text(cell(row, "Qty_Unit")),
            "cost": _gp_num(cell(row, "Cost")),
        })
    return {
        "selling_price": selling_price,
        "qty_sold": qty_sold,
        "typed_cost_per_cup": typed_cost,
        "ingredients": ingredients,
    }


def read_gp_workbook(uploaded_file):
    """Reads a GP tracker workbook: an optional 'Menu Mapping' sheet plus
    one sheet per menu. Returns (menus, warnings, diagnostics)."""
    if hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file)

    names, menus, warnings, diagnostics = {}, [], [], []
    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if raw.empty:
            diagnostics.append({"sheet": sheet_name, "kind": "Skipped", "detail": "the sheet is empty"})
            continue

        mapping = read_gp_mapping(raw)
        if mapping:
            names.update(mapping)
            diagnostics.append({
                "sheet": sheet_name, "kind": "Menu mapping",
                "detail": f"{len(mapping)} menu names mapped to menu IDs",
            })
            continue

        sheet = read_gp_sheet(raw)
        if sheet is None:
            diagnostics.append({
                "sheet": sheet_name, "kind": "Skipped",
                "detail": "no header row with Ingredient / Item, Selling Price Per Cup and QTY Sold",
            })
            continue

        sheet["id"] = normalize_text(sheet_name)
        menus.append(sheet)
        diagnostics.append({
            "sheet": sheet_name, "kind": "Menu",
            "detail": f"{len(sheet['ingredients'])} ingredient lines",
        })

    for menu in menus:
        menu["name"] = names.get(menu["id"].upper(), menu["id"])

    on_file = {menu["id"].upper() for menu in menus}
    for menu_id, name in names.items():
        if menu_id not in on_file:
            warnings.append(f"'{name}' ({menu_id}) is in the Menu Mapping but has no sheet of its own.")
    return menus, warnings, diagnostics


def gp_detect_period(filename):
    """(year, month) guessed from a file name, or None if no month name is
    in it. 'GP_Tracker_September.xlsx' carries no year, so the year is
    assumed to be the current one — or last year, when that month hasn't
    happened yet this year (a December file uploaded in January)."""
    both = parse_month_year(filename)
    if both:
        return both
    tokens = re.findall(r"[a-z]+", str(filename).lower())
    month = next((MONTH_NAME_TO_NUM[t] for t in tokens if t in MONTH_NAME_TO_NUM), None)
    if month is None:
        return None
    today = date.today()
    return (today.year - 1 if month > today.month else today.year, month)


# ---------- Storage: one JSON file per month ----------

def gp_path(year, month):
    return os.path.join(GP_DIR, f"{month_key(year, month)}.json")


def save_gp_month(year, month, menus, source_name):
    """Writes (or replaces) one month. Written to a temporary file first
    and swapped in, so an interrupted save can't leave a half-written
    month behind."""
    os.makedirs(GP_DIR, exist_ok=True)
    payload = {
        "year": int(year),
        "month": int(month),
        "source_file": source_name,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "menus": menus,
    }
    path = gp_path(year, month)
    tmp_path = path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    os.replace(tmp_path, path)


def load_gp_month(year, month):
    path = gp_path(year, month)
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            payload = json.load(f)
    except (json.JSONDecodeError, OSError):
        return None
    if not isinstance(payload, dict) or not isinstance(payload.get("menus"), list):
        return None
    return payload


def list_gp_months():
    if not os.path.isdir(GP_DIR):
        return []
    months = []
    for fname in os.listdir(GP_DIR):
        match = re.fullmatch(r"(\d{4})-(\d{2})\.json", fname)
        if match:
            y, m = int(match.group(1)), int(match.group(2))
            if 1 <= m <= 12:
                months.append((y, m))
    return sorted(months)


def delete_gp_month(year, month):
    path = gp_path(year, month)
    if os.path.exists(path):
        os.remove(path)


# ---------- Calculations (all from the typed inputs) ----------

def gp_menu_metrics(menu):
    """Cost, price, profit and volume figures for one menu.

    Cost per cup is the sum of the ingredient costs — the same thing the
    workbook's own SUM does — unless a different total was typed over it
    (a few menus have unpriced ingredients and a hand-typed total).
    A menu with no selling price has no revenue or profit, so it is left
    out of every money total rather than being counted as free."""
    ingredients = menu.get("ingredients") or []
    listed = [i["cost"] for i in ingredients if i.get("cost") is not None]
    listed_sum = float(sum(listed))
    typed = menu.get("typed_cost_per_cup")
    if typed is not None and abs(typed - listed_sum) > GP_COST_TOLERANCE:
        cost_per_cup, cost_source = float(typed), "typed"
    else:
        cost_per_cup, cost_source = listed_sum, "sum"

    price = menu.get("selling_price")
    qty = menu.get("qty_sold")
    out = {
        "id": menu.get("id", ""),
        "name": menu.get("name") or menu.get("id", ""),
        "qty_sold": float(qty) if qty is not None else 0.0,
        "qty_blank": qty is None,
        "selling_price": price,
        "has_price": price is not None,
        "cost_per_cup": cost_per_cup,
        "cost_source": cost_source,
        "listed_cost": listed_sum,
        "missing_costs": [i["name"] for i in ingredients if i.get("cost") is None],
        "ingredient_count": len(ingredients),
        "gp_per_cup": None, "margin": None,
        "total_cost": None, "total_revenue": None, "gross_profit": None,
    }
    if price is not None:
        out["gp_per_cup"] = price - cost_per_cup
        out["margin"] = (price - cost_per_cup) / price if price else None
        out["total_cost"] = out["qty_sold"] * cost_per_cup
        out["total_revenue"] = out["qty_sold"] * price
        out["gross_profit"] = out["total_revenue"] - out["total_cost"]
    return out


GP_NUMERIC_COLUMNS = [
    "qty_sold", "selling_price", "cost_per_cup", "listed_cost", "gp_per_cup",
    "margin", "total_cost", "total_revenue", "gross_profit",
]


def gp_frame(menus):
    """One row per menu, with each menu's share of the month's totals."""
    df = pd.DataFrame([gp_menu_metrics(m) for m in menus])
    if df.empty:
        return df
    for col in GP_NUMERIC_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["missing_count"] = df["missing_costs"].apply(len)
    total_revenue = df["total_revenue"].sum()
    total_profit = df["gross_profit"].sum()
    df["revenue_share"] = df["total_revenue"] / total_revenue if total_revenue else np.nan
    df["profit_share"] = df["gross_profit"] / total_profit if total_profit else np.nan
    return df


def gp_totals(df):
    """Month totals. Averages are per cup actually sold (weighted by
    volume), so a slow seller can't drag the average around."""
    empty = {
        "revenue": 0.0, "cost": 0.0, "profit": 0.0, "margin": None, "cups": 0.0,
        "avg_price": None, "avg_cost": None, "avg_gp": None,
        "menus": 0, "menus_sold": 0,
    }
    if df is None or df.empty:
        return empty
    priced = df[df["has_price"]]
    revenue = float(priced["total_revenue"].sum())
    cost = float(priced["total_cost"].sum())
    cups = float(priced["qty_sold"].sum())
    profit = revenue - cost
    return {
        "revenue": revenue,
        "cost": cost,
        "profit": profit,
        "margin": profit / revenue if revenue else None,
        "cups": cups,
        "avg_price": revenue / cups if cups else None,
        "avg_cost": cost / cups if cups else None,
        "avg_gp": profit / cups if cups else None,
        "menus": int(len(df)),
        "menus_sold": int((df["qty_sold"] > 0).sum()),
    }


def gp_ingredient_drivers(menus):
    """Total Tk spent on each ingredient across every cup sold this month
    (its cost in each menu x that menu's cups). Only ingredient lines
    that have a cost typed in are counted."""
    totals = {}
    for menu in menus:
        qty = menu.get("qty_sold") or 0.0
        if menu.get("selling_price") is None or qty <= 0:
            continue
        for ing in menu.get("ingredients") or []:
            if ing.get("cost") is None:
                continue
            key = ing["name"].strip().lower()
            row = totals.setdefault(
                key, {"ingredient": ing["name"], "total_cost": 0.0, "menus": 0, "cups": 0.0}
            )
            row["total_cost"] += ing["cost"] * qty
            row["menus"] += 1
            row["cups"] += qty
    if not totals:
        return pd.DataFrame(columns=["ingredient", "total_cost", "menus", "cups"])
    return pd.DataFrame(list(totals.values())).sort_values("total_cost", ascending=False)


def gp_waterfall_steps(menu, metrics):
    """Selling price -> each ingredient cost -> gross profit per cup.
    A menu whose cost was typed over its ingredient lines gets one extra
    'other cost' step so the walk still lands exactly on GP per cup."""
    steps = [("Selling price", metrics["selling_price"], "absolute")]
    for ing in menu.get("ingredients") or []:
        if ing.get("cost") is not None:
            steps.append((ing["name"], -ing["cost"], "relative"))
    gap = metrics["cost_per_cup"] - metrics["listed_cost"]
    if abs(gap) > GP_COST_TOLERANCE:
        steps.append(("Other cost (typed total)", -gap, "relative"))
    steps.append(("GP per cup", metrics["gp_per_cup"], "total"))
    return steps


def gp_history():
    """One row per stored month (oldest first) for the trend charts."""
    rows = []
    for year, month in list_gp_months():
        payload = load_gp_month(year, month)
        if not payload:
            continue
        totals = gp_totals(gp_frame(payload["menus"]))
        totals.update({
            "key": month_key(year, month),
            "label": month_label(year, month),
            "short": f"{calendar.month_abbr[month]} {year}",
        })
        rows.append(totals)
    return pd.DataFrame(rows)



def process_uploaded_workbook(uploaded_file):
    """Reads every sheet in the uploaded workbook, saves any recognizable
    monthly inventory sheet to disk, and matches any item-sales sheet(s)
    to the months they belong to. Also returns a per-sheet diagnostics log
    so the user can see exactly how each sheet in the workbook was (or
    wasn't) classified, instead of a silent '—' on the dashboard."""
    xls = pd.ExcelFile(uploaded_file)
    added_months = []
    sales_sheets = []
    warnings = []
    diagnostics = []  # list of dicts: sheet, kind, detail

    for sheet_name in xls.sheet_names:
        # Read once, reused for whichever pattern the header actually
        # matches — sheet NAME alone can't tell inventory and item-sales
        # workbooks apart, since both commonly use month-named tabs like
        # "January26" with no "sales" in the name at all.
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        inv_header_row = find_header_row(raw)
        sales_header_row = find_sales_header_row(raw)
        name_hints_sales = bool(re.search(r"sales", sheet_name, re.I))

        if inv_header_row is not None and sales_header_row is None:
            kind = "inventory"
        elif sales_header_row is not None and inv_header_row is None:
            kind = "sales"
        elif inv_header_row is not None and sales_header_row is not None:
            kind = "sales" if name_hints_sales else "inventory"
        else:
            kind = None

        if kind is None:
            diagnostics.append({
                "sheet": sheet_name, "kind": "⏭️ Skipped",
                "detail": (
                    "no recognizable inventory header (Ingredient / Item) or "
                    "item-sales header (Item Name + Quantity) found in the "
                    "first 25 rows"
                ),
            })
            continue

        if kind == "sales":
            sales_df, fail_reason = read_item_sales_sheet(raw, sheet_name)
            if sales_df is not None:
                sales_sheets.append((parse_month_year(sheet_name) or parse_month_year(uploaded_file.name), sales_df))
                diagnostics.append({
                    "sheet": sheet_name, "kind": "✅ Item sales report",
                    "detail": f"{len(sales_df)} rows read (Total row included: "
                              f"{sales_df['_Is_Total_Row'].any() if '_Is_Total_Row' in sales_df else 'n/a'})",
                })
            else:
                warnings.append(f"Could not read '{sheet_name}' as an item sales report ({fail_reason}).")
                diagnostics.append({
                    "sheet": sheet_name, "kind": "❌ Sales sheet — not read", "detail": fail_reason,
                })
            continue

        # kind == "inventory"
        parsed = parse_month_year(sheet_name) or parse_month_year(uploaded_file.name)
        if parsed is None:
            diagnostics.append({
                "sheet": sheet_name, "kind": "⏭️ Skipped",
                "detail": "looked like an inventory sheet, but its name has no recognizable month + year",
            })
            continue

        year, month = parsed
        try:
            inv_df, _title = read_inventory_sheet(raw, sheet_name)
        except Exception as e:
            warnings.append(f"Could not read sheet '{sheet_name}': {e}")
            diagnostics.append({
                "sheet": sheet_name, "kind": "❌ Inventory sheet — not read", "detail": str(e),
            })
            continue

        save_month_inventory(year, month, inv_df, uploaded_file.name)
        added_months.append((year, month))
        diagnostics.append({
            "sheet": sheet_name, "kind": f"✅ Inventory — {month_label(year, month)}",
            "detail": f"{len(inv_df)} ingredient rows saved",
        })

    for parsed, sales_df in sales_sheets:
        if "Month" in sales_df.columns and sales_df["Month"].notna().any():
            for month_text, group in sales_df.groupby("Month"):
                my = parse_month_year(str(month_text))
                if my is None:
                    continue
                y, m = my
                cups_sold, menu_items = summarize_item_sales(group)
                set_month_sales(y, m, cups_sold, menu_items)
                if (y, m) not in added_months:
                    added_months.append((y, m))
        elif parsed is not None:
            y, m = parsed
            cups_sold, menu_items = summarize_item_sales(sales_df)
            set_month_sales(y, m, cups_sold, menu_items)
        elif len(added_months) == 1:
            y, m = added_months[0]
            cups_sold, menu_items = summarize_item_sales(sales_df)
            set_month_sales(y, m, cups_sold, menu_items)
        else:
            stored = list_stored_months()
            if stored:
                y, m = max(stored)
                cups_sold, menu_items = summarize_item_sales(sales_df)
                set_month_sales(y, m, cups_sold, menu_items)
                if (y, m) not in added_months:
                    added_months.append((y, m))
            else:
                warnings.append(
                    "Found an item sales report but couldn't tell which month it "
                    "belongs to (add a 'Month' column, or name the sheet with a "
                    "month, e.g. 'January26 Item Sales')."
                )

    return sorted(set(added_months)), warnings, diagnostics


# ============================================================
# UI HELPERS
# ============================================================

def handle_workbook_upload(uploaded_file, slot):
    """Reads one uploaded monthly inventory/sales workbook.

    `slot` namespaces the session-state keys so several uploaders on the
    same page each remember their own last file independently, and a page
    rerun never re-imports a file that was already read.

    Returns a result dict (or None when there is nothing new to report).
    """
    if uploaded_file is None:
        return st.session_state.get(f"_result_{slot}")

    signature = (uploaded_file.name, uploaded_file.size)
    if st.session_state.get(f"_sig_{slot}") == signature:
        return st.session_state.get(f"_result_{slot}")

    with st.spinner(f"Reading {uploaded_file.name}..."):
        try:
            months, warnings, diagnostics = process_uploaded_workbook(uploaded_file)
            result = {
                "kind": "inventory", "saved": 0, "months": months,
                "warnings": warnings, "diagnostics": diagnostics,
                "name": uploaded_file.name,
            }
        except Exception as exc:
            result = {
                "kind": "error", "saved": 0, "months": [],
                "warnings": [f"Could not read '{uploaded_file.name}': {exc}"],
                "diagnostics": [], "name": uploaded_file.name,
            }

    st.session_state[f"_sig_{slot}"] = signature
    st.session_state[f"_result_{slot}"] = result
    return result


def render_upload_result(result):
    """Prints one upload's outcome: what was saved, any warnings, and the
    per-sheet breakdown behind an expander."""
    if not result:
        return

    if result["kind"] == "inventory" and result["months"]:
        labels = ", ".join(month_label(y, m) for y, m in result["months"])
        st.success(f"**{result['name']}** — saved: {labels}")
    elif result["kind"] == "error":
        pass
    elif not result["warnings"]:
        st.warning(
            f"**{result['name']}** — nothing recognizable was found. Check that "
            "the sheet has an 'Ingredient / Item' header, or that an item-sales "
            "sheet has 'Item Name' and 'Quantity' columns."
        )

    for w in result["warnings"]:
        st.warning(w)

    if result["diagnostics"]:
        with st.expander(f"Sheet-by-sheet breakdown — {result['name']}"):
            for d in result["diagnostics"]:
                st.markdown(f"**{d['sheet']}** — {d['kind']}")
                st.caption(d["detail"])


def quad_card(col, icon, label, value, sub="", warn=False):
    cls = "quad-card quad-warn" if warn else "quad-card"
    col.markdown(
        f"""
        <div class="{cls}">
            <div class="quad-icon">{icon}</div>
            <p class="quad-label">{label}</p>
            <p class="quad-value">{value}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# SIDEBAR — BRAND + NAVIGATION
# ============================================================

PAGE_DASHBOARD = "Dashboard"
PAGE_GP = "Gross Profit Tracker"
PAGE_UPLOAD = "Upload monthly files"

# (page key, label, Material Symbols icon name, emoji fallback)
NAV_ITEMS = [
    (PAGE_DASHBOARD, "Dashboard", "grid_view", "📊"),
    (PAGE_GP, "Gross Profit Tracker", "monitoring", "📈"),
    (PAGE_UPLOAD, "Upload monthly files", "upload", "📤"),
]


def sidebar_nav_button(label, material_icon, emoji_icon, key, is_active):
    """A sidebar button with an icon, using Streamlit's own `icon`
    support (Material Symbols) when available and falling back to a
    plain emoji-prefixed label on older Streamlit builds that don't
    accept the `icon` argument — so a version mismatch never crashes
    the page, it just loses the nicer icon."""
    kind = "primary" if is_active else "secondary"
    try:
        return st.sidebar.button(
            label, icon=f":material/{material_icon}:", key=key,
            use_container_width=True, type=kind,
        )
    except TypeError:
        return st.sidebar.button(
            f"{emoji_icon}  {label}", key=key,
            use_container_width=True, type=kind,
        )

if "_active_page" not in st.session_state:
    st.session_state["_active_page"] = PAGE_DASHBOARD

st.sidebar.markdown(
    """
    <div class="tl-brand">
        <div class="tl-avatar">TL</div>
        <div>
            <div class="tl-name">Tea Lounge</div>
            <div class="tl-sub">Ingredient usage</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.sidebar.markdown('<div style="height: 14px;"></div>', unsafe_allow_html=True)

for page_key, page_label, page_icon, page_emoji in NAV_ITEMS:
    is_active = st.session_state["_active_page"] == page_key
    if sidebar_nav_button(
        page_label, page_icon, page_emoji,
        "tlnav_" + re.sub(r"\W+", "_", page_key.lower()), is_active,
    ):
        st.session_state["_active_page"] = page_key
        st.rerun()

st.sidebar.markdown(
    """
    <div class="tl-note">
        <b>File-based system</b>
        <span>Reports calculate directly from stored Excel files.</span>
    </div>
    """,
    unsafe_allow_html=True,
)

active_page = st.session_state["_active_page"]

# ============================================================
# PAGE — UPLOAD MONTHLY FILES
# Two uploaders, one per monthly workbook. Each file is inspected and
# routed on its own, so the order they are dropped in does not matter.
# Ends with st.stop(), so nothing below runs while this page is open.
# ============================================================

if active_page == PAGE_UPLOAD:
    st.markdown(
        """
        <div class="app-header">
            <p class="eyebrow">Tea Lounge</p>
            <div class="dashboard-title">Upload monthly files</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    upload_col, help_col = st.columns([1.6, 1], gap="large")

    with upload_col:
        st.markdown(
            '<div class="section-title">Select both Excel workbooks</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="small-note">Upload them in any order. Each workbook is '
            'inspected on its own, and the month it belongs to is read from the '
            'sheet name, the file name, or a Month column inside the sheet.</div>',
            unsafe_allow_html=True,
        )

        sales_file = st.file_uploader(
            "Monthly Item Sales Report",
            type=["xlsx", "xls"],
            key="_upload_sales",
            help="The item-sales workbook — needs Item Name and Quantity columns.",
        )
        sales_result = handle_workbook_upload(sales_file, "sales")
        render_upload_result(sales_result)

        st.markdown("")

        inventory_file = st.file_uploader(
            "Tea Lounge Ingredient Summary",
            type=["xlsx", "xls"],
            key="_upload_inventory",
            help=(
                "The inventory workbook — needs an 'Ingredient / Item' header. "
                "It may hold a single month or all twelve month sheets."
            ),
        )
        inventory_result = handle_workbook_upload(inventory_file, "inventory")
        render_upload_result(inventory_result)

    with help_col:
        st.markdown(
            '<div class="section-title">File requirements</div>',
            unsafe_allow_html=True,
        )
        requirements = [
            ("Two Excel files",
             "Upload one sales report and one ingredient summary, both in "
             ".xlsx or .xls format."),
            ("Matching month",
             "Both files should cover the same month and year, so their "
             "figures are filed together."),
            ("Mapped headers",
             "Column order can change — headings are matched by name, not "
             "by position."),
            ("History is kept",
             "Every month you upload is stored permanently. Adding a new "
             "month never overwrites the ones already saved."),
        ]
        for i, (title, body) in enumerate(requirements, start=1):
            st.markdown(
                f'<div class="recipe-note"><b>{i}. {title}</b><br>{body}</div>',
                unsafe_allow_html=True,
            )

        saved_months = list_stored_months()
        if saved_months:
            st.markdown(
                '<div class="section-title">Months on file</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                '<div class="small-note">'
                + " &nbsp;·&nbsp; ".join(
                    month_label(y, m) for y, m in sorted(saved_months, reverse=True)
                )
                + "</div>",
                unsafe_allow_html=True,
            )

    st.stop()

# ============================================================
# PAGE — GROSS PROFIT TRACKER
# Its own self-contained page, drawn in the palette's LIGHT theme while
# the sidebar and the header banner stay deep green. Chart colours are
# the ones the dashboard's own charts use. It ends with st.stop(), so
# none of the dashboard sections below run while this page is open.
# Tables are built as HTML (not st.dataframe) because a dataframe takes
# its colours from Streamlit's theme and would stay dark on this page.
# ============================================================

GL = PALETTES["light"]

GP_PAGE_CSS = """
<style>
    /* ---- light canvas (overrides the dark canvas for this page only) ---- */
    html, body,
    [data-testid="stApp"],
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewContainer"] > .main,
    [data-testid="stMain"],
    [data-testid="stMainBlockContainer"],
    .main, .block-container { background-color: __BG__ !important; }
    [data-testid="stHeader"] { background-color: __BG__ !important; }
    [data-testid="stHeader"] *, [data-testid="stToolbar"] * { color: __TXT__ !important; }

    /* ---- native widgets, repainted light ---- */
    [data-testid="stMain"] label,
    [data-testid="stMain"] [data-testid="stWidgetLabel"] * { color: __TXT__ !important; font-weight: 600; }
    [data-testid="stMain"] div[data-baseweb="select"] > div {
        background-color: #ffffff !important; border: 1px solid __LINE__ !important; border-radius: 10px !important;
    }
    [data-testid="stMain"] div[data-baseweb="select"] * { color: __TXT__ !important; }
    [data-testid="stMain"] div[data-baseweb="select"] svg { fill: __TXT__ !important; }
    div[data-baseweb="popover"] ul, div[data-baseweb="popover"] [role="listbox"] { background-color: #ffffff !important; }
    div[data-baseweb="popover"] li, div[data-baseweb="popover"] [role="option"] {
        background-color: #ffffff !important; color: __TXT__ !important;
    }
    div[data-baseweb="popover"] li:hover, div[data-baseweb="popover"] [role="option"]:hover {
        background-color: __SOFT__ !important;
    }
    [data-testid="stMain"] [data-testid="stFileUploader"] section,
    [data-testid="stMain"] [data-testid="stFileUploaderDropzone"] {
        background-color: #ffffff !important; border: 1.5px dashed __GREEN__ !important; border-radius: 12px !important;
    }
    [data-testid="stMain"] [data-testid="stFileUploader"] * { color: __TXT__ !important; }
    [data-testid="stMain"] [data-testid="stFileUploader"] small { display: none !important; }
    [data-testid="stMain"] [data-testid="stFileUploader"] svg { fill: __TXT__ !important; }
    [data-testid="stMain"] button[kind="secondary"],
    [data-testid="stMain"] button[data-testid="stBaseButton-secondary"],
    [data-testid="stMain"] [data-testid="stDownloadButton"] button,
    [data-testid="stMain"] [data-testid="stFileUploader"] button {
        background-color: #ffffff !important; color: __TXT__ !important;
        border: 1px solid __LINE__ !important; border-radius: 10px !important;
    }
    [data-testid="stMain"] button[kind="secondary"] *,
    [data-testid="stMain"] [data-testid="stDownloadButton"] button * { color: __TXT__ !important; }

    /* ---- tracker components ---- */
    .gp-title {
        font-size: 20px; font-weight: 750; margin: 30px 0 4px 0; color: __TXT__;
        border-left: 5px solid __GREEN__; padding-left: 12px;
    }
    .gp-sub { color: __MUTED__; font-size: 13px; margin: 0 0 12px 0; line-height: 1.5; }
    .gp-panel {
        background: #ffffff; border: 1px solid __LINE__; border-radius: 14px;
        padding: 16px 18px; margin-bottom: 12px; color: __TXT__;
    }
    .gp-note {
        border-radius: 10px; padding: 10px 14px; margin: 6px 0 10px 0; font-size: 13px; line-height: 1.5;
        background: __SOFT__; border-left: 4px solid __GREEN__; color: __TXT__;
    }
    .gp-note b, .gp-note span { color: inherit; }
    .gp-note.warn { background: #fff7ea; border-left-color: #e0a526; color: #6b4a10; }
    .gp-note.err  { background: #fdeeee; border-left-color: __LOSS__; color: #7a1d1f; }
    .gp-chips { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; margin: 8px 0 14px 0; }
    .gp-chip { background: #ffffff; border: 1px solid __LINE__; border-radius: 12px; padding: 10px 14px; }
    .gp-chip .k { color: __LABEL__; font-size: 11px; font-weight: 700; letter-spacing: .05em; text-transform: uppercase; }
    .gp-chip .v { color: __TXT__; font-size: 19px; font-weight: 800; margin-top: 2px; }
    .gp-chip .d { font-size: 12px; font-weight: 700; margin-top: 2px; }
    .gp-chip .d.up { color: #0b6e4f; } .gp-chip .d.down { color: __LOSS__; } .gp-chip .d.flat { color: __MUTED__; }
    .gp-scroll { overflow-x: auto; background: #ffffff; border: 1px solid __LINE__; border-radius: 14px; margin-bottom: 12px; }
    .gp-table { width: 100%; border-collapse: collapse; font-size: 13px; color: __TXT__; }
    .gp-table th {
        background: __SOFT__; color: __LABEL__; font-size: 11px; font-weight: 700; text-transform: uppercase;
        letter-spacing: .04em; padding: 10px 12px; white-space: nowrap; border-bottom: 1px solid __LINE__;
    }
    .gp-table td { padding: 9px 12px; border-bottom: 1px solid #e9f4ee; white-space: nowrap; }
    .gp-table tr:last-child td { border-bottom: none; }
    .gp-table td.wrap { white-space: normal; min-width: 190px; }
    .gp-barwrap { display: flex; align-items: center; gap: 8px; justify-content: flex-end; }
    .gp-bar { width: 70px; height: 7px; background: #e3f1ea; border-radius: 4px; overflow: hidden; }
    .gp-bar span { display: block; height: 100%; background: __GREEN__; border-radius: 4px; }
    .gp-bar.neg span { background: __LOSS__; }
    .gp-badge {
        display: inline-block; font-size: 10.5px; font-weight: 700; padding: 2px 8px; border-radius: 999px;
        margin-left: 4px; background: #e3f1ea; color: #0b6e4f;
    }
    .gp-badge.loss { background: #fdeeee; color: __LOSS__; }
    .gp-badge.warn { background: #fff3d6; color: #8a5a00; }
    .gp-badge.grey { background: #edf1ef; color: #5b6b63; }
    .gp-list { margin: 4px 0 0 0; padding-left: 18px; color: __TXT__; font-size: 13px; line-height: 1.7; }
</style>
"""
for _token, _value in {
    "__BG__": GL["main_bg"], "__TXT__": GL["section_title"], "__MUTED__": GL["small_note"],
    "__LABEL__": GL["card_label"], "__SOFT__": GL["card_bg"], "__LINE__": GL["card_border"],
    "__GREEN__": GP_GREEN, "__LOSS__": GP_LOSS,
}.items():
    GP_PAGE_CSS = GP_PAGE_CSS.replace(_token, _value)


# ---------- small formatting / layout helpers ----------

def gp_tk(x, decimals=0):
    if x is None or pd.isna(x):
        return "—"
    return f"Tk {x:,.{decimals}f}"


def gp_pct(x, decimals=1):
    if x is None or pd.isna(x):
        return "—"
    return f"{x * 100:.{decimals}f}%"


def gp_short(text, limit=34):
    text = str(text)
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def gp_title(text, sub=""):
    st.markdown(f'<div class="gp-title">{html_escape(text)}</div>', unsafe_allow_html=True)
    if sub:
        st.markdown(f'<div class="gp-sub">{sub}</div>', unsafe_allow_html=True)


def gp_note(kind, text):
    """kind: '' (info), 'warn' or 'err'. `text` may hold simple HTML."""
    st.markdown(f'<div class="gp-note {kind}">{text}</div>', unsafe_allow_html=True)


def gp_chips(items):
    """items: [(label, value, delta_html_or_'')] -> a responsive chip row."""
    cells = "".join(
        f'<div class="gp-chip"><div class="k">{html_escape(k)}</div><div class="v">{v}</div>{d}</div>'
        for k, v, d in items
    )
    st.markdown(f'<div class="gp-chips">{cells}</div>', unsafe_allow_html=True)


def gp_delta(cur, prev, kind="pct"):
    """'▲ 12.3% vs last month' chip line. kind='pp' compares two ratios."""
    if cur is None or prev is None or pd.isna(cur) or pd.isna(prev):
        return ""
    if kind == "pp":
        diff, text = (cur - prev) * 100, None
        text = f"{abs(diff):.1f} pts"
    else:
        if prev == 0:
            return ""
        diff = (cur - prev) / abs(prev) * 100
        text = f"{abs(diff):.1f}%"
    if abs(diff) < 0.05:
        return '<div class="d flat">no change</div>'
    arrow, cls = ("▲", "up") if diff > 0 else ("▼", "down")
    return f'<div class="d {cls}">{arrow} {text}</div>'


def gp_bar(fraction, label, negative=False):
    width = max(0.0, min(100.0, abs(fraction) * 100))
    cls = "gp-bar neg" if negative else "gp-bar"
    return f'<div class="gp-barwrap"><div class="{cls}"><span style="width:{width:.1f}%"></span></div><b>{label}</b></div>'


def gp_table(headers, rows, aligns=None, wrap_cols=()):
    """Static HTML table. Header text is fixed by us; cells are either
    numbers/pre-escaped text or small HTML fragments built here."""
    aligns = aligns or ["left"] + ["right"] * (len(headers) - 1)
    head = "".join(f'<th style="text-align:{a}">{h}</th>' for h, a in zip(headers, aligns))
    body = ""
    for row in rows:
        cells = "".join(
            f'<td class="{"wrap" if i in wrap_cols else ""}" style="text-align:{aligns[i]}">{c}</td>'
            for i, c in enumerate(row)
        )
        body += f"<tr>{cells}</tr>"
    st.markdown(
        f'<div class="gp-scroll"><table class="gp-table"><thead><tr>{head}</tr></thead>'
        f"<tbody>{body}</tbody></table></div>",
        unsafe_allow_html=True,
    )


def gp_style_fig(fig, height, **layout):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=GL["chart_font"], family="Inter, sans-serif"),
        margin=dict(l=10, r=10, t=30, b=10),
    )
    fig.update_xaxes(gridcolor="#dcefe5", zeroline=False, linecolor="#bfe9d3")
    fig.update_yaxes(gridcolor="#dcefe5", zeroline=False, linecolor="#bfe9d3")
    if layout:
        fig.update_layout(**layout)
    return fig


def gp_chart(fig):
    # theme=None: draw the figure exactly as styled here rather than letting
    # Streamlit's (dark) theme recolour it.
    st.plotly_chart(fig, use_container_width=True, theme=None)


# ---------- page sections ----------

def gp_upload_panel(existing_months):
    """Upload -> read -> confirm month -> save. Nothing is written until
    'Save' is pressed, so a file can never land in the wrong month by
    accident. The file's name proposes the month; the year is assumed."""
    n = st.session_state.get("_gp_up_n", 0)
    upload = st.file_uploader(
        "GP tracker workbook", type=["xlsx", "xls"], key=f"_gp_up_{n}",
        label_visibility="collapsed",
    )
    if upload is None:
        st.markdown(
            '<div class="gp-sub">Upload the GP tracker workbook once, at the end of each month. '
            "Each month is stored on its own and gets its own tracker.</div>",
            unsafe_allow_html=True,
        )
        return

    data = upload.getvalue()
    sig = f"{upload.name}|{len(data)}"
    cache = st.session_state.get("_gp_parsed")
    if not cache or cache.get("sig") != sig:
        try:
            menus, warnings, diagnostics = read_gp_workbook(io.BytesIO(data))
            cache = {"sig": sig, "menus": menus, "warnings": warnings, "diagnostics": diagnostics, "error": None}
        except Exception as exc:
            cache = {"sig": sig, "menus": [], "warnings": [], "diagnostics": [], "error": str(exc)}
        st.session_state["_gp_parsed"] = cache

    if cache["error"]:
        gp_note("err", f"Could not open <b>{html_escape(upload.name)}</b>: {html_escape(cache['error'])}")
        return
    if not cache["menus"]:
        gp_note(
            "err",
            f"<b>{html_escape(upload.name)}</b> doesn't look like a GP tracker workbook. Each menu sheet needs "
            "<b>Ingredient / Item</b>, <b>Selling Price Per Cup (Tk)</b> and <b>QTY Sold</b> columns.",
        )
        return

    detected = gp_detect_period(upload.name)
    today = date.today()
    d_year, d_month = detected if detected else (today.year, today.month)
    gp_note(
        "",
        f"Read <b>{len(cache['menus'])} menus</b> from <b>{html_escape(upload.name)}</b>. "
        + (f"Filed under <b>{month_label(d_year, d_month)}</b> from the file name — change it below if that's wrong."
           if detected else "No month found in the file name — choose the month below."),
    )
    years = sorted({today.year - 3, today.year - 2, today.year - 1, today.year, today.year + 1, d_year})
    pick_m, pick_y = st.columns(2)
    month = pick_m.selectbox(
        "Month", list(range(1, 13)), index=d_month - 1,
        format_func=lambda m: calendar.month_name[m], key=f"_gp_pm_{sig}",
    )
    year = pick_y.selectbox("Year", years, index=years.index(d_year), key=f"_gp_py_{sig}")

    replacing = (year, month) in existing_months
    if replacing:
        gp_note("warn", f"<b>{month_label(year, month)}</b> is already saved — saving will replace it.")
    for w in cache["warnings"]:
        gp_note("warn", html_escape(w))

    if st.button(
        "Replace saved month" if replacing else "Save to tracker",
        type="primary", key=f"_gp_save_{sig}",
    ):
        save_gp_month(year, month, cache["menus"], upload.name)
        st.session_state["_gp_pending_select"] = (year, month)
        st.session_state["_gp_flash"] = f"Saved <b>{month_label(year, month)}</b> — {len(cache['menus'])} menus."
        st.session_state["_gp_up_n"] = n + 1
        st.session_state.pop("_gp_parsed", None)
        st.rerun()


def gp_render_kpis(totals, prev):
    """Option C: two grouped panels. Left = gross profit with a margin ring,
    plus revenue and ingredient cost. Right = cups sold and the averages."""
    margin = totals["margin"]
    has_margin = margin is not None and not pd.isna(margin)
    ring_pct = max(0.0, min(100.0, margin * 100)) if has_margin else 0.0
    ring_color = GP_LOSS if has_margin and margin < 0 else GP_GREEN
    profit_color = GP_LOSS if totals["profit"] < 0 else GP_GREEN_DARK

    kpi_css = """<style>
.gp-kpi { display: grid; grid-template-columns: 1.1fr 1fr; gap: 12px; align-items: stretch; margin: 8px 0 4px 0; }
.gp-kpi-panel { background: #ffffff; border: 1px solid __LINE__; border-radius: 14px; padding: 16px 20px; display: flex; flex-direction: column; }
.gp-kpi-hero { display: flex; align-items: center; gap: 18px; margin-bottom: 8px; }
.gp-kpi-ring { position: relative; width: 96px; height: 96px; flex: 0 0 96px; border-radius: 50%; }
.gp-kpi-ring::before { content: ""; position: absolute; inset: 10px; background: #ffffff; border-radius: 50%; }
.gp-kpi-ring span { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; font-size: 16px; font-weight: 800; color: __TXT__; z-index: 1; }
.gp-kpi-k { color: __LABEL__; font-size: 12.5px; font-weight: 600; margin: 0; }
.gp-kpi-big { font-size: 28px; font-weight: 800; line-height: 1.2; margin: 2px 0 2px 0; }
.gp-kpi-cap { color: __MUTED__; font-size: 12px; margin: 0; }
.gp-kpi-row { display: flex; justify-content: space-between; align-items: center; flex: 1 1 0; min-height: 40px; padding: 8px 0; border-top: 1px solid #e9f4ee; font-size: 14px; }
.gp-kpi-row.first { border-top: none; }
.gp-kpi-row span:first-child { color: __LABEL__; }
.gp-kpi-row span:last-child { color: __TXT__; font-weight: 800; }
.gp-kpi-row span.good { color: __GOOD__; }
@media (max-width: 760px) { .gp-kpi { grid-template-columns: 1fr; } }
</style>"""
    for _token, _value in {
        "__LINE__": GL["card_border"], "__TXT__": GL["section_title"], "__LABEL__": GL["card_label"],
        "__MUTED__": GL["small_note"], "__GOOD__": GP_GREEN_DARK,
    }.items():
        kpi_css = kpi_css.replace(_token, _value)

    ring_style = (
        f"background: conic-gradient({ring_color} 0 {ring_pct:.2f}%, #e3f1ea {ring_pct:.2f}% 100%);"
    )
    kpi_html = "".join([
        '<div class="gp-kpi">',
        '<div class="gp-kpi-panel">',
        '<div class="gp-kpi-hero">',
        f'<div class="gp-kpi-ring" style="{ring_style}"><span>{gp_pct(margin)}</span></div>',
        '<div>',
        '<p class="gp-kpi-k">Gross profit</p>',
        f'<p class="gp-kpi-big" style="color:{profit_color}">{gp_tk(totals["profit"])}</p>',
        '<p class="gp-kpi-cap">gross margin shown in ring</p>',
        '</div>',
        '</div>',
        f'<div class="gp-kpi-row"><span>Total revenue</span><span>{gp_tk(totals["revenue"])}</span></div>',
        f'<div class="gp-kpi-row"><span>Total ingredient cost</span><span>{gp_tk(totals["cost"])}</span></div>',
        '</div>',
        '<div class="gp-kpi-panel">',
        f'<div class="gp-kpi-row first"><span>Cups sold</span><span>{fmt_num(totals["cups"])}</span></div>',
        f'<div class="gp-kpi-row"><span>Avg selling price</span><span>{gp_tk(totals["avg_price"], 1)}</span></div>',
        f'<div class="gp-kpi-row"><span>Avg cost</span><span>{gp_tk(totals["avg_cost"], 1)}</span></div>',
        f'<div class="gp-kpi-row"><span>Avg profit</span><span class="good">{gp_tk(totals["avg_gp"], 1)}</span></div>',
        '</div>',
        '</div>',
    ])
    st.markdown(kpi_css + kpi_html, unsafe_allow_html=True)

    if prev is not None:
        label, p = prev
        st.markdown(f'<div class="gp-sub" style="margin-top:14px">Change vs {label}</div>', unsafe_allow_html=True)
        gp_chips([
            ("Revenue", gp_tk(totals["revenue"]), gp_delta(totals["revenue"], p["revenue"])),
            ("Gross profit", gp_tk(totals["profit"]), gp_delta(totals["profit"], p["profit"])),
            ("Gross margin", gp_pct(totals["margin"]), gp_delta(totals["margin"], p["margin"], "pp")),
            ("Cups sold", fmt_num(totals["cups"]), gp_delta(totals["cups"], p["cups"])),
            ("Avg profit / cup", gp_tk(totals["avg_gp"], 1), gp_delta(totals["avg_gp"], p["avg_gp"])),
        ])


def gp_render_ranking(df):
    """Section 1 — every menu ranked by QTY sold as a horizontal bar graph:
    best seller at the top, slowest at the bottom. Bars are labelled with
    the item name only (no menu ID); the top 3 are the darkest green and
    menus with no sales are drawn pale. Hover shows rank and share."""
    gp_title(
        "1. Best sellers — ranking by cups sold",
        "Menus ranked by QTY sold: the best-selling item at the top, the least-selling at the bottom. "
        "The top 3 are the darkest bars.",
    )
    ranked = df.sort_values(["qty_sold", "id"], ascending=[False, True]).reset_index(drop=True)
    ranked["rank"] = ranked["qty_sold"].rank(method="min", ascending=False).astype(int)
    total_qty = float(ranked["qty_sold"].sum())
    top_qty = float(ranked["qty_sold"].max())
    share_pct = (ranked["qty_sold"] / total_qty * 100) if total_qty else pd.Series(0.0, index=ranked.index)

    # Item names as the axis labels; if two menus ever share a name, the
    # second gets a "(2)" so the bars don't merge into one row.
    seen, labels = {}, []
    for name in ranked["name"]:
        short = gp_short(name, 42)
        seen[short] = seen.get(short, 0) + 1
        labels.append(short if seen[short] == 1 else f"{short} ({seen[short]})")

    colors = [
        GP_GREEN_PALE if q <= 0 else (GP_GREEN_DARK if r <= 3 else GP_GREEN)
        for q, r in zip(ranked["qty_sold"], ranked["rank"])
    ]
    fig = go.Figure(go.Bar(
        x=ranked["qty_sold"], y=labels, orientation="h",
        marker=dict(color=colors, line=dict(color="#ffffff", width=0.5)),
        text=[fmt_num(q) for q in ranked["qty_sold"]],
        textposition="outside", cliponaxis=False, textfont=dict(size=12),
        customdata=np.stack([ranked["name"], ranked["rank"], share_pct], axis=-1),
        hovertemplate="<b>%{customdata[0]}</b><br>Rank: #%{customdata[1]}"
                      "<br>Cups sold: %{x:,.0f}<br>Share of cups: %{customdata[2]:.1f}%<extra></extra>",
    ))
    gp_style_fig(
        fig, max(380, 30 * len(ranked) + 90), showlegend=False, bargap=0.3,
        xaxis=dict(title="Cups sold", range=[0, top_qty * 1.15 if top_qty else 1]),
        yaxis=dict(autorange="reversed", automargin=True, tickfont=dict(size=12), title=None),
        margin=dict(l=10, r=40, t=10, b=40),
    )
    gp_chart(fig)


def gp_render_profitability(df):
    gp_title(
        "3. Which menus earn their place",
        "Volume against margin shows who is carrying the month; the map shows who brings in the revenue.",
    )
    d = df[(df["qty_sold"] > 0) & df["has_price"] & df["margin"].notna()].copy()
    if d.empty:
        gp_note("", "No menu has both a selling price and cups sold, so there is nothing to compare.")
        return
    d["label"] = d["id"] + " · " + d["name"].apply(gp_short)
    left, right = st.columns(2, gap="large")
    with left:
        sizes = d["gross_profit"].clip(lower=0) + 1.0
        median_margin = float(d["margin"].median() * 100)
        fig = go.Figure(go.Scatter(
            x=d["qty_sold"], y=d["margin"] * 100, mode="markers+text",
            text=d["id"], textposition="top center", textfont=dict(size=10),
            customdata=np.stack([d["label"], d["gross_profit"]], axis=-1),
            hovertemplate="<b>%{customdata[0]}</b><br>Cups sold: %{x:,.0f}<br>Margin: %{y:.1f}%"
                          "<br>Gross profit: Tk %{customdata[1]:,.0f}<extra></extra>",
            marker=dict(
                size=sizes, sizemode="area", sizeref=2.0 * float(sizes.max()) / (44.0 ** 2), sizemin=5,
                color=[GP_LOSS if v < 0 else GP_GREEN for v in d["gross_profit"]],
                line=dict(color="#ffffff", width=1), opacity=0.85,
            ),
        ))
        fig.add_hline(y=median_margin, line_dash="dot", line_color=GP_GREEN_DARK)
        gp_style_fig(
            fig, 400, showlegend=False,
            title=dict(text=f"Cups sold vs margin (dotted line = median margin {median_margin:.0f}%)", font=dict(size=14)),
            xaxis=dict(title="Cups sold (log scale)", type="log"),
            yaxis=dict(title="Margin (%)"),
            margin=dict(l=10, r=10, t=44, b=10),
        )
        gp_chart(fig)
    with right:
        tree = d[d["total_revenue"] > 0]
        if tree.empty:
            gp_note("", "No revenue to map.")
        else:
            fig = px.treemap(
                tree, path=["label"], values="total_revenue", color="margin",
                color_continuous_scale=[GP_GREEN_PALE, GP_GREEN, GP_GREEN_DARK],
            )
            fig.update_traces(textinfo="label+value", root_color="rgba(0,0,0,0)")
            fig.update_coloraxes(colorbar_title="Margin", colorbar_tickformat=".0%")
            gp_style_fig(
                fig, 400, margin=dict(l=4, r=4, t=44, b=4),
                title=dict(text="Revenue by menu (colour = margin)", font=dict(size=14)),
            )
            gp_chart(fig)


def gp_menu_flags(row):
    flags = ""
    if not row["has_price"]:
        flags += '<span class="gp-badge warn">No price</span>'
    elif row["gp_per_cup"] is not None and row["gp_per_cup"] < 0:
        flags += '<span class="gp-badge loss">Loss</span>'
    if row["qty_sold"] <= 0:
        flags += '<span class="gp-badge grey">No sales</span>'
    if row["cost_source"] == "typed":
        flags += '<span class="gp-badge warn">Typed cost</span>'
    if row["missing_count"] and row["cost_source"] == "sum":
        flags += f'<span class="gp-badge warn">{int(row["missing_count"])} unpriced</span>'
    return flags


GP_SORTS = {
    "Gross profit — high to low": ("gross_profit", False),
    "Margin — high to low": ("margin", False),
    "Margin — low to high": ("margin", True),
    "Revenue — high to low": ("total_revenue", False),
    "Cups sold — high to low": ("qty_sold", False),
    "Cost per cup — high to low": ("cost_per_cup", False),
    "Menu ID": ("id", True),
}


def gp_render_table(df, month_key_text):
    gp_title("4. Every menu, side by side", "Sort it any way you like. Flags point out anything worth a second look.")
    sort_label = st.selectbox("Sort by", list(GP_SORTS), key="_gp_sort")
    column, ascending = GP_SORTS[sort_label]
    ordered = df.sort_values(column, ascending=ascending, na_position="last")
    rows = []
    for r in ordered.to_dict("records"):
        margin = r["margin"]
        rows.append([
            html_escape(r["id"]),
            f'{html_escape(r["name"])}{gp_menu_flags(r)}',
            gp_tk(r["selling_price"], 0),
            gp_tk(r["cost_per_cup"], 2),
            gp_tk(r["gp_per_cup"], 2),
            gp_bar(margin, gp_pct(margin), negative=margin < 0) if margin is not None and not pd.isna(margin) else "—",
            fmt_num(r["qty_sold"]),
            gp_tk(r["total_revenue"]),
            gp_tk(r["total_cost"]),
            gp_tk(r["gross_profit"]),
            gp_pct(r["profit_share"]),
        ])
    gp_table(
        ["ID", "Menu", "Price / cup", "Cost / cup", "Profit / cup", "Margin", "Cups", "Revenue", "Total cost",
         "Gross profit", "Share of profit"],
        rows,
        aligns=["left", "left", "right", "right", "right", "right", "right", "right", "right", "right", "right"],
        wrap_cols=(1,),
    )
    export = ordered[[
        "id", "name", "selling_price", "cost_per_cup", "gp_per_cup", "margin", "qty_sold",
        "total_revenue", "total_cost", "gross_profit", "profit_share",
    ]].rename(columns={
        "id": "Menu ID", "name": "Menu", "selling_price": "Selling price per cup", "cost_per_cup": "Cost per cup",
        "gp_per_cup": "Gross profit per cup", "margin": "Margin", "qty_sold": "Cups sold",
        "total_revenue": "Total revenue", "total_cost": "Total cost", "gross_profit": "Gross profit",
        "profit_share": "Share of profit",
    })
    st.download_button(
        "Download this table (CSV)", data=export.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"gp_menu_table_{month_key_text}.csv", mime="text/csv", key="_gp_csv",
    )


def gp_menu_history(menu_id):
    """This menu's figures in every stored month."""
    rows = []
    for year, month in list_gp_months():
        payload = load_gp_month(year, month)
        if not payload:
            continue
        for menu in payload["menus"]:
            if str(menu.get("id", "")).upper() == menu_id.upper():
                m = gp_menu_metrics(menu)
                rows.append({"label": f"{calendar.month_abbr[month]} {year}", **m})
    return pd.DataFrame(rows)


def gp_render_deep_dive(payload, df, month_key_text):
    gp_title(
        "2. One menu, ingredient by ingredient",
        "Pick a menu to see exactly how its selling price is spent, cup by cup.",
    )
    ids = df["id"].tolist()
    names = dict(zip(df["id"], df["name"]))
    pick = st.selectbox(
        "Menu", ids, format_func=lambda i: f"{i} · {names[i]}", key=f"_gp_menu_{month_key_text}",
    )
    menu = next(m for m in payload["menus"] if m.get("id") == pick)
    met = gp_menu_metrics(menu)

    gp_chips([
        ("Selling price / cup", gp_tk(met["selling_price"], 0), ""),
        ("Cost / cup", gp_tk(met["cost_per_cup"], 2), ""),
        ("Profit / cup", gp_tk(met["gp_per_cup"], 2), ""),
        ("Margin", gp_pct(met["margin"]), ""),
        ("Cups sold", fmt_num(met["qty_sold"]), ""),
        ("Gross profit", gp_tk(met["gross_profit"]), ""),
    ])

    if not met["has_price"]:
        gp_note("warn", "This menu has no selling price in the workbook, so profit can't be worked out.")
        return

    left, right = st.columns([1.15, 1], gap="large")
    with left:
        steps = gp_waterfall_steps(menu, met)
        losing = met["gp_per_cup"] < 0
        fig = go.Figure(go.Waterfall(
            orientation="v",
            measure=[s[2] for s in steps],
            x=[gp_short(s[0], 18) for s in steps],
            y=[s[1] for s in steps],
            text=[f"{s[1]:,.1f}" for s in steps], textposition="outside",
            increasing=dict(marker=dict(color=GP_GREEN_LIGHT)),
            decreasing=dict(marker=dict(color=GP_GREEN_DARK)),
            totals=dict(marker=dict(color=GP_LOSS if losing else GP_GREEN)),
            connector=dict(line=dict(color="#9cc9b4", width=1)),
            hovertemplate="%{x}<br>Tk %{y:,.2f}<extra></extra>",
        ))
        gp_style_fig(
            fig, 420, showlegend=False,
            title=dict(text="From selling price to profit, per cup (Tk)", font=dict(size=14)),
            xaxis=dict(tickangle=-40), margin=dict(l=10, r=10, t=44, b=10),
        )
        gp_chart(fig)
    with right:
        cost_total = met["cost_per_cup"] or 0.0
        rows = []
        for ing in menu.get("ingredients") or []:
            qty = ing["qty_raw"] or (
                f'{ing["qty_value"]:g} {ing["qty_unit"]}'.strip() if ing.get("qty_value") is not None else "—"
            )
            if ing.get("cost") is None:
                rows.append([html_escape(ing["name"]), html_escape(qty), '<span class="gp-badge warn">not entered</span>', "—"])
            else:
                share = ing["cost"] / cost_total if cost_total else 0.0
                rows.append([html_escape(ing["name"]), html_escape(qty), gp_tk(ing["cost"], 2), gp_bar(share, gp_pct(share))])
        gap = met["cost_per_cup"] - met["listed_cost"]
        if abs(gap) > GP_COST_TOLERANCE:
            share = gap / cost_total if cost_total else 0.0
            rows.append([
                "<i>Other cost (typed total)</i>", "—", gp_tk(gap, 2),
                gp_bar(share, gp_pct(share), negative=gap < 0),
            ])
        rows.append(["<b>Cost per cup</b>", "", f"<b>{gp_tk(met['cost_per_cup'], 2)}</b>", ""])
        gp_table(
            ["Ingredient", "Quantity", "Cost", "Share of cost"], rows,
            aligns=["left", "left", "right", "right"], wrap_cols=(0,),
        )
    if met["cost_source"] == "typed":
        gp_note(
            "warn",
            f"The workbook has a typed cost per cup of <b>{gp_tk(met['cost_per_cup'], 2)}</b>, while the ingredient lines "
            f"add up to <b>{gp_tk(met['listed_cost'], 2)}</b>. The typed total is used.",
        )
    elif met["missing_costs"]:
        gp_note(
            "warn",
            "No cost entered for: " + html_escape(", ".join(met["missing_costs"]))
            + ". The cost per cup above leaves them out, so profit may be overstated.",
        )

    history = gp_menu_history(pick)
    if len(history) >= 2:
        fig = go.Figure()
        fig.add_bar(x=history["label"], y=history["gross_profit"], name="Gross profit (Tk)", marker_color=GP_GREEN)
        fig.add_scatter(
            x=history["label"], y=history["margin"] * 100, name="Margin (%)", yaxis="y2",
            mode="lines+markers", line=dict(color=GP_GREEN_DARK, width=2),
        )
        gp_style_fig(
            fig, 320, title=dict(text=f"{pick} across months", font=dict(size=14)),
            yaxis=dict(title="Gross profit (Tk)"),
            yaxis2=dict(title="Margin (%)", overlaying="y", side="right", showgrid=False),
            legend=dict(orientation="h", y=-0.2), margin=dict(l=10, r=10, t=44, b=10),
        )
        gp_chart(fig)


def gp_render_checks(df):
    gp_title("5. Worth a second look", "Things in the workbook that can make the numbers look better or worse than they are.")
    items = []
    loss = df[df["has_price"] & (df["gp_per_cup"] < 0)]
    for r in loss.itertuples():
        items.append(f"<b>{html_escape(r.id)} · {html_escape(r.name)}</b> sells below its cost — it loses {gp_tk(-r.gp_per_cup, 2)} on every cup.")
    for r in df[~df["has_price"]].itertuples():
        items.append(f"<b>{html_escape(r.id)} · {html_escape(r.name)}</b> has no selling price, so it is left out of every total.")
    for r in df[df["has_price"] & (df["qty_sold"] <= 0)].itertuples():
        items.append(f"<b>{html_escape(r.id)} · {html_escape(r.name)}</b> sold no cups this month.")
    for r in df[df["cost_source"] == "typed"].itertuples():
        items.append(
            f"<b>{html_escape(r.id)} · {html_escape(r.name)}</b> uses a typed cost of {gp_tk(r.cost_per_cup, 2)} "
            f"(its ingredient lines add up to {gp_tk(r.listed_cost, 2)}; {int(r.missing_count)} have no cost entered)."
        )
    for r in df[(df["cost_source"] == "sum") & (df["missing_count"] > 0)].itertuples():
        items.append(
            f"<b>{html_escape(r.id)} · {html_escape(r.name)}</b> has no cost for {html_escape(', '.join(r.missing_costs))} — "
            "its cost per cup leaves them out, so its profit may be overstated."
        )
    if not items:
        gp_note("", "Nothing to flag: every menu has a price and complete costs, and none sells at a loss.")
        return
    lis = "".join(f"<li>{i}</li>" for i in items)
    st.markdown(f'<div class="gp-panel"><ul class="gp-list">{lis}</ul></div>', unsafe_allow_html=True)


def gp_render_trend(history):
    gp_title("6. Month by month", "Every month you upload is added here, so the trend builds itself.")
    if len(history) < 2:
        gp_note("", "Only one month is saved so far. Upload next month's workbook and the month-by-month trend appears here.")
        return
    fig = go.Figure()
    fig.add_bar(x=history["short"], y=history["revenue"], name="Revenue", marker_color=GP_GREEN_PALE)
    fig.add_bar(x=history["short"], y=history["cost"], name="Ingredient cost", marker_color=GP_GREEN_DARK)
    fig.add_bar(x=history["short"], y=history["profit"], name="Gross profit", marker_color=GP_GREEN)
    fig.add_scatter(
        x=history["short"], y=history["margin"] * 100, name="Margin (%)", yaxis="y2",
        mode="lines+markers", line=dict(color=GP_GREEN_LIGHT, width=3),
    )
    gp_style_fig(
        fig, 400, barmode="group",
        yaxis=dict(title="Tk"),
        yaxis2=dict(title="Margin (%)", overlaying="y", side="right", showgrid=False),
        legend=dict(orientation="h", y=-0.18), margin=dict(l=10, r=10, t=20, b=10),
    )
    gp_chart(fig)
    rows = [
        [html_escape(r.label), gp_tk(r.revenue), gp_tk(r.cost), gp_tk(r.profit), gp_pct(r.margin),
         fmt_num(r.cups), gp_tk(r.avg_gp, 1)]
        for r in history.iloc[::-1].itertuples()
    ]
    gp_table(["Month", "Revenue", "Ingredient cost", "Gross profit", "Margin", "Cups sold", "Profit / cup"], rows)


def gp_render_month(selected, months):
    year, month = selected
    key_text = month_key(year, month)
    payload = load_gp_month(year, month)
    if payload is None:
        gp_note("err", f"The saved file for {month_label(year, month)} could not be read.")
        return
    df = gp_frame(payload["menus"])
    if df.empty:
        gp_note("warn", f"{month_label(year, month)} has no menus in it.")
        return
    totals = gp_totals(df)

    earlier = [ym for ym in months if ym < selected]
    prev = None
    if earlier:
        py, pm = max(earlier)
        prev_payload = load_gp_month(py, pm)
        if prev_payload and prev_payload["menus"]:
            prev = (month_label(py, pm), gp_totals(gp_frame(prev_payload["menus"])))

    st.markdown(
        f'<div class="gp-sub" style="margin-top:6px">Showing <b>{month_label(year, month)}</b> · '
        f'{totals["menus_sold"]} of {totals["menus"]} menus sold · from <i>{html_escape(payload.get("source_file", ""))}</i> '
        f'(saved {html_escape(payload.get("saved_at", ""))})</div>',
        unsafe_allow_html=True,
    )
    gp_render_kpis(totals, prev)
    gp_render_ranking(df)                              # Section 1
    gp_render_deep_dive(payload, df, key_text)         # Section 2
    gp_render_profitability(df)                        # Section 3
    gp_render_table(df, key_text)                      # Section 4
    gp_render_checks(df)                               # Section 5
    gp_render_trend(gp_history())                      # Section 6

    # "Add a month" box lives at the bottom of Section 6.
    gp_title("Add a month")
    gp_upload_panel(months)

    st.markdown('<div style="height:18px"></div>', unsafe_allow_html=True)
    if st.button("Remove this month", key="_gp_rm"):
        st.session_state["_gp_confirm_rm"] = key_text
    if st.session_state.get("_gp_confirm_rm") == key_text:
        gp_note("warn", f"Remove <b>{month_label(year, month)}</b> from the tracker for good? The saved data for this month is deleted.")
        yes, no, _ = st.columns([1, 1, 3])
        if yes.button("Yes, remove it", type="primary", key="_gp_rm_yes"):
            delete_gp_month(year, month)
            st.session_state.pop("_gp_confirm_rm", None)
            st.rerun()
        if no.button("Keep it", key="_gp_rm_no"):
            st.session_state.pop("_gp_confirm_rm", None)
            st.rerun()


if active_page == PAGE_GP:
    st.markdown(GP_PAGE_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="app-header">
            <p class="eyebrow">Tea Lounge</p>
            <div class="dashboard-title">Gross Profit Tracker</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    gp_months = list_gp_months()

    # A save (or removal) asks for a particular month to be shown next.
    # It is applied here, before the month picker exists, because a widget's
    # value can't be changed after the widget has been drawn.
    pending = st.session_state.pop("_gp_pending_select", None)
    if pending in gp_months:
        st.session_state["_gp_month_select"] = pending
    if st.session_state.get("_gp_month_select") not in gp_months:
        st.session_state.pop("_gp_month_select", None)

    flash = st.session_state.pop("_gp_flash", None)
    if flash:
        gp_note("", flash)

    # Report month picker now lives in the sidebar.
    if gp_months:
        st.sidebar.markdown("---")
        st.sidebar.markdown("### Report month")
        selected_month = st.sidebar.selectbox(
            "Report month", sorted(gp_months, reverse=True),
            format_func=lambda ym: month_label(*ym), key="_gp_month_select",
            label_visibility="collapsed",
        )
    else:
        selected_month = None

    if selected_month:
        gp_render_month(selected_month, gp_months)
    else:
        gp_note("", "No month saved yet. Upload your first GP tracker workbook to begin.")
        gp_title("Add a month")
        gp_upload_panel(gp_months)

    st.stop()

st.sidebar.markdown("---")
st.sidebar.markdown("### Select Period")

stored_months = list_stored_months()

if not stored_months:
    st.markdown(
        """
        <div class="app-header">
            <p class="eyebrow">Tea Lounge</p>
            <div class="dashboard-title">Tea Lounge Inventory Tracker</div>
            <div class="dashboard-subtitle">Open <b>Upload monthly files</b> in the sidebar to add your first month.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    c1.info("📤 Upload Excel")
    c2.info("🗓️ Sheets named per month")
    c3.info("📊 History saved automatically")

    st.markdown("### Expected workbook structure")
    st.write(
        "Name each monthly sheet with a month and year, e.g. **January26**, "
        "**February26**, **March26**. Each of those sheets should have the "
        "columns: Ingredient / Item, Total Units Sold, Total Required Qty, "
        "**Projected Use**, Pre. Month Stock, Received 1, Received 2, "
        "Total Stock, Month End Stock, and **Total Use**."
    )
    st.write(
        "Optionally include a sheet named **Monthly Item Sales Report** with "
        "**Item** and **Total Quantity** columns (and a **Month** column if "
        "it covers more than one month) — this powers the 'Cups Sold' and "
        "'Menu Items' cards."
    )
    st.stop()

years_available = sorted({y for y, _ in stored_months}, reverse=True)
selected_year = st.sidebar.selectbox("Year", years_available, index=0)

months_for_year = sorted({m for y, m in stored_months if y == selected_year})
month_name_options = [calendar.month_name[m] for m in months_for_year]
default_index = len(month_name_options) - 1  # most recent month for that year
selected_month_name = st.sidebar.selectbox("Month", month_name_options, index=default_index)
selected_month = months_for_year[month_name_options.index(selected_month_name)]

df = load_month_inventory(selected_year, selected_month)
if df is None:
    st.error("Could not load stored data for the selected period.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.markdown("### Filters")

category_list = sorted(df["Category"].dropna().unique().tolist())
selected_categories = st.sidebar.multiselect(
    "Category",
    options=category_list,
    default=category_list,
)
search_term = st.sidebar.text_input("Search ingredient", "")

if not selected_categories:
    st.warning("Select at least one category from the sidebar.")
    st.stop()

filtered = df[df["Category"].isin(selected_categories)].copy()
if search_term.strip():
    filtered = filtered[
        filtered["Ingredient"].str.contains(search_term.strip(), case=False, na=False)
    ]

with st.sidebar.expander(f"📁 Stored months ({len(stored_months)})"):
    for y, m in sorted(stored_months, reverse=True):
        st.write(f"• {month_label(y, m)}")
    st.markdown("---")
    if st.button("🗑️ Clear all stored data", use_container_width=True):
        st.session_state["_confirm_clear"] = True
    if st.session_state.get("_confirm_clear"):
        st.warning("This permanently deletes every saved month.")
        cc1, cc2 = st.columns(2)
        if cc1.button("Yes, delete", use_container_width=True):
            shutil.rmtree(DATA_DIR, ignore_errors=True)
            st.session_state["_confirm_clear"] = False
            st.rerun()
        if cc2.button("Cancel", use_container_width=True):
            st.session_state["_confirm_clear"] = False

# ============================================================
# HEADER
# ============================================================

month_meta = get_month_meta(selected_year, selected_month)
source_file = month_meta.get("source_file", "—")
period = month_label(selected_year, selected_month)

st.markdown(
    f"""
    <div class="app-header">
        <p class="eyebrow">Tea Lounge</p>
        <div class="dashboard-title">Tea Lounge Inventory Tracker</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 4 KPI CARDS
# ============================================================

menu_items = month_meta.get("menu_items")
cups_sold = month_meta.get("total_quantity")
total_matched_ingredients = len(df)

out_of_stock = filtered[
    filtered["Month_End_Stock"].notna() & (filtered["Month_End_Stock"] <= 0)
]
oos_names = out_of_stock["Ingredient"].tolist()
if len(oos_names) > 4:
    oos_sub = ", ".join(oos_names[:4]) + f", +{len(oos_names) - 4} more"
elif oos_names:
    oos_sub = ", ".join(oos_names)
else:
    oos_sub = "All ingredients have healthy stock this month."

q1, q2, q3, q4 = st.columns(4)
quad_card(
    q1, "🍵", "Menu Items",
    fmt_num(menu_items) if menu_items is not None else "—",
    "Distinct items — Monthly Item Sales Report",
)
quad_card(
    q2, "🧾", "Cups Sold This Month",
    fmt_num(cups_sold) if cups_sold is not None else "—",
    "Total Quantity — Monthly Item Sales Report",
)
quad_card(
    q3, "🌿", "Matched Ingredients",
    fmt_num(total_matched_ingredients),
    "Ingredients found in this month's inventory report",
)
quad_card(
    q4, "⚠️", "Items Out of Stock",
    fmt_num(len(out_of_stock)),
    oos_sub,
    warn=len(out_of_stock) > 0,
)

if menu_items is None or cups_sold is None:
    st.caption(
        f"ℹ️ No item-sales data is stored for {period} yet, so Menu Items / "
        "Cups Sold show '—'. Upload a workbook that includes a sheet with "
        "'sales' in its name (e.g. 'Monthly Item Sales Report') for this "
        "month — check the sidebar's **🔍 Last upload — sheet-by-sheet** "
        "panel after uploading to see exactly how each sheet was read."
    )

# ============================================================
# SECTION 1 — INGREDIENT USAGE TABLE + SINGLE-INGREDIENT LOOKUP
# Projected Use (column D) is compared against each ingredient's own
# Total Use (column J) — not against other ingredients, since units
# differ across ingredients. Rows missing either value are excluded
# from the table but stay in the detailed report further down.
# ============================================================

st.markdown(
    '<div class="section-title">1. Ingredient Usage — Projected vs Actual</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="small-note">Each ingredient\'s Projected Use is compared '
    "against its own Total Use. Sort any column by clicking its header, or "
    "pick a single ingredient below to see it as a chart.</div>",
    unsafe_allow_html=True,
)

compare_data = filtered.dropna(subset=["Projected_Use", "Total_Use"]).sort_values(
    "Total_Use", ascending=False
).copy()

if compare_data.empty:
    st.info("No ingredients have both a Projected Use and a Total Use value to compare.")
else:
    compare_data["Variance"] = compare_data["Total_Use"] - compare_data["Projected_Use"]
    compare_table = compare_data[
        ["Ingredient", "Category", "Projected_Use", "Total_Use", "Variance"]
    ].rename(
        columns={
            "Ingredient": "Ingredient / Item",
            "Projected_Use": "Projected Use",
            "Total_Use": "Total Use",
            "Variance": "Variance (Actual − Projected)",
        }
    )

    proj_max = compare_table["Projected Use"].max(skipna=True)
    total_max = compare_table["Total Use"].max(skipna=True)
    scale_max = max(
        proj_max if pd.notna(proj_max) else 0,
        total_max if pd.notna(total_max) else 0,
        1,
    )

    st.dataframe(
        compare_table,
        use_container_width=True,
        hide_index=True,
        height=420,
        column_config={
            "Projected Use": st.column_config.ProgressColumn(
                "Projected Use", format="%.2f", min_value=0, max_value=float(scale_max),
            ),
            "Total Use": st.column_config.ProgressColumn(
                "Total Use", format="%.2f", min_value=0, max_value=float(scale_max),
            ),
            "Variance (Actual − Projected)": st.column_config.NumberColumn(
                "Variance (Actual − Projected)", format="%.2f",
            ),
        },
    )

    st.markdown("**🔍 Look up a single ingredient**")
    ingredient_choice = st.selectbox(
        "Select an ingredient to inspect",
        ["— Select an ingredient —"] + sorted(filtered["Ingredient"].dropna().unique().tolist()),
        label_visibility="collapsed",
    )

    if ingredient_choice != "— Select an ingredient —":
        row = filtered[filtered["Ingredient"] == ingredient_choice].iloc[0]
        bar_fig = px.bar(
            x=["Projected Use", "Total Use"],
            y=[row["Projected_Use"], row["Total_Use"]],
            color=["Projected Use", "Total Use"],
            color_discrete_map={"Projected Use": "#38d996", "Total Use": "#0b6e4f"},
            labels={"x": "", "y": "Quantity"},
            text=[fmt_num(row["Projected_Use"]), fmt_num(row["Total_Use"])],
        )
        bar_fig.update_traces(textposition="outside", cliponaxis=False)
        bar_fig.update_layout(
            height=340,
            showlegend=False,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color=T["chart_font"],
            margin=dict(l=10, r=10, t=30, b=10),
            title=f"{ingredient_choice}: Projected vs Actual Use",
        )
        col_chart, col_metrics = st.columns([2, 1])
        col_chart.plotly_chart(bar_fig, use_container_width=True)
        with col_metrics:
            st.metric("Total Stock", fmt_num(row.get("Total_Stock")))
            st.metric("Month End Stock", fmt_num(row.get("Month_End_Stock")))
            st.metric("Received (Total)", fmt_num(row.get("Received_Total")))

# ============================================================
# SECTION 2 — CATEGORY BREAKDOWN (PIE)
# ============================================================

st.markdown(
    '<div class="section-title">2. Category Breakdown</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="small-note">Share of Total Use across ingredient categories.</div>',
    unsafe_allow_html=True,
)

cat_summary = (
    filtered.groupby("Category")["Total_Use"].sum(min_count=1).reset_index().dropna()
)
if cat_summary.empty:
    st.info("No Total Use data available to break down by category.")
else:
    donut_fig = px.pie(
        cat_summary,
        names="Category",
        values="Total_Use",
        hole=0.62,
        color_discrete_sequence=px.colors.sequential.Greens_r,
    )
    donut_fig.update_traces(
        textinfo="percent+label",
        textposition="outside",
        marker=dict(line=dict(color="#f2fbf6", width=2)),
    )
    donut_fig.update_layout(
        height=380,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        font_color=T["chart_font"],
        showlegend=True,
        annotations=[dict(
            text=f"<b>{fmt_num(cat_summary['Total_Use'].sum())}</b><br>Total Use",
            x=0.5, y=0.5, showarrow=False, font_size=14, font_color=T["chart_font"],
        )],
    )
    st.plotly_chart(donut_fig, use_container_width=True)

# ============================================================
# SECTION 3 — STOCK MOVEMENT (TOP ITEMS BY TOTAL STOCK)
# ============================================================

st.markdown(
    '<div class="section-title">3. Stock Movement — Top Items</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="small-note">Pre-month stock, quantity received and month-end '
    "stock for the items with the largest total stock.</div>",
    unsafe_allow_html=True,
)

top_stock = (
    filtered.dropna(subset=["Total_Stock"])
    .sort_values("Total_Stock", ascending=False)
    .head(12)
)

if top_stock.empty:
    st.info("No Total Stock data available for the current filters.")
else:
    fig_stock = px.bar(
        top_stock,
        x="Ingredient",
        y=["Pre_Month_Stock", "Received_Total", "Month_End_Stock"],
        barmode="group",
        labels={"value": "Quantity", "Ingredient": "", "variable": "Metric"},
        color_discrete_map={
            "Pre_Month_Stock": "#a7e8c8",
            "Received_Total": "#2fbf87",
            "Month_End_Stock": "#0b6e4f",
        },
    )
    newnames = {
        "Pre_Month_Stock": "Pre. Month Stock",
        "Received_Total": "Received (Total)",
        "Month_End_Stock": "Month End Stock",
    }
    fig_stock.for_each_trace(lambda t: t.update(name=newnames.get(t.name, t.name)))
    fig_stock.update_layout(
        height=450,
        margin=dict(l=10, r=20, t=20, b=100),
        xaxis_tickangle=-45,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font_color=T["chart_font"],
        legend_title_text="",
    )
    st.plotly_chart(fig_stock, use_container_width=True)

# ============================================================
# SECTION 4 — DETAILED REPORT
# (No matplotlib anywhere — ProgressColumn gives the same "highlighted
# column" effect natively, and it can't throw an ImportError.)
# ============================================================

st.markdown(
    '<div class="section-title">4. Detailed Inventory Report</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="small-note">Projected Use and Total Use are highlighted below.</div>',
    unsafe_allow_html=True,
)

detail_cols = [
    "Ingredient", "Category", "Projected_Use", "Total_Use",
    "Total_Units_Sold", "Total_Required_Qty", "Pre_Month_Stock",
    "Received_1", "Received_2", "Total_Stock", "Month_End_Stock",
]
detail_cols = [c for c in detail_cols if c in filtered.columns]
detail = filtered[detail_cols].copy()

rename_detail = {
    "Ingredient": "Ingredient / Item",
    "Total_Units_Sold": "Total Units Sold",
    "Total_Required_Qty": "Total Required Qty",
    "Projected_Use": "Projected Use",
    "Pre_Month_Stock": "Pre. Month Stock",
    "Received_1": "Received 1",
    "Received_2": "Received 2",
    "Total_Stock": "Total Stock",
    "Month_End_Stock": "Month End Stock",
    "Total_Use": "Total Use",
}
detail = detail.rename(columns=rename_detail)

proj_max_d = detail["Projected Use"].max(skipna=True) if "Projected Use" in detail else np.nan
total_max_d = detail["Total Use"].max(skipna=True) if "Total Use" in detail else np.nan
scale_max_d = max(
    proj_max_d if pd.notna(proj_max_d) else 0,
    total_max_d if pd.notna(total_max_d) else 0,
    1,
)

detail_column_config = {}
if "Projected Use" in detail.columns:
    detail_column_config["Projected Use"] = st.column_config.ProgressColumn(
        "Projected Use", format="%.2f", min_value=0, max_value=float(scale_max_d),
    )
if "Total Use" in detail.columns:
    detail_column_config["Total Use"] = st.column_config.ProgressColumn(
        "Total Use", format="%.2f", min_value=0, max_value=float(scale_max_d),
    )

st.dataframe(
    detail,
    use_container_width=True,
    hide_index=True,
    height=520,
    column_config=detail_column_config,
)

# ============================================================
# EXPORT
# ============================================================

st.markdown('<div class="section-title">5. Export</div>', unsafe_allow_html=True)

export_df = filtered.copy()
csv_data = export_df.to_csv(index=False).encode("utf-8-sig")

st.download_button(
    "⬇ Download Filtered Report (CSV)",
    data=csv_data,
    file_name=f"Tea_Lounge_Inventory_{month_key(selected_year, selected_month)}.csv",
    mime="text/csv",
)

# ============================================================
# FOOTER
# ============================================================

st.markdown("---")
st.caption(
    "Tea Lounge Inventory Tracker • Upload a new monthly workbook at any time — "
    "past months stay saved in tea_lounge_store/ and remain browsable from the "
    "sidebar's Year / Month selectors."
)
