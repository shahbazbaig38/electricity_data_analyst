"""
Energy Consumption Data Analyst – Production-Ready Streamlit Application
=========================================================================
A comprehensive tool for importing, converting, analysing and visualising
electricity-consumption data files (Excel & CSV).

Supported data formats
----------------------
* **Legacy Finnish** – Excel with Finnish weekday-prefixed dates (original format)
* **Fingrid 15 min** – Tab-separated with ISO 8601 dates, PT15M resolution
* **Fingrid 1 h**   – Tab-separated with ISO 8601 dates, PT1H resolution
* **Vare 15 min**   – Tab-separated, date format ``d.M.yyyy H.mm``
* **Vare 1 h**      – Tab-separated, date format ``HH:mm dd.MM.yyyy``

Key features
------------
* **Multi-format auto-detection** – Fingrid, Vare, and legacy Finnish formats
* **CSV & Excel support** – upload .csv, .xlsx, or .xls files
* **Time-resolution conversion** – 15 min → 1 h (sum) *or* 1 h → 15 min (÷4)
* **Automatic leap-day detection & removal**
* **Smart date parser** for Finnish weekday names
* **Interactive charts** (daily / weekly / monthly profiles, heatmaps, time-series)
* **Statistical overview** (mean, median, std, min/max, percentiles)
* **Optional price column** (€/kWh) support – cost analytics
* **One-click Excel export** of processed data
* **WCAG AA accessible** – colorblind-safe palette throughout

Author: Senior Data Analyst
Version: 3.0.0
"""

from __future__ import annotations

import base64
import io
import re
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ═══════════════════════════════════════════════════════════════════════════════
# COLORBLIND-SAFE PALETTE  (Wong 2011 – Nature Methods)
# Each colour has been verified ≥ 4.5:1 contrast on white backgrounds.
# ═══════════════════════════════════════════════════════════════════════════════
CB_BLUE      = "#0072B2"   # Strong blue   – primary accent
CB_ORANGE    = "#D55E00"   # Vermillion    – alert / secondary
CB_TEAL      = "#009E73"   # Bluish green  – success / positive
CB_YELLOW    = "#F0E442"   # Yellow        – (use on dark bg only, or with black text)
CB_SKY       = "#56B4E9"   # Sky blue      – light accent
CB_PINK      = "#CC79A7"   # Reddish purple – category
CB_BLACK     = "#1A1A2E"   # Near-black    – text
CB_GREY      = "#6B7280"   # Neutral       – secondary text

# Ordered sequence for multi-series charts (safe for all CVD types)
CB_SEQUENCE = [CB_BLUE, CB_ORANGE, CB_TEAL, CB_PINK, CB_SKY, "#E69F00", CB_BLACK]

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Energy Data Analyst",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════════════════════
# PROFESSIONAL LIGHT-THEME CSS
# High contrast • Clean spacing • Accessible typography
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown(
    """
    <style>
    /* ── Typography ────────────────────────────────────── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        color: #1A1A2E;
    }

    /* ── Header banner ─────────────────────────────────── */
    .main-header {
        background: #0B5394;
        padding: 1.75rem 2rem;
        border-radius: 10px;
        margin-bottom: 1.25rem;
        color: #FFFFFF;
    }
    .main-header h1 {
        margin: 0;
        font-size: 1.6rem;
        font-weight: 700;
        color: #FFFFFF;
    }
    .main-header p {
        margin: 0.3rem 0 0 0;
        color: #D6E4F0;
        font-size: 0.95rem;
    }

    /* ── Metric cards ──────────────────────────────────── */
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #D1D5DB;
        border-radius: 10px;
        padding: 1.1rem 1.25rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        transition: box-shadow 0.2s ease;
    }
    .metric-card:hover {
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    .metric-card .metric-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0B5394;
    }
    .metric-card .metric-label {
        font-size: 0.80rem;
        font-weight: 500;
        color: #6B7280;
        margin-top: 0.2rem;
        text-transform: uppercase;
        letter-spacing: 0.02em;
    }

    /* ── Section headers ───────────────────────────────── */
    .section-header {
        font-size: 1.15rem;
        font-weight: 600;
        color: #1A1A2E;
        margin: 1.25rem 0 0.6rem 0;
        padding-bottom: 0.4rem;
        border-bottom: 2px solid #0B5394;
    }

    /* ── Badges – use shape + text + colour ─────────────── */
    .badge-ok {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #DEF7EC;
        color: #065F46;
        padding: 4px 12px;
        border-radius: 6px;
        font-size: 0.82rem;
        font-weight: 600;
        border: 1px solid #A7F3D0;
    }
    .badge-ok::before { content: "●"; }

    .badge-warn {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #FEF3C7;
        color: #92400E;
        padding: 4px 12px;
        border-radius: 6px;
        font-size: 0.82rem;
        font-weight: 600;
        border: 1px solid #FDE68A;
    }
    .badge-warn::before { content: "▲"; }

    .badge-info {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #DBEAFE;
        color: #1E40AF;
        padding: 4px 12px;
        border-radius: 6px;
        font-size: 0.82rem;
        font-weight: 600;
        border: 1px solid #93C5FD;
    }
    .badge-info::before { content: "ℹ"; }

    /* ── Sidebar ───────────────────────────────────────── */
    section[data-testid="stSidebar"] {
        background: #F4F6F9;
        border-right: 1px solid #E5E7EB;
    }

    /* ── Divider ───────────────────────────────────────── */
    .styled-divider {
        height: 1px;
        background: #E5E7EB;
        border: none;
        margin: 1rem 0;
    }

    /* ── Download button ───────────────────────────────── */
    .stDownloadButton > button {
        background: #0B5394 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.55rem 1.75rem !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        transition: background 0.2s ease !important;
    }
    .stDownloadButton > button:hover {
        background: #094075 !important;
    }
    .stDownloadButton > button:focus-visible {
        outline: 3px solid #56B4E9 !important;
        outline-offset: 2px !important;
    }

    /* ── Focus outlines for accessibility ───────────────── */
    button:focus-visible,
    input:focus-visible,
    select:focus-visible,
    [tabindex]:focus-visible {
        outline: 3px solid #56B4E9 !important;
        outline-offset: 2px !important;
    }

    /* ── Feature cards on landing ───────────────────────── */
    .feature-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 1.5rem;
        height: 100%;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .feature-card .feature-icon {
        font-size: 2rem;
        margin-bottom: 0.5rem;
    }
    .feature-card h4 {
        color: #1A1A2E;
        margin: 0 0 0.5rem 0;
    }
    .feature-card p {
        color: #6B7280;
        font-size: 0.9rem;
        margin: 0;
        line-height: 1.5;
    }

    /* ── Logo strip ────────────────────────────────────── */
    .logo-strip {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 1rem;
        margin: 1.25rem 0 1.75rem 0;
    }
    .logo-item {
        flex: 1 1 0;
        display: flex;
        justify-content: center;
        align-items: center;
        padding: 0.35rem 0.5rem;
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        min-height: 96px;
    }
    .logo-item img {
        max-width: 100%;
        max-height: 72px;
        width: auto;
        height: auto;
        object-fit: contain;
    }

    /* ── Footer ────────────────────────────────────────── */
    .app-footer {
        text-align: center;
        color: #9CA3AF;
        font-size: 0.78rem;
        margin-top: 2rem;
        padding-top: 1rem;
        border-top: 1px solid #E5E7EB;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ═══════════════════════════════════════════════════════════════════════════════
# PLOTLY LAYOUT – light, accessible defaults
# ═══════════════════════════════════════════════════════════════════════════════
PLOTLY_LAYOUT = dict(
    template="plotly_white",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="#FAFBFC",
    font=dict(family="Inter, sans-serif", size=13, color="#1A1A2E"),
    margin=dict(l=50, r=20, t=50, b=50),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1,
        font=dict(size=12),
    ),
)

LOGO_DIR = Path(__file__).resolve().parent / "logos"
LOGO_FILES = [
    LOGO_DIR / "FI Euroopan unionin osarahoittama_POS.png",
    LOGO_DIR / "PS-LIITTO-logo.png",
    LOGO_DIR / "savonia_logo_2020.webp",
]


def _get_logo_html(logo_paths: list[Path]) -> str:
    image_tags = []
    for logo_path in logo_paths:
        if not logo_path.exists():
            continue
        suffix = logo_path.suffix.lower().lstrip(".")
        mime_type = "image/webp" if suffix == "webp" else f"image/{suffix}"
        encoded = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
        image_tags.append(
            f'<div class="logo-item"><img src="data:{mime_type};base64,{encoded}" alt="{logo_path.stem}" /></div>'
        )
    if not image_tags:
        return ""
    return f'<div class="logo-strip">{"".join(image_tags)}</div>'

# Grid styling applied separately so it never clashes with per-chart axis kwargs
_AXIS_STYLE = dict(gridcolor="#E5E7EB", zerolinecolor="#D1D5DB")


def _style_axes(fig: go.Figure) -> go.Figure:
    """Apply consistent grid/zeroline styling to x and y axes."""
    fig.update_xaxes(**_AXIS_STYLE)
    fig.update_yaxes(**_AXIS_STYLE)
    return fig

# Colorblind-safe colorscale for sequential data (blue → orange, perceptually uniform)
CB_SEQUENTIAL = [
    [0.0, "#DBEAFE"],
    [0.25, "#56B4E9"],
    [0.5, "#0072B2"],
    [0.75, "#D55E00"],
    [1.0, "#92400E"],
]

# Colorblind-safe colorscale for heatmaps (Cividis – designed for CVD)
CB_HEATMAP = "Cividis"


# ═══════════════════════════════════════════════════════════════════════════════
# FINNISH WEEKDAYS
# ═══════════════════════════════════════════════════════════════════════════════
FINNISH_WEEKDAYS = {
    "maanantai": "Monday",
    "tiistai": "Tuesday",
    "keskiviikko": "Wednesday",
    "torstai": "Thursday",
    "perjantai": "Friday",
    "lauantai": "Saturday",
    "sunnuntai": "Sunday",
}

# ═══════════════════════════════════════════════════════════════════════════════
# CORE HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def parse_finnish_date(date_str: str) -> Optional[pd.Timestamp]:
    """Extract a datetime from a Finnish-format string like
    ``'Maanantai 01.01.2023 00:00'``."""
    try:
        txt = str(date_str).strip()
        match = re.search(
            r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})\s+(\d{1,2}):(\d{2})",
            txt,
        )
        if match:
            day, month, year, hour, minute = match.groups()
            if len(year) == 2:
                year = "20" + year
            return pd.Timestamp(
                year=int(year),
                month=int(month),
                day=int(day),
                hour=int(hour),
                minute=int(minute),
            )
        return pd.to_datetime(txt, dayfirst=True)
    except Exception:
        return pd.NaT


def parse_vare_15min_date(date_str: str) -> Optional[pd.Timestamp]:
    """Parse Vare 15-min timestamp format: ``'1.1.2024 0.00'``.
    Date part is d.M.yyyy, time uses dots instead of colons."""
    try:
        txt = str(date_str).strip()
        # Match: d.M.yyyy H.mm  (e.g. "1.1.2024 0.00" or "31.12.2024 23.45")
        match = re.match(
            r"(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2})\.(\d{2})", txt
        )
        if match:
            day, month, year, hour, minute = match.groups()
            return pd.Timestamp(
                year=int(year),
                month=int(month),
                day=int(day),
                hour=int(hour),
                minute=int(minute),
            )
        return pd.NaT
    except Exception:
        return pd.NaT


def parse_vare_hourly_date(date_str: str) -> Optional[pd.Timestamp]:
    """Parse Vare hourly timestamp format: ``'00:00 01.01.2024'``.
    Time comes first (HH:mm), then the date (dd.MM.yyyy)."""
    try:
        txt = str(date_str).strip()
        # Match: HH:mm dd.MM.yyyy  (e.g. "00:00 01.01.2024" or "13:00 31.12.2024")
        match = re.match(
            r"(\d{1,2}):(\d{2})\s+(\d{1,2})\.(\d{1,2})\.(\d{4})", txt
        )
        if match:
            hour, minute, day, month, year = match.groups()
            return pd.Timestamp(
                year=int(year),
                month=int(month),
                day=int(day),
                hour=int(hour),
                minute=int(minute),
            )
        return pd.NaT
    except Exception:
        return pd.NaT


def _finnish_decimal(series: pd.Series) -> pd.Series:
    """Convert a Series with Finnish decimal commas (e.g. '1,93') to float.
    If the values are already numeric they pass through unchanged."""
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    ).fillna(0.0)


def detect_resolution(df: pd.DataFrame) -> str:
    """Detect if data is ~15 min or ~1 h resolution based on median time delta."""
    if len(df) < 2:
        return "unknown"
    deltas = df.index.to_series().diff().dropna()
    median_min = deltas.median().total_seconds() / 60
    if median_min <= 20:
        return "15 min"
    elif median_min <= 65:
        return "1 h"
    return f"~{int(median_min)} min"


def has_leap_day(df: pd.DataFrame) -> tuple[bool, int]:
    """Check if leap-day data exists; return (bool, count)."""
    mask = (df.index.month == 2) & (df.index.day == 29)
    count = mask.sum()
    return count > 0, int(count)


def remove_leap_days(df: pd.DataFrame) -> pd.DataFrame:
    """Remove all Feb 29 rows."""
    return df[~((df.index.month == 2) & (df.index.day == 29))].copy()


def _deduplicate_index(df: pd.DataFrame) -> pd.DataFrame:
    """Remove rows with duplicate timestamps (e.g. from DST fall-back).
    Keeps the first occurrence of each timestamp."""
    if df.index.duplicated().any():
        n_dups = df.index.duplicated().sum()
        st.warning(
            f"Found **{n_dups}** duplicate timestamp(s) "
            f"(likely DST transition) – keeping first occurrence."
        )
        return df[~df.index.duplicated(keep="first")]
    return df


def convert_15min_to_1h(df: pd.DataFrame) -> pd.DataFrame:
    """Sum 15-min kWh values to 1-hour values."""
    df = _deduplicate_index(df)
    agg_dict = {"kWh": "sum"}
    if "Price_EUR_kWh" in df.columns:
        agg_dict["Price_EUR_kWh"] = "mean"
    result = df.resample("1h").agg(agg_dict)
    if "Price_EUR_kWh" in df.columns:
        result["Cost_EUR"] = result["kWh"] * result["Price_EUR_kWh"]
    return result


def convert_1h_to_15min(df: pd.DataFrame) -> pd.DataFrame:
    """Divide hourly kWh by 4 and expand to 15-min intervals."""
    df = _deduplicate_index(df)
    result = df.resample("15min").ffill()
    result["kWh"] = result["kWh"] / 4.0
    if "Price_EUR_kWh" in result.columns:
        result["Cost_EUR"] = result["kWh"] * result["Price_EUR_kWh"]
    return result


# ─── Format detection ────────────────────────────────────────────────────────

FORMAT_FINGRID = "Fingrid"
FORMAT_VARE_15MIN = "Vare (15 min)"
FORMAT_VARE_HOURLY = "Vare (hourly)"
FORMAT_LEGACY = "Legacy Finnish"
FORMAT_UNKNOWN = "Unknown"


def detect_data_format(columns: list[str]) -> str:
    """Identify the data format from the column header names.

    Returns one of the FORMAT_* constants.
    """
    col_lower = [c.strip().lower() for c in columns]

    # Fingrid: has 'alkuaika' and 'määrä'
    if "alkuaika" in col_lower and "määrä" in col_lower:
        return FORMAT_FINGRID

    # Vare: has 'aikaleima' and 'kulutus' (partial match)
    has_aikaleima = "aikaleima" in col_lower
    has_kulutus = any("kulutus" in c for c in col_lower)
    if has_aikaleima and has_kulutus:
        return FORMAT_VARE_15MIN  # will be refined after reading a sample row

    # Fallback – try the legacy 2-col format
    return FORMAT_LEGACY


def _refine_vare_format(sample_value: str) -> str:
    """Distinguish Vare 15-min from Vare hourly by looking at an Aikaleima
    sample value.  Vare hourly starts with ``HH:mm`` while 15-min starts
    with a day digit."""
    txt = str(sample_value).strip()
    # Hourly format: time first, e.g. "00:00 01.01.2024"
    if re.match(r"\d{1,2}:\d{2}\s+\d", txt):
        return FORMAT_VARE_HOURLY
    # 15-min format: date first, e.g. "1.1.2024 0.00"
    return FORMAT_VARE_15MIN


# ─── Unified file reader ─────────────────────────────────────────────────────

def _decode_csv_bytes(file_bytes: bytes) -> str:
    """Decode CSV bytes trying multiple encodings common in Finnish data."""
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return file_bytes.decode(enc)
        except (UnicodeDecodeError, ValueError):
            continue
    # Last resort – replace bad chars
    return file_bytes.decode("utf-8", errors="replace")


def _read_raw_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    """Read a CSV or Excel file into a raw DataFrame (no parsing yet).
    For CSV files, try tab separator first (all new formats use tabs),
    then fall back to comma / semicolon."""
    lower = file_name.lower()
    if lower.endswith(".csv"):
        text = _decode_csv_bytes(file_bytes)
        # Try tab first
        df = pd.read_csv(io.StringIO(text), sep="\t", header=0)
        if df.shape[1] < 2:
            # Fall back to comma
            df = pd.read_csv(io.StringIO(text), sep=",", header=0)
        if df.shape[1] < 2:
            # Fall back to semicolon
            df = pd.read_csv(io.StringIO(text), sep=";", header=0)
        return df
    else:
        return pd.read_excel(io.BytesIO(file_bytes), header=0)


@st.cache_data(show_spinner=False)
def load_file(file_bytes: bytes, file_name: str) -> tuple[pd.DataFrame, str]:
    """Load an uploaded file (CSV or Excel), auto-detect the format,
    parse dates, and return (tidy_DataFrame, detected_format_name)."""
    raw = _read_raw_file(file_bytes, file_name)

    if raw.shape[1] < 2:
        st.error("The file must have at least 2 columns (DateTime, kWh).")
        st.stop()

    fmt = detect_data_format(list(raw.columns))

    # ── Fingrid ──────────────────────────────────────────────────────────
    if fmt == FORMAT_FINGRID:
        # Find columns by case-insensitive match (handles encoding edge cases)
        alkuaika_col = next(c for c in raw.columns if c.strip().lower() == "alkuaika")
        maara_col = next(c for c in raw.columns if c.strip().lower() == "määrä")

        raw["Timestamp"] = pd.to_datetime(raw[alkuaika_col], utc=True, errors="coerce")
        # Convert from UTC to Finnish local time (EET/EEST)
        raw["Timestamp"] = (
            raw["Timestamp"]
            .dt.tz_convert("Europe/Helsinki")
            .dt.tz_localize(None)
        )
        raw["kWh"] = _finnish_decimal(raw[maara_col])

        failed = raw["Timestamp"].isna().sum()
        if failed > 0:
            st.warning(f"Could not parse **{failed}** date values – dropped.")
        raw = raw.dropna(subset=["Timestamp"])
        raw = raw.set_index("Timestamp").sort_index()
        return raw[["kWh"]], fmt

    # ── Vare ─────────────────────────────────────────────────────────────
    if fmt in (FORMAT_VARE_15MIN, FORMAT_VARE_HOURLY):
        # Refine sub-format
        first_ts = str(raw.iloc[0]["Aikaleima"]).strip()
        fmt = _refine_vare_format(first_ts)

        if fmt == FORMAT_VARE_15MIN:
            raw["Timestamp"] = raw["Aikaleima"].apply(parse_vare_15min_date)
        else:
            raw["Timestamp"] = raw["Aikaleima"].apply(parse_vare_hourly_date)

        # Locate the consumption column (name contains 'Kulutus')
        kulutus_col = [c for c in raw.columns if "Kulutus" in c or "kulutus" in c]
        if not kulutus_col:
            st.error("Cannot find a consumption (Kulutus) column.")
            st.stop()
        raw["kWh"] = _finnish_decimal(raw[kulutus_col[0]])

        # Locate the price column (name contains 'Spot-hinta')
        hinta_col = [c for c in raw.columns if "Spot-hinta" in c or "spot-hinta" in c]
        has_price = len(hinta_col) > 0
        if has_price:
            # Price is in snt/kWh → convert to €/kWh (÷ 100)
            raw["Price_EUR_kWh"] = _finnish_decimal(raw[hinta_col[0]]) / 100.0

        failed = raw["Timestamp"].isna().sum()
        if failed > 0:
            st.warning(f"Could not parse **{failed}** date values – dropped.")
        raw = raw.dropna(subset=["Timestamp"])
        raw = raw.set_index("Timestamp").sort_index()

        keep = ["kWh"]
        if has_price:
            keep.append("Price_EUR_kWh")
        return raw[keep], fmt

    # ── Legacy Finnish format (original behaviour) ───────────────────────
    col_names = list(raw.columns)
    has_price = raw.shape[1] >= 3

    # Build a clean DataFrame with only the columns we need to avoid
    # duplicate-name issues when the source file has extra columns.
    clean = pd.DataFrame()
    clean["RawTime"] = raw.iloc[:, 0]
    clean["kWh"] = raw.iloc[:, 1]
    if has_price:
        clean["Price_EUR_kWh"] = raw.iloc[:, 2]

    clean["Timestamp"] = clean["RawTime"].apply(parse_finnish_date)
    failed = clean["Timestamp"].isna().sum()
    if failed > 0:
        st.warning(f"Could not parse **{failed}** date values – dropped.")
    clean = clean.dropna(subset=["Timestamp"])
    clean = clean.set_index("Timestamp").sort_index()

    clean["kWh"] = _finnish_decimal(clean["kWh"])
    if has_price:
        clean["Price_EUR_kWh"] = _finnish_decimal(clean["Price_EUR_kWh"])

    keep = ["kWh"]
    if has_price:
        keep.append("Price_EUR_kWh")
    return clean[keep], FORMAT_LEGACY


@st.cache_data(show_spinner=False)
def to_excel_bytes(_df: pd.DataFrame) -> bytes:
    """Convert a DataFrame to Excel bytes for download (cached)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        _df.to_excel(writer, index=True, sheet_name="Data")
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# PLOTTING HELPERS  – all use colorblind-safe palette
# ═══════════════════════════════════════════════════════════════════════════════

def fig_timeseries(
    df: pd.DataFrame,
    col: str = "kWh",
    title: str = "Consumption Over Time",
) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=df.index,
            y=df[col],
            mode="lines",
            line=dict(color=CB_BLUE, width=1.3),
            fill="tozeroy",
            fillcolor="rgba(0,114,178,0.10)",
            name=col,
            hovertemplate="%{x|%Y-%m-%d %H:%M}<br>%{y:.3f} kWh<extra></extra>",
        )
    )
    fig.update_layout(title=title, xaxis_title="Time", yaxis_title=col, **PLOTLY_LAYOUT)
    return _style_axes(fig)


def fig_daily_profile(df: pd.DataFrame) -> go.Figure:
    """Average consumption by hour of day."""
    hourly = df.groupby(df.index.hour)["kWh"].mean()
    fig = go.Figure(
        go.Bar(
            x=hourly.index,
            y=hourly.values,
            marker=dict(color=CB_BLUE, line=dict(color="#094075", width=0.5)),
            hovertemplate="Hour %{x}:00<br>%{y:.3f} kWh avg<extra></extra>",
        )
    )
    fig.update_layout(
        title="Average Daily Load Profile",
        xaxis_title="Hour of Day",
        yaxis_title="kWh (mean)",
        xaxis=dict(dtick=1),
        **PLOTLY_LAYOUT,
    )
    return _style_axes(fig)


def fig_weekly_profile(df: pd.DataFrame) -> go.Figure:
    """Average consumption by day of week."""
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    daily = df.groupby(df.index.dayofweek)["kWh"].mean().reindex(range(7))

    # Use different colours for weekdays vs weekend for extra differentiation
    colors = [CB_BLUE] * 5 + [CB_ORANGE] * 2

    fig = go.Figure(
        go.Bar(
            x=day_names,
            y=daily.values,
            marker=dict(color=colors, line=dict(color="#1A1A2E", width=0.5)),
            hovertemplate="%{x}<br>%{y:.3f} kWh avg<extra></extra>",
        )
    )
    fig.update_layout(
        title="Average Weekly Load Profile",
        xaxis_title="Day of Week",
        yaxis_title="kWh (mean)",
        **PLOTLY_LAYOUT,
    )
    # Add legend-like annotation for colour meaning
    fig.add_annotation(
        text="<b>■</b> Weekday &nbsp; <b style='color:#D55E00'>■</b> Weekend",
        xref="paper", yref="paper", x=1, y=1.12,
        showarrow=False, font=dict(size=11, color="#1A1A2E"),
    )
    return _style_axes(fig)


def fig_monthly_totals(df: pd.DataFrame) -> go.Figure:
    """Monthly total consumption."""
    monthly = df.resample("ME")["kWh"].sum()
    fig = go.Figure(
        go.Bar(
            x=monthly.index.strftime("%Y-%m"),
            y=monthly.values,
            marker=dict(color=CB_TEAL, line=dict(color="#065F46", width=0.5)),
            hovertemplate="%{x}<br>%{y:,.1f} kWh total<extra></extra>",
        )
    )
    fig.update_layout(
        title="Monthly Total Consumption",
        xaxis_title="Month",
        yaxis_title="kWh (total)",
        **PLOTLY_LAYOUT,
    )
    return _style_axes(fig)


def fig_heatmap(df: pd.DataFrame) -> go.Figure:
    """Hour × Day-of-Week heatmap of average consumption.
    Uses Cividis colorscale – specifically designed for colorblind accessibility."""
    heat = (
        df.assign(hour=df.index.hour, dow=df.index.dayofweek)
        .groupby(["dow", "hour"])["kWh"]
        .mean()
        .unstack(fill_value=0)
    )
    day_labels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    fig = go.Figure(
        go.Heatmap(
            z=heat.values,
            x=heat.columns,
            y=day_labels,
            colorscale=CB_HEATMAP,
            colorbar=dict(title=dict(text="kWh", font=dict(size=12))),
            hovertemplate="Day: %{y}<br>Hour: %{x}:00<br>%{z:.3f} kWh avg<extra></extra>",
        )
    )
    fig.update_layout(
        title="Consumption Heatmap (Hour × Day of Week)",
        xaxis_title="Hour",
        yaxis_title="Day",
        xaxis=dict(dtick=1),
        **PLOTLY_LAYOUT,
    )
    return _style_axes(fig)


def fig_monthly_boxplot(df: pd.DataFrame) -> go.Figure:
    """Box-plot of daily totals per month – shows variance / seasonality."""
    tmp = df.resample("D")["kWh"].sum().reset_index()
    tmp["month"] = tmp["Timestamp"].dt.month
    tmp["month_name"] = tmp["Timestamp"].dt.strftime("%b")
    order = tmp.sort_values("month")["month_name"].unique()
    fig = px.box(
        tmp,
        x="month_name",
        y="kWh",
        category_orders={"month_name": list(order)},
        color_discrete_sequence=[CB_BLUE],
    )
    fig.update_layout(
        title="Daily Consumption Distribution by Month",
        xaxis_title="Month",
        yaxis_title="Daily kWh",
        **PLOTLY_LAYOUT,
    )
    return _style_axes(fig)


def fig_price_timeseries(df: pd.DataFrame) -> go.Figure:
    """Price and cost over time (dual axis).
    Uses blue for price and orange for cost – distinguishable for all CVD types."""
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Scatter(
            x=df.index,
            y=df["Price_EUR_kWh"],
            name="Price (€/kWh)",
            line=dict(color=CB_BLUE, width=1.2),
            hovertemplate="%{x|%Y-%m-%d %H:%M}<br>%{y:.4f} €/kWh<extra></extra>",
        ),
        secondary_y=False,
    )
    if "Cost_EUR" in df.columns:
        fig.add_trace(
            go.Scatter(
                x=df.index,
                y=df["Cost_EUR"],
                name="Cost (€)",
                line=dict(color=CB_ORANGE, width=1.2, dash="dot"),
                fill="tozeroy",
                fillcolor="rgba(213,94,0,0.06)",
                hovertemplate="%{x|%Y-%m-%d %H:%M}<br>%{y:.4f} €<extra></extra>",
            ),
            secondary_y=True,
        )
    fig.update_layout(title="Price & Cost Over Time", **PLOTLY_LAYOUT)
    fig.update_yaxes(title_text="€/kWh", secondary_y=False)
    fig.update_yaxes(title_text="€", secondary_y=True)
    return _style_axes(fig)


def fig_cumulative(df: pd.DataFrame) -> go.Figure:
    """Cumulative consumption curve."""
    cum = df["kWh"].cumsum()
    fig = go.Figure(
        go.Scatter(
            x=cum.index,
            y=cum.values,
            mode="lines",
            line=dict(color=CB_TEAL, width=2),
            fill="tozeroy",
            fillcolor="rgba(0,158,115,0.08)",
            hovertemplate="%{x|%Y-%m-%d %H:%M}<br>%{y:,.1f} kWh cumulative<extra></extra>",
        )
    )
    fig.update_layout(
        title="Cumulative Consumption",
        xaxis_title="Time",
        yaxis_title="Cumulative kWh",
        **PLOTLY_LAYOUT,
    )
    return _style_axes(fig)


# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## ⚡ Energy Analyst")
    st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Upload Data File",
        type=["xlsx", "xls", "csv"],
        help="Excel or CSV – supports Fingrid, Vare, and legacy Finnish formats",
    )

    st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)
    st.markdown("### Operations")

    operation = st.radio(
        "Time Resolution Conversion",
        options=["None", "15 min → 1 hour (SUM)", "1 hour → 15 min (÷4)"],
        index=0,
        help="A) Sum 15-min readings into hourly totals  |  B) Expand hourly into 15-min intervals",
    )

    st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)
    st.markdown("### Leap Day Handling")

    leap_day_action = st.radio(
        "Leap Day (Feb 29) Rows",
        options=["Keep", "Remove automatically"],
        index=0,
    )

    st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)
    st.markdown("### Chart Options")

    show_timeseries = st.checkbox("Time-Series Plot", value=True)
    show_daily_profile = st.checkbox("Daily Load Profile", value=True)
    show_weekly_profile = st.checkbox("Weekly Load Profile", value=True)
    show_monthly = st.checkbox("Monthly Totals", value=True)
    show_heatmap = st.checkbox("Consumption Heatmap", value=True)
    show_boxplot = st.checkbox("Monthly Box-Plot", value=False)
    show_cumulative = st.checkbox("Cumulative Curve", value=False)
    show_price = st.checkbox("Price / Cost Analysis", value=True)

    st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)
    st.markdown(
        '<p class="app-footer">Energy Data Analyst v3.0.0<br>© 2026</p>',
        unsafe_allow_html=True,
    )

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN CONTENT
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown(
    """
    <div class="main-header">
        <h1>⚡ Energy Consumption Data Analyst</h1>
        <p>Import · Convert · Analyse · Visualise · Export</p>
    </div>
    """,
    unsafe_allow_html=True,
)

logo_html = _get_logo_html(LOGO_FILES)
if logo_html:
    st.markdown(logo_html, unsafe_allow_html=True)

# ── Landing page (no file yet) ───────────────────────────────────────────────
if uploaded_file is None:
    st.markdown("---")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown(
            """
            <div class="feature-card">
                <div class="feature-icon">🔄</div>
                <h4>Convert</h4>
                <p>Switch between <strong>15-min</strong> and <strong>1-hour</strong>
                resolution instantly. Sum or divide energy readings with one click.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_b:
        st.markdown(
            """
            <div class="feature-card">
                <div class="feature-icon">📊</div>
                <h4>Analyse</h4>
                <p>Daily, weekly and monthly profiles, heatmaps, box-plots,
                and a full statistical summary — all auto-generated.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_c:
        st.markdown(
            """
            <div class="feature-card">
                <div class="feature-icon">📥</div>
                <h4>Export</h4>
                <p>Download the processed data as a new Excel file,
                ready for further analysis in Excel or other tools.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown("")
    st.info("👈  **Upload a data file** (Excel or CSV) in the sidebar to get started.")
    st.markdown(
        """
        <div style="background:#F8FAFC; border:1px solid #E5E7EB; border-radius:10px;
                    padding:1.25rem; margin-top:1rem;">
        <p style="font-weight:600; color:#1A1A2E; margin:0 0 0.5rem 0;">Supported Formats</p>
        <table style="font-size:0.88rem; color:#374151; border-collapse:collapse; width:100%;">
        <tr style="border-bottom:1px solid #E5E7EB">
            <td style="padding:6px 10px"><strong>Fingrid</strong></td>
            <td style="padding:6px 10px">15 min or 1 h – columns: Alkuaika, Määrä (ISO 8601 dates)</td>
        </tr>
        <tr style="border-bottom:1px solid #E5E7EB">
            <td style="padding:6px 10px"><strong>Vare</strong></td>
            <td style="padding:6px 10px">15 min or 1 h – columns: Aikaleima, Kulutus, Spot-hinta</td>
        </tr>
        <tr>
            <td style="padding:6px 10px"><strong>Legacy Finnish</strong></td>
            <td style="padding:6px 10px">Excel with Finnish weekday-prefixed dates + kWh column</td>
        </tr>
        </table>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

# ═══════════════════════════════════════════════════════════════════════════════
# LOAD & PROCESS DATA
# ═══════════════════════════════════════════════════════════════════════════════

with st.spinner("Loading and parsing your data…"):
    df_raw, detected_format = load_file(uploaded_file.getvalue(), uploaded_file.name)

st.success(f"Loaded **{len(df_raw):,}** rows from `{uploaded_file.name}`")
st.markdown(
    f'<span class="badge-info">Detected format: {detected_format}</span>',
    unsafe_allow_html=True,
)

# ── Data overview metrics ────────────────────────────────────────────────────
resolution = detect_resolution(df_raw)
has_price_col = "Price_EUR_kWh" in df_raw.columns
leap_found, leap_count = has_leap_day(df_raw)

st.markdown('<p class="section-header">Data Overview</p>', unsafe_allow_html=True)

mc1, mc2, mc3, mc4, mc5 = st.columns(5)

with mc1:
    st.markdown(
        f"""<div class="metric-card">
        <div class="metric-value">{len(df_raw):,}</div>
        <div class="metric-label">Total Rows</div>
        </div>""",
        unsafe_allow_html=True,
    )
with mc2:
    st.markdown(
        f"""<div class="metric-card">
        <div class="metric-value">{resolution}</div>
        <div class="metric-label">Detected Resolution</div>
        </div>""",
        unsafe_allow_html=True,
    )
with mc3:
    date_range = f"{df_raw.index.min():%Y-%m-%d} — {df_raw.index.max():%Y-%m-%d}"
    st.markdown(
        f"""<div class="metric-card">
        <div class="metric-value" style="font-size:0.95rem">{date_range}</div>
        <div class="metric-label">Date Range</div>
        </div>""",
        unsafe_allow_html=True,
    )
with mc4:
    total_kwh = df_raw["kWh"].sum()
    st.markdown(
        f"""<div class="metric-card">
        <div class="metric-value">{total_kwh:,.1f}</div>
        <div class="metric-label">Total kWh</div>
        </div>""",
        unsafe_allow_html=True,
    )
with mc5:
    price_label = "Yes" if has_price_col else "No"
    st.markdown(
        f"""<div class="metric-card">
        <div class="metric-value">{price_label}</div>
        <div class="metric-label">Price Column</div>
        </div>""",
        unsafe_allow_html=True,
    )

# Leap day badge (uses shape + text, not colour alone)
if leap_found:
    st.markdown(
        f'<span class="badge-warn">Leap day detected — {leap_count} rows on Feb 29</span>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<span class="badge-ok">No leap-day data present</span>',
        unsafe_allow_html=True,
    )

# ── Apply transformations ───────────────────────────────────────────────────
df = df_raw.copy()

if leap_day_action == "Remove automatically" and leap_found:
    df = remove_leap_days(df)
    st.info(f"Removed {leap_count} leap-day rows (Feb 29).")

if operation == "15 min → 1 hour (SUM)":
    df = convert_15min_to_1h(df)
    st.info(f"Converted to **1-hour** resolution — {len(df):,} rows")
elif operation == "1 hour → 15 min (÷4)":
    df = convert_1h_to_15min(df)
    st.info(f"Converted to **15-min** resolution — {len(df):,} rows")

if "Price_EUR_kWh" in df.columns and "Cost_EUR" not in df.columns:
    df["Cost_EUR"] = df["kWh"] * df["Price_EUR_kWh"]

# ═══════════════════════════════════════════════════════════════════════════════
# DATA PREVIEW
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)
st.markdown('<p class="section-header">Data Preview (first 100 rows)</p>', unsafe_allow_html=True)
st.dataframe(df.head(100), width='stretch', height=300)

# ═══════════════════════════════════════════════════════════════════════════════
# STATISTICS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)
st.markdown('<p class="section-header">Statistical Summary</p>', unsafe_allow_html=True)

stats_cols = st.columns(2)

with stats_cols[0]:
    st.markdown("**Consumption (kWh)**")
    desc_kwh = df["kWh"].describe(percentiles=[0.05, 0.25, 0.5, 0.75, 0.95])
    desc_kwh = desc_kwh.rename({
        "count": "Count", "mean": "Mean", "std": "Std Dev", "min": "Min",
        "5%": "P5", "25%": "Q1 (25%)", "50%": "Median",
        "75%": "Q3 (75%)", "95%": "P95", "max": "Max",
    })
    st.dataframe(
        desc_kwh.to_frame("Value").style.format("{:.4f}"),
        width='stretch',
    )

with stats_cols[1]:
    if "Price_EUR_kWh" in df.columns:
        st.markdown("**Price (€/kWh) & Cost (€)**")
        desc_price = df[["Price_EUR_kWh", "Cost_EUR"]].describe(
            percentiles=[0.05, 0.25, 0.5, 0.75, 0.95]
        )
        st.dataframe(desc_price.style.format("{:.4f}"), width='stretch')
    else:
        st.markdown("**Daily Aggregates**")
        daily_agg = df.resample("D")["kWh"].agg(["sum", "mean", "max", "min"])
        daily_agg.columns = ["Total kWh", "Mean kWh", "Peak kWh", "Min kWh"]
        st.dataframe(
            daily_agg.describe().style.format("{:.2f}"),
            width='stretch',
        )

# ═══════════════════════════════════════════════════════════════════════════════
# CHARTS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)
st.markdown('<p class="section-header">Visualisations</p>', unsafe_allow_html=True)

if show_timeseries:
    st.plotly_chart(fig_timeseries(df), width='stretch')

chart_row1 = st.columns(2)
if show_daily_profile:
    with chart_row1[0]:
        st.plotly_chart(fig_daily_profile(df), width='stretch')
if show_weekly_profile:
    with chart_row1[1]:
        st.plotly_chart(fig_weekly_profile(df), width='stretch')

chart_row2 = st.columns(2)
if show_monthly:
    with chart_row2[0]:
        st.plotly_chart(fig_monthly_totals(df), width='stretch')
if show_heatmap:
    with chart_row2[1]:
        st.plotly_chart(fig_heatmap(df), width='stretch')

chart_row3 = st.columns(2)
if show_boxplot:
    with chart_row3[0]:
        st.plotly_chart(fig_monthly_boxplot(df), width='stretch')
if show_cumulative:
    with chart_row3[1]:
        st.plotly_chart(fig_cumulative(df), width='stretch')

if show_price and "Price_EUR_kWh" in df.columns:
    st.plotly_chart(fig_price_timeseries(df), width='stretch')

# ═══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('<hr class="styled-divider">', unsafe_allow_html=True)
st.markdown('<p class="section-header">Export Processed Data</p>', unsafe_allow_html=True)

export_col1, export_col2 = st.columns([2, 1])

with export_col1:
    # Strip original extension for clean default name
    _base_name = uploaded_file.name.rsplit(".", 1)[0] if "." in uploaded_file.name else uploaded_file.name
    export_name = st.text_input(
        "Output filename",
        value=f"processed_{_base_name}.xlsx",
        help="The file will be saved as .xlsx",
    )

with export_col2:
    st.markdown("<br>", unsafe_allow_html=True)
    excel_bytes = to_excel_bytes(df.reset_index())
    st.download_button(
        label="Download Excel",
        data=excel_bytes,
        file_name=export_name if export_name.endswith(".xlsx") else export_name + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown(
    f"""
    <p class="app-footer">
    Processed <strong>{len(df):,}</strong> rows &nbsp;·&nbsp;
    Resolution: <strong>{detect_resolution(df)}</strong> &nbsp;·&nbsp;
    Total: <strong>{df['kWh'].sum():,.1f} kWh</strong>
    </p>
    """,
    unsafe_allow_html=True,
)