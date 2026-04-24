"""
GST Reconciliation App
Version: 9.0 — Multi-file GSTR-2B, Month-wise analysis, Action columns
"""

import streamlit as st
import pandas as pd
import io
import logging
import os
from datetime import datetime
from reconciliation_engine import (
    parse_tally, parse_gstr2b,
    parse_tally_purchase_register, parse_gstr2b_excel,
    detect_file_format,
    reconcile, post_processing_cleaner, strict_numeric_cleaner
)

# =================== CONFIG =================== #

DEFAULT_TOLERANCE = 1.0
LOG_DIR = "logs"

if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

log_file = f"{LOG_DIR}/reconcile_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler(log_file), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# =================== PAGE CONFIG =================== #

st.set_page_config(page_title="GST Reconciliation", layout="wide")

st.markdown("""
<style>
    .stApp, .stMarkdown, .stText, .stCaption, .stMetric,
    .stTabs [data-baseweb="tab"], button, label, input {
        font-family: 'Aptos Narrow', 'Aptos', sans-serif !important;
    }
    .insight-card {
        border-radius: 12px; padding: 20px 24px;
        text-align: center; margin-bottom: 8px;
    }
    .insight-card .card-number {
        font-size: 2.4rem; font-weight: 700; line-height: 1.1;
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
    }
    .insight-card .card-label {
        font-size: 0.85rem; font-weight: 600; margin-top: 4px;
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
    }
    .card-green  { background:#DCFCE7; }
    .card-green  .card-number, .card-green  .card-label { color:#166534; }
    .card-red    { background:#FEE2E2; }
    .card-red    .card-number, .card-red    .card-label { color:#991B1B; }
    .card-orange { background:#FEF3C7; }
    .card-orange .card-number, .card-orange .card-label { color:#92400E; }
    .card-yellow { background:#FFEDD5; }
    .card-yellow .card-number, .card-yellow .card-label { color:#9A3412; }
    .warning-box {
        background-color:#FEF2F2; padding:1rem;
        border-left:4px solid #DC2626; border-radius:4px; margin:1rem 0;
    }
    .month-summary-box {
        background:#F0F9FF; border:1px solid #BAE6FD;
        border-radius:10px; padding:16px 20px; margin:12px 0;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 2px; }
    .stTabs [data-baseweb="tab"]      { padding: 8px 16px; }
    .streamlit-expanderHeader          { font-size: 1rem !important; }
    .stDataFrame:has(td:empty)         { display: none !important; }

    /* ── Expander anti-shake ───────────────────────────────────────────────── */
    /* Kill the open/close animation so the page doesn't jump */
    details > div[data-testid="stExpanderDetails"] {
        animation: none !important;
        transition: none !important;
    }
    /* Reserve vertical space so content below doesn't shift when expander opens */
    div[data-testid="stExpander"] {
        contain: layout;
    }
    /* Prevent scroll-position jump on re-render */
    html { scroll-behavior: auto !important; }
</style>
""", unsafe_allow_html=True)

# =================== HEADER =================== #

st.title("GST Reconciliation")
st.caption("Books vs GSTR-2B · Multi-month · Tally PR, GSTR-2B Excel & Standard Template")

# =================== SAFE DISPLAY HELPER =================== #

def safe_dataframe(df, column_config=None, empty_message="No data to display", caption=None):
    """Render a dataframe only when it has real rows."""
    if df is None or df.empty:
        st.info(empty_message); return False
    df = df.dropna(how="all")
    for col in ["Invoice_No", "Invoice No"]:
        if col in df.columns:
            mask = (df[col].astype(str).str.strip().str.lower() != "nan") & \
                   (df[col].astype(str).str.strip() != "")
            df = df[mask]
    if df.empty:
        st.info(empty_message); return False
    numeric_cols = df.select_dtypes(include=["number"]).columns
    str_cols     = df.select_dtypes(exclude=["number"]).columns
    has_numbers  = len(numeric_cols) > 0 and df[numeric_cols].abs().sum().sum() > 0
    _EMPTIES     = {"", "nan", "none", "0", "0.0"}
    has_strings  = False
    if len(str_cols) > 0:
        for col in str_cols:
            if df[col].astype(str).str.strip().str.lower().apply(lambda v: v not in _EMPTIES).any():
                has_strings = True; break
    if not has_numbers and not has_strings:
        st.info(empty_message); return False
    if caption:
        st.caption(caption)
    st.dataframe(df, use_container_width=True, hide_index=True, column_config=column_config)
    return True

# =================== UPLOAD =================== #

col1, col2 = st.columns(2)
with col1:
    books_file = st.file_uploader("📤 Upload Books", type=["xlsx", "xls", "csv"])
with col2:
    # FEATURE 1: Multiple GSTR-2B files
    gstr_files = st.file_uploader(
        "📤 Upload GSTR-2B (multiple months supported)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
    )

tolerance = st.number_input(
    "Tolerance (₹)", value=DEFAULT_TOLERANCE, step=0.5, min_value=0.0,
    help="Max acceptable tax difference (₹) for 'Matched'. Default: ₹1.00"
)

# =================== PARSER ROUTER =================== #

def load_raw(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, header=None)
    else:
        return pd.read_excel(uploaded_file, header=None)

def parse_books(uploaded_file):
    raw = load_raw(uploaded_file)
    fmt = detect_file_format(raw, uploaded_file.name)
    logger.info(f"Books format detected: {fmt}")
    if fmt == "tally_pr":
        return (*parse_tally_purchase_register(raw), fmt)
    uploaded_file.seek(0)
    df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith(".csv") else \
         pd.read_excel(uploaded_file)
    return (*parse_tally(df), "standard")

def parse_gstr_single(uploaded_file):
    """Parse a single GSTR-2B file and return (df, fmt)."""
    raw = load_raw(uploaded_file)
    fmt = detect_file_format(raw, uploaded_file.name)
    logger.info(f"GSTR-2B format detected: {fmt} for {uploaded_file.name}")
    if fmt == "gstr2b_excel":
        return parse_gstr2b_excel(raw), fmt
    uploaded_file.seek(0)
    df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith(".csv") else \
         pd.read_excel(uploaded_file)
    return parse_gstr2b(df), "standard"

# FEATURE 2: Month column helper
def add_month_column(df: pd.DataFrame) -> pd.DataFrame:
    """Add Month column derived from Invoice_Date."""
    if df.empty or "Invoice_Date" not in df.columns:
        df["Month"] = ""
        return df
    df = df.copy()
    df["Month"] = pd.to_datetime(df["Invoice_Date"], errors="coerce") \
                    .dt.to_period("M").astype(str)
    return df

# =================== SAFE EXCEL HELPERS =================== #

def safe_write_number(ws, row, col, value, fmt):
    try:
        v = float(value)
        ws.write_number(row, col, 0.0 if pd.isna(v) else v, fmt)
    except Exception:
        ws.write_number(row, col, 0.0, fmt)

def safe_write_text(ws, row, col, value, fmt):
    try:
        ws.write_string(row, col, "" if (value is None or pd.isna(value)) else str(value).strip(), fmt)
    except Exception:
        ws.write_string(row, col, "", fmt)

# =================== PROCESS =================== #

if st.button("🚀 Run Reconciliation", use_container_width=True, type="primary"):
    if not books_file or not gstr_files:
        st.error("Please upload both Books and at least one GSTR-2B file.")
    else:
        try:
            with st.spinner("Processing files…"):
                # Parse Books
                books_clean, no_itc, issues, books_fmt = parse_books(books_file)

                # FEATURE 1: Parse and combine multiple GSTR-2B files
                gstr_parts = []
                for gf in gstr_files:
                    gdf, gfmt = parse_gstr_single(gf)
                    gdf["Source_File"] = gf.name   # tag source file name
                    gstr_parts.append(gdf)
                    logger.info(f"Parsed {gf.name}: {len(gdf)} rows")

                gstr_clean = pd.concat(gstr_parts, ignore_index=True) if gstr_parts else pd.DataFrame()

                # FEATURE 2: Add Month columns
                books_clean = add_month_column(books_clean)
                gstr_clean  = add_month_column(gstr_clean)

                # Reconcile (engine works on standard columns, Month is extra)
                results = reconcile(gstr_clean, books_clean, tolerance)
                results.update({
                    "no_itc":    no_itc,
                    "issues":    issues,
                    "books_raw": books_clean,
                    "gstr_raw":  gstr_clean,
                    "books_fmt": books_fmt,
                    "gstr_fmts": [gf.name for gf in gstr_files],
                    "n_gstr_files": len(gstr_files),
                })
                st.session_state["results"] = results

            n = len(gstr_files)
            st.success(f"✅ Reconciliation completed! ({n} GSTR-2B file{'s' if n>1 else ''} processed)")
        except Exception as e:
            logger.error(f"Reconciliation failed: {e}", exc_info=True)
            st.error(f"Error: {str(e)}")

# =================== DISPLAY =================== #

if "results" not in st.session_state:
    st.stop()

r   = st.session_state["results"]
s   = r["summary"]
tol = tolerance

# Stable cache key: hash of ITC totals + file counts. Changes only on new reconciliation.
_cache_key = (
    s.get("ITC_Books", 0),
    s.get("ITC_GSTR", 0),
    s.get("Matched", 0),
    s.get("Missing_2B", 0),
    s.get("Missing_Books", 0),
    r.get("n_gstr_files", 1),
    len(r.get("books_raw", pd.DataFrame())),
    len(r.get("gstr_raw",  pd.DataFrame())),
)

# ── helpers ─────────────────────────────────────────────────────────────────

def fmt_date(val) -> str:
    try:
        return pd.to_datetime(val).strftime("%d-%b-%Y") if pd.notna(val) else ""
    except Exception:
        return ""

def _fmt_month_label(m: str) -> str:
    """Convert '2025-02' → 'February 2025' for display. Falls back to raw string."""
    try:
        return pd.to_datetime(m + "-01", format="%Y-%m-%d").strftime("%B %Y")
    except Exception:
        return m

def coerce_str_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].apply(lambda v: "" if (v is None or pd.isna(v)) else str(v))
    return df

def apply_filters(df: pd.DataFrame, f_gstin: str, f_supplier: str) -> pd.DataFrame:
    if f_gstin and "GSTIN" in df.columns:
        df = df[df["GSTIN"].str.contains(f_gstin, case=False, na=False)]
    if f_supplier:
        for col in ["Supplier", "Trade_Name"]:
            if col in df.columns:
                df = df[df[col].str.contains(f_supplier, case=False, na=False)]
                break
    return df

# FEATURE 5: Action Required mapping
ACTION_MAP = {
    "❌ Missing in GST":   "Follow up with supplier",
    "📕 Missing in Books": "Record purchase entry",
    "⚠️ Tax Difference":   "Verify invoice values",
    "✅ Matched":          "No action",
}

def add_action_column(df: pd.DataFrame) -> pd.DataFrame:
    if "Remarks" in df.columns:
        df = df.copy()
        df["Action Required"] = df["Remarks"].map(ACTION_MAP).fillna("")
    return df

# ── DATA ISSUES ──────────────────────────────────────────────────────────────

all_issues = r["issues"].copy() if not r["issues"].empty else pd.DataFrame()
if "duplicate_issues" in r and not r["duplicate_issues"].empty:
    all_issues = pd.concat(
        [x for x in [all_issues, r["duplicate_issues"]] if not x.empty],
        ignore_index=True
    )

# ── DATA ISSUES — passive banner only, detail lives in the Issues tab ────────
# No widget here = no re-run triggered = no shake.
# The dataframe is shown inside st.tabs below (tab switching is CSS-only, zero re-run).
if not all_issues.empty:
    st.markdown(f"""
    <div class="warning-box">
        <strong>⚠️ {len(all_issues)} Data Issues Found</strong>
        — See the <strong>⚠️ Data Issues</strong> tab below for details.
    </div>""", unsafe_allow_html=True)

# ── QUICK INSIGHT CARDS ───────────────────────────────────────────────────────

st.markdown("## 📊 Reconciliation Summary")

n_files = r.get("n_gstr_files", 1)
if n_files > 1:
    st.caption(f"Across {n_files} GSTR-2B files: {', '.join(r.get('gstr_fmts', []))}")

c1, c2, c3, c4 = st.columns(4)
cards = [
    (c1, "card-green",  "✅ Matched",            s["Matched"]),
    (c2, "card-red",    "❌ Missing in GSTR-2B",  s["Missing_2B"]),
    (c3, "card-orange", "📕 Missing in Books",    s["Missing_Books"]),
    (c4, "card-yellow", "⚠️ Tax Difference",      s["Tax_Diff"]),
]
for _col, _css, _label, _value in cards:
    with _col:
        st.markdown(f"""
        <div class="insight-card {_css}">
            <div class="card-number">{_value}</div>
            <div class="card-label">{_label}</div>
        </div>""", unsafe_allow_html=True)
del _col, _css, _label, _value

# ── FEATURE 3 & 6: MONTH-WISE SUMMARY ────────────────────────────────────────

def _build_month_summary(books_raw, gstr_raw, missing_2b, missing_books, matched_df):
    """Build month-wise ITC summary. Runs in function scope."""
    rows = []
    # Collect all months from both sources
    all_months = set()
    for df in [books_raw, gstr_raw]:
        if not df.empty and "Month" in df.columns:
            all_months.update(df["Month"].dropna().unique())
    all_months.discard("")
    all_months.discard("NaT")

    for month in sorted(all_months):
        b_tax = books_raw[books_raw["Month"] == month]["TOTAL_TAX"].sum() \
                if not books_raw.empty and "Month" in books_raw.columns else 0
        g_tax = gstr_raw[gstr_raw["Month"] == month]["TOTAL_TAX"].sum() \
                if not gstr_raw.empty and "Month" in gstr_raw.columns else 0

        # Missing in 2B — books invoices not in GSTR (use Invoice_Date month)
        m2b_count = 0
        if not missing_2b.empty and "Invoice_Date" in missing_2b.columns:
            m2b = missing_2b.copy()
            m2b["Month"] = pd.to_datetime(m2b["Invoice_Date"], errors="coerce") \
                             .dt.to_period("M").astype(str)
            m2b_count = int((m2b["Month"] == month).sum())

        mb_count = 0
        if not missing_books.empty and "Invoice_Date" in missing_books.columns:
            mb = missing_books.copy()
            mb["Month"] = pd.to_datetime(mb["Invoice_Date"], errors="coerce") \
                            .dt.to_period("M").astype(str)
            mb_count = int((mb["Month"] == month).sum())

        matched_count = 0
        if not matched_df.empty:
            for col in ["Invoice_Date_2B", "Invoice_Date_Books"]:
                if col in matched_df.columns:
                    mc = matched_df.copy()
                    mc["_m"] = pd.to_datetime(mc[col], errors="coerce") \
                                  .dt.to_period("M").astype(str)
                    matched_count = int((mc["_m"] == month).sum())
                    break

        rows.append({
            "Month":          month,
            "Books ITC":      round(float(b_tax), 2),
            "GSTR ITC":       round(float(g_tax), 2),
            "Difference":     round(float(g_tax) - float(b_tax), 2),
            "Missing 2B":     m2b_count,
            "Missing Books":  mb_count,
            "Matched":        matched_count,
        })

    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    # Sort chronologically
    df["_sort"] = pd.to_datetime(df["Month"], format="%Y-%m", errors="coerce")
    df = df.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
    return df

# Cache month_summary — avoid recomputing on every UI interaction (expander, button, etc.)
if st.session_state.get("_ms_cache_key") != _cache_key:
    st.session_state["_month_summary"] = _build_month_summary(
        r["books_raw"], r["gstr_raw"],
        r.get("missing_2b", pd.DataFrame()),
        r.get("missing_books", pd.DataFrame()),
        r.get("matched", pd.DataFrame()),
    )
    st.session_state["_ms_cache_key"] = _cache_key
month_summary = st.session_state["_month_summary"]

if not month_summary.empty:
    st.markdown("## 📅 Month-wise Summary")

    # Header row
    h0,h1,h2,h3,h4,h5,h6,h7 = st.columns([1.2,1.4,1.4,1.4,1.0,1.2,1.0,1.0])
    h0.markdown("**Month**");      h1.markdown("**📚 Books ITC**")
    h2.markdown("**📊 GSTR ITC**"); h3.markdown("**📉 Difference**")
    h4.markdown("**❌ Miss 2B**");  h5.markdown("**📕 Miss Books**")
    h6.markdown("**✅ Matched**");  h7.markdown("**Action**")
    st.divider()
    for _ms_idx, _ms_row in month_summary.iterrows():
        c0,c1,c2,c3,c4,c5,c6,c7 = st.columns([1.2,1.4,1.4,1.4,1.0,1.2,1.0,1.0])
        c0.markdown(f"**{_fmt_month_label(_ms_row['Month'])}**")
        c1.markdown(f"{_ms_row['Books ITC']:,.2f}")
        c2.markdown(f"{_ms_row['GSTR ITC']:,.2f}")
        _diff_color = "#991B1B" if _ms_row["Difference"] < 0 else "#166534"
        c3.markdown(f"<span style='color:{_diff_color}'>{_ms_row['Difference']:,.2f}</span>",
                    unsafe_allow_html=True)
        c4.markdown(str(int(_ms_row["Missing 2B"])))
        c5.markdown(str(int(_ms_row["Missing Books"])))
        c6.markdown(str(int(_ms_row["Matched"])))
        if c7.button("🔍 View", key=f"view_month_{_ms_row['Month']}",
                     help=f"Filter Invoice Details to {_fmt_month_label(_ms_row['Month'])}"):
            st.session_state["selected_month"] = _ms_row["Month"]
            st.rerun()
    st.caption(f"{len(month_summary)} month(s) of data")

# ── FEATURE 4: MONTH DRILL-DOWN FILTER ────────────────────────────────────────

sorted_months = list(month_summary["Month"]) if not month_summary.empty else []

st.markdown("### 🔍 Filter Results")
drill_col, gstin_col, sup_col, status_col = st.columns([1.5, 1.5, 1.5, 1.5])

with drill_col:
    if sorted_months:
        # Build label→raw mapping: "February 2025" → "2025-02"
        _month_labels   = ["All months"] + [_fmt_month_label(m) for m in sorted_months]
        _month_raw      = ["All months"] + sorted_months
        _label_to_raw   = dict(zip(_month_labels, _month_raw))
        _raw_to_label   = dict(zip(_month_raw, _month_labels))
        # Sync with session_state (stores raw "2025-02")
        _ss_month = st.session_state.get("selected_month", "All months")
        _ss_label = _raw_to_label.get(_ss_month, "All months")
        _default_idx = _month_labels.index(_ss_label) if _ss_label in _month_labels else 0
        _chosen_label = st.selectbox(
            "📅 Filter by Month", _month_labels, index=_default_idx,
            key="month_selectbox",
        )
        selected_month = _label_to_raw.get(_chosen_label, "All months")
        st.session_state["selected_month"] = selected_month
    else:
        selected_month = "All months"
        st.selectbox("📅 Filter by Month", ["All months"], disabled=True)

with gstin_col:
    f_gstin = st.text_input("Filter by GSTIN", placeholder="Enter GSTIN…")
with sup_col:
    f_supplier = st.text_input("Filter by Supplier", placeholder="Enter supplier name…")
with status_col:
    f_status = st.multiselect(
        "Filter by Status",
        options=["✅ Matched", "❌ Missing in GST", "📕 Missing in Books", "⚠️ Tax Difference"],
        default=[], placeholder="All statuses…",
    )

def apply_month_filter(df: pd.DataFrame, month_col: str = "Month") -> pd.DataFrame:
    """Filter rows to selected_month. Applies to df that already has a Month column."""
    if selected_month == "All months" or df.empty:
        return df
    if month_col in df.columns:
        return df[df[month_col] == selected_month]
    return df

def apply_month_filter_by_date(df: pd.DataFrame, date_col: str = "Invoice_Date") -> pd.DataFrame:
    """Filter rows by deriving Month from a date column."""
    if selected_month == "All months" or df.empty:
        return df
    if date_col not in df.columns:
        return df
    df = df.copy()
    df["_m"] = pd.to_datetime(df[date_col], errors="coerce").dt.to_period("M").astype(str)
    result = df[df["_m"] == selected_month].drop(columns=["_m"])
    return result

# ── BUILD INVOICE LEVEL DETAIL ────────────────────────────────────────────────

trade_name_map = r.get("trade_name_mapping", {})

def _build_detail_df(books_raw, gstr_raw, trade_name_map, tol, fmt_date_fn):
    """Build invoice-level detail. Function scope prevents loop var magic display."""
    data = []
    seen = set()

    for _, row in books_raw.iterrows():
        key = f"{row['GSTIN']}|{row['Invoice_No']}"
        if key in seen: continue
        seen.add(key)
        gstin    = row["GSTIN"]
        supplier = trade_name_map.get(gstin, row["Trade_Name"])
        month    = row.get("Month", "")
        m_row    = None
        if not gstr_raw.empty:
            m = gstr_raw[(gstr_raw["GSTIN"] == gstin) & (gstr_raw["Invoice_No"] == row["Invoice_No"])]
            if not m.empty: m_row = m.iloc[0]

        if m_row is not None:
            diff   = float(m_row["TOTAL_TAX"]) - float(row["TOTAL_TAX"])
            remark = "✅ Matched" if abs(diff) <= tol else "⚠️ Tax Difference"
            data.append({"GSTIN": gstin, "Supplier": supplier, "Month": month,
                "Invoice No": row["Invoice_No"], "Date": fmt_date_fn(row["Invoice_Date"]),
                "ITC Books": float(row["TOTAL_TAX"]), "ITC 2B": float(m_row["TOTAL_TAX"]),
                "Difference": diff, "Remarks": remark,
                "Action Required": ACTION_MAP.get(remark, "")})
        else:
            remark = "❌ Missing in GST"
            data.append({"GSTIN": gstin, "Supplier": supplier, "Month": month,
                "Invoice No": row["Invoice_No"], "Date": fmt_date_fn(row["Invoice_Date"]),
                "ITC Books": float(row["TOTAL_TAX"]), "ITC 2B": 0.0,
                "Difference": -float(row["TOTAL_TAX"]), "Remarks": remark,
                "Action Required": ACTION_MAP.get(remark, "")})

    for _, row in gstr_raw.iterrows():
        key = f"{row['GSTIN']}|{row['Invoice_No']}"
        if key in seen: continue
        seen.add(key)
        remark = "📕 Missing in Books"
        data.append({"GSTIN": row["GSTIN"],
            "Supplier": trade_name_map.get(row["GSTIN"], row["Trade_Name"]),
            "Month": row.get("Month", ""),
            "Invoice No": row["Invoice_No"], "Date": fmt_date_fn(row["Invoice_Date"]),
            "ITC Books": 0.0, "ITC 2B": float(row["TOTAL_TAX"]),
            "Difference": float(row["TOTAL_TAX"]), "Remarks": remark,
            "Action Required": ACTION_MAP.get(remark, "")})

    return pd.DataFrame(data).sort_values(["Supplier", "Date"]) if data else pd.DataFrame()

# Cache detail_df — iterating 100s of invoices on every UI interaction causes the shake
if st.session_state.get("_dd_cache_key") != _cache_key:
    st.session_state["_detail_df"] = _build_detail_df(
        r["books_raw"], r["gstr_raw"], trade_name_map, tol, fmt_date
    )
    st.session_state["_dd_cache_key"] = _cache_key
detail_df = st.session_state["_detail_df"]

def filter_detail(df):
    if df.empty: return df
    df = apply_month_filter(df)
    if f_gstin:
        df = df[df["GSTIN"].str.contains(f_gstin, case=False, na=False)]
    if f_supplier:
        df = df[df["Supplier"].str.contains(f_supplier, case=False, na=False)]
    if f_status:
        status_map = {
            "✅ Matched":          "✅ Matched",
            "❌ Missing in GST":   "❌ Missing in GST",
            "📕 Missing in Books": "📕 Missing in Books",
            "⚠️ Tax Difference":   "⚠️ Tax Difference",
        }
        allowed = [status_map[s] for s in f_status if s in status_map]
        df = df[df["Remarks"].isin(allowed)]
    # Invoice number search (applied last, used only in Invoice Details tab)
    if "Invoice No" in df.columns and st.session_state.get("_f_inv_no","").strip():
        _inv_q = st.session_state["_f_inv_no"].strip()
        df = df[df["Invoice No"].astype(str).str.contains(_inv_q, case=False, na=False)]
    return df

# Column configs
DETAIL_COL_CFG = {
    "Month":           st.column_config.TextColumn("Month",           width=90),
    "GSTIN":           st.column_config.TextColumn("GSTIN",           width=180),
    "Supplier":        st.column_config.TextColumn("Supplier",        width=250),
    "Invoice No":      st.column_config.TextColumn("Invoice No",      width=140),
    "Date":            st.column_config.TextColumn("Date",            width=105),
    "ITC Books":       st.column_config.NumberColumn("📚 Books",      width=110, format="%.2f"),
    "ITC 2B":          st.column_config.NumberColumn("📊 GSTR-2B",    width=110, format="%.2f"),
    "Difference":      st.column_config.NumberColumn("📉 Diff",       width=100, format="%.2f"),
    "Remarks":         st.column_config.TextColumn("Remarks",         width=160),
    "Action Required": st.column_config.TextColumn("Action Required", width=180),
}

STD_BOOK_CFG = {
    "GSTIN":         st.column_config.TextColumn("GSTIN",          width=180),
    "Trade_Name":    st.column_config.TextColumn("Trade Name",     width=260),
    "Invoice_No":    st.column_config.TextColumn("Invoice No",     width=140),
    "Invoice_Date":  st.column_config.TextColumn("Invoice Date",   width=110),
    "Month":         st.column_config.TextColumn("Month",          width=90),
    "Taxable_Value": st.column_config.NumberColumn("Taxable",      width=110, format="%.2f"),
    "CGST":          st.column_config.NumberColumn("CGST",         width=100, format="%.2f"),
    "SGST":          st.column_config.NumberColumn("SGST",         width=100, format="%.2f"),
    "IGST":          st.column_config.NumberColumn("IGST",         width=100, format="%.2f"),
    "CESS":          st.column_config.NumberColumn("CESS",         width=80,  format="%.2f"),
    "TOTAL_TAX":     st.column_config.NumberColumn("Total Tax",    width=110, format="%.2f"),
    "Invoice_Value": st.column_config.NumberColumn("Invoice Value",width=120, format="%.2f"),
    "Source_File":   st.column_config.TextColumn("Source File",    width=160),
}

MISS_CFG = {
    "GSTIN":      st.column_config.TextColumn("GSTIN",      width=180),
    "Supplier":   st.column_config.TextColumn("Supplier",   width=260),
    "Invoice No": st.column_config.TextColumn("Invoice No", width=140),
    "Date":       st.column_config.TextColumn("Date",       width=105),
    "Taxable":    st.column_config.NumberColumn("Taxable",  width=110, format="%.2f"),
    "ITC":        st.column_config.NumberColumn("ITC",      width=110, format="%.2f"),
}

# ── TABS ─────────────────────────────────────────────────────────────────────

_issues_tab_label = f"⚠️ Data Issues ({len(all_issues)})" if not all_issues.empty else "⚠️ Data Issues"
tabs = st.tabs([
    "📋 Invoice Details",
    "📚 Books",
    "📊 GSTR-2B",
    "❌ Missing in 2B",
    "📕 Missing in Books",
    "📋 Supplier Summary",
    _issues_tab_label,
])

# Tab 0: Invoice Level Details (with Month, Action Required)
with tabs[0]:
    if not detail_df.empty:
        # Invoice number search bar — stored in session_state so filter_detail can read it
        _inv_search = st.text_input(
            "🔎 Search Invoice Number",
            value=st.session_state.get("_f_inv_no", ""),
            placeholder="Type invoice number to search…",
            key="_inv_no_widget",
        )
        st.session_state["_f_inv_no"] = _inv_search

        display_df = filter_detail(detail_df.copy())
        safe_dataframe(
            coerce_str_cols(display_df),
            column_config=DETAIL_COL_CFG,
            empty_message="No invoices match the current filters.",
            caption=f"Showing {len(display_df)} of {len(detail_df)} invoices",
        )
    else:
        st.info("No invoice data available.")

# Tab 1: Books
with tabs[1]:
    if not r["books_raw"].empty:
        df_b = r["books_raw"].copy()
        df_b["Invoice_Date"] = pd.to_datetime(
            df_b["Invoice_Date"], errors="coerce"
        ).dt.strftime("%d-%b-%Y").fillna("")
        df_b = apply_month_filter(df_b)
        df_b = apply_filters(coerce_str_cols(df_b), f_gstin, f_supplier)
        safe_dataframe(df_b, column_config=STD_BOOK_CFG,
            empty_message="No Books data matches the current filters.",
            caption=f"{len(df_b)} records")
    else:
        st.info("No Books data loaded.")

# Tab 2: GSTR-2B
with tabs[2]:
    if not r["gstr_raw"].empty:
        df_g = r["gstr_raw"].copy()
        df_g["Invoice_Date"] = pd.to_datetime(
            df_g["Invoice_Date"], errors="coerce"
        ).dt.strftime("%d-%b-%Y").fillna("")
        df_g = apply_month_filter(df_g)
        df_g = apply_filters(coerce_str_cols(df_g), f_gstin, f_supplier)
        src_note = f" · Source file shown in last column" if "Source_File" in df_g.columns else ""
        safe_dataframe(df_g, column_config=STD_BOOK_CFG,
            empty_message="No GSTR-2B data matches the current filters.",
            caption=f"{len(df_g)} records{src_note}")
    else:
        st.info("No GSTR-2B data loaded.")

# Tab 3: Missing in 2B
with tabs[3]:
    if not r["missing_2b"].empty:
        df_m = r["missing_2b"][
            ["GSTIN","Trade_Name","Invoice_No","Invoice_Date","Taxable_Value","TOTAL_TAX"]
        ].copy()
        df_m["Invoice_Date"] = pd.to_datetime(
            df_m["Invoice_Date"], errors="coerce"
        ).dt.strftime("%d-%b-%Y").fillna("")
        df_m.columns = ["GSTIN","Supplier","Invoice No","Date","Taxable","ITC"]
        df_m = apply_month_filter_by_date(df_m, "Date")
        df_m = apply_filters(coerce_str_cols(df_m), f_gstin, f_supplier)
        safe_dataframe(df_m, column_config=MISS_CFG,
            empty_message="No missing invoices match the current filters.",
            caption=f"💰 ITC at Risk: ₹{df_m['ITC'].sum():,.2f} across {len(df_m)} invoices")
    else:
        st.success("✅ No invoices missing in GSTR-2B.")

# Tab 4: Missing in Books
with tabs[4]:
    if not r["missing_books"].empty:
        df_mb = r["missing_books"][
            ["GSTIN","Trade_Name","Invoice_No","Invoice_Date","Taxable_Value","TOTAL_TAX"]
        ].copy()
        df_mb["Invoice_Date"] = pd.to_datetime(
            df_mb["Invoice_Date"], errors="coerce"
        ).dt.strftime("%d-%b-%Y").fillna("")
        df_mb.columns = ["GSTIN","Supplier","Invoice No","Date","Taxable","ITC"]
        df_mb = apply_month_filter_by_date(df_mb, "Date")
        df_mb = apply_filters(coerce_str_cols(df_mb), f_gstin, f_supplier)
        safe_dataframe(df_mb, column_config=MISS_CFG,
            empty_message="No missing invoices match the current filters.",
            caption=f"{len(df_mb)} invoices in GSTR-2B not found in Books")
    else:
        st.success("✅ No invoices missing in Books.")

# Tab 5: Supplier Summary
with tabs[5]:
    all_gstins = set()
    if not r["books_raw"].empty: all_gstins.update(r["books_raw"]["GSTIN"].unique())
    if not r["gstr_raw"].empty:  all_gstins.update(r["gstr_raw"]["GSTIN"].unique())
    sup_rows = []
    for _gstin in all_gstins:
        _b_raw = apply_month_filter(r["books_raw"], "Month") if not r["books_raw"].empty else pd.DataFrame()
        _g_raw = apply_month_filter(r["gstr_raw"],  "Month") if not r["gstr_raw"].empty  else pd.DataFrame()
        ib = _b_raw[_b_raw["GSTIN"] == _gstin]["TOTAL_TAX"].sum() if not _b_raw.empty else 0
        ig = _g_raw[_g_raw["GSTIN"] == _gstin]["TOTAL_TAX"].sum() if not _g_raw.empty else 0
        sup_rows.append({
            "GSTIN":            str(_gstin),
            "Supplier":         str(trade_name_map.get(_gstin, "Unknown")),
            "ITC as per Books": round(float(ib), 2),
            "ITC as per 2B":    round(float(ig), 2),
            "ITC Difference":   round(float(ig) - float(ib), 2),
        })
    if sup_rows:
        sup_df = pd.DataFrame(sup_rows).sort_values("Supplier")
        sup_df = sup_df[sup_df["GSTIN"].str.strip().replace("nan","") != ""]
        if f_gstin:
            sup_df = sup_df[sup_df["GSTIN"].str.contains(f_gstin, case=False, na=False)]
        if f_supplier:
            sup_df = sup_df[sup_df["Supplier"].str.contains(f_supplier, case=False, na=False)]
        safe_dataframe(sup_df, column_config={
            "GSTIN":            st.column_config.TextColumn("GSTIN",        width=180),
            "Supplier":         st.column_config.TextColumn("Supplier",     width=260),
            "ITC as per Books": st.column_config.NumberColumn("📚 Books",   width=140, format="%.2f"),
            "ITC as per 2B":    st.column_config.NumberColumn("📊 GSTR-2B", width=140, format="%.2f"),
            "ITC Difference":   st.column_config.NumberColumn("📉 Diff",    width=130, format="%.2f"),
        }, empty_message="No supplier data matches the current filters.",
           caption=f"Total Difference: ₹{sup_df['ITC Difference'].sum():,.2f}" if not sup_df.empty else None)

# Tab 6: Data Issues — zero Python re-run when switching to this tab
with tabs[6]:
    if all_issues.empty:
        st.success("✅ No data issues found.")
    else:
        _df_iss = all_issues.copy()
        if "Invoice_Date" in _df_iss.columns:
            _df_iss["Invoice_Date"] = pd.to_datetime(
                _df_iss["Invoice_Date"], errors="coerce"
            ).dt.strftime("%d-%b-%Y").fillna("")
        _df_iss = coerce_str_cols(_df_iss)
        _cols_order = ["Issue"] + [c for c in _df_iss.columns if c != "Issue"]
        _df_iss = _df_iss[_cols_order]

        # Issue count summary (compact)
        _ic = _df_iss["Issue"].value_counts().reset_index()
        _ic.columns = ["Issue Type", "Count"]
        st.caption(f"{len(_df_iss)} issues across {len(_ic)} categories")
        st.dataframe(_ic, use_container_width=True, hide_index=True,
                     column_config={
                         "Issue Type": st.column_config.TextColumn("Issue Type", width=300),
                         "Count":      st.column_config.NumberColumn("Count",     width=100),
                     })
        st.markdown("##### All Issues")
        st.dataframe(_df_iss, use_container_width=True, hide_index=True,
                     column_config={
                         "Issue":         st.column_config.TextColumn("Issue",        width=220),
                         "GSTIN":         st.column_config.TextColumn("GSTIN",        width=180),
                         "Trade_Name":    st.column_config.TextColumn("Trade Name",   width=260),
                         "Invoice_No":    st.column_config.TextColumn("Invoice No",   width=150),
                         "Invoice_Date":  st.column_config.TextColumn("Invoice Date", width=120),
                         "Taxable_Value": st.column_config.NumberColumn("Taxable",    width=110, format="%.2f"),
                         "TOTAL_TAX":     st.column_config.NumberColumn("Total Tax",  width=110, format="%.2f"),
                     })

# ── ZERO ITC ─────────────────────────────────────────────────────────────────

if not r["no_itc"].empty:
    df_ni = r["no_itc"][
        ["GSTIN","Trade_Name","Invoice_No","Invoice_Date","Taxable_Value","Invoice_Value"]
    ].copy()
    df_ni["Invoice_Date"] = pd.to_datetime(
        df_ni["Invoice_Date"], errors="coerce"
    ).dt.strftime("%d-%b-%Y").fillna("")
    df_ni.columns = ["GSTIN","Supplier","Invoice No","Date","Taxable","Invoice Value"]
    df_ni = coerce_str_cols(df_ni)
    df_ni = df_ni[
        (df_ni["Invoice No"].astype(str).str.strip() != "") &
        (df_ni["Invoice No"].astype(str).str.strip().str.lower() != "nan") &
        (df_ni["Invoice Value"].astype(float) > 0)
    ]
    if not df_ni.empty:
        st.divider()
        st.markdown("## 🟡 Zero ITC Invoices")
        safe_dataframe(df_ni, column_config={
            "GSTIN":         st.column_config.TextColumn("GSTIN",          width=180),
            "Supplier":      st.column_config.TextColumn("Supplier",       width=260),
            "Invoice No":    st.column_config.TextColumn("Invoice No",     width=140),
            "Date":          st.column_config.TextColumn("Date",           width=105),
            "Taxable":       st.column_config.NumberColumn("Taxable",      width=110, format="%.2f"),
            "Invoice Value": st.column_config.NumberColumn("Invoice Value",width=120, format="%.2f"),
        }, empty_message="No zero ITC invoices.", caption=f"{len(df_ni)} invoices with zero tax")

# ── EXCEL EXPORT HELPERS ──────────────────────────────────────────────────────

def _build_workbook_formats(wb):
    return {
        "title":  wb.add_format({'bold':True,'font_size':14,'font_name':'Aptos Narrow'}),
        "header": wb.add_format({'bold':True,'font_name':'Aptos Narrow',
                                  'font_color':'white','bg_color':'#1F4E78',
                                  'align':'center','valign':'vcenter','border':1}),
        "number": wb.add_format({'font_name':'Aptos Narrow','num_format':'#,##0.00'}),
        "date":   wb.add_format({'font_name':'Aptos Narrow','num_format':'dd-mmm-yyyy'}),
        "text":   wb.add_format({'font_name':'Aptos Narrow'}),
    }

def _xls_write_date(ws, r, c, val, fmts):
    try:
        dt = pd.to_datetime(val)
        if pd.notna(dt):
            ws.write_datetime(r, c, dt.to_pydatetime(), fmts["date"])
        else:
            safe_write_text(ws, r, c, "", fmts["text"])
    except Exception:
        safe_write_text(ws, r, c, str(val) if pd.notna(val) else "", fmts["text"])

def _xls_write_header(ws, row, col, text, fmt):
    _ = ws.write(row, col, text, fmt)

def _write_sheet(ws, title, headers, data_rows, col_types, fmts):
    _ = ws.write(0, 0, title, fmts["title"])
    _ = ws.set_row(1, 20)
    for ci, h in enumerate(headers):
        _xls_write_header(ws, 1, ci, h, fmts["header"])
    for ri, row_data in enumerate(data_rows):
        for ci, val in enumerate(row_data):
            ct = col_types.get(headers[ci], "text")
            if ct == "date" and pd.notna(val):
                _xls_write_date(ws, ri+2, ci, val, fmts)
            elif ct == "number":
                safe_write_number(ws, ri+2, ci, val, fmts["number"])
            else:
                safe_write_text(ws, ri+2, ci, val, fmts["text"])
    for ci, h in enumerate(headers):
        ct = col_types.get(h, "text")
        _ = ws.set_column(ci, ci, 15 if ct in ("date","number") else 22,
                          fmts["date"] if ct=="date" else fmts["number"] if ct=="number" else fmts["text"])
    _ = ws.freeze_panes(2, 0)
    _ = ws.autofit()

BOOK_CT = {"GSTIN":"text","Trade_Name":"text","Invoice_No":"text","Invoice_Date":"date",
           "Month":"text","Month Label":"text","Taxable_Value":"number","CGST":"number","SGST":"number",
           "IGST":"number","CESS":"number","TOTAL_TAX":"number","Invoice_Value":"number",
           "Source_File":"text"}
MISS_CT  = {"GSTIN":"text","Trade_Name":"text","Invoice_No":"text","Invoice_Date":"date",
            "Taxable_Value":"number","TOTAL_TAX":"number"}

def _build_full_excel(r, s, detail_df, sup_rows, trade_name_map, tol, month_summary):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb   = writer.book
        fmts = _build_workbook_formats(wb)

        # Summary
        summary_df = post_processing_cleaner(pd.DataFrame({
            "Particulars": ["ITC - Books","ITC - GSTR-2B","Difference","ITC at Risk","Match %",
                            "Total Books","Total GSTR","Matched","Tax Diff","Missing 2B","Missing Books"],
            "Value":       [s["ITC_Books"],s["ITC_GSTR"],s["ITC_Diff"],s["ITC_at_Risk"],s["Match_%"],
                            s["Total_Books"],s["Total_GSTR"],s["Matched"],s["Tax_Diff"],
                            s["Missing_2B"],s["Missing_Books"]],
        }))
        _write_sheet(wb.add_worksheet("Summary"), "Reconciliation Summary",
                     list(summary_df.columns), summary_df.values.tolist(),
                     {"Particulars":"text","Value":"number"}, fmts)

        # Month-wise summary sheet
        if not month_summary.empty:
            ms_ct = {"Month":"text","Books ITC":"number","GSTR ITC":"number",
                     "Difference":"number","Missing 2B":"number","Missing Books":"number","Matched":"number"}
            _write_sheet(wb.add_worksheet("Month-wise Summary"), "Month-wise ITC Summary",
                         list(month_summary.columns), month_summary.values.tolist(), ms_ct, fmts)

        if not r["books_raw"].empty:
            bdf = post_processing_cleaner(r["books_raw"].copy())
            bdf_cols = [c for c in BOOK_CT if c in bdf.columns]
            _write_sheet(wb.add_worksheet("Books"), "Books Data",
                         bdf_cols, bdf[bdf_cols].values.tolist(),
                         {c: BOOK_CT[c] for c in bdf_cols}, fmts)

        if not r["gstr_raw"].empty:
            gdf = post_processing_cleaner(r["gstr_raw"].copy())
            gdf_cols = [c for c in BOOK_CT if c in gdf.columns]
            _write_sheet(wb.add_worksheet("GSTR-2B"), "GSTR-2B Data",
                         gdf_cols, gdf[gdf_cols].values.tolist(),
                         {c: BOOK_CT[c] for c in gdf_cols}, fmts)

        if not r["missing_2b"].empty:
            mdf = post_processing_cleaner(r["missing_2b"][list(MISS_CT)].copy())
            _write_sheet(wb.add_worksheet("Missing in 2B"), "Missing in 2B",
                         list(mdf.columns), mdf.values.tolist(), MISS_CT, fmts)

        if not r["missing_books"].empty:
            mbdf = post_processing_cleaner(r["missing_books"][list(MISS_CT)].copy())
            _write_sheet(wb.add_worksheet("Missing in Books"), "Missing in Books",
                         list(mbdf.columns), mbdf.values.tolist(), MISS_CT, fmts)

        if sup_rows:
            sdf = post_processing_cleaner(pd.DataFrame(sup_rows))
            _write_sheet(wb.add_worksheet("Supplier Wise ITC Summary"), "Supplier Wise ITC Summary",
                         ["GSTIN","Supplier","ITC as per Books","ITC as per 2B","ITC Difference"],
                         sdf[["GSTIN","Supplier","ITC as per Books","ITC as per 2B","ITC Difference"]].values.tolist(),
                         {"GSTIN":"text","Supplier":"text","ITC as per Books":"number",
                          "ITC as per 2B":"number","ITC Difference":"number"}, fmts)

        if not detail_df.empty:
            ddf = detail_df.copy()
            ddf["Invoice Date"] = pd.to_datetime(ddf["Date"], format="%d-%b-%Y", errors="coerce")
            # Add human-readable Month Label column for Excel readability
            ddf["Month Label"] = ddf["Month"].apply(
                lambda m: pd.to_datetime(str(m)+"-01", format="%Y-%m-%d").strftime("%B %Y")
                if pd.notna(m) and str(m) not in ("","nan") else ""
            )
            ddf = ddf.sort_values(["Month","Supplier","Date"]).reset_index(drop=True)

            # ── Main Supplier Drill Down sheet (all months, with autofilter) ──
            ws_dd = wb.add_worksheet("Supplier Drill Down")
            _xls_write_header(ws_dd, 0, 0, "Supplier Drill Down — Invoice Level Details (use Month filter to drill down)", fmts["title"])
            _ = ws_dd.set_row(1, 20)
            dd_hdrs = ["Month Label","GSTIN","Supplier","Invoice No","Invoice Date",
                       "ITC Books","ITC 2B","Difference","Remarks","Action Required"]
            for ci, h in enumerate(dd_hdrs):
                _xls_write_header(ws_dd, 1, ci, h, fmts["header"])
            for ri, row_data in ddf.iterrows():
                safe_write_text(ws_dd,   ri+2, 0, row_data.get("Month Label",""),  fmts["text"])
                safe_write_text(ws_dd,   ri+2, 1, row_data["GSTIN"],               fmts["text"])
                safe_write_text(ws_dd,   ri+2, 2, row_data["Supplier"],            fmts["text"])
                safe_write_text(ws_dd,   ri+2, 3, row_data["Invoice No"],          fmts["text"])
                _xls_write_date(ws_dd,   ri+2, 4, row_data["Invoice Date"],        fmts)
                safe_write_number(ws_dd, ri+2, 5, row_data["ITC Books"],           fmts["number"])
                safe_write_number(ws_dd, ri+2, 6, row_data["ITC 2B"],             fmts["number"])
                safe_write_number(ws_dd, ri+2, 7, row_data["Difference"],          fmts["number"])
                safe_write_text(ws_dd,   ri+2, 8, row_data["Remarks"],             fmts["text"])
                safe_write_text(ws_dd,   ri+2, 9, row_data.get("Action Required",""), fmts["text"])
            _ = ws_dd.autofilter(1, 0, 1, 9)
            for ci, w in enumerate([16,20,28,18,13,13,13,13,16,20]):
                _ = ws_dd.set_column(ci, ci, w)
            _ = ws_dd.freeze_panes(2, 0)
            _ = ws_dd.autofit()

            # ── Per-month sheets: one sheet per month for direct drill-down ──
            for _m in ddf["Month"].dropna().unique():
                _m_df  = ddf[ddf["Month"] == _m].reset_index(drop=True)
                _label = pd.to_datetime(str(_m)+"-01", format="%Y-%m-%d").strftime("%b-%Y")                          if str(_m) not in ("","nan") else "Unknown"
                # Sheet name max 31 chars, must be unique
                _sheet_name = _label[:31]
                ws_m = wb.add_worksheet(_sheet_name)
                _xls_write_header(ws_m, 0, 0,
                    f"Invoice Details — {pd.to_datetime(str(_m)+'-01', format='%Y-%m-%d').strftime('%B %Y') if str(_m) not in ('','nan') else _m}",
                    fmts["title"])
                _ = ws_m.set_row(1, 20)
                m_hdrs = ["GSTIN","Supplier","Invoice No","Invoice Date",
                          "ITC Books","ITC 2B","Difference","Remarks","Action Required"]
                for ci, h in enumerate(m_hdrs):
                    _xls_write_header(ws_m, 1, ci, h, fmts["header"])
                for ri, row_data in _m_df.iterrows():
                    safe_write_text(ws_m,   ri+2, 0, row_data["GSTIN"],               fmts["text"])
                    safe_write_text(ws_m,   ri+2, 1, row_data["Supplier"],            fmts["text"])
                    safe_write_text(ws_m,   ri+2, 2, row_data["Invoice No"],          fmts["text"])
                    _xls_write_date(ws_m,   ri+2, 3, row_data["Invoice Date"],        fmts)
                    safe_write_number(ws_m, ri+2, 4, row_data["ITC Books"],           fmts["number"])
                    safe_write_number(ws_m, ri+2, 5, row_data["ITC 2B"],             fmts["number"])
                    safe_write_number(ws_m, ri+2, 6, row_data["Difference"],          fmts["number"])
                    safe_write_text(ws_m,   ri+2, 7, row_data["Remarks"],             fmts["text"])
                    safe_write_text(ws_m,   ri+2, 8, row_data.get("Action Required",""), fmts["text"])
                _ = ws_m.autofilter(1, 0, 1, 8)
                for ci, w in enumerate([20,28,18,13,13,13,13,16,20]):
                    _ = ws_m.set_column(ci, ci, w)
                _ = ws_m.freeze_panes(2, 0)
                _ = ws_m.autofit()

        if not r["no_itc"].empty:
            nidf = post_processing_cleaner(
                r["no_itc"][["GSTIN","Trade_Name","Invoice_No","Invoice_Date","Taxable_Value","Invoice_Value"]].copy())
            _write_sheet(wb.add_worksheet("NO ITC"), "Zero ITC Invoices",
                         list(nidf.columns), nidf.values.tolist(),
                         {"GSTIN":"text","Trade_Name":"text","Invoice_No":"text",
                          "Invoice_Date":"date","Taxable_Value":"number","Invoice_Value":"number"}, fmts)

    out.seek(0)
    return out.read()


def _build_issues_excel(r, all_issues, trade_name_map):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb2   = writer.book
        fmts2 = _build_workbook_formats(wb2)

        if not r["missing_2b"].empty:
            mdf = post_processing_cleaner(r["missing_2b"][list(MISS_CT)].copy())
            _write_sheet(wb2.add_worksheet("Missing in 2B"), "Invoices Missing in GSTR-2B",
                         list(mdf.columns), mdf.values.tolist(), MISS_CT, fmts2)
        else:
            _ = wb2.add_worksheet("Missing in 2B").write(0, 0, "No invoices missing in GSTR-2B ✓", fmts2["title"])

        if not r["missing_books"].empty:
            mbdf = post_processing_cleaner(r["missing_books"][list(MISS_CT)].copy())
            _write_sheet(wb2.add_worksheet("Missing in Books"), "Invoices Missing in Books",
                         list(mbdf.columns), mbdf.values.tolist(), MISS_CT, fmts2)
        else:
            _ = wb2.add_worksheet("Missing in Books").write(0, 0, "No invoices missing in Books ✓", fmts2["title"])

        tax_diff_df = r.get("tax_diff", pd.DataFrame())
        if not tax_diff_df.empty:
            td_rows = []
            for _, row in tax_diff_df.iterrows():
                gstin    = row.get("GSTIN_2B") or row.get("GSTIN_Books", "")
                inv_no   = row.get("Invoice_No_2B") or row.get("Invoice_No_Books", "")
                inv_date = row.get("Invoice_Date_2B") or row.get("Invoice_Date_Books")
                td_rows.append({
                    "GSTIN":       str(gstin),
                    "Supplier":    str(trade_name_map.get(gstin, "")),
                    "Invoice_No":  str(inv_no),
                    "Invoice_Date":inv_date,
                    "ITC Books":   float(row.get("TOTAL_TAX_Books", 0) or 0),
                    "ITC 2B":      float(row.get("TOTAL_TAX_2B", 0) or 0),
                    "Difference":  float(row.get("TAX_DIFF", 0) or 0),
                })
            tdf = pd.DataFrame(td_rows)
            _write_sheet(wb2.add_worksheet("Tax Differences"), "Tax Amount Differences",
                         list(tdf.columns), tdf.values.tolist(),
                         {"GSTIN":"text","Supplier":"text","Invoice_No":"text","Invoice_Date":"date",
                          "ITC Books":"number","ITC 2B":"number","Difference":"number"}, fmts2)
        else:
            _ = wb2.add_worksheet("Tax Differences").write(0, 0, "No tax differences found ✓", fmts2["title"])

        if not all_issues.empty:
            idf = all_issues.copy()
            for col in idf.select_dtypes(include=["object"]).columns:
                idf[col] = idf[col].apply(lambda v: "" if (v is None or pd.isna(v)) else str(v))
            issue_ct = {c: ("date" if c=="Invoice_Date" else
                            "number" if c in ("Taxable_Value","Invoice_Value","TOTAL_TAX") else "text")
                        for c in idf.columns}
            _write_sheet(wb2.add_worksheet("Data Issues"), "Data Quality Issues",
                         list(idf.columns), idf.values.tolist(), issue_ct, fmts2)
        else:
            _ = wb2.add_worksheet("Data Issues").write(0, 0, "No data issues found ✓", fmts2["title"])

    out.seek(0)
    return out.read()


# Cache Excel bytes — building xlsxwriter workbooks on every UI interaction
# is the primary cause of the screen shake when expanding sections.
if st.session_state.get("_xl_cache_key") != _cache_key:
    st.session_state["_full_excel"]   = _build_full_excel(
        r, s, detail_df, sup_rows, trade_name_map, tol, month_summary
    )
    st.session_state["_issues_excel"] = _build_issues_excel(r, all_issues, trade_name_map)
    st.session_state["_xl_cache_key"] = _cache_key
full_excel_bytes   = st.session_state["_full_excel"]
issues_excel_bytes = st.session_state["_issues_excel"]

# ── DOWNLOAD BUTTONS ─────────────────────────────────────────────────────────
st.divider()
col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    st.download_button(
        "📥 Download Full Report",
        data=full_excel_bytes,
        file_name=f"reconciliation_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_dl2:
    st.download_button(
        "📥 Download Issues Only",
        data=issues_excel_bytes,
        file_name=f"issues_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── CSS safety net ────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stDataFrame:has(td:empty) { display: none !important; }
</style>
""", unsafe_allow_html=True)
