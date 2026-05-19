"""
GST Reconciliation App
Version: 10.1 — Professional Summary Report Design (Old Style)
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

# Professional styling for summary report
st.markdown("""
<style>
    /* Main font */
    .stApp, .stMarkdown, .stText, .stCaption, .stMetric,
    .stTabs [data-baseweb="tab"], button, label, input {
        font-family: 'Segoe UI', 'Aptos', 'Calibri', sans-serif !important;
    }
    
    /* Summary Card Styles - Professional Accounting Look */
    .summary-container {
        background: white;
        border-radius: 8px;
        padding: 0;
        margin-bottom: 20px;
        border: 1px solid #e0e0e0;
        overflow: hidden;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    
    .summary-header {
        background: #1e3a5f;
        color: white;
        padding: 12px 16px;
        font-weight: 600;
        font-size: 16px;
        border-bottom: 1px solid #2c4e7a;
    }
    
    .summary-row {
        display: flex;
        border-bottom: 1px solid #f0f0f0;
        padding: 10px 16px;
    }
    
    .summary-row:last-child {
        border-bottom: none;
    }
    
    .summary-label {
        width: 45%;
        font-weight: 500;
        color: #333;
        font-size: 14px;
    }
    
    .summary-value {
        width: 55%;
        text-align: right;
        font-weight: 600;
        font-size: 15px;
        color: #1e3a5f;
    }
    
    .summary-value-positive {
        color: #2e7d32;
        font-weight: 700;
    }
    
    .summary-value-negative {
        color: #c62828;
        font-weight: 700;
    }
    
    .summary-total-row {
        background: #f8f9fa;
        display: flex;
        border-top: 2px solid #1e3a5f;
        padding: 12px 16px;
        font-weight: 700;
    }
    
    .summary-total-label {
        width: 45%;
        font-size: 15px;
        font-weight: 700;
        color: #1e3a5f;
    }
    
    .summary-total-value {
        width: 55%;
        text-align: right;
        font-size: 16px;
        font-weight: 700;
        color: #1e3a5f;
    }
    
    /* Insight Cards - Clean and Professional */
    .insight-card {
        border-radius: 8px;
        padding: 16px 20px;
        text-align: center;
        margin-bottom: 10px;
        border: 1px solid #e0e0e0;
        background: white;
        transition: all 0.2s;
    }
    
    .insight-card .card-number {
        font-size: 28px;
        font-weight: 700;
        line-height: 1.2;
        font-family: 'Segoe UI', 'Aptos', sans-serif;
    }
    
    .insight-card .card-label {
        font-size: 12px;
        font-weight: 600;
        margin-top: 6px;
        letter-spacing: 0.3px;
        text-transform: uppercase;
        color: #666;
    }
    
    .card-green { border-top: 3px solid #2e7d32; }
    .card-green .card-number { color: #2e7d32; }
    
    .card-red { border-top: 3px solid #c62828; }
    .card-red .card-number { color: #c62828; }
    
    .card-orange { border-top: 3px solid #ed6c02; }
    .card-orange .card-number { color: #ed6c02; }
    
    .card-yellow { border-top: 3px solid #f9a825; }
    .card-yellow .card-number { color: #f9a825; }
    
    /* Month-wise table styling - Grid/Table format */
    .month-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
        margin: 10px 0;
    }
    
    .month-table th {
        background: #f5f5f5;
        border: 1px solid #ddd;
        padding: 10px 12px;
        text-align: center;
        font-weight: 600;
        color: #333;
    }
    
    .month-table td {
        border: 1px solid #ddd;
        padding: 8px 12px;
        text-align: center;
    }
    
    .month-table tr:hover {
        background-color: #fafafa;
    }
    
    .month-header {
        font-weight: 600;
        background-color: #f9f9f9;
    }
    
    /* Dataframe styling */
    .stDataFrame {
        border: 1px solid #e0e0e0 !important;
        border-radius: 8px !important;
        overflow: hidden !important;
    }
    
    .stDataFrame table {
        font-size: 13px !important;
    }
    
    .stDataFrame thead tr th {
        background: #f5f5f5 !important;
        font-weight: 600 !important;
        border-bottom: 1px solid #ddd !important;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: #f8f9fa;
        padding: 6px 6px 0 6px;
        border-radius: 8px 8px 0 0;
    }
    
    .stTabs [data-baseweb="tab"] {
        padding: 8px 20px;
        border-radius: 6px 6px 0 0;
        font-weight: 500;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #1e3a5f !important;
        color: white !important;
    }
    
    /* Divider */
    hr {
        margin: 20px 0;
        border: none;
        border-top: 1px solid #e0e0e0;
    }
    
    /* Warning box */
    .warning-box {
        background-color: #fff3e0;
        border-left: 4px solid #ed6c02;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 12px 0;
        font-size: 13px;
    }
    
    /* Info box */
    .info-box {
        background-color: #e3f2fd;
        border-left: 4px solid #1976d2;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 12px 0;
        font-size: 13px;
    }
    
    /* Success box */
    .success-box {
        background-color: #e8f5e9;
        border-left: 4px solid #2e7d32;
        padding: 12px 16px;
        border-radius: 4px;
        margin: 12px 0;
        font-size: 13px;
    }
</style>
""", unsafe_allow_html=True)

# =================== HEADER =================== #

st.title("GST Reconciliation")
st.caption("Books vs GSTR-2B · Multi-month · Tally PR, GSTR-2B Excel & Standard Template")

# =================== SAFE DISPLAY HELPER =================== #

def safe_dataframe(df, column_config=None, empty_message="No data to display", caption=None):
    if df is None or df.empty:
        st.info(empty_message)
        return False
    df = df.dropna(how="all")
    if df.empty:
        st.info(empty_message)
        return False
    if caption:
        st.caption(caption)
    st.dataframe(df, use_container_width=True, hide_index=True, column_config=column_config)
    return True

# =================== UPLOAD =================== #

col1, col2 = st.columns(2)
with col1:
    books_file = st.file_uploader("📤 Upload Books", type=["xlsx", "xls", "csv"])
with col2:
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
    raw = load_raw(uploaded_file)
    fmt = detect_file_format(raw, uploaded_file.name)
    logger.info(f"GSTR-2B format detected: {fmt} for {uploaded_file.name}")
    if fmt == "gstr2b_excel":
        return parse_gstr2b_excel(raw), fmt
    uploaded_file.seek(0)
    df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith(".csv") else \
         pd.read_excel(uploaded_file)
    return parse_gstr2b(df), "standard"

def add_month_column(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "Invoice_Date" not in df.columns:
        df["Month"] = ""
        return df
    df = df.copy()
    df["Month"] = pd.to_datetime(df["Invoice_Date"], errors="coerce") \
                    .dt.to_period("M").astype(str)
    return df

# =================== PROCESS =================== #

if st.button("🚀 Run Reconciliation", use_container_width=True, type="primary"):
    if not books_file or not gstr_files:
        st.error("Please upload both Books and at least one GSTR-2B file.")
    else:
        try:
            with st.spinner("Processing files…"):
                books_clean, no_itc, issues, books_fmt = parse_books(books_file)

                gstr_parts = []
                for gf in gstr_files:
                    gdf, gfmt = parse_gstr_single(gf)
                    gdf["Source_File"] = gf.name
                    gstr_parts.append(gdf)
                    logger.info(f"Parsed {gf.name}: {len(gdf)} rows")

                gstr_clean = pd.concat(gstr_parts, ignore_index=True) if gstr_parts else pd.DataFrame()

                books_clean = add_month_column(books_clean)
                gstr_clean = add_month_column(gstr_clean)

                results = reconcile(gstr_clean, books_clean, tolerance)
                results.update({
                    "no_itc": no_itc,
                    "issues": issues,
                    "books_raw": books_clean,
                    "gstr_raw": gstr_clean,
                    "books_fmt": books_fmt,
                    "gstr_fmts": [gf.name for gf in gstr_files],
                    "n_gstr_files": len(gstr_files),
                })
                st.session_state["results"] = results

            st.success(f"✅ Reconciliation completed! ({len(gstr_files)} GSTR-2B file(s) processed)")
        except Exception as e:
            logger.error(f"Reconciliation failed: {e}", exc_info=True)
            st.error(f"Error: {str(e)}")

# =================== DISPLAY =================== #

if "results" not in st.session_state:
    st.stop()

r = st.session_state["results"]
s = r["summary"]
tol = tolerance

# Action mapping
ACTION_MAP = {
    "✅ Matched": "No action",
    "❌ Missing in GST": "Follow up with supplier",
    "📕 Missing in Books": "Record purchase entry",
    "⚠️ Tax Difference": "Verify invoice values",
}

def add_action_column(df: pd.DataFrame) -> pd.DataFrame:
    if "Remarks" in df.columns:
        df = df.copy()
        df["Action Required"] = df["Remarks"].map(ACTION_MAP).fillna("")
    return df

def fmt_date(val) -> str:
    try:
        return pd.to_datetime(val).strftime("%d-%b-%Y") if pd.notna(val) else ""
    except:
        return ""

# =================== BUILD DETAIL DF FROM RECONCILIATION RESULTS =================== #

def build_detail_df_from_results(matched_df, missing_2b_df, missing_books_df, trade_name_map, tolerance):
    """Build invoice detail dataframe using reconciliation results (NOT raw data)"""
    rows = []
    
    # Process matched invoices
    if not matched_df.empty:
        for _, row in matched_df.iterrows():
            gstin = row.get("GSTIN_2B") or row.get("GSTIN_Books", "")
            supplier = trade_name_map.get(gstin, row.get("Trade_Name_2B") or row.get("Trade_Name_Books", ""))
            month = row.get("Month_2B") or row.get("Month_Books", "")
            inv_no = row.get("Invoice_No_2B") or row.get("Invoice_No_Books", "")
            inv_date = row.get("Invoice_Date_2B") or row.get("Invoice_Date_Books")
            itc_books = float(row.get("TOTAL_TAX_Books", 0) or 0)
            itc_gstr = float(row.get("TOTAL_TAX_2B", 0) or 0)
            diff = itc_gstr - itc_books
            remark = "✅ Matched" if abs(diff) <= tolerance else "⚠️ Tax Difference"
            
            rows.append({
                "GSTIN": gstin,
                "Supplier": supplier,
                "Month": month,
                "Invoice No": inv_no,
                "Date": fmt_date(inv_date),
                "ITC Books": itc_books,
                "ITC 2B": itc_gstr,
                "Difference": diff,
                "Remarks": remark,
                "Action Required": ACTION_MAP.get(remark, "")
            })
    
    # Process missing in GSTR-2B (in Books but not in GSTR)
    if not missing_2b_df.empty:
        for _, row in missing_2b_df.iterrows():
            gstin = row.get("GSTIN", "")
            supplier = trade_name_map.get(gstin, row.get("Trade_Name", ""))
            itc_books = float(row.get("TOTAL_TAX", 0) or 0)
            remark = "❌ Missing in GST"
            
            rows.append({
                "GSTIN": gstin,
                "Supplier": supplier,
                "Month": row.get("Month", ""),
                "Invoice No": row.get("Invoice_No", ""),
                "Date": fmt_date(row.get("Invoice_Date")),
                "ITC Books": itc_books,
                "ITC 2B": 0.0,
                "Difference": -itc_books,
                "Remarks": remark,
                "Action Required": ACTION_MAP.get(remark, "")
            })
    
    # Process missing in Books (in GSTR but not in Books)
    if not missing_books_df.empty:
        for _, row in missing_books_df.iterrows():
            gstin = row.get("GSTIN", "")
            supplier = trade_name_map.get(gstin, row.get("Trade_Name", ""))
            itc_gstr = float(row.get("TOTAL_TAX", 0) or 0)
            remark = "📕 Missing in Books"
            
            rows.append({
                "GSTIN": gstin,
                "Supplier": supplier,
                "Month": row.get("Month", ""),
                "Invoice No": row.get("Invoice_No", ""),
                "Date": fmt_date(row.get("Invoice_Date")),
                "ITC Books": 0.0,
                "ITC 2B": itc_gstr,
                "Difference": itc_gstr,
                "Remarks": remark,
                "Action Required": ACTION_MAP.get(remark, "")
            })
    
    return pd.DataFrame(rows)

# =================== BUILD MONTH SUMMARY =================== #

def build_month_summary(matched_df, missing_2b_df, missing_books_df, books_raw, gstr_raw):
    """Build month-wise summary using reconciliation results"""
    rows = []
    
    # Get all months from both raw data
    all_months = set()
    if not books_raw.empty and "Month" in books_raw.columns:
        all_months.update(books_raw["Month"].dropna().unique())
    if not gstr_raw.empty and "Month" in gstr_raw.columns:
        all_months.update(gstr_raw["Month"].dropna().unique())
    all_months.discard("")
    all_months.discard("NaT")
    
    for month in sorted(all_months):
        # Books ITC from raw data
        b_tax = books_raw[books_raw["Month"] == month]["TOTAL_TAX"].sum() if not books_raw.empty else 0
        # GSTR ITC from raw data
        g_tax = gstr_raw[gstr_raw["Month"] == month]["TOTAL_TAX"].sum() if not gstr_raw.empty else 0
        
        # Missing counts from reconciliation results
        m2b_count = 0
        if not missing_2b_df.empty and "Invoice_Date" in missing_2b_df.columns:
            m2b = missing_2b_df.copy()
            m2b["_m"] = pd.to_datetime(m2b["Invoice_Date"], errors="coerce").dt.to_period("M").astype(str)
            m2b_count = int((m2b["_m"] == month).sum())
        
        mb_count = 0
        if not missing_books_df.empty and "Invoice_Date" in missing_books_df.columns:
            mb = missing_books_df.copy()
            mb["_m"] = pd.to_datetime(mb["Invoice_Date"], errors="coerce").dt.to_period("M").astype(str)
            mb_count = int((mb["_m"] == month).sum())
        
        matched_count = 0
        if not matched_df.empty:
            for col in ["Invoice_Date_2B", "Invoice_Date_Books"]:
                if col in matched_df.columns:
                    mc = matched_df.copy()
                    mc["_m"] = pd.to_datetime(mc[col], errors="coerce").dt.to_period("M").astype(str)
                    matched_count = int((mc["_m"] == month).sum())
                    break
        
        rows.append({
            "Month": month,
            "Books ITC": round(float(b_tax), 2),
            "GSTR ITC": round(float(g_tax), 2),
            "Difference": round(float(g_tax) - float(b_tax), 2),
            "Missing 2B": m2b_count,
            "Missing Books": mb_count,
            "Matched": matched_count,
        })
    
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df["_sort"] = pd.to_datetime(df["Month"], format="%Y-%m", errors="coerce")
    df = df.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
    return df

# =================== FILTER HELPER =================== #

def apply_filters(df: pd.DataFrame, f_gstin: str, f_supplier: str, f_status: list, selected_month: str) -> pd.DataFrame:
    if df.empty:
        return df
    
    if selected_month != "All months" and "Month" in df.columns:
        df = df[df["Month"] == selected_month]
    
    if f_gstin and "GSTIN" in df.columns:
        df = df[df["GSTIN"].str.contains(f_gstin, case=False, na=False)]
    
    if f_supplier and "Supplier" in df.columns:
        df = df[df["Supplier"].str.contains(f_supplier, case=False, na=False)]
    
    if f_status and "Remarks" in df.columns:
        df = df[df["Remarks"].isin(f_status)]
    
    return df

# =================== SUMMARY REPORT - PROFESSIONAL STYLE =================== #

st.markdown("## 📊 Reconciliation Summary")

n_files = r.get("n_gstr_files", 1)
if n_files > 1:
    st.caption(f"Across {n_files} GSTR-2B files: {', '.join(r.get('gstr_fmts', []))}")

# Create two columns for layout
col_left, col_right = st.columns([1, 1])

with col_left:
    # Summary Card - Professional boxed style
    st.markdown("""
    <div class="summary-container">
        <div class="summary-header">RECONCILIATION SUMMARY</div>
        <div class="summary-row">
            <div class="summary-label">ITC - Books</div>
            <div class="summary-value">₹ {:,.2f}</div>
        </div>
        <div class="summary-row">
            <div class="summary-label">ITC - GSTR-2B</div>
            <div class="summary-value">₹ {:,.2f}</div>
        </div>
        <div class="summary-row">
            <div class="summary-label">Difference</div>
            <div class="summary-value {}">₹ {:,.2f}</div>
        </div>
        <div class="summary-row">
            <div class="summary-label">ITC at Risk</div>
            <div class="summary-value summary-value-negative">₹ {:,.2f}</div>
        </div>
        <div class="summary-total-row">
            <div class="summary-total-label">Match %</div>
            <div class="summary-total-value">{:.2f}%</div>
        </div>
    </div>
    """.format(
        s["ITC_Books"],
        s["ITC_GSTR"],
        "summary-value-positive" if s["ITC_Diff"] >= 0 else "summary-value-negative",
        abs(s["ITC_Diff"]),
        s["ITC_at_Risk"],
        s["Match_%"]
    ), unsafe_allow_html=True)

with col_right:
    # Counts Card
    st.markdown("""
    <div class="summary-container">
        <div class="summary-header">INVOICE COUNTS</div>
        <div class="summary-row">
            <div class="summary-label">Total Books</div>
            <div class="summary-value">{:,}</div>
        </div>
        <div class="summary-row">
            <div class="summary-label">Total GSTR-2B</div>
            <div class="summary-value">{:,}</div>
        </div>
        <div class="summary-row">
            <div class="summary-label">✅ Matched</div>
            <div class="summary-value summary-value-positive">{:,}</div>
        </div>
        <div class="summary-row">
            <div class="summary-label">⚠️ Tax Difference</div>
            <div class="summary-value summary-value-negative">{:,}</div>
        </div>
        <div class="summary-row">
            <div class="summary-label">❌ Missing in 2B</div>
            <div class="summary-value summary-value-negative">{:,}</div>
        </div>
        <div class="summary-row">
            <div class="summary-label">📕 Missing in Books</div>
            <div class="summary-value summary-value-negative">{:,}</div>
        </div>
    </div>
    """.format(
        s["Total_Books"],
        s["Total_GSTR"],
        s["Matched"],
        s["Tax_Diff"],
        s["Missing_2B"],
        s["Missing_Books"]
    ), unsafe_allow_html=True)

# Quick insight cards (4 across)
st.markdown("---")
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown("""
    <div class="insight-card card-green">
        <div class="card-number">{}</div>
        <div class="card-label">✅ MATCHED</div>
    </div>
    """.format(s["Matched"]), unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div class="insight-card card-red">
        <div class="card-number">{}</div>
        <div class="card-label">❌ MISSING IN 2B</div>
    </div>
    """.format(s["Missing_2B"]), unsafe_allow_html=True)

with c3:
    st.markdown("""
    <div class="insight-card card-orange">
        <div class="card-number">{}</div>
        <div class="card-label">📕 MISSING IN BOOKS</div>
    </div>
    """.format(s["Missing_Books"]), unsafe_allow_html=True)

with c4:
    st.markdown("""
    <div class="insight-card card-yellow">
        <div class="card-number">{}</div>
        <div class="card-label">⚠️ TAX DIFFERENCES</div>
    </div>
    """.format(s["Tax_Diff"]), unsafe_allow_html=True)

# =================== BUILD DATA FROM RECONCILIATION RESULTS =================== #

trade_name_map = r.get("trade_name_mapping", {})
detail_df = build_detail_df_from_results(
    r.get("matched", pd.DataFrame()),
    r.get("missing_2b", pd.DataFrame()),
    r.get("missing_books", pd.DataFrame()),
    trade_name_map,
    tolerance
)

month_summary = build_month_summary(
    r.get("matched", pd.DataFrame()),
    r.get("missing_2b", pd.DataFrame()),
    r.get("missing_books", pd.DataFrame()),
    r.get("books_raw", pd.DataFrame()),
    r.get("gstr_raw", pd.DataFrame())
)

# =================== MONTH-WISE SUMMARY (TABLE FORMAT) =================== #

if not month_summary.empty:
    st.markdown("---")
    st.markdown("## 📅 Month-wise Summary")
    
    # Convert month format for display
    month_summary_display = month_summary.copy()
    month_summary_display["Month"] = month_summary_display["Month"].apply(
        lambda m: pd.to_datetime(m + "-01", format="%Y-%m-%d").strftime("%B %Y") if pd.notna(m) and m not in ("", "NaT") else m
    )
    
    # Style the dataframe for professional table look
    styled_df = month_summary_display.style.format({
        "Books ITC": "{:,.2f}",
        "GSTR ITC": "{:,.2f}",
        "Difference": "{:,.2f}"
    }).set_properties(**{
        'text-align': 'center',
        'padding': '8px 12px',
        'font-size': '13px'
    }).set_table_styles([
        {'selector': 'thead tr th', 'props': [('background', '#f5f5f5'), ('font-weight', '600'), ('border', '1px solid #ddd'), ('padding', '10px 12px')]},
        {'selector': 'tbody tr td', 'props': [('border', '1px solid #e0e0e0'), ('padding', '8px 12px')]},
        {'selector': 'tbody tr:hover', 'props': [('background', '#fafafa')]},
        {'selector': 'table', 'props': [('border-collapse', 'collapse'), ('width', '100%'), ('border', '1px solid #ddd'), ('border-radius', '8px')]}
    ])
    
    st.dataframe(styled_df, use_container_width=True, hide_index=True)
    st.caption(f"{len(month_summary)} month(s) of data")

# =================== FILTER CONTROLS =================== #

st.markdown("---")
st.markdown("### 🔍 Filter Results")
filter_col1, filter_col2, filter_col3, filter_col4 = st.columns([1, 1, 1, 1])

with filter_col1:
    month_options = ["All months"] + list(month_summary["Month"]) if not month_summary.empty else ["All months"]
    selected_month = st.selectbox("📅 Month", month_options)

with filter_col2:
    f_gstin = st.text_input("GSTIN", placeholder="Enter GSTIN...")

with filter_col3:
    f_supplier = st.text_input("Supplier", placeholder="Enter supplier name...")

with filter_col4:
    f_status = st.multiselect(
        "Status",
        options=["✅ Matched", "❌ Missing in GST", "📕 Missing in Books", "⚠️ Tax Difference"],
        default=[]
    )

# =================== TABS =================== #

# Data issues
all_issues = r["issues"].copy() if not r["issues"].empty else pd.DataFrame()
if "duplicate_issues" in r and not r["duplicate_issues"].empty:
    all_issues = pd.concat([x for x in [all_issues, r["duplicate_issues"]] if not x.empty], ignore_index=True)

tabs = st.tabs([
    "📋 Invoice Details",
    "❌ Missing in 2B",
    "📕 Missing in Books",
    "📚 Books Raw",
    "📊 GSTR-2B Raw",
    "⚠️ Data Issues"
])

# Tab 0: Invoice Details
with tabs[0]:
    filtered_df = apply_filters(detail_df, f_gstin, f_supplier, f_status, selected_month)
    filtered_df = add_action_column(filtered_df)
    
    col_config = {
        "Month": st.column_config.TextColumn("Month", width=90),
        "GSTIN": st.column_config.TextColumn("GSTIN", width=180),
        "Supplier": st.column_config.TextColumn("Supplier", width=250),
        "Invoice No": st.column_config.TextColumn("Invoice No", width=140),
        "Date": st.column_config.TextColumn("Date", width=105),
        "ITC Books": st.column_config.NumberColumn("📚 Books", width=110, format="%.2f"),
        "ITC 2B": st.column_config.NumberColumn("📊 GSTR-2B", width=110, format="%.2f"),
        "Difference": st.column_config.NumberColumn("📉 Diff", width=100, format="%.2f"),
        "Remarks": st.column_config.TextColumn("Remarks", width=160),
        "Action Required": st.column_config.TextColumn("Action Required", width=180),
    }
    safe_dataframe(filtered_df, column_config=col_config, 
                   empty_message="No invoices match the current filters.",
                   caption=f"Showing {len(filtered_df)} of {len(detail_df)} invoices")

# Tab 1: Missing in 2B
with tabs[1]:
    missing_2b_df = r.get("missing_2b", pd.DataFrame()).copy()
    if not missing_2b_df.empty:
        missing_2b_df["Date"] = missing_2b_df["Invoice_Date"].apply(fmt_date)
        missing_2b_df["Supplier"] = missing_2b_df["GSTIN"].map(trade_name_map).fillna(missing_2b_df["Trade_Name"])
        display_df = missing_2b_df[["GSTIN", "Supplier", "Invoice_No", "Date", "Taxable_Value", "TOTAL_TAX"]]
        display_df.columns = ["GSTIN", "Supplier", "Invoice No", "Date", "Taxable", "ITC"]
        
        if selected_month != "All months":
            display_df = display_df[display_df["Month"] == selected_month] if "Month" in display_df.columns else display_df
        
        safe_dataframe(display_df, empty_message="No invoices missing in GSTR-2B",
                       caption=f"💰 ITC at Risk: ₹{display_df['ITC'].sum():,.2f} across {len(display_df)} invoices")
    else:
        st.success("✅ No invoices missing in GSTR-2B")

# Tab 2: Missing in Books
with tabs[2]:
    missing_books_df = r.get("missing_books", pd.DataFrame()).copy()
    if not missing_books_df.empty:
        missing_books_df["Date"] = missing_books_df["Invoice_Date"].apply(fmt_date)
        missing_books_df["Supplier"] = missing_books_df["GSTIN"].map(trade_name_map).fillna(missing_books_df["Trade_Name"])
        display_df = missing_books_df[["GSTIN", "Supplier", "Invoice_No", "Date", "Taxable_Value", "TOTAL_TAX"]]
        display_df.columns = ["GSTIN", "Supplier", "Invoice No", "Date", "Taxable", "ITC"]
        
        safe_dataframe(display_df, empty_message="No invoices missing in Books",
                       caption=f"{len(display_df)} invoices in GSTR-2B not found in Books")
    else:
        st.success("✅ No invoices missing in Books")

# Tab 3: Books Raw
with tabs[3]:
    books_raw = r.get("books_raw", pd.DataFrame()).copy()
    if not books_raw.empty:
        books_raw["Invoice_Date"] = books_raw["Invoice_Date"].apply(fmt_date)
        safe_dataframe(books_raw, empty_message="No Books data")

# Tab 4: GSTR-2B Raw
with tabs[4]:
    gstr_raw = r.get("gstr_raw", pd.DataFrame()).copy()
    if not gstr_raw.empty:
        gstr_raw["Invoice_Date"] = gstr_raw["Invoice_Date"].apply(fmt_date)
        safe_dataframe(gstr_raw, empty_message="No GSTR-2B data")

# Tab 5: Data Issues
with tabs[5]:
    if all_issues.empty:
        st.success("✅ No data issues found.")
    else:
        display_issues = all_issues.copy()
        if "Invoice_Date" in display_issues.columns:
            display_issues["Invoice_Date"] = display_issues["Invoice_Date"].apply(fmt_date)
        st.dataframe(display_issues, use_container_width=True, hide_index=True)

# =================== ZERO ITC =================== #

if not r["no_itc"].empty:
    no_itc_df = r["no_itc"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "Invoice_Value"]].copy()
    no_itc_df["Invoice_Date"] = no_itc_df["Invoice_Date"].apply(fmt_date)
    no_itc_df = no_itc_df[no_itc_df["Invoice_Value"].astype(float) > 0]
    if not no_itc_df.empty:
        st.markdown("---")
        st.markdown("## 🟡 Zero ITC Invoices")
        st.dataframe(no_itc_df, use_container_width=True, hide_index=True)

# =================== DOWNLOAD BUTTONS =================== #

st.markdown("---")
col_dl1, col_dl2 = st.columns(2)

def export_to_excel(r, detail_df, month_summary, trade_name_map):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_df.to_excel(writer, sheet_name="Invoice Details", index=False)
        month_summary.to_excel(writer, sheet_name="Month-wise Summary", index=False)
        
        if not r.get("matched", pd.DataFrame()).empty:
            r["matched"].to_excel(writer, sheet_name="Matched", index=False)
        if not r.get("missing_2b", pd.DataFrame()).empty:
            r["missing_2b"].to_excel(writer, sheet_name="Missing in 2B", index=False)
        if not r.get("missing_books", pd.DataFrame()).empty:
            r["missing_books"].to_excel(writer, sheet_name="Missing in Books", index=False)
        
        summary_df = pd.DataFrame([r["summary"]])
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
    
    return output.getvalue()

with col_dl1:
    st.download_button(
        "📥 Download Full Report",
        data=export_to_excel(r, detail_df, month_summary, trade_name_map),
        file_name=f"reconciliation_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_dl2:
    def export_issues(r, all_issues):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            if not r.get("missing_2b", pd.DataFrame()).empty:
                r["missing_2b"].to_excel(writer, sheet_name="Missing in 2B", index=False)
            if not r.get("missing_books", pd.DataFrame()).empty:
                r["missing_books"].to_excel(writer, sheet_name="Missing in Books", index=False)
            if not r.get("tax_diff", pd.DataFrame()).empty:
                r["tax_diff"].to_excel(writer, sheet_name="Tax Differences", index=False)
            if not all_issues.empty:
                all_issues.to_excel(writer, sheet_name="Data Issues", index=False)
        return output.getvalue()
    
    st.download_button(
        "📥 Download Issues Only",
        data=export_issues(r, all_issues),
        file_name=f"issues_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
