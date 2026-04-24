"""
GST Reconciliation App
Version: 8.0 — Premium UI redesign
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
    /* Base font */
    .stApp, .stMarkdown, .stText, .stCaption, .stMetric,
    .stTabs [data-baseweb="tab"], button, label, input {
        font-family: 'Aptos Narrow', 'Aptos', sans-serif !important;
    }

    /* Insight cards */
    .insight-card {
        border-radius: 12px;
        padding: 20px 24px;
        text-align: center;
        margin-bottom: 8px;
    }
    .insight-card .card-number {
        font-size: 2.4rem;
        font-weight: 700;
        line-height: 1.1;
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
    }
    .insight-card .card-label {
        font-size: 0.85rem;
        font-weight: 600;
        margin-top: 4px;
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
    }
    .card-green  { background:#DCFCE7; }
    .card-green  .card-number { color:#166534; }
    .card-green  .card-label  { color:#166534; }
    .card-red    { background:#FEE2E2; }
    .card-red    .card-number { color:#991B1B; }
    .card-red    .card-label  { color:#991B1B; }
    .card-orange { background:#FEF3C7; }
    .card-orange .card-number { color:#92400E; }
    .card-orange .card-label  { color:#92400E; }
    .card-yellow { background:#FFEDD5; }
    .card-yellow .card-number { color:#9A3412; }
    .card-yellow .card-label  { color:#9A3412; }

    /* Warning box for data issues */
    .warning-box {
        background-color: #FEF2F2;
        padding: 1rem;
        border-left: 4px solid #DC2626;
        border-radius: 4px;
        margin: 1rem 0;
    }

    /* Filter bar */
    .filter-section {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 16px 20px;
        margin: 12px 0;
    }

    /* Remark chips in tables */
    .remark-matched   { background:#DCFCE7; color:#166534; padding:2px 8px; border-radius:8px; font-size:0.8rem; }
    .remark-missing-g { background:#FEE2E2; color:#991B1B; padding:2px 8px; border-radius:8px; font-size:0.8rem; }
    .remark-missing-b { background:#FEF3C7; color:#92400E; padding:2px 8px; border-radius:8px; font-size:0.8rem; }
    .remark-taxdiff   { background:#FFEDD5; color:#9A3412; padding:2px 8px; border-radius:8px; font-size:0.8rem; }

    .stTabs [data-baseweb="tab-list"] { gap: 2px; }
    .stTabs [data-baseweb="tab"]      { padding: 8px 16px; }
    .streamlit-expanderHeader          { font-size: 1rem !important; }

    /* Download buttons row */
    .download-row { display:flex; gap:12px; margin-top:16px; }
</style>
""", unsafe_allow_html=True)

# =================== HEADER =================== #

st.title("GST Reconciliation")
st.caption("Books vs GSTR-2B · Supports Tally PR, GSTR-2B Excel & Standard Template")

# =================== SAFE DISPLAY HELPER =================== #

def safe_dataframe(df, column_config=None, empty_message="No data to display", caption=None):
    """Render a dataframe only when it has real rows. Shows info message otherwise."""
    if df is None or df.empty:
        st.info(empty_message)
        return False
    # Drop rows that are entirely empty / all-NaN
    df = df.dropna(how="all")
    # Drop rows where Invoice_No / GSTIN are blank placeholders
    for col in ["Invoice_No", "Invoice No"]:
        if col in df.columns:
            mask = df[col].astype(str).str.strip().str.lower() != "nan"
            mask = mask & (df[col].astype(str).str.strip() != "")
            df = df[mask]
    if df.empty:
        st.info(empty_message)
        return False
    # Guard: don't render a table that is nothing but zeros/empty values
    numeric_cols = df.select_dtypes(include=["number"]).columns
    str_cols     = df.select_dtypes(exclude=["number"]).columns
    has_numbers  = len(numeric_cols) > 0 and df[numeric_cols].abs().sum().sum() > 0
    _EMPTIES     = {"", "nan", "none", "0", "0.0"}
    has_strings  = False
    if len(str_cols) > 0:
        for col in str_cols:
            if df[col].astype(str).str.strip().str.lower().apply(lambda v: v not in _EMPTIES).any():
                has_strings = True
                break
    if not has_numbers and not has_strings:
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
    gstr_file = st.file_uploader("📤 Upload GSTR-2B", type=["xlsx", "xls", "csv"])

tolerance = st.number_input(
    "Tolerance (₹)",
    value=DEFAULT_TOLERANCE,
    step=0.5,
    min_value=0.0,
    help="Maximum acceptable tax difference (in ₹) between Books and GSTR-2B for an invoice to be considered 'Matched'. Default: ₹1.00"
)

# =================== PARSER ROUTER =================== #

def load_raw(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, header=None)
    elif name.endswith(".xls"):
        return pd.read_excel(uploaded_file, header=None)
    else:
        return pd.read_excel(uploaded_file, header=None)

def parse_books(uploaded_file):
    raw = load_raw(uploaded_file)
    fmt = detect_file_format(raw, uploaded_file.name)
    logger.info(f"Books format detected: {fmt}")          # log only — no UI badge
    if fmt == "tally_pr":
        return (*parse_tally_purchase_register(raw), fmt)
    uploaded_file.seek(0)
    df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith(".csv") else \
         pd.read_excel(uploaded_file)
    return (*parse_tally(df), "standard")

def parse_gstr(uploaded_file):
    raw = load_raw(uploaded_file)
    fmt = detect_file_format(raw, uploaded_file.name)
    logger.info(f"GSTR-2B format detected: {fmt}")        # log only — no UI badge
    if fmt == "gstr2b_excel":
        return parse_gstr2b_excel(raw), fmt
    uploaded_file.seek(0)
    df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith(".csv") else \
         pd.read_excel(uploaded_file)
    return parse_gstr2b(df), "standard"

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
    if not books_file or not gstr_file:
        st.error("Please upload both files.")
    else:
        try:
            with st.spinner("Processing files…"):
                books_clean, no_itc, issues, books_fmt = parse_books(books_file)
                gstr_clean, gstr_fmt                   = parse_gstr(gstr_file)
                results = reconcile(gstr_clean, books_clean, tolerance)
                results.update({
                    "no_itc":    no_itc,
                    "issues":    issues,
                    "books_raw": books_clean,
                    "gstr_raw":  gstr_clean,
                    "books_fmt": books_fmt,
                    "gstr_fmt":  gstr_fmt,
                })
                st.session_state["results"] = results
            st.success("✅ Reconciliation completed!")
        except Exception as e:
            logger.error(f"Reconciliation failed: {e}", exc_info=True)
            st.error(f"Error: {str(e)}")

# =================== DISPLAY =================== #

if "results" not in st.session_state:
    st.stop()

r   = st.session_state["results"]
s   = r["summary"]
tol = tolerance   # use current slider value for display logic

# ── helpers ─────────────────────────────────────────────────────────────────

def fmt_date(val) -> str:
    try:
        return pd.to_datetime(val).strftime("%d-%b-%Y") if pd.notna(val) else ""
    except Exception:
        return ""

def coerce_str_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Convert only truly missing values (None/NaN) to empty strings. Real data is untouched."""
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

# ── DATA ISSUES ──────────────────────────────────────────────────────────────

all_issues = r["issues"].copy() if not r["issues"].empty else pd.DataFrame()
if "duplicate_issues" in r and not r["duplicate_issues"].empty:
    all_issues = pd.concat(
        [x for x in [all_issues, r["duplicate_issues"]] if not x.empty],
        ignore_index=True
    )

if not all_issues.empty:
    st.markdown(f"""
    <div class="warning-box">
        <strong>⚠️ {len(all_issues)} Data Issues Found</strong> — Fix these in source data before reconciling.
    </div>""", unsafe_allow_html=True)
    with st.expander("🔍 View Data Issues", expanded=False):
        df_iss = all_issues.copy()
        if "Invoice_Date" in df_iss.columns:
            df_iss["Invoice_Date"] = pd.to_datetime(df_iss["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y").fillna("")
        df_iss = coerce_str_cols(df_iss)
        cols_order = ["Issue"] + [c for c in df_iss.columns if c != "Issue"]
        safe_dataframe(
            df_iss[cols_order],
            column_config={
                "Issue":         st.column_config.TextColumn("Issue",        width=220),
                "GSTIN":         st.column_config.TextColumn("GSTIN",        width=180),
                "Trade_Name":    st.column_config.TextColumn("Trade Name",   width=280),
                "Invoice_No":    st.column_config.TextColumn("Invoice No",   width=150),
                "Invoice_Date":  st.column_config.TextColumn("Invoice Date", width=120),
                "Taxable_Value": st.column_config.NumberColumn("Taxable",    width=110, format="%.2f"),
                "TOTAL_TAX":     st.column_config.NumberColumn("Total Tax",  width=110, format="%.2f"),
            },
            empty_message="No issues to display.",
        )
        ic = df_iss["Issue"].value_counts().reset_index()
        ic.columns = ["Issue Type", "Count"]
        if not ic.empty:
            st.dataframe(ic, use_container_width=True, hide_index=True)

# ── QUICK INSIGHT CARDS ───────────────────────────────────────────────────────

st.markdown("## 📊 Reconciliation Summary")

c1, c2, c3, c4 = st.columns(4)
cards = [
    (c1, "card-green",  "✅ Matched",            s["Matched"]),
    (c2, "card-red",    "❌ Missing in GSTR-2B",  s["Missing_2B"]),
    (c3, "card-orange", "📕 Missing in Books",    s["Missing_Books"]),
    (c4, "card-yellow", "⚠️ Tax Difference",      s["Tax_Diff"]),
]
for col, css, label, value in cards:
    with col:
        st.markdown(f"""
        <div class="insight-card {css}">
            <div class="card-number">{value}</div>
            <div class="card-label">{label}</div>
        </div>""", unsafe_allow_html=True)

# ── FILTER BAR ────────────────────────────────────────────────────────────────

st.markdown("### 🔍 Filter Results")
with st.container():
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        f_gstin    = st.text_input("Filter by GSTIN",     placeholder="Enter GSTIN…",         label_visibility="visible")
    with fc2:
        f_supplier = st.text_input("Filter by Supplier",  placeholder="Enter supplier name…",  label_visibility="visible")
    with fc3:
        f_status   = st.multiselect(
            "Filter by Status",
            options=["✅ Matched", "❌ Missing in GST", "📕 Missing in Books", "⚠️ Tax Difference"],
            default=[],
            placeholder="All statuses…",
        )

# ── BUILD INVOICE LEVEL DETAIL ────────────────────────────────────────────────

trade_name_map = r.get("trade_name_mapping", {})
detail_data    = []
processed_keys = set()

for _, row in r["books_raw"].iterrows():
    key = f"{row['GSTIN']}|{row['Invoice_No']}"
    if key in processed_keys:
        continue
    processed_keys.add(key)
    gstin    = row["GSTIN"]
    supplier = trade_name_map.get(gstin, row["Trade_Name"])
    m_row    = None
    if not r["gstr_raw"].empty:
        m = r["gstr_raw"][(r["gstr_raw"]["GSTIN"] == gstin) & (r["gstr_raw"]["Invoice_No"] == row["Invoice_No"])]
        if not m.empty:
            m_row = m.iloc[0]

    if m_row is not None:
        diff   = float(m_row["TOTAL_TAX"]) - float(row["TOTAL_TAX"])
        remark = "✅ Matched" if abs(diff) <= tol else "⚠️ Tax Difference"
        detail_data.append({"GSTIN": gstin, "Supplier": supplier,
            "Invoice No": row["Invoice_No"], "Date": fmt_date(row["Invoice_Date"]),
            "ITC Books": float(row["TOTAL_TAX"]), "ITC 2B": float(m_row["TOTAL_TAX"]),
            "Difference": diff, "Remarks": remark})
    else:
        detail_data.append({"GSTIN": gstin, "Supplier": supplier,
            "Invoice No": row["Invoice_No"], "Date": fmt_date(row["Invoice_Date"]),
            "ITC Books": float(row["TOTAL_TAX"]), "ITC 2B": 0.0,
            "Difference": -float(row["TOTAL_TAX"]), "Remarks": "❌ Missing in GST"})

for _, row in r["gstr_raw"].iterrows():
    key = f"{row['GSTIN']}|{row['Invoice_No']}"
    if key in processed_keys:
        continue
    processed_keys.add(key)
    detail_data.append({"GSTIN": row["GSTIN"],
        "Supplier": trade_name_map.get(row["GSTIN"], row["Trade_Name"]),
        "Invoice No": row["Invoice_No"], "Date": fmt_date(row["Invoice_Date"]),
        "ITC Books": 0.0, "ITC 2B": float(row["TOTAL_TAX"]),
        "Difference": float(row["TOTAL_TAX"]), "Remarks": "📕 Missing in Books"})

detail_df = pd.DataFrame(detail_data).sort_values(["Supplier", "Date"]) if detail_data else pd.DataFrame()

def filter_detail(df):
    """Apply GSTIN, supplier, and status filters to detail dataframe."""
    if df.empty:
        return df
    if f_gstin:
        df = df[df["GSTIN"].str.contains(f_gstin, case=False, na=False)]
    if f_supplier:
        df = df[df["Supplier"].str.contains(f_supplier, case=False, na=False)]
    if f_status:
        status_map = {
            "✅ Matched":           "✅ Matched",
            "❌ Missing in GST":    "❌ Missing in GST",
            "📕 Missing in Books":  "📕 Missing in Books",
            "⚠️ Tax Difference":    "⚠️ Tax Difference",
        }
        allowed = [status_map[s] for s in f_status if s in status_map]
        df = df[df["Remarks"].isin(allowed)]
    return df

DETAIL_COL_CFG = {
    "GSTIN":       st.column_config.TextColumn("GSTIN",       width=180),
    "Supplier":    st.column_config.TextColumn("Supplier",    width=280),
    "Invoice No":  st.column_config.TextColumn("Invoice No",  width=150),
    "Date":        st.column_config.TextColumn("Date",        width=115),
    "ITC Books":   st.column_config.NumberColumn("📚 Books",  width=120, format="%.2f"),
    "ITC 2B":      st.column_config.NumberColumn("📊 GSTR-2B",width=120, format="%.2f"),
    "Difference":  st.column_config.NumberColumn("📉 Diff",   width=110, format="%.2f"),
    "Remarks":     st.column_config.TextColumn("Remarks",     width=170),
}

STD_BOOK_CFG = {
    "GSTIN":         st.column_config.TextColumn("GSTIN",         width=180),
    "Trade_Name":    st.column_config.TextColumn("Trade Name",    width=280),
    "Invoice_No":    st.column_config.TextColumn("Invoice No",    width=150),
    "Invoice_Date":  st.column_config.TextColumn("Invoice Date",  width=115),
    "Taxable_Value": st.column_config.NumberColumn("Taxable",     width=110, format="%.2f"),
    "CGST":          st.column_config.NumberColumn("CGST",        width=100, format="%.2f"),
    "SGST":          st.column_config.NumberColumn("SGST",        width=100, format="%.2f"),
    "IGST":          st.column_config.NumberColumn("IGST",        width=100, format="%.2f"),
    "CESS":          st.column_config.NumberColumn("CESS",        width=80,  format="%.2f"),
    "TOTAL_TAX":     st.column_config.NumberColumn("Total Tax",   width=110, format="%.2f"),
    "Invoice_Value": st.column_config.NumberColumn("Invoice Value",width=120, format="%.2f"),
}

MISS_CFG = {
    "GSTIN":      st.column_config.TextColumn("GSTIN",      width=180),
    "Supplier":   st.column_config.TextColumn("Supplier",   width=280),
    "Invoice No": st.column_config.TextColumn("Invoice No", width=150),
    "Date":       st.column_config.TextColumn("Date",       width=115),
    "Taxable":    st.column_config.NumberColumn("Taxable",  width=110, format="%.2f"),
    "ITC":        st.column_config.NumberColumn("ITC",      width=110, format="%.2f"),
}

# ── TABS ─────────────────────────────────────────────────────────────────────

tabs = st.tabs([
    "📋 Invoice Details",
    "📚 Books",
    "📊 GSTR-2B",
    "❌ Missing in 2B",
    "📕 Missing in Books",
    "📋 Supplier Summary",
])

# Tab 0: Invoice Level Details
with tabs[0]:
    if not detail_df.empty:
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
        df_b["Invoice_Date"] = pd.to_datetime(df_b["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y").fillna("")
        df_b = apply_filters(coerce_str_cols(df_b), f_gstin, f_supplier)
        safe_dataframe(
            df_b,
            column_config=STD_BOOK_CFG,
            empty_message="No Books data matches the current filters.",
            caption=f"{len(df_b)} records",
        )
    else:
        st.info("No Books data loaded.")

# Tab 2: GSTR-2B
with tabs[2]:
    if not r["gstr_raw"].empty:
        df_g = r["gstr_raw"].copy()
        df_g["Invoice_Date"] = pd.to_datetime(df_g["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y").fillna("")
        df_g = apply_filters(coerce_str_cols(df_g), f_gstin, f_supplier)
        safe_dataframe(
            df_g,
            column_config=STD_BOOK_CFG,
            empty_message="No GSTR-2B data matches the current filters.",
            caption=f"{len(df_g)} records",
        )
    else:
        st.info("No GSTR-2B data loaded.")

# Tab 3: Missing in 2B
with tabs[3]:
    if not r["missing_2b"].empty:
        df_m = r["missing_2b"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "TOTAL_TAX"]].copy()
        df_m["Invoice_Date"] = pd.to_datetime(df_m["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y").fillna("")
        df_m.columns = ["GSTIN", "Supplier", "Invoice No", "Date", "Taxable", "ITC"]
        df_m = apply_filters(coerce_str_cols(df_m), f_gstin, f_supplier)
        safe_dataframe(
            df_m,
            column_config=MISS_CFG,
            empty_message="No missing invoices match the current filters.",
            caption=f"💰 ITC at Risk: ₹{df_m['ITC'].sum():,.2f} across {len(df_m)} invoices",
        )
    else:
        st.success("✅ No invoices missing in GSTR-2B.")

# Tab 4: Missing in Books
with tabs[4]:
    if not r["missing_books"].empty:
        df_mb = r["missing_books"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "TOTAL_TAX"]].copy()
        df_mb["Invoice_Date"] = pd.to_datetime(df_mb["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y").fillna("")
        df_mb.columns = ["GSTIN", "Supplier", "Invoice No", "Date", "Taxable", "ITC"]
        df_mb = apply_filters(coerce_str_cols(df_mb), f_gstin, f_supplier)
        safe_dataframe(
            df_mb,
            column_config=MISS_CFG,
            empty_message="No missing invoices match the current filters.",
            caption=f"{len(df_mb)} invoices in GSTR-2B not found in Books",
        )
    else:
        st.success("✅ No invoices missing in Books.")

# Tab 5: Supplier Summary
with tabs[5]:
    all_gstins = set()
    if not r["books_raw"].empty: all_gstins.update(r["books_raw"]["GSTIN"].unique())
    if not r["gstr_raw"].empty:  all_gstins.update(r["gstr_raw"]["GSTIN"].unique())
    sup_rows = []
    for gstin in all_gstins:
        ib = r["books_raw"][r["books_raw"]["GSTIN"] == gstin]["TOTAL_TAX"].sum() if not r["books_raw"].empty else 0
        ig = r["gstr_raw"][r["gstr_raw"]["GSTIN"] == gstin]["TOTAL_TAX"].sum()   if not r["gstr_raw"].empty  else 0
        sup_rows.append({
            "GSTIN":            str(gstin),
            "Supplier":         str(trade_name_map.get(gstin, "Unknown")),
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
        safe_dataframe(
            sup_df,
            column_config={
                "GSTIN":            st.column_config.TextColumn("GSTIN",       width=180),
                "Supplier":         st.column_config.TextColumn("Supplier",    width=280),
                "ITC as per Books": st.column_config.NumberColumn("📚 Books",  width=140, format="%.2f"),
                "ITC as per 2B":    st.column_config.NumberColumn("📊 GSTR-2B",width=140, format="%.2f"),
                "ITC Difference":   st.column_config.NumberColumn("📉 Diff",   width=130, format="%.2f"),
            },
            empty_message="No supplier data matches the current filters.",
            caption=f"Total Difference: ₹{sup_df['ITC Difference'].sum():,.2f}" if not sup_df.empty else None,
        )

# ── ZERO ITC ─────────────────────────────────────────────────────────────────

if not r["no_itc"].empty:
    df_ni = r["no_itc"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "Invoice_Value"]].copy()
    df_ni["Invoice_Date"] = pd.to_datetime(df_ni["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y").fillna("")
    df_ni.columns = ["GSTIN", "Supplier", "Invoice No", "Date", "Taxable", "Invoice Value"]
    df_ni = coerce_str_cols(df_ni)
    df_ni = df_ni[
        (df_ni["Invoice No"].astype(str).str.strip() != "") &
        (df_ni["Invoice No"].astype(str).str.strip().str.lower() != "nan") &
        (df_ni["Invoice Value"].astype(float) > 0)
    ]
    if not df_ni.empty:
        st.divider()
        st.markdown("## 🟡 Zero ITC Invoices")
        safe_dataframe(
            df_ni,
            column_config={
                "GSTIN":         st.column_config.TextColumn("GSTIN",          width=180),
                "Supplier":      st.column_config.TextColumn("Supplier",       width=280),
                "Invoice No":    st.column_config.TextColumn("Invoice No",     width=150),
                "Date":          st.column_config.TextColumn("Date",           width=115),
                "Taxable":       st.column_config.NumberColumn("Taxable",      width=110, format="%.2f"),
                "Invoice Value": st.column_config.NumberColumn("Invoice Value",width=120, format="%.2f"),
            },
            empty_message="No zero ITC invoices.",
            caption=f"{len(df_ni)} invoices with zero tax",
        )

# ── EXCEL EXPORT HELPERS ──────────────────────────────────────────────────────

def _build_workbook_formats(wb):
    return {
        "title":  wb.add_format({'bold': True, 'font_size': 14, 'font_name': 'Aptos Narrow'}),
        "header": wb.add_format({'bold': True, 'font_name': 'Aptos Narrow',
                                  'font_color': 'white', 'bg_color': '#1F4E78',
                                  'align': 'center', 'valign': 'vcenter', 'border': 1}),
        "number": wb.add_format({'font_name': 'Aptos Narrow', 'num_format': '#,##0.00'}),
        "date":   wb.add_format({'font_name': 'Aptos Narrow', 'num_format': 'dd-mmm-yyyy'}),
        "text":   wb.add_format({'font_name': 'Aptos Narrow'}),
    }

def _write_sheet(ws, title, headers, data_rows, col_types, fmts):
    ws.write(0, 0, title, fmts["title"])
    ws.set_row(1, 20)
    for ci, h in enumerate(headers):
        ws.write(1, ci, h, fmts["header"])
    for ri, row_data in enumerate(data_rows):
        for ci, val in enumerate(row_data):
            ct = col_types.get(headers[ci], "text")
            if ct == "date" and pd.notna(val):
                try:
                    dt = pd.to_datetime(val)
                    ws.write_datetime(ri+2, ci, dt.to_pydatetime(), fmts["date"]) if not pd.isna(dt) else \
                        safe_write_text(ws, ri+2, ci, "", fmts["text"])
                except Exception:
                    safe_write_text(ws, ri+2, ci, str(val) if pd.notna(val) else "", fmts["text"])
            elif ct == "number":
                safe_write_number(ws, ri+2, ci, val, fmts["number"])
            else:
                safe_write_text(ws, ri+2, ci, val, fmts["text"])
    for ci, h in enumerate(headers):
        ct = col_types.get(h, "text")
        ws.set_column(ci, ci, 15 if ct in ("date","number") else 22,
                      fmts["date"] if ct=="date" else fmts["number"] if ct=="number" else fmts["text"])
    ws.freeze_panes(2, 0)
    ws.autofit()

BOOK_CT = {"GSTIN":"text","Trade_Name":"text","Invoice_No":"text","Invoice_Date":"date",
           "Taxable_Value":"number","CGST":"number","SGST":"number","IGST":"number",
           "CESS":"number","TOTAL_TAX":"number","Invoice_Value":"number"}
MISS_CT  = {"GSTIN":"text","Trade_Name":"text","Invoice_No":"text","Invoice_Date":"date",
            "Taxable_Value":"number","TOTAL_TAX":"number"}

# ── FULL EXCEL REPORT ─────────────────────────────────────────────────────────

st.divider()
col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    full_output = io.BytesIO()
    with pd.ExcelWriter(full_output, engine="xlsxwriter") as writer:
        wb   = writer.book
        fmts = _build_workbook_formats(wb)

        # Summary sheet
        summary_df = post_processing_cleaner(pd.DataFrame({
            "Particulars": ["ITC - Books","ITC - GSTR-2B","Difference","ITC at Risk","Match %",
                            "Total Books","Total GSTR","Matched","Tax Diff","Missing 2B","Missing Books"],
            "Value":       [s["ITC_Books"],s["ITC_GSTR"],s["ITC_Diff"],s["ITC_at_Risk"],s["Match_%"],
                            s["Total_Books"],s["Total_GSTR"],s["Matched"],s["Tax_Diff"],
                            s["Missing_2B"],s["Missing_Books"]],
        }))
        ws_s = wb.add_worksheet("Summary")
        _write_sheet(ws_s, "Reconciliation Summary", list(summary_df.columns),
                     summary_df.values.tolist(), {"Particulars":"text","Value":"number"}, fmts)

        # Books
        if not r["books_raw"].empty:
            bdf = post_processing_cleaner(r["books_raw"].copy())
            _write_sheet(wb.add_worksheet("Books"), "Books Data",
                         list(bdf.columns), bdf.values.tolist(), BOOK_CT, fmts)

        # GSTR-2B
        if not r["gstr_raw"].empty:
            gdf = post_processing_cleaner(r["gstr_raw"].copy())
            _write_sheet(wb.add_worksheet("GSTR-2B"), "GSTR-2B Data",
                         list(gdf.columns), gdf.values.tolist(), BOOK_CT, fmts)

        # Missing in 2B
        if not r["missing_2b"].empty:
            mdf = post_processing_cleaner(r["missing_2b"][list(MISS_CT)].copy())
            _write_sheet(wb.add_worksheet("Missing in 2B"), "Missing in 2B",
                         list(mdf.columns), mdf.values.tolist(), MISS_CT, fmts)

        # Missing in Books
        if not r["missing_books"].empty:
            mbdf = post_processing_cleaner(r["missing_books"][list(MISS_CT)].copy())
            _write_sheet(wb.add_worksheet("Missing in Books"), "Missing in Books",
                         list(mbdf.columns), mbdf.values.tolist(), MISS_CT, fmts)

        # Supplier Wise ITC Summary
        if sup_rows:
            sdf = post_processing_cleaner(pd.DataFrame(sup_rows))
            _write_sheet(wb.add_worksheet("Supplier Wise ITC Summary"), "Supplier Wise ITC Summary",
                         ["GSTIN","Supplier","ITC as per Books","ITC as per 2B","ITC Difference"],
                         sdf[["GSTIN","Supplier","ITC as per Books","ITC as per 2B","ITC Difference"]].values.tolist(),
                         {"GSTIN":"text","Supplier":"text","ITC as per Books":"number",
                          "ITC as per 2B":"number","ITC Difference":"number"}, fmts)

        # Supplier Drill Down
        if not detail_df.empty:
            ddf = detail_df.copy()
            ddf["Invoice Date"] = pd.to_datetime(ddf["Date"], format="%d-%b-%Y", errors="coerce")
            ddf = ddf.sort_values(["Supplier","Date"]).reset_index(drop=True)
            ws_dd = wb.add_worksheet("Supplier Drill Down")
            ws_dd.write(0, 0, "Supplier Drill Down — Invoice Level Details", fmts["title"])
            ws_dd.set_row(1, 20)
            dd_hdrs = ["GSTIN","Supplier","Invoice No","Invoice Date",
                       "ITC Books","ITC 2B","Difference","Remarks"]
            for ci, h in enumerate(dd_hdrs):
                ws_dd.write(1, ci, h, fmts["header"])
            for ri, row_data in ddf.iterrows():
                safe_write_text(ws_dd, ri+2, 0, row_data["GSTIN"],       fmts["text"])
                safe_write_text(ws_dd, ri+2, 1, row_data["Supplier"],    fmts["text"])
                safe_write_text(ws_dd, ri+2, 2, row_data["Invoice No"],  fmts["text"])
                try:
                    idt = row_data["Invoice Date"]
                    ws_dd.write_datetime(ri+2, 3, pd.Timestamp(idt).to_pydatetime(), fmts["date"]) \
                        if pd.notna(idt) else safe_write_text(ws_dd, ri+2, 3, "", fmts["text"])
                except Exception:
                    safe_write_text(ws_dd, ri+2, 3, "", fmts["text"])
                safe_write_number(ws_dd, ri+2, 4, row_data["ITC Books"],  fmts["number"])
                safe_write_number(ws_dd, ri+2, 5, row_data["ITC 2B"],    fmts["number"])
                safe_write_number(ws_dd, ri+2, 6, row_data["Difference"],fmts["number"])
                safe_write_text(ws_dd, ri+2, 7, row_data["Remarks"],     fmts["text"])
            ws_dd.autofilter(1, 0, 1, 7)
            for ci, w in enumerate([20,30,20,15,15,15,15,18]):
                ws_dd.set_column(ci, ci, w)
            ws_dd.freeze_panes(2, 0)
            ws_dd.autofit()

        # Zero ITC
        if not r["no_itc"].empty:
            nidf = post_processing_cleaner(
                r["no_itc"][["GSTIN","Trade_Name","Invoice_No","Invoice_Date","Taxable_Value","Invoice_Value"]].copy())
            _write_sheet(wb.add_worksheet("NO ITC"), "Zero ITC Invoices",
                         list(nidf.columns), nidf.values.tolist(),
                         {"GSTIN":"text","Trade_Name":"text","Invoice_No":"text",
                          "Invoice_Date":"date","Taxable_Value":"number","Invoice_Value":"number"}, fmts)

    full_output.seek(0)
    st.download_button(
        "📥 Download Full Report",
        data=full_output,
        file_name=f"reconciliation_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── ISSUES ONLY EXCEL ─────────────────────────────────────────────────────────

with col_dl2:
    issues_output = io.BytesIO()
    with pd.ExcelWriter(issues_output, engine="xlsxwriter") as writer:
        wb2   = writer.book
        fmts2 = _build_workbook_formats(wb2)

        # Sheet 1: Missing in GSTR-2B
        if not r["missing_2b"].empty:
            mdf = post_processing_cleaner(r["missing_2b"][list(MISS_CT)].copy())
            _write_sheet(wb2.add_worksheet("Missing in 2B"), "Invoices Missing in GSTR-2B",
                         list(mdf.columns), mdf.values.tolist(), MISS_CT, fmts2)
        else:
            ws_e = wb2.add_worksheet("Missing in 2B")
            ws_e.write(0, 0, "No invoices missing in GSTR-2B ✓", fmts2["title"])

        # Sheet 2: Missing in Books
        if not r["missing_books"].empty:
            mbdf = post_processing_cleaner(r["missing_books"][list(MISS_CT)].copy())
            _write_sheet(wb2.add_worksheet("Missing in Books"), "Invoices Missing in Books",
                         list(mbdf.columns), mbdf.values.tolist(), MISS_CT, fmts2)
        else:
            ws_e = wb2.add_worksheet("Missing in Books")
            ws_e.write(0, 0, "No invoices missing in Books ✓", fmts2["title"])

        # Sheet 3: Tax Differences
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
            ws_e = wb2.add_worksheet("Tax Differences")
            ws_e.write(0, 0, "No tax differences found ✓", fmts2["title"])

        # Sheet 4: Data Issues
        if not all_issues.empty:
            idf = all_issues.copy()
            for col in idf.select_dtypes(include=["object"]).columns:
                idf[col] = idf[col].apply(lambda v: "" if (v is None or pd.isna(v)) else str(v))
            issue_ct = {c: ("date" if c=="Invoice_Date" else "number" if c in ("Taxable_Value","Invoice_Value","TOTAL_TAX") else "text")
                        for c in idf.columns}
            _write_sheet(wb2.add_worksheet("Data Issues"), "Data Quality Issues",
                         list(idf.columns), idf.values.tolist(), issue_ct, fmts2)
        else:
            ws_e = wb2.add_worksheet("Data Issues")
            ws_e.write(0, 0, "No data issues found ✓", fmts2["title"])

    issues_output.seek(0)
    st.download_button(
        "📥 Download Issues Only",
        data=issues_output,
        file_name=f"issues_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── HIDE ANY RESIDUAL ZERO-ONLY TABLES ───────────────────────────────────────
st.markdown("""
<style>
    /* Belt-and-suspenders: hide any dataframe that rendered with only empty/zero cells */
    .stDataFrame:has(td:empty) { display: none !important; }
</style>
""", unsafe_allow_html=True)
