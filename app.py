"""
GST Reconciliation App
Minimal, clean, production-ready
Version: 6.0 (Tally PR + GSTR-2B Excel Native Parsers)
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
    .stApp {
        background-color: white;
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
    }
    .warning-box {
        background-color: #FEF2F2;
        padding: 1rem;
        border-left: 4px solid #DC2626;
        margin: 1rem 0;
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
    }
    .format-badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: bold;
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
        margin-top: 4px;
    }
    .badge-tally  { background:#DCFCE7; color:#166534; }
    .badge-gstr2b { background:#DBEAFE; color:#1E40AF; }
    .badge-std    { background:#F3F4F6; color:#374151; }
    .badge-unk    { background:#FEF3C7; color:#92400E; }
    .stTabs [data-baseweb="tab-list"] { gap: 2px; }
    .stTabs [data-baseweb="tab"] {
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
        padding: 8px 16px;
    }
    .streamlit-expanderHeader {
        font-family: 'Aptos Narrow', 'Aptos', sans-serif !important;
        font-size: 1rem !important;
    }
    .stMarkdown, .stText, .stCaption, .stMetric {
        font-family: 'Aptos Narrow', 'Aptos', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

st.title("GST Reconciliation")
st.caption("Books vs GSTR-2B | Auto-detects Tally PR / GSTR-2B Excel / Standard Template")

# =================== TEMPLATES =================== #

st.markdown("### Download Templates")
st.divider()

col1, col2 = st.columns(2)

template = pd.DataFrame({
    "GSTIN":         ["27AAAAA0000A1Z5"],
    "Trade_Name":    ["Sample Supplier"],
    "Invoice_No":    ["INV001"],
    "Invoice_Date":  ["01/04/2025"],
    "Taxable_Value": [8474.58],
    "CGST":          [762.71],
    "SGST":          [762.71],
    "IGST":          [0.00],
    "CESS":          [0.00],
})

with col1:
    st.download_button(
        "📥 Books Template",
        template.to_csv(index=False).encode(),
        "books_template.csv",
        "text/csv"
    )

with col2:
    st.download_button(
        "📥 GSTR-2B Template",
        template.to_csv(index=False).encode(),
        "gstr2b_template.csv",
        "text/csv"
    )

st.divider()

# =================== SUPPORTED FORMATS INFO =================== #

with st.expander("ℹ️ Supported File Formats", expanded=False):
    st.markdown("""
    **Books (Left Upload)**
    - 🟢 **Tally Purchase Register** — Upload XLS/XLSX directly from Tally Prime (Purchase Register report). TDS is automatically added back to Gross Total.
    - ⬜ **Standard Template** — CSV/XLSX with columns: GSTIN, Trade_Name, Invoice_No, Invoice_Date, Taxable_Value, CGST, SGST, IGST, CESS

    **GSTR-2B (Right Upload)**
    - 🔵 **GSTR-2B Excel Download** — Upload the Excel file directly downloaded from the GST Portal (GSTR-2B section). No reformatting needed.
    - ⬜ **Standard Template** — Same 9-column format as above

    The app **auto-detects** the format — no manual selection required.
    """)

# =================== UPLOAD =================== #

col1, col2 = st.columns(2)

with col1:
    books_file = st.file_uploader("📤 Upload Books", type=["xlsx", "xls", "csv"])

with col2:
    gstr_file = st.file_uploader("📤 Upload GSTR-2B", type=["xlsx", "xls", "csv"])

tolerance = st.number_input("Tolerance (₹)", value=DEFAULT_TOLERANCE, step=0.5, min_value=0.0)

# =================== FORMAT BADGE HELPER =================== #

FORMAT_LABELS = {
    "tally_pr":    ("🟢 Tally Purchase Register", "badge-tally"),
    "gstr2b_excel":("🔵 GSTR-2B Excel (Portal)",  "badge-gstr2b"),
    "standard":    ("⬜ Standard Template",         "badge-std"),
    "unknown":     ("⚠️ Unknown Format",            "badge-unk"),
}

def show_format_badge(fmt: str):
    label, css_class = FORMAT_LABELS.get(fmt, ("⚠️ Unknown", "badge-unk"))
    st.markdown(
        f'<span class="format-badge {css_class}">{label}</span>',
        unsafe_allow_html=True
    )

# =================== SMART PARSER ROUTER =================== #

def load_raw(uploaded_file) -> pd.DataFrame:
    """Load file into DataFrame without any header assumptions."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, header=None, dtype=str)
    elif name.endswith(".xls"):
        return pd.read_excel(uploaded_file, engine="xlrd", header=None, dtype=str)
    else:
        return pd.read_excel(uploaded_file, header=None, dtype=str)

def parse_books(uploaded_file) -> tuple:
    """
    Auto-detect and parse Books file.
    Returns: (books_clean_df, no_itc_df, issues_df, detected_format)
    """
    raw = load_raw(uploaded_file)
    fmt = detect_file_format(raw, uploaded_file.name)

    if fmt == "tally_pr":
        books_clean, no_itc, issues = parse_tally_purchase_register(raw)
    elif fmt == "standard":
        # Re-read with proper header for standard template
        uploaded_file.seek(0)
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.lower().endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file)
        books_clean, no_itc, issues = parse_tally(df)
    else:
        # Try standard as fallback
        uploaded_file.seek(0)
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.lower().endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file)
        books_clean, no_itc, issues = parse_tally(df)
        fmt = "standard"

    return books_clean, no_itc, issues, fmt


def parse_gstr(uploaded_file) -> tuple:
    """
    Auto-detect and parse GSTR-2B file.
    Returns: (gstr_clean_df, detected_format)
    """
    raw = load_raw(uploaded_file)
    fmt = detect_file_format(raw, uploaded_file.name)

    if fmt == "gstr2b_excel":
        gstr_clean = parse_gstr2b_excel(raw)
    elif fmt == "standard":
        uploaded_file.seek(0)
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.lower().endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file)
        gstr_clean = parse_gstr2b(df)
    else:
        # Fallback — try as standard
        uploaded_file.seek(0)
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.lower().endswith(".xls"):
            df = pd.read_excel(uploaded_file, engine="xlrd")
        else:
            df = pd.read_excel(uploaded_file)
        gstr_clean = parse_gstr2b(df)
        fmt = "standard"

    return gstr_clean, fmt

# =================== PROCESS =================== #

if st.button("🚀 Run Reconciliation", use_container_width=True, type="primary"):
    if not books_file or not gstr_file:
        st.error("Please upload both files")
    else:
        try:
            with st.spinner("Detecting formats and processing..."):

                # ── Parse Books ──
                books_clean, no_itc, issues, books_fmt = parse_books(books_file)

                # ── Parse GSTR-2B ──
                gstr_clean, gstr_fmt = parse_gstr(gstr_file)

                # ── Reconcile ──
                results = reconcile(gstr_clean, books_clean, tolerance)

                results["no_itc"]     = no_itc
                results["issues"]     = issues
                results["books_raw"]  = books_clean
                results["gstr_raw"]   = gstr_clean
                results["books_fmt"]  = books_fmt
                results["gstr_fmt"]   = gstr_fmt

                st.session_state["results"] = results

            st.success("✅ Reconciliation completed successfully!")

            # Show detected formats
            col_b, col_g = st.columns(2)
            with col_b:
                st.caption("Books format detected:")
                show_format_badge(books_fmt)
            with col_g:
                st.caption("GSTR-2B format detected:")
                show_format_badge(gstr_fmt)

        except Exception as e:
            logger.error(f"Failed: {e}", exc_info=True)
            st.error(f"Error: {str(e)}")

# =================== DATE FORMATTING =================== #

def format_date_display(date_val):
    if pd.isna(date_val):
        return ""
    if isinstance(date_val, (datetime, pd.Timestamp)):
        return date_val.strftime("%d-%b-%Y")
    try:
        return pd.to_datetime(date_val).strftime("%d-%b-%Y")
    except:
        return str(date_val)

# =================== SAFE EXCEL WRITERS =================== #

def safe_write_number(worksheet, row, col, value, format_obj):
    try:
        if isinstance(value, (int, float)) and not pd.isna(value):
            worksheet.write_number(row, col, float(value), format_obj)
            return
        if value is None or pd.isna(value):
            worksheet.write_number(row, col, 0.0, format_obj)
            return
        str_val = str(value).strip()
        if str_val in ["", "-", "null", "none", "nil", "na", "nan"]:
            worksheet.write_number(row, col, 0.0, format_obj)
            return
        str_val = str_val.replace('₹', '').replace(',', '').replace(' ', '').strip()
        try:
            worksheet.write_number(row, col, float(str_val), format_obj)
        except (ValueError, TypeError):
            worksheet.write_number(row, col, 0.0, format_obj)
    except Exception:
        worksheet.write_number(row, col, 0.0, format_obj)

def safe_write_text(worksheet, row, col, value, format_obj):
    try:
        if value is None or pd.isna(value):
            worksheet.write_string(row, col, "", format_obj)
            return
        worksheet.write_string(row, col, str(value).strip(), format_obj)
    except Exception:
        worksheet.write_string(row, col, "", format_obj)

# =================== DISPLAY =================== #

if "results" in st.session_state:
    r = st.session_state["results"]
    s = r["summary"]

    # Show format badges if available
    if "books_fmt" in r and "gstr_fmt" in r:
        cb, cg = st.columns(2)
        with cb:
            st.caption("Books format:")
            show_format_badge(r["books_fmt"])
        with cg:
            st.caption("GSTR-2B format:")
            show_format_badge(r["gstr_fmt"])
        st.markdown("")

    # ===== DATA ISSUES ===== #

    all_issues = r["issues"].copy() if not r["issues"].empty else pd.DataFrame()
    if "duplicate_issues" in r and not r["duplicate_issues"].empty:
        if all_issues.empty:
            all_issues = r["duplicate_issues"].copy()
        else:
            all_issues = pd.concat([all_issues, r["duplicate_issues"]], ignore_index=True)

    if not all_issues.empty:
        st.markdown(f"""
        <div class="warning-box">
            <strong>⚠️ {len(all_issues)} Data Issues Found</strong><br>
            Fix these in source data
        </div>
        """, unsafe_allow_html=True)

        with st.expander("🔍 View Issues", expanded=True):
            df_issues = all_issues.copy()
            if "Invoice_Date" in df_issues.columns:
                df_issues["Invoice_Date"] = pd.to_datetime(df_issues["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y")
            cols = ["Issue"] + [c for c in df_issues.columns if c != "Issue"]
            df_issues = df_issues[cols]
            st.dataframe(df_issues, use_container_width=True, hide_index=True,
                column_config={
                    "Issue":         st.column_config.TextColumn("Issue",         width=200),
                    "GSTIN":         st.column_config.TextColumn("GSTIN",         width=180),
                    "Trade_Name":    st.column_config.TextColumn("Trade Name",    width=300),
                    "Invoice_No":    st.column_config.TextColumn("Invoice No",    width=150),
                    "Invoice_Date":  st.column_config.TextColumn("Invoice Date",  width=120),
                    "Taxable_Value": st.column_config.NumberColumn("Taxable",     width=100, format="%.2f"),
                    "Invoice_Value": st.column_config.NumberColumn("Invoice Value",width=120, format="%.2f"),
                    "TOTAL_TAX":     st.column_config.NumberColumn("Total Tax",   width=100, format="%.2f"),
                })
            st.markdown("### Issue Summary")
            ic = df_issues["Issue"].value_counts().reset_index()
            ic.columns = ["Issue Type", "Count"]
            st.dataframe(ic, use_container_width=True, hide_index=True)

    # ===== SUMMARY ===== #

    st.markdown("## Summary")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("📚 ITC - Books",   f"{s['ITC_Books']:,.2f}")
    m2.metric("📊 ITC - GSTR-2B", f"{s['ITC_GSTR']:,.2f}")
    m3.metric("📉 Difference",    f"{s['ITC_Diff']:,.2f}")
    m4.metric("⚠️ ITC at Risk",   f"{s['ITC_at_Risk']:,.2f}")
    m5.metric("✅ Match %",        f"{s['Match_%']}%")

    st.divider()

    # ===== TABS ===== #

    tabs = st.tabs(["📚 Books", "📊 GSTR-2B", "❌ Missing in 2B", "📕 Missing in Books", "📋 Supplier Summary"])

    STD_COL_CFG = {
        "GSTIN":         st.column_config.TextColumn("GSTIN",         width=180),
        "Trade_Name":    st.column_config.TextColumn("Trade Name",    width=300),
        "Invoice_No":    st.column_config.TextColumn("Invoice No",    width=150),
        "Invoice_Date":  st.column_config.TextColumn("Invoice Date",  width=120),
        "Taxable_Value": st.column_config.NumberColumn("Taxable",     width=100, format="%.2f"),
        "CGST":          st.column_config.NumberColumn("CGST",        width=100, format="%.2f"),
        "SGST":          st.column_config.NumberColumn("SGST",        width=100, format="%.2f"),
        "IGST":          st.column_config.NumberColumn("IGST",        width=100, format="%.2f"),
        "CESS":          st.column_config.NumberColumn("CESS",        width=100, format="%.2f"),
        "TOTAL_TAX":     st.column_config.NumberColumn("Total Tax",   width=100, format="%.2f"),
        "Invoice_Value": st.column_config.NumberColumn("Invoice Value",width=120, format="%.2f"),
    }

    with tabs[0]:
        if not r["books_raw"].empty:
            df_d = r["books_raw"].copy()
            if "Invoice_Date" in df_d.columns:
                df_d["Invoice_Date"] = pd.to_datetime(df_d["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y")
            st.dataframe(df_d, use_container_width=True, hide_index=True, column_config=STD_COL_CFG)

    with tabs[1]:
        if not r["gstr_raw"].empty:
            df_d = r["gstr_raw"].copy()
            if "Invoice_Date" in df_d.columns:
                df_d["Invoice_Date"] = pd.to_datetime(df_d["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y")
            st.dataframe(df_d, use_container_width=True, hide_index=True, column_config=STD_COL_CFG)

    with tabs[2]:
        if not r["missing_2b"].empty:
            df = r["missing_2b"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "TOTAL_TAX"]].copy()
            if "Invoice_Date" in df.columns:
                df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y")
            df.columns = ["GSTIN", "Supplier", "Invoice No", "Date", "Taxable", "ITC"]
            st.dataframe(df, use_container_width=True, hide_index=True,
                column_config={
                    "GSTIN":      st.column_config.TextColumn("GSTIN",      width=180),
                    "Supplier":   st.column_config.TextColumn("Supplier",   width=300),
                    "Invoice No": st.column_config.TextColumn("Invoice No", width=150),
                    "Date":       st.column_config.TextColumn("Date",       width=120),
                    "Taxable":    st.column_config.NumberColumn("Taxable",  width=100, format="%.2f"),
                    "ITC":        st.column_config.NumberColumn("ITC",      width=100, format="%.2f"),
                })
            st.caption(f"💰 Total ITC Impact: {df['ITC'].sum():,.2f}")

    with tabs[3]:
        if not r["missing_books"].empty:
            df = r["missing_books"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "TOTAL_TAX"]].copy()
            if "Invoice_Date" in df.columns:
                df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y")
            df.columns = ["GSTIN", "Supplier", "Invoice No", "Date", "Taxable", "ITC"]
            st.dataframe(df, use_container_width=True, hide_index=True,
                column_config={
                    "GSTIN":      st.column_config.TextColumn("GSTIN",      width=180),
                    "Supplier":   st.column_config.TextColumn("Supplier",   width=300),
                    "Invoice No": st.column_config.TextColumn("Invoice No", width=150),
                    "Date":       st.column_config.TextColumn("Date",       width=120),
                    "Taxable":    st.column_config.NumberColumn("Taxable",  width=100, format="%.2f"),
                    "ITC":        st.column_config.NumberColumn("ITC",      width=100, format="%.2f"),
                })
            st.caption(f"💰 Total ITC Impact: {df['ITC'].sum():,.2f}")

    with tabs[4]:
        if not r["books_raw"].empty or not r["gstr_raw"].empty:
            trade_name_map = r.get("trade_name_mapping", {})
            all_gstins = set()
            if not r["books_raw"].empty:
                all_gstins.update(r["books_raw"]["GSTIN"].unique())
            if not r["gstr_raw"].empty:
                all_gstins.update(r["gstr_raw"]["GSTIN"].unique())

            supplier_data = []
            for gstin in all_gstins:
                itc_books = r["books_raw"][r["books_raw"]["GSTIN"] == gstin]["TOTAL_TAX"].sum() \
                    if not r["books_raw"].empty else 0
                itc_gstr  = r["gstr_raw"][r["gstr_raw"]["GSTIN"] == gstin]["TOTAL_TAX"].sum() \
                    if not r["gstr_raw"].empty else 0
                supplier_data.append({
                    "GSTIN":            gstin,
                    "Supplier":         trade_name_map.get(gstin, "Unknown"),
                    "ITC as per Books": round(itc_books, 2),
                    "ITC as per 2B":    round(itc_gstr, 2),
                    "ITC Difference":   round(itc_gstr - itc_books, 2),
                })

            if supplier_data:
                df = pd.DataFrame(supplier_data).sort_values("Supplier")
                st.dataframe(df, use_container_width=True, hide_index=True,
                    column_config={
                        "GSTIN":             st.column_config.TextColumn("GSTIN",      width=180),
                        "Supplier":          st.column_config.TextColumn("Supplier",   width=300),
                        "ITC as per Books":  st.column_config.NumberColumn("📚 Books", width=150, format="%.2f"),
                        "ITC as per 2B":     st.column_config.NumberColumn("📊 GSTR-2B",width=150, format="%.2f"),
                        "ITC Difference":    st.column_config.NumberColumn("📉 Difference",width=150, format="%.2f"),
                    })
                st.caption(f"💰 Total Difference: {df['ITC Difference'].sum():,.2f}")

    # ===== INVOICE LEVEL DETAILS ===== #

    if not r["books_raw"].empty or not r["gstr_raw"].empty:
        st.divider()
        st.markdown("## 📋 Invoice Level Details")
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
            matching_gstr = None
            if not r["gstr_raw"].empty:
                m = r["gstr_raw"][(r["gstr_raw"]["GSTIN"] == gstin) & (r["gstr_raw"]["Invoice_No"] == row["Invoice_No"])]
                if not m.empty:
                    matching_gstr = m.iloc[0]

            if matching_gstr is not None:
                itc_books  = row["TOTAL_TAX"]
                itc_gstr   = matching_gstr["TOTAL_TAX"]
                difference = itc_gstr - itc_books
                remark = "✅ Matched" if abs(difference) <= tolerance else "⚠️ Tax Difference"
                detail_data.append({
                    "GSTIN": gstin, "Supplier": supplier,
                    "Invoice No": row["Invoice_No"],
                    "Date": pd.to_datetime(row["Invoice_Date"]).strftime("%d-%b-%Y"),
                    "ITC as per Books": itc_books, "ITC as per 2B": itc_gstr,
                    "Difference": difference, "Remarks": remark,
                })
            else:
                detail_data.append({
                    "GSTIN": gstin, "Supplier": supplier,
                    "Invoice No": row["Invoice_No"],
                    "Date": pd.to_datetime(row["Invoice_Date"]).strftime("%d-%b-%Y"),
                    "ITC as per Books": row["TOTAL_TAX"], "ITC as per 2B": 0.0,
                    "Difference": -row["TOTAL_TAX"], "Remarks": "❌ Missing in GST",
                })

        for _, row in r["gstr_raw"].iterrows():
            key = f"{row['GSTIN']}|{row['Invoice_No']}"
            if key in processed_keys:
                continue
            processed_keys.add(key)
            gstin    = row["GSTIN"]
            supplier = trade_name_map.get(gstin, row["Trade_Name"])
            detail_data.append({
                "GSTIN": gstin, "Supplier": supplier,
                "Invoice No": row["Invoice_No"],
                "Date": pd.to_datetime(row["Invoice_Date"]).strftime("%d-%b-%Y"),
                "ITC as per Books": 0.0, "ITC as per 2B": row["TOTAL_TAX"],
                "Difference": row["TOTAL_TAX"], "Remarks": "📕 Missing in Books",
            })

        if detail_data:
            df = pd.DataFrame(detail_data).sort_values(["Supplier", "Date"])
            st.dataframe(df, use_container_width=True, hide_index=True,
                column_config={
                    "GSTIN":            st.column_config.TextColumn("GSTIN",       width=180),
                    "Supplier":         st.column_config.TextColumn("Supplier",    width=300),
                    "Invoice No":       st.column_config.TextColumn("Invoice No",  width=150),
                    "Date":             st.column_config.TextColumn("Date",        width=120),
                    "ITC as per Books": st.column_config.NumberColumn("📚 Books",  width=130, format="%.2f"),
                    "ITC as per 2B":    st.column_config.NumberColumn("📊 GSTR-2B",width=130, format="%.2f"),
                    "Difference":       st.column_config.NumberColumn("📉 Diff",   width=120, format="%.2f"),
                    "Remarks":          st.column_config.TextColumn("Remarks",     width=150),
                })

    # ===== ZERO ITC ===== #

    if not r["no_itc"].empty:
        st.divider()
        st.markdown("## 🟡 Zero ITC Invoices")
        df = r["no_itc"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "Invoice_Value"]].copy()
        if "Invoice_Date" in df.columns:
            df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce").dt.strftime("%d-%b-%Y")
        df.columns = ["GSTIN", "Supplier", "Invoice No", "Date", "Taxable", "Invoice Value"]
        st.dataframe(df, use_container_width=True, hide_index=True,
            column_config={
                "GSTIN":         st.column_config.TextColumn("GSTIN",         width=180),
                "Supplier":      st.column_config.TextColumn("Supplier",      width=300),
                "Invoice No":    st.column_config.TextColumn("Invoice No",    width=150),
                "Date":          st.column_config.TextColumn("Date",          width=120),
                "Taxable":       st.column_config.NumberColumn("Taxable",     width=100, format="%.2f"),
                "Invoice Value": st.column_config.NumberColumn("Invoice Value",width=120, format="%.2f"),
            })

    # ===== EXCEL EXPORT ===== #

    st.divider()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        fmt_title  = wb.add_format({'bold': True, 'font_size': 14, 'font_name': 'Aptos Narrow'})
        fmt_header = wb.add_format({
            'bold': True, 'font_name': 'Aptos Narrow',
            'font_color': 'white', 'bg_color': '#1F4E78',
            'align': 'center', 'valign': 'vcenter', 'border': 1,
        })
        fmt_number = wb.add_format({'font_name': 'Aptos Narrow', 'num_format': '#,##0.00'})
        fmt_date   = wb.add_format({'font_name': 'Aptos Narrow', 'num_format': 'dd-mmm-yyyy'})
        fmt_text   = wb.add_format({'font_name': 'Aptos Narrow'})

        def write_sheet(ws, title, headers, data_rows, col_types):
            """
            col_types: dict {col_name: 'text'|'number'|'date'}
            data_rows: list of lists
            """
            ws.write(0, 0, title, fmt_title)
            ws.set_row(1, 20)
            for ci, h in enumerate(headers):
                ws.write(1, ci, h, fmt_header)
            for ri, row_data in enumerate(data_rows):
                for ci, val in enumerate(row_data):
                    ct = col_types.get(headers[ci], "text")
                    if ct == "date" and pd.notna(val):
                        try:
                            ws.write_datetime(ri + 2, ci, pd.to_datetime(val), fmt_date)
                        except Exception:
                            safe_write_text(ws, ri + 2, ci, val, fmt_text)
                    elif ct == "number":
                        safe_write_number(ws, ri + 2, ci, val, fmt_number)
                    else:
                        safe_write_text(ws, ri + 2, ci, val, fmt_text)
            for ci, h in enumerate(headers):
                ct = col_types.get(h, "text")
                if ct == "date":
                    ws.set_column(ci, ci, 15, fmt_date)
                elif ct == "number":
                    ws.set_column(ci, ci, 15, fmt_number)
                else:
                    ws.set_column(ci, ci, 20, fmt_text)
            ws.freeze_panes(2, 0)
            ws.autofit()

        # Summary
        summary_df = post_processing_cleaner(pd.DataFrame({
            "Particulars": ["ITC - Books", "ITC - GSTR-2B", "Difference", "ITC at Risk", "Match %"],
            "Amount":      [s["ITC_Books"], s["ITC_GSTR"], s["ITC_Diff"], s["ITC_at_Risk"], s["Match_%"]],
        }))
        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=1)
        ws_s = writer.sheets["Summary"]
        ws_s.write(0, 0, "Reconciliation Summary", fmt_title)
        ws_s.set_row(1, 20)
        for ci, col in enumerate(summary_df.columns):
            ws_s.write(1, ci, col, fmt_header)
        ws_s.set_column('A:A', 25, fmt_text)
        ws_s.set_column('B:B', 20, fmt_number)
        ws_s.freeze_panes(2, 0)
        ws_s.autofit()

        # Shared column type maps
        BOOKS_COL_TYPES = {
            "GSTIN": "text", "Trade_Name": "text", "Invoice_No": "text",
            "Invoice_Date": "date", "Taxable_Value": "number", "CGST": "number",
            "SGST": "number", "IGST": "number", "CESS": "number",
            "TOTAL_TAX": "number", "Invoice_Value": "number",
        }

        # Books sheet
        if not r["books_raw"].empty:
            bdf = post_processing_cleaner(r["books_raw"].copy())
            ws_b = wb.add_worksheet("Books")
            write_sheet(ws_b, "Books Data", list(bdf.columns), bdf.values.tolist(), BOOKS_COL_TYPES)

        # GSTR-2B sheet
        if not r["gstr_raw"].empty:
            gdf = post_processing_cleaner(r["gstr_raw"].copy())
            ws_g = wb.add_worksheet("GSTR-2B")
            write_sheet(ws_g, "GSTR-2B Data", list(gdf.columns), gdf.values.tolist(), BOOKS_COL_TYPES)

        # Missing in 2B
        if not r["missing_2b"].empty:
            mdf = post_processing_cleaner(
                r["missing_2b"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "TOTAL_TAX"]].copy()
            )
            ws_m = wb.add_worksheet("Missing in 2B")
            write_sheet(ws_m, "Missing in 2B", list(mdf.columns), mdf.values.tolist(),
                {"GSTIN": "text", "Trade_Name": "text", "Invoice_No": "text",
                 "Invoice_Date": "date", "Taxable_Value": "number", "TOTAL_TAX": "number"})

        # Missing in Books
        if not r["missing_books"].empty:
            mbdf = post_processing_cleaner(
                r["missing_books"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "TOTAL_TAX"]].copy()
            )
            ws_mb = wb.add_worksheet("Missing in Books")
            write_sheet(ws_mb, "Missing in Books", list(mbdf.columns), mbdf.values.tolist(),
                {"GSTIN": "text", "Trade_Name": "text", "Invoice_No": "text",
                 "Invoice_Date": "date", "Taxable_Value": "number", "TOTAL_TAX": "number"})

        # Supplier Wise ITC Summary
        if not r["books_raw"].empty or not r["gstr_raw"].empty:
            tmap      = r.get("trade_name_mapping", {})
            all_gstns = set()
            if not r["books_raw"].empty: all_gstns.update(r["books_raw"]["GSTIN"].unique())
            if not r["gstr_raw"].empty:  all_gstns.update(r["gstr_raw"]["GSTIN"].unique())
            sup_data = []
            for gstin in all_gstns:
                ib = r["books_raw"][r["books_raw"]["GSTIN"] == gstin]["TOTAL_TAX"].sum() if not r["books_raw"].empty else 0
                ig = r["gstr_raw"][r["gstr_raw"]["GSTIN"] == gstin]["TOTAL_TAX"].sum()  if not r["gstr_raw"].empty  else 0
                sup_data.append({
                    "GSTIN": gstin, "Supplier": tmap.get(gstin, "Unknown"),
                    "ITC as per Books": round(ib, 2), "ITC as per 2B": round(ig, 2),
                    "ITC Difference": round(ig - ib, 2),
                })
            if sup_data:
                sdf = post_processing_cleaner(pd.DataFrame(sup_data))
                ws_sup = wb.add_worksheet("Supplier Wise ITC Summary")
                write_sheet(ws_sup, "Supplier Wise ITC Summary",
                    ["GSTIN", "Supplier", "ITC as per Books", "ITC as per 2B", "ITC Difference"],
                    sdf[["GSTIN", "Supplier", "ITC as per Books", "ITC as per 2B", "ITC Difference"]].values.tolist(),
                    {"GSTIN": "text", "Supplier": "text",
                     "ITC as per Books": "number", "ITC as per 2B": "number", "ITC Difference": "number"})

        # Supplier Drill Down
        if not r["books_raw"].empty or not r["gstr_raw"].empty:
            drilldown_data = []
            tmap2          = r.get("trade_name_mapping", {})
            proc_keys      = set()
            matched_pairs  = set()

            for label, df_src in [("matched", r.get("matched", pd.DataFrame())),
                                   ("tax_diff", r.get("tax_diff", pd.DataFrame()))]:
                if df_src.empty:
                    continue
                for _, row in df_src.iterrows():
                    gstin    = row.get("GSTIN_2B") if pd.notna(row.get("GSTIN_2B")) else row.get("GSTIN_Books")
                    inv_2b   = row.get("Invoice_No_2B")  if pd.notna(row.get("Invoice_No_2B"))  else None
                    inv_bk   = row.get("Invoice_No_Books") if pd.notna(row.get("Invoice_No_Books")) else None
                    inv_no   = inv_2b if inv_2b else inv_bk
                    inv_date = row.get("Invoice_Date_2B") if pd.notna(row.get("Invoice_Date_2B")) else row.get("Invoice_Date_Books")
                    itc_bk   = row.get("TOTAL_TAX_Books", 0)
                    itc_gstr = row.get("TOTAL_TAX_2B", 0)
                    diff     = row.get("TAX_DIFF", itc_gstr - itc_bk)
                    remark   = "Matched" if label == "matched" else "Tax Difference"
                    pk       = f"{gstin}|{inv_2b}|{inv_bk}"
                    if pk not in matched_pairs:
                        matched_pairs.add(pk)
                        drilldown_data.append({
                            "GSTIN": gstin, "Supplier": tmap2.get(gstin, "Unknown"),
                            "Invoice No": inv_no,
                            "Invoice Date": pd.to_datetime(inv_date) if pd.notna(inv_date) else None,
                            "ITC as per Books": itc_bk, "ITC as per 2B": itc_gstr,
                            "Difference": diff, "Remarks": remark,
                        })
                    if inv_2b: proc_keys.add(f"{gstin}|{inv_2b}")
                    if inv_bk: proc_keys.add(f"{gstin}|{inv_bk}")

            for _, row in r["books_raw"].iterrows():
                k = f"{row['GSTIN']}|{row['Invoice_No']}"
                if k not in proc_keys:
                    proc_keys.add(k)
                    drilldown_data.append({
                        "GSTIN": row["GSTIN"], "Supplier": tmap2.get(row["GSTIN"], row["Trade_Name"]),
                        "Invoice No": row["Invoice_No"],
                        "Invoice Date": pd.to_datetime(row["Invoice_Date"]),
                        "ITC as per Books": row["TOTAL_TAX"], "ITC as per 2B": 0.0,
                        "Difference": -row["TOTAL_TAX"], "Remarks": "Missing in GST",
                    })

            for _, row in r["gstr_raw"].iterrows():
                k = f"{row['GSTIN']}|{row['Invoice_No']}"
                if k not in proc_keys:
                    proc_keys.add(k)
                    drilldown_data.append({
                        "GSTIN": row["GSTIN"], "Supplier": tmap2.get(row["GSTIN"], row["Trade_Name"]),
                        "Invoice No": row["Invoice_No"],
                        "Invoice Date": pd.to_datetime(row["Invoice_Date"]),
                        "ITC as per Books": 0.0, "ITC as per 2B": row["TOTAL_TAX"],
                        "Difference": row["TOTAL_TAX"], "Remarks": "Missing in Books",
                    })

            if drilldown_data:
                ddf = pd.DataFrame(drilldown_data)
                for nc in ["ITC as per Books", "ITC as per 2B", "Difference"]:
                    if nc in ddf.columns:
                        ddf[nc] = ddf[nc].apply(strict_numeric_cleaner)
                ddf = ddf.sort_values(["Supplier", "Invoice Date"])

                ws_dd = wb.add_worksheet("Supplier Drill Down")
                ws_dd.write(0, 0, "Supplier Drill Down - Invoice Level Details", fmt_title)
                ws_dd.set_row(1, 20)
                dd_headers = ["GSTIN", "Supplier", "Invoice No", "Invoice Date",
                              "ITC as per Books", "ITC as per 2B", "Difference", "Remarks"]
                for ci, h in enumerate(dd_headers):
                    ws_dd.write(1, ci, h, fmt_header)
                for ri, row_data in ddf.iterrows():
                    safe_write_text(ws_dd,   ri + 2, 0, row_data["GSTIN"],            fmt_text)
                    safe_write_text(ws_dd,   ri + 2, 1, row_data["Supplier"],         fmt_text)
                    safe_write_text(ws_dd,   ri + 2, 2, row_data["Invoice No"],       fmt_text)
                    try:
                        ws_dd.write_datetime(ri + 2, 3, row_data["Invoice Date"],     fmt_date)
                    except Exception:
                        safe_write_text(ws_dd, ri + 2, 3, str(row_data["Invoice Date"]), fmt_text)
                    safe_write_number(ws_dd, ri + 2, 4, row_data["ITC as per Books"], fmt_number)
                    safe_write_number(ws_dd, ri + 2, 5, row_data["ITC as per 2B"],   fmt_number)
                    safe_write_number(ws_dd, ri + 2, 6, row_data["Difference"],       fmt_number)
                    safe_write_text(ws_dd,   ri + 2, 7, row_data["Remarks"],          fmt_text)
                ws_dd.autofilter(1, 0, 1, 7)
                ws_dd.set_column(0, 0, 20, fmt_text)
                ws_dd.set_column(1, 1, 30, fmt_text)
                ws_dd.set_column(2, 2, 20, fmt_text)
                ws_dd.set_column(3, 3, 15, fmt_date)
                ws_dd.set_column(4, 4, 15, fmt_number)
                ws_dd.set_column(5, 5, 15, fmt_number)
                ws_dd.set_column(6, 6, 15, fmt_number)
                ws_dd.set_column(7, 7, 15, fmt_text)
                ws_dd.freeze_panes(2, 0)
                ws_dd.autofit()

        # NO ITC sheet
        if not r["no_itc"].empty:
            nidf = post_processing_cleaner(
                r["no_itc"][["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Taxable_Value", "Invoice_Value"]].copy()
            )
            ws_ni = wb.add_worksheet("NO ITC")
            write_sheet(ws_ni, "Zero ITC Invoices", list(nidf.columns), nidf.values.tolist(),
                {"GSTIN": "text", "Trade_Name": "text", "Invoice_No": "text",
                 "Invoice_Date": "date", "Taxable_Value": "number", "Invoice_Value": "number"})

    output.seek(0)
    st.download_button(
        "📥 Download Excel Report",
        data=output,
        file_name=f"reconciliation_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
