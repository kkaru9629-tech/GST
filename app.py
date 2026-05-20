"""
GST Reconciliation App
Version: 12.0 — EXACT MATCH to Old Summary Workbook Structure
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

# Professional accounting style
st.markdown("""
<style>
    .stApp, .stMarkdown, .stText, .stCaption, .stMetric,
    .stTabs [data-baseweb="tab"], button, label, input {
        font-family: 'Segoe UI', 'Aptos Narrow', 'Calibri', 'Arial', sans-serif !important;
    }
    
    .dashboard-title {
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
        gap: 16px;
        margin: 8px 0 14px 0;
    }

    .dashboard-title h2 {
        margin: 0;
        color: #17324d;
        font-size: 24px;
        font-weight: 700;
    }

    .dashboard-subtitle {
        color: #607080;
        font-size: 13px;
    }

    .kpi-card, .count-card {
        background: #ffffff;
        border: 1px solid #dbe3ea;
        border-radius: 6px;
        padding: 16px 18px;
        min-height: 104px;
        box-shadow: 0 1px 2px rgba(23, 50, 77, 0.05);
    }

    .kpi-card {
        border-top: 4px solid #1f4e78;
    }

    .kpi-card.positive {
        border-top-color: #2e7d32;
    }

    .kpi-card.warning {
        border-top-color: #ed6c02;
    }

    .kpi-label, .count-label {
        color: #607080;
        font-size: 12px;
        font-weight: 600;
        letter-spacing: .02em;
        text-transform: uppercase;
        margin-bottom: 8px;
    }

    .kpi-value {
        color: #17324d;
        font-size: 25px;
        font-weight: 750;
        line-height: 1.15;
    }

    .kpi-note {
        color: #7a8793;
        font-size: 12px;
        margin-top: 8px;
    }

    .count-card {
        min-height: 88px;
        border-left: 4px solid #1f4e78;
    }

    .count-card.good {
        border-left-color: #2e7d32;
    }

    .count-card.warn {
        border-left-color: #ed6c02;
    }

    .count-card.danger {
        border-left-color: #c62828;
    }

    .count-value {
        color: #17324d;
        font-size: 24px;
        font-weight: 750;
        line-height: 1;
    }
    
    .stDataFrame table {
        font-size: 12px !important;
    }
    
    .stDataFrame thead tr th {
        background-color: #1e3a5f !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 10px 8px !important;
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
    help="Max acceptable tax difference (₹) for 'Matched'"
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
    parsed = pd.to_datetime(df["Invoice_Date"], errors="coerce")
    df["Month"] = parsed.dt.to_period("M").astype(str)
    df.loc[parsed.isna(), "Month"] = "Unknown"
    return df

def normalize_report_df(df: pd.DataFrame) -> pd.DataFrame:
    """Keep imported rows for reporting, with stable GSTIN and numeric totals."""
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    for col in ["GSTIN", "Trade_Name", "Invoice_No"]:
        if col not in df.columns:
            df[col] = ""
    for col in ["Taxable_Value", "CGST", "SGST", "IGST", "CESS", "TOTAL_TAX", "Invoice_Value"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = df[col].apply(strict_numeric_cleaner)
    df["GSTIN"] = df["GSTIN"].fillna("").astype(str).str.strip().str.upper()
    df["Trade_Name"] = df["Trade_Name"].fillna("").astype(str).str.strip()
    df["Invoice_No"] = df["Invoice_No"].fillna("").astype(str).str.strip()
    return add_month_column(df)

def build_books_report_df(books_clean: pd.DataFrame, no_itc: pd.DataFrame, issues: pd.DataFrame) -> pd.DataFrame:
    parts = [df for df in [books_clean, no_itc, issues] if df is not None and not df.empty]
    return normalize_report_df(pd.concat(parts, ignore_index=True, sort=False)) if parts else pd.DataFrame()

def apply_report_totals(summary: dict, books_report: pd.DataFrame, gstr_report: pd.DataFrame) -> dict:
    summary = summary.copy()
    books_itc = books_report["TOTAL_TAX"].sum() if not books_report.empty and "TOTAL_TAX" in books_report.columns else 0.0
    gstr_itc = gstr_report["TOTAL_TAX"].sum() if not gstr_report.empty and "TOTAL_TAX" in gstr_report.columns else 0.0
    summary["ITC_Books"] = round(float(books_itc), 2)
    summary["ITC_GSTR"] = round(float(gstr_itc), 2)
    summary["ITC_Diff"] = round(float(gstr_itc - books_itc), 2)
    summary["Total_Books"] = len(books_report)
    summary["Total_GSTR"] = len(gstr_report)
    summary["Match_%"] = round(summary.get("Matched", 0) / len(gstr_report) * 100, 2) if len(gstr_report) else 0
    return summary

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
                gstr_clean = normalize_report_df(gstr_clean)
                books_report = build_books_report_df(books_clean, no_itc, issues)
                gstr_report = normalize_report_df(gstr_clean)

                results = reconcile(gstr_clean, books_clean, tolerance)
                results["summary"] = apply_report_totals(results["summary"], books_report, gstr_report)
                results.update({
                    "no_itc": no_itc,
                    "issues": issues,
                    "books_raw": books_report,
                    "gstr_raw": gstr_report,
                    "books_recon": books_clean,
                    "gstr_recon": gstr_clean,
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
    "⚠️ Date Mismatch": "Verify invoice date",
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

def same_date(left, right) -> bool:
    try:
        ldt = pd.to_datetime(left, errors="coerce")
        rdt = pd.to_datetime(right, errors="coerce")
        return pd.notna(ldt) and pd.notna(rdt) and ldt.normalize() == rdt.normalize()
    except Exception:
        return False

def is_true_flag(value) -> bool:
    return value is True or str(value).strip().lower() == "true"

def fmt_month_label(val, fmt="%B %Y") -> str:
    try:
        if pd.isna(val) or val in ("", "NaT", "Unknown"):
            return val or ""
        return pd.to_datetime(str(val) + "-01", format="%Y-%m-%d").strftime(fmt)
    except Exception:
        return str(val)

def month_label_from_row(row) -> str:
    month = row.get("Month", "")
    if pd.notna(month) and str(month).strip() not in ("", "NaT", "Unknown"):
        return fmt_month_label(month)
    try:
        date_val = pd.to_datetime(row.get("Date"), errors="coerce")
        if pd.notna(date_val):
            return date_val.strftime("%B %Y")
    except Exception:
        pass
    return ""

def autofit_worksheet(worksheet):
    try:
        worksheet.autofit()
    except Exception:
        pass

# =================== BUILD DETAIL DF FROM RECONCILIATION RESULTS =================== #

def build_detail_df_from_results(matched_df, missing_2b_df, missing_books_df, trade_name_map, tolerance):
    rows = []
    
    if not matched_df.empty:
        for _, row in matched_df.iterrows():
            gstin = row.get("GSTIN_2B") or row.get("GSTIN_Books", "")
            supplier = trade_name_map.get(gstin, row.get("Trade_Name_2B") or row.get("Trade_Name_Books", ""))
            month = row.get("Month_2B") or row.get("Month_Books", "")
            inv_no = row.get("Invoice_No_2B") or row.get("Invoice_No_Books", "")
            gstr_date = row.get("Invoice_Date_2B")
            books_date = row.get("Invoice_Date_Books")
            inv_date = gstr_date or books_date
            itc_books = float(row.get("TOTAL_TAX_Books", 0) or 0)
            itc_gstr = float(row.get("TOTAL_TAX_2B", 0) or 0)
            diff = itc_gstr - itc_books
            date_mismatch = bool(row.get("DATE_MISMATCH", False)) and not same_date(gstr_date, books_date)
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
                "Date Mismatch": date_mismatch,
                "Action Required": ACTION_MAP.get(remark, "")
            })
    
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
    rows = []
    
    all_months = set()
    if not books_raw.empty and "Month" in books_raw.columns:
        all_months.update(books_raw["Month"].dropna().unique())
    if not gstr_raw.empty and "Month" in gstr_raw.columns:
        all_months.update(gstr_raw["Month"].dropna().unique())
    all_months.discard("")
    all_months.discard("NaT")
    
    for month in sorted(all_months):
        b_tax = books_raw[books_raw["Month"] == month]["TOTAL_TAX"].sum() if not books_raw.empty else 0
        g_tax = gstr_raw[gstr_raw["Month"] == month]["TOTAL_TAX"].sum() if not gstr_raw.empty else 0
        
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

n_files = r.get("n_gstr_files", 1)
file_caption = f"{n_files} GSTR-2B file(s) processed" if n_files > 1 else "Books vs GSTR-2B"

st.markdown(f"""
<div style="background:#f6f8fb;border:1px solid #dbe3ea;border-radius:8px;padding:18px;margin:8px 0 18px 0;">
    <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:14px;">
        <div>
            <div style="font-size:24px;font-weight:750;color:#17324d;line-height:1.1;">Reconciliation Summary</div>
            <div style="font-size:13px;color:#607080;margin-top:4px;">{file_caption}</div>
        </div>
        <div style="font-size:13px;color:#607080;">Tolerance: &#8377; {tol:,.2f}</div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-bottom:14px;">
        <div style="background:white;border:1px solid #dbe3ea;border-top:5px solid #2e7d32;border-radius:6px;padding:16px 18px;box-shadow:0 2px 8px rgba(23,50,77,.06);">
            <div style="font-size:12px;font-weight:700;color:#607080;text-transform:uppercase;letter-spacing:.03em;">ITC - Books</div>
            <div style="font-size:27px;font-weight:800;color:#17324d;margin-top:8px;">&#8377; {s['ITC_Books']:,.2f}</div>
            <div style="font-size:12px;color:#7a8793;margin-top:8px;">Purchase register value</div>
        </div>
        <div style="background:white;border:1px solid #dbe3ea;border-top:5px solid #1f4e78;border-radius:6px;padding:16px 18px;box-shadow:0 2px 8px rgba(23,50,77,.06);">
            <div style="font-size:12px;font-weight:700;color:#607080;text-transform:uppercase;letter-spacing:.03em;">ITC - GSTR-2B</div>
            <div style="font-size:27px;font-weight:800;color:#17324d;margin-top:8px;">&#8377; {s['ITC_GSTR']:,.2f}</div>
            <div style="font-size:12px;color:#7a8793;margin-top:8px;">GST portal value</div>
        </div>
        <div style="background:white;border:1px solid #dbe3ea;border-top:5px solid #ed6c02;border-radius:6px;padding:16px 18px;box-shadow:0 2px 8px rgba(23,50,77,.06);">
            <div style="font-size:12px;font-weight:700;color:#607080;text-transform:uppercase;letter-spacing:.03em;">Difference</div>
            <div style="font-size:27px;font-weight:800;color:{'#c62828' if s['ITC_Diff'] < 0 else '#17324d'};margin-top:8px;">&#8377; {s['ITC_Diff']:,.2f}</div>
            <div style="font-size:12px;color:#7a8793;margin-top:8px;">GSTR-2B minus Books</div>
        </div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;">
        <div style="background:white;border:1px solid #dbe3ea;border-left:5px solid #1f4e78;border-radius:6px;padding:13px 14px;">
            <div style="font-size:11px;font-weight:700;color:#607080;text-transform:uppercase;">Invoices in Books</div>
            <div style="font-size:25px;font-weight:800;color:#17324d;margin-top:7px;">{s['Total_Books']:,}</div>
        </div>
        <div style="background:white;border:1px solid #dbe3ea;border-left:5px solid #1f4e78;border-radius:6px;padding:13px 14px;">
            <div style="font-size:11px;font-weight:700;color:#607080;text-transform:uppercase;">Invoices in GST</div>
            <div style="font-size:25px;font-weight:800;color:#17324d;margin-top:7px;">{s['Total_GSTR']:,}</div>
        </div>
        <div style="background:white;border:1px solid #dbe3ea;border-left:5px solid #c62828;border-radius:6px;padding:13px 14px;">
            <div style="font-size:11px;font-weight:700;color:#607080;text-transform:uppercase;">Missing in 2B</div>
            <div style="font-size:25px;font-weight:800;color:#17324d;margin-top:7px;">{s['Missing_2B']:,}</div>
        </div>
        <div style="background:white;border:1px solid #dbe3ea;border-left:5px solid #c62828;border-radius:6px;padding:13px 14px;">
            <div style="font-size:11px;font-weight:700;color:#607080;text-transform:uppercase;">Missing in Books</div>
            <div style="font-size:25px;font-weight:800;color:#17324d;margin-top:7px;">{s['Missing_Books']:,}</div>
        </div>
        <div style="background:white;border:1px solid #dbe3ea;border-left:5px solid #ed6c02;border-radius:6px;padding:13px 14px;">
            <div style="font-size:11px;font-weight:700;color:#607080;text-transform:uppercase;">Tax Difference</div>
            <div style="font-size:25px;font-weight:800;color:#17324d;margin-top:7px;">{s['Tax_Diff']:,}</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# =================== BUILD DATA =================== #

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

# =================== MONTH-WISE SUMMARY =================== #

if not month_summary.empty:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("## Month-wise Summary")
    
    month_summary_display = month_summary.copy()
    month_summary_display["Month"] = month_summary_display["Month"].apply(fmt_month_label)
    
    st.dataframe(month_summary_display, use_container_width=True, hide_index=True)
    st.caption(f"{len(month_summary)} month(s) of data")

# =================== FILTER CONTROLS =================== #

st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("### Filter Results")
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
    filtered_display_df = filtered_df.drop(columns=["Date Mismatch"], errors="ignore")
    safe_dataframe(filtered_display_df, column_config=col_config,
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
        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown("## Zero ITC Invoices")
        st.dataframe(no_itc_df, use_container_width=True, hide_index=True)

# =================== EXCEL EXPORT - EXACT MATCH TO OLD WORKBOOK ===================

def create_summary_sheet_exact_match(writer, r):
    """Create Summary sheet EXACTLY matching old workbook (Row 0: Title, Row 1: header, Row 2+: data)"""
    workbook = writer.book
    worksheet = workbook.add_worksheet('Summary')
    
    # Define formats - Aptos Narrow, row height 18
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left',
        'valign': 'vcenter'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 10,
        'font_name': 'Aptos Narrow',
        'align': 'left',
        'valign': 'vcenter',
        'bg_color': '#F0F0F0'
    })
    
    label_format = workbook.add_format({
        'bold': False,
        'font_size': 10,
        'font_name': 'Aptos Narrow',
        'align': 'left',
        'valign': 'vcenter'
    })
    
    value_format = workbook.add_format({
        'bold': False,
        'font_size': 10,
        'font_name': 'Aptos Narrow',
        'align': 'right',
        'valign': 'vcenter',
        'num_format': '#,##0.00'
    })
    
    value_int_format = workbook.add_format({
        'bold': False,
        'font_size': 10,
        'font_name': 'Aptos Narrow',
        'align': 'right',
        'valign': 'vcenter',
        'num_format': '#,##0'
    })
    
    worksheet.set_default_row(18)
    
    # Row 0: Title
    worksheet.write(0, 0, 'Reconciliation Summary', title_format)
    
    # Row 1: Headers
    worksheet.write(1, 0, 'Particulars', header_format)
    worksheet.write(1, 1, 'Value', header_format)
    
    # Row 2: ITC - Books
    worksheet.write(2, 0, 'ITC - Books', label_format)
    worksheet.write(2, 1, r['summary']['ITC_Books'], value_format)
    
    # Row 3: ITC - GSTR-2B
    worksheet.write(3, 0, 'ITC - GSTR-2B', label_format)
    worksheet.write(3, 1, r['summary']['ITC_GSTR'], value_format)
    
    # Row 4: Difference
    worksheet.write(4, 0, 'Difference', label_format)
    worksheet.write(4, 1, r['summary']['ITC_Diff'], value_format)
    
    # Blank row
    worksheet.write(5, 0, '', label_format)
    
    # Row 6: No of invoices in Books
    worksheet.write(6, 0, 'No of invoices in Books', label_format)
    worksheet.write(6, 1, r['summary']['Total_Books'], value_int_format)
    
    # Row 7: No of invoices in GST
    worksheet.write(7, 0, 'No of invoices in GST', label_format)
    worksheet.write(7, 1, r['summary']['Total_GSTR'], value_int_format)
    
    # Set column widths
    worksheet.set_column(0, 0, 22)
    worksheet.set_column(1, 1, 18)
    
    # Freeze pane at A3 (row 3, col 0)
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_monthwise_summary_exact_match(writer, month_summary):
    """Create Month-wise Summary sheet EXACTLY matching old workbook"""
    if month_summary.empty:
        return
    
    workbook = writer.book
    worksheet = workbook.add_worksheet('Month-wise Summary')
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    money_format = workbook.add_format({
        'num_format': '#,##0.00',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    int_format = workbook.add_format({
        'num_format': '#,##0',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })

    worksheet.set_default_row(18)
    
    # Row 0: Title
    worksheet.write(0, 0, 'Month-wise ITC Summary', title_format)
    
    # Prepare display data with formatted month names (Mar-2026, Apr-2026 format)
    display_df = month_summary.copy()
    display_df['Month'] = display_df['Month'].apply(lambda m: fmt_month_label(m, "%b-%Y"))
    
    # Row 2: Headers (old workbook has headers at row 2)
    headers = list(display_df.columns)
    for col_idx, header in enumerate(headers):
        worksheet.write(2, col_idx, header, header_format)
    
    # Data starting at row 3
    for row_idx, (_, row) in enumerate(display_df.iterrows()):
        for col_idx, header in enumerate(headers):
            val = row[header]
            col_lower = header.lower()
            
            if 'itc' in col_lower or 'difference' in col_lower:
                worksheet.write(3 + row_idx, col_idx, float(val) if pd.notna(val) else 0, money_format)
            elif 'missing' in col_lower or 'matched' in col_lower:
                worksheet.write(3 + row_idx, col_idx, int(val) if pd.notna(val) else 0, int_format)
            else:
                worksheet.write(3 + row_idx, col_idx, str(val) if pd.notna(val) else "", text_format)
    
    # Auto-fit columns
    for col_idx, header in enumerate(headers):
        max_len = len(header)
        if not display_df.empty:
            for _, row in display_df.iterrows():
                val = str(row[header]) if pd.notna(row[header]) else ""
                max_len = max(max_len, min(len(val), 25))
        worksheet.set_column(col_idx, col_idx, max_len + 2)
    
    # Freeze pane at A3
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_books_sheet_exact_match(writer, books_raw):
    """Create Books sheet EXACTLY matching old workbook"""
    if books_raw.empty:
        return
    
    workbook = writer.book
    worksheet = workbook.add_worksheet('Books')
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    money_format = workbook.add_format({
        'num_format': '#,##0.00',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    date_format = workbook.add_format({
        'num_format': 'dd-mmm-yyyy',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })

    worksheet.set_default_row(18)
    
    # Row 0: Title
    worksheet.write(0, 0, 'Books Data', title_format)
    
    # Prepare data - ensure date format
    books_export = books_raw.copy()
    if 'Invoice_Date' in books_export.columns:
        books_export['Invoice_Date'] = pd.to_datetime(books_export['Invoice_Date'], errors='coerce')
    
    # Row 2: Headers
    headers = list(books_export.columns)
    for col_idx, header in enumerate(headers):
        worksheet.write(2, col_idx, header, header_format)
    
    # Data starting at row 3
    for row_idx, (_, row) in enumerate(books_export.iterrows()):
        for col_idx, header in enumerate(headers):
            val = row[header]
            col_lower = header.lower()
            
            if any(term in col_lower for term in ['taxable', 'value', 'cgst', 'sgst', 'igst', 'cess', 'total_tax', 'invoice_value']):
                worksheet.write(3 + row_idx, col_idx, float(val) if pd.notna(val) else 0, money_format)
            elif 'date' in col_lower:
                if pd.notna(val):
                    try:
                        dt = pd.to_datetime(val)
                        worksheet.write_datetime(3 + row_idx, col_idx, dt.to_pydatetime(), date_format)
                    except:
                        worksheet.write(3 + row_idx, col_idx, str(val), text_format)
                else:
                    worksheet.write(3 + row_idx, col_idx, "", text_format)
            else:
                worksheet.write(3 + row_idx, col_idx, str(val) if pd.notna(val) else "", text_format)
    
    # Set column widths
    for col_idx, header in enumerate(headers):
        max_len = len(header)
        if not books_export.empty:
            for row_idx in range(min(100, len(books_export))):
                val = str(books_export.iloc[row_idx][header]) if pd.notna(books_export.iloc[row_idx][header]) else ""
                max_len = max(max_len, min(len(val), 25))
        worksheet.set_column(col_idx, col_idx, max_len + 2)
    
    # Freeze pane at A3
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_gstr_sheet_exact_match(writer, gstr_raw):
    """Create GSTR-2B sheet EXACTLY matching old workbook"""
    if gstr_raw.empty:
        return
    
    workbook = writer.book
    worksheet = workbook.add_worksheet('GSTR-2B')
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    money_format = workbook.add_format({
        'num_format': '#,##0.00',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    date_format = workbook.add_format({
        'num_format': 'dd-mmm-yyyy',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    worksheet.set_default_row(18)
    
    # Row 0: Title
    worksheet.write(0, 0, 'GSTR-2B Data', title_format)
    
    # Prepare data
    gstr_export = gstr_raw.copy()
    if 'Invoice_Date' in gstr_export.columns:
        gstr_export['Invoice_Date'] = pd.to_datetime(gstr_export['Invoice_Date'], errors='coerce')
    
    # Row 2: Headers
    headers = list(gstr_export.columns)
    for col_idx, header in enumerate(headers):
        worksheet.write(2, col_idx, header, header_format)
    
    # Data starting at row 3
    for row_idx, (_, row) in enumerate(gstr_export.iterrows()):
        for col_idx, header in enumerate(headers):
            val = row[header]
            col_lower = header.lower()
            
            if any(term in col_lower for term in ['taxable', 'value', 'cgst', 'sgst', 'igst', 'cess', 'total_tax', 'invoice_value']):
                worksheet.write(3 + row_idx, col_idx, float(val) if pd.notna(val) else 0, money_format)
            elif 'date' in col_lower:
                if pd.notna(val):
                    try:
                        dt = pd.to_datetime(val)
                        worksheet.write_datetime(3 + row_idx, col_idx, dt.to_pydatetime(), date_format)
                    except:
                        worksheet.write(3 + row_idx, col_idx, str(val), text_format)
                else:
                    worksheet.write(3 + row_idx, col_idx, "", text_format)
            else:
                worksheet.write(3 + row_idx, col_idx, str(val) if pd.notna(val) else "", text_format)
    
    # Set column widths
    for col_idx, header in enumerate(headers):
        max_len = len(header)
        if not gstr_export.empty:
            for row_idx in range(min(100, len(gstr_export))):
                val = str(gstr_export.iloc[row_idx][header]) if pd.notna(gstr_export.iloc[row_idx][header]) else ""
                max_len = max(max_len, min(len(val), 25))
        worksheet.set_column(col_idx, col_idx, max_len + 2)
    
    # Freeze pane at A3
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_missing_in_2b_sheet_exact_match(writer, r, trade_name_map):
    """Create Missing in 2B sheet EXACTLY matching old workbook"""
    missing_2b = r.get('missing_2b', pd.DataFrame())
    if missing_2b.empty:
        return
    
    workbook = writer.book
    worksheet = workbook.add_worksheet('Missing in 2B')
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    money_format = workbook.add_format({
        'num_format': '#,##0.00',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    date_format = workbook.add_format({
        'num_format': 'dd-mmm-yyyy',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    worksheet.set_default_row(18)
    
    # Row 0: Title
    worksheet.write(0, 0, 'Missing in 2B', title_format)
    
    # Prepare data
    m2b = missing_2b.copy()
    if 'Invoice_Date' in m2b.columns:
        m2b['Invoice_Date'] = pd.to_datetime(m2b['Invoice_Date'], errors='coerce')
    m2b['Supplier'] = m2b['GSTIN'].map(trade_name_map).fillna(m2b.get('Trade_Name', ''))
    
    # Row 2: Headers
    headers = ['GSTIN', 'Trade_Name', 'Invoice_No', 'Invoice_Date', 'Taxable_Value', 'TOTAL_TAX']
    header_labels = ['GSTIN', 'Trade_Name', 'Invoice_No', 'Invoice_Date', 'Taxable_Value', 'TOTAL_TAX']
    
    for col_idx, header in enumerate(headers):
        worksheet.write(2, col_idx, header_labels[col_idx], header_format)
    
    # Data starting at row 3
    for row_idx, (_, row) in enumerate(m2b.iterrows()):
        worksheet.write(3 + row_idx, 0, str(row.get('GSTIN', '')), text_format)
        worksheet.write(3 + row_idx, 1, str(row.get('Supplier', '')), text_format)
        worksheet.write(3 + row_idx, 2, str(row.get('Invoice_No', '')), text_format)
        
        if pd.notna(row.get('Invoice_Date')):
            try:
                dt = pd.to_datetime(row['Invoice_Date'])
                worksheet.write_datetime(3 + row_idx, 3, dt.to_pydatetime(), date_format)
            except:
                worksheet.write(3 + row_idx, 3, str(row.get('Invoice_Date', '')), text_format)
        else:
            worksheet.write(3 + row_idx, 3, "", text_format)
        
        worksheet.write(3 + row_idx, 4, float(row.get('Taxable_Value', 0)), money_format)
        worksheet.write(3 + row_idx, 5, float(row.get('TOTAL_TAX', 0)), money_format)
    
    # Set column widths
    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 35)
    worksheet.set_column(2, 2, 20)
    worksheet.set_column(3, 3, 12)
    worksheet.set_column(4, 4, 15)
    worksheet.set_column(5, 5, 12)
    
    # Freeze pane at A3
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_missing_in_books_sheet_exact_match(writer, r, trade_name_map):
    """Create Missing in Books sheet EXACTLY matching old workbook"""
    missing_books = r.get('missing_books', pd.DataFrame())
    if missing_books.empty:
        return
    
    workbook = writer.book
    worksheet = workbook.add_worksheet('Missing in Books')
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    money_format = workbook.add_format({
        'num_format': '#,##0.00',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    date_format = workbook.add_format({
        'num_format': 'dd-mmm-yyyy',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    worksheet.set_default_row(18)
    
    # Row 0: Title
    worksheet.write(0, 0, 'Missing in Books', title_format)
    
    # Prepare data
    mb = missing_books.copy()
    if 'Invoice_Date' in mb.columns:
        mb['Invoice_Date'] = pd.to_datetime(mb['Invoice_Date'], errors='coerce')
    mb['Supplier'] = mb['GSTIN'].map(trade_name_map).fillna(mb.get('Trade_Name', ''))
    
    # Row 2: Headers
    headers = ['GSTIN', 'Trade_Name', 'Invoice_No', 'Invoice_Date', 'Taxable_Value', 'TOTAL_TAX']
    
    for col_idx, header in enumerate(headers):
        worksheet.write(2, col_idx, header, header_format)
    
    # Data starting at row 3
    for row_idx, (_, row) in enumerate(mb.iterrows()):
        worksheet.write(3 + row_idx, 0, str(row.get('GSTIN', '')), text_format)
        worksheet.write(3 + row_idx, 1, str(row.get('Supplier', '')), text_format)
        worksheet.write(3 + row_idx, 2, str(row.get('Invoice_No', '')), text_format)
        
        if pd.notna(row.get('Invoice_Date')):
            try:
                dt = pd.to_datetime(row['Invoice_Date'])
                worksheet.write_datetime(3 + row_idx, 3, dt.to_pydatetime(), date_format)
            except:
                worksheet.write(3 + row_idx, 3, str(row.get('Invoice_Date', '')), text_format)
        else:
            worksheet.write(3 + row_idx, 3, "", text_format)
        
        worksheet.write(3 + row_idx, 4, float(row.get('Taxable_Value', 0)), money_format)
        worksheet.write(3 + row_idx, 5, float(row.get('TOTAL_TAX', 0)), money_format)
    
    # Set column widths
    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 35)
    worksheet.set_column(2, 2, 20)
    worksheet.set_column(3, 3, 12)
    worksheet.set_column(4, 4, 15)
    worksheet.set_column(5, 5, 12)
    
    # Freeze pane at A3
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_supplier_wise_itc_summary_exact_match(writer, r, trade_name_map):
    """Create Supplier Wise ITC Summary sheet EXACTLY matching old workbook"""
    books_raw = normalize_report_df(r.get('books_raw', pd.DataFrame()))
    gstr_raw = normalize_report_df(r.get('gstr_raw', pd.DataFrame()))
    
    suppliers = set()
    if not books_raw.empty:
        suppliers.update(books_raw['GSTIN'].dropna().unique())
    if not gstr_raw.empty:
        suppliers.update(gstr_raw['GSTIN'].dropna().unique())
    
    rows = []
    for gstin in sorted(g for g in suppliers if str(g).strip()):
        supplier_name = trade_name_map.get(gstin, '')
        if not supplier_name and not books_raw.empty:
            names = books_raw.loc[books_raw['GSTIN'] == gstin, 'Trade_Name'].dropna().astype(str).str.strip()
            supplier_name = next((n for n in names if n), '')
        if not supplier_name and not gstr_raw.empty:
            names = gstr_raw.loc[gstr_raw['GSTIN'] == gstin, 'Trade_Name'].dropna().astype(str).str.strip()
            supplier_name = next((n for n in names if n), '')
        books_itc = books_raw.loc[books_raw['GSTIN'] == gstin, 'TOTAL_TAX'].sum() if not books_raw.empty else 0
        gstr_itc = gstr_raw.loc[gstr_raw['GSTIN'] == gstin, 'TOTAL_TAX'].sum() if not gstr_raw.empty else 0
        diff = gstr_itc - books_itc
        
        rows.append({
            'GSTIN': gstin,
            'Supplier': supplier_name,
            'ITC as per Books': round(float(books_itc), 2),
            'ITC as per 2B': round(float(gstr_itc), 2),
            'ITC Difference': round(float(diff), 2),
        })
    
    if not rows:
        return
    
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values('ITC as per 2B', ascending=False)
    
    workbook = writer.book
    worksheet = workbook.add_worksheet('Supplier Wise ITC Summary')
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    money_format = workbook.add_format({
        'num_format': '#,##0.00',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    worksheet.set_default_row(18)
    
    # Row 0: Title
    worksheet.write(0, 0, 'Supplier Wise ITC Summary', title_format)
    
    # Row 2: Headers
    headers = ['GSTIN', 'Supplier', 'ITC as per Books', 'ITC as per 2B', 'ITC Difference']
    for col_idx, header in enumerate(headers):
        worksheet.write(2, col_idx, header, header_format)
    
    # Data starting at row 3
    for row_idx, (_, row) in enumerate(df.iterrows()):
        worksheet.write(3 + row_idx, 0, str(row['GSTIN']), text_format)
        worksheet.write(3 + row_idx, 1, str(row['Supplier']), text_format)
        worksheet.write(3 + row_idx, 2, row['ITC as per Books'], money_format)
        worksheet.write(3 + row_idx, 3, row['ITC as per 2B'], money_format)
        worksheet.write(3 + row_idx, 4, row['ITC Difference'], money_format)
    
    # Set column widths
    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 35)
    worksheet.set_column(2, 2, 15)
    worksheet.set_column(3, 3, 15)
    worksheet.set_column(4, 4, 15)
    
    # Freeze pane at A3
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_supplier_drill_down_exact_match(writer, detail_df, r, trade_name_map):
    """Create Supplier Drill Down sheet EXACTLY matching old workbook"""
    if detail_df.empty:
        return
    
    workbook = writer.book
    worksheet = workbook.add_worksheet('Supplier Drill Down')
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    money_format = workbook.add_format({
        'num_format': '#,##0.00',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    date_format = workbook.add_format({
        'num_format': 'dd-mmm-yyyy',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })

    worksheet.set_default_row(18)
    
    date_mismatch_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow',
        'bg_color': '#FFF2CC'
    })

    # Row 0: Title
    worksheet.write(0, 0, 'Supplier Drill Down', title_format)
    
    # Prepare data
    drill_data = detail_df.copy()
    if 'Date' in drill_data.columns:
        drill_data['Date'] = pd.to_datetime(drill_data['Date'], errors='coerce')
    
    # Row 2: Headers
    headers = ['Month', 'GSTIN', 'Supplier', 'Invoice No', 'Invoice Date', 'ITC Books', 'ITC 2B', 'Difference', 'Remarks', 'Action Required']
    for col_idx, header in enumerate(headers):
        worksheet.write(2, col_idx, header, header_format)
    
    # Data starting at row 3
    for row_idx, (_, row) in enumerate(drill_data.iterrows()):
        worksheet.write(3 + row_idx, 0, month_label_from_row(row), text_format)
        worksheet.write(3 + row_idx, 1, str(row.get('GSTIN', '')), text_format)
        worksheet.write(3 + row_idx, 2, str(row.get('Supplier', '')), text_format)
        worksheet.write(3 + row_idx, 3, str(row.get('Invoice No', '')), text_format)
        
        if pd.notna(row.get('Date')):
            try:
                dt = pd.to_datetime(row['Date'])
                worksheet.write_datetime(3 + row_idx, 4, dt.to_pydatetime(), date_format)
            except:
                worksheet.write(3 + row_idx, 4, str(row.get('Date', '')), text_format)
        else:
            worksheet.write(3 + row_idx, 4, "", text_format)
        
        worksheet.write(3 + row_idx, 5, row.get('ITC Books', 0), money_format)
        worksheet.write(3 + row_idx, 6, row.get('ITC 2B', 0), money_format)
        worksheet.write(3 + row_idx, 7, row.get('Difference', 0), money_format)
        date_mismatch = is_true_flag(row.get('Date Mismatch', False))
        remarks = 'Date Mismatch' if date_mismatch else str(row.get('Remarks', ''))
        action_required = 'Verify invoice date' if date_mismatch else str(row.get('Action Required', ''))
        remark_format = date_mismatch_format if date_mismatch else text_format
        worksheet.write(3 + row_idx, 8, remarks, remark_format)
        worksheet.write(3 + row_idx, 9, action_required, text_format)
    
    # Set column widths
    worksheet.set_column(0, 0, 14)
    worksheet.set_column(1, 1, 18)
    worksheet.set_column(2, 2, 35)
    worksheet.set_column(3, 3, 20)
    worksheet.set_column(4, 4, 12)
    worksheet.set_column(5, 5, 12)
    worksheet.set_column(6, 6, 12)
    worksheet.set_column(7, 7, 12)
    worksheet.set_column(8, 8, 15)
    worksheet.set_column(9, 9, 20)
    
    # Freeze pane at A3
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_no_itc_sheet_exact_match(writer, r):
    """Create NO ITC sheet EXACTLY matching old workbook"""
    no_itc = r.get('no_itc', pd.DataFrame())
    if no_itc.empty:
        return
    
    # Filter for Invoice_Value > 0
    no_itc = no_itc[no_itc['Invoice_Value'].astype(float) > 0]
    if no_itc.empty:
        return
    
    workbook = writer.book
    worksheet = workbook.add_worksheet('NO ITC')
    
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Aptos Narrow',
        'align': 'left'
    })
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#1F4E78',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    money_format = workbook.add_format({
        'num_format': '#,##0.00',
        'align': 'right',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    date_format = workbook.add_format({
        'num_format': 'dd-mmm-yyyy',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    text_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Aptos Narrow'
    })
    
    worksheet.set_default_row(18)
    
    # Row 0: Title
    worksheet.write(0, 0, 'Zero ITC Invoices', title_format)
    
    # Prepare data
    no_itc_export = no_itc[['GSTIN', 'Trade_Name', 'Invoice_No', 'Invoice_Date', 'Taxable_Value', 'Invoice_Value']].copy()
    if 'Invoice_Date' in no_itc_export.columns:
        no_itc_export['Invoice_Date'] = pd.to_datetime(no_itc_export['Invoice_Date'], errors='coerce')
    
    # Row 2: Headers
    headers = ['GSTIN', 'Trade_Name', 'Invoice_No', 'Invoice_Date', 'Taxable_Value', 'Invoice_Value']
    for col_idx, header in enumerate(headers):
        worksheet.write(2, col_idx, header, header_format)
    
    # Data starting at row 3
    for row_idx, (_, row) in enumerate(no_itc_export.iterrows()):
        worksheet.write(3 + row_idx, 0, str(row.get('GSTIN', '')), text_format)
        worksheet.write(3 + row_idx, 1, str(row.get('Trade_Name', '')), text_format)
        worksheet.write(3 + row_idx, 2, str(row.get('Invoice_No', '')), text_format)
        
        if pd.notna(row.get('Invoice_Date')):
            try:
                dt = pd.to_datetime(row['Invoice_Date'])
                worksheet.write_datetime(3 + row_idx, 3, dt.to_pydatetime(), date_format)
            except:
                worksheet.write(3 + row_idx, 3, str(row.get('Invoice_Date', '')), text_format)
        else:
            worksheet.write(3 + row_idx, 3, "", text_format)
        
        worksheet.write(3 + row_idx, 4, float(row.get('Taxable_Value', 0)), money_format)
        worksheet.write(3 + row_idx, 5, float(row.get('Invoice_Value', 0)), money_format)
    
    # Set column widths
    worksheet.set_column(0, 0, 18)
    worksheet.set_column(1, 1, 35)
    worksheet.set_column(2, 2, 20)
    worksheet.set_column(3, 3, 12)
    worksheet.set_column(4, 4, 15)
    worksheet.set_column(5, 5, 15)
    
    # Freeze pane at A3
    autofit_worksheet(worksheet)
    worksheet.freeze_panes(3, 0)


def create_monthly_detail_sheets_exact_match(writer, r, trade_name_map, tolerance, detail_df):
    """Create monthly detail sheets (Mar-2026, Apr-2026, etc.) EXACTLY matching old workbook"""
    if detail_df.empty or 'Month' not in detail_df.columns:
        return

    all_months = [m for m in detail_df['Month'].dropna().unique() if m and m not in ('', 'NaT', 'Unknown')]
    
    if not all_months:
        return
    
    workbook = writer.book
    
    # Create detail data for each month
    for month in sorted(all_months):
        month_name = pd.to_datetime(month + '-01', format='%Y-%m-%d').strftime('%b-%Y')
        month_df = detail_df[detail_df['Month'] == month].copy()
        if not month_df.empty:
            month_df['Invoice Date'] = pd.to_datetime(month_df.get('Date', ''), errors='coerce')
            
            worksheet = workbook.add_worksheet(month_name)
            
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'font_name': 'Aptos Narrow',
                'align': 'left'
            })
            
            header_format = workbook.add_format({
                'bold': True,
                'font_color': '#FFFFFF',
                'bg_color': '#1F4E78',
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Aptos Narrow'
            })
            
            money_format = workbook.add_format({
                'num_format': '#,##0.00',
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Aptos Narrow'
            })
            
            date_format = workbook.add_format({
                'num_format': 'dd-mmm-yyyy',
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Aptos Narrow'
            })
            
            text_format = workbook.add_format({
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Aptos Narrow'
            })

            date_mismatch_format = workbook.add_format({
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Aptos Narrow',
                'bg_color': '#FFF2CC'
            })
            
            worksheet.set_default_row(18)
            
            # Row 0: Title
            worksheet.write(0, 0, f'Invoice Details — {month_name}', title_format)
            
            # Row 2: Headers
            headers = ['GSTIN', 'Supplier', 'Invoice No', 'Invoice Date', 'ITC Books', 'ITC 2B', 'Difference', 'Remarks', 'Action Required']
            for col_idx, header in enumerate(headers):
                worksheet.write(2, col_idx, header, header_format)
            
            # Data starting at row 3
            for row_idx, (_, row) in enumerate(month_df.iterrows()):
                worksheet.write(3 + row_idx, 0, str(row.get('GSTIN', '')), text_format)
                worksheet.write(3 + row_idx, 1, str(row.get('Supplier', '')), text_format)
                worksheet.write(3 + row_idx, 2, str(row.get('Invoice No', '')), text_format)
                
                if pd.notna(row.get('Invoice Date')):
                    try:
                        dt = pd.to_datetime(row['Invoice Date'])
                        worksheet.write_datetime(3 + row_idx, 3, dt.to_pydatetime(), date_format)
                    except:
                        worksheet.write(3 + row_idx, 3, str(row['Invoice Date']), text_format)
                else:
                    worksheet.write(3 + row_idx, 3, "", text_format)
                
                worksheet.write(3 + row_idx, 4, row.get('ITC Books', 0), money_format)
                worksheet.write(3 + row_idx, 5, row.get('ITC 2B', 0), money_format)
                worksheet.write(3 + row_idx, 6, row.get('Difference', 0), money_format)
                worksheet.write(3 + row_idx, 7, str(row.get('Remarks', '')), text_format)
                worksheet.write(3 + row_idx, 8, str(row.get('Action Required', '')), text_format)
            
            # Set column widths
            worksheet.set_column(0, 0, 18)
            worksheet.set_column(1, 1, 35)
            worksheet.set_column(2, 2, 20)
            worksheet.set_column(3, 3, 12)
            worksheet.set_column(4, 4, 12)
            worksheet.set_column(5, 5, 12)
            worksheet.set_column(6, 6, 12)
            worksheet.set_column(7, 7, 15)
            worksheet.set_column(8, 8, 20)
            
            # Freeze pane at A3
            autofit_worksheet(worksheet)
            worksheet.freeze_panes(3, 0)


def export_to_excel_old_workbook_exact_match(r, detail_df, month_summary, trade_name_map, books_raw, gstr_raw, tolerance):
    """Export Excel EXACTLY matching old workbook structure"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        
        # 1. Summary sheet
        create_summary_sheet_exact_match(writer, r)
        
        # 2. Month-wise Summary sheet
        if not month_summary.empty:
            create_monthwise_summary_exact_match(writer, month_summary)
        
        # 3. Books sheet
        if not books_raw.empty:
            create_books_sheet_exact_match(writer, books_raw)
        
        # 4. GSTR-2B sheet
        if not gstr_raw.empty:
            create_gstr_sheet_exact_match(writer, gstr_raw)
        
        # 5. Missing in 2B sheet
        create_missing_in_2b_sheet_exact_match(writer, r, trade_name_map)
        
        # 6. Missing in Books sheet
        create_missing_in_books_sheet_exact_match(writer, r, trade_name_map)
        
        # 7. Supplier Wise ITC Summary
        create_supplier_wise_itc_summary_exact_match(writer, r, trade_name_map)
        
        # 8. Supplier Drill Down
        if not detail_df.empty:
            create_supplier_drill_down_exact_match(writer, detail_df, r, trade_name_map)
        
        # 9. Monthly detail sheets (Mar-2026, Apr-2026, etc.)
        create_monthly_detail_sheets_exact_match(writer, r, trade_name_map, tolerance, detail_df)
        
        # 10. NO ITC sheet
        create_no_itc_sheet_exact_match(writer, r)
        
        # 11. Data Issues (if any)
        all_issues_local = r["issues"].copy() if not r["issues"].empty else pd.DataFrame()
        if "duplicate_issues" in r and not r["duplicate_issues"].empty:
            all_issues_local = pd.concat([x for x in [all_issues_local, r["duplicate_issues"]] if not x.empty], ignore_index=True)
        
        if not all_issues_local.empty:
            workbook = writer.book
            worksheet = workbook.add_worksheet('Data Issues')
            
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'font_name': 'Aptos Narrow',
                'align': 'left'
            })
            
            header_format = workbook.add_format({
                'bold': True,
                'font_color': '#FFFFFF',
                'bg_color': '#1F4E78',
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Aptos Narrow'
            })
            
            text_format = workbook.add_format({
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Aptos Narrow'
            })

            issue_date_format = workbook.add_format({
                'num_format': 'dd-mmm-yyyy',
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Aptos Narrow'
            })
            
            worksheet.set_default_row(18)
            worksheet.write(0, 0, 'Data Quality Issues', title_format)
            
            issues_export = all_issues_local.copy()
            if "Invoice_Date" in issues_export.columns:
                issues_export["Invoice_Date"] = pd.to_datetime(issues_export["Invoice_Date"], errors='coerce')
            
            headers = list(issues_export.columns)
            for col_idx, header in enumerate(headers):
                worksheet.write(2, col_idx, header, header_format)
            
            for row_idx, (_, row) in enumerate(issues_export.iterrows()):
                for col_idx, header in enumerate(headers):
                    val = row[header]
                    if header == "Invoice_Date" and pd.notna(val):
                        try:
                            worksheet.write_datetime(3 + row_idx, col_idx, pd.to_datetime(val).to_pydatetime(), issue_date_format)
                        except Exception:
                            worksheet.write(3 + row_idx, col_idx, fmt_date(val), text_format)
                    else:
                        worksheet.write(3 + row_idx, col_idx, str(val) if pd.notna(val) else "", text_format)
            
            for col_idx, header in enumerate(headers):
                max_len = len(header)
                if not issues_export.empty:
                    for row_idx in range(min(100, len(issues_export))):
                        if header == "Invoice_Date":
                            val = fmt_date(issues_export.iloc[row_idx][header])
                        else:
                            val = str(issues_export.iloc[row_idx][header]) if pd.notna(issues_export.iloc[row_idx][header]) else ""
                        max_len = max(max_len, min(len(val), 25))
                worksheet.set_column(col_idx, col_idx, max_len + 2)
            
            autofit_worksheet(worksheet)
            worksheet.freeze_panes(3, 0)
    
    output.seek(0)
    return output.getvalue()


# =================== DOWNLOAD BUTTONS =================== #

st.markdown("<hr>", unsafe_allow_html=True)
col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    excel_data = export_to_excel_old_workbook_exact_match(
        r, detail_df, month_summary, trade_name_map,
        r.get("books_raw", pd.DataFrame()),
        r.get("gstr_raw", pd.DataFrame()),
        tolerance
    )
    st.download_button(
        "📥 Download Full Report (Excel)",
        data=excel_data,
        file_name="reconciliation_20260519_180337.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
