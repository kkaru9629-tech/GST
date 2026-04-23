"""
GST Reconciliation Engine
Strict key-based matching with tax difference only
Version: 6.0 (Tally Purchase Register + GSTR-2B Excel Native Parsers)
"""

import pandas as pd
import re
import logging
from typing import Tuple, Dict, Any, List, Optional
import numpy as np

# =================== CONFIGURATION =================== #

GSTIN_PATTERN = r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}Z[0-9A-Z]{1}$'
logger = logging.getLogger(__name__)

# =================== DATA CLEANING FUNCTIONS =================== #

def strict_numeric_cleaner(val: Any) -> float:
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        if np.isnan(val) or np.isinf(val):
            return 0.0
        return float(val)
    str_val = str(val).strip().lower()
    if str_val in ["", "-", "nan", "null", "none", "nil", "na"]:
        return 0.0
    str_val = str_val.replace(',', '').replace(' ', '')
    try:
        float_val = float(str_val)
        if np.isnan(float_val) or np.isinf(float_val):
            return 0.0
        return float_val
    except (ValueError, TypeError):
        return 0.0

def pre_processing_cleaner(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    financial_cols = ["Taxable_Value", "CGST", "SGST", "IGST", "CESS", "TOTAL_TAX", "Invoice_Value"]
    for col in df.columns:
        if col in financial_cols:
            df[col] = df[col].apply(strict_numeric_cleaner)
        else:
            if df[col].dtype == 'object':
                df[col] = df[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
    return df

def post_processing_cleaner(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    for col in numeric_cols:
        df[col] = df[col].apply(strict_numeric_cleaner)
    return df

# =================== NORMALIZATION FUNCTIONS =================== #

def normalize_invoice_number(inv: str) -> str:
    if pd.isna(inv) or inv == "":
        return ""
    str_inv = str(inv).strip().upper()
    for char in ["/", "-", "_", ".", " ", "\\", "|", ","]:
        str_inv = str_inv.replace(char, "")
    return str_inv

def extract_numeric_core(inv: str) -> str:
    if pd.isna(inv) or inv == "":
        return ""
    str_inv = str(inv).strip()
    numeric_sequences = re.findall(r'\d+', str_inv)
    if not numeric_sequences:
        return ""
    return max(numeric_sequences, key=len)

# =================== CLEANERS =================== #

def clean_string(val: Any) -> str:
    if pd.isna(val):
        return ""
    return str(val).strip().upper()

def clean_invoice(inv: Any) -> str:
    if pd.isna(inv):
        return ""
    return str(inv).strip()

def clean_amount(val: Any) -> float:
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    cleaned = str(val).replace(',', '').replace(' ', '').strip()
    try:
        return float(cleaned)
    except:
        return 0.0

def validate_gstin(gstin: Any) -> bool:
    if pd.isna(gstin):
        return False
    return bool(re.match(GSTIN_PATTERN, str(gstin).strip().upper()))

def calculate_invoice_value(row: pd.Series) -> float:
    return (row.get('Taxable_Value', 0) +
            row.get('CGST', 0) +
            row.get('SGST', 0) +
            row.get('IGST', 0) +
            row.get('CESS', 0))

# =================== FORMAT DETECTION =================== #

def _safe_str_list(row_series) -> list:
    """Convert a pandas Series row to a safe list of lowercase strings. 
    Handles floats, NaN, dates — anything that isn't already a string."""
    result = []
    for v in row_series.tolist():
        try:
            if v is None or (isinstance(v, float) and np.isnan(v)):
                result.append("")
            else:
                result.append(str(v).strip().lower())
        except Exception:
            result.append("")
    return result

def detect_file_format(df: pd.DataFrame, filename: str = "") -> str:
    """
    Detect file format from raw DataFrame.
    Returns: 'tally_pr' | 'gstr2b_excel' | 'standard' | 'unknown'
    """
    try:
        # Scan first 15 rows as flat text — safe against float/NaN cells
        preview_text = ""
        for i in range(min(15, len(df))):
            row_vals = _safe_str_list(df.iloc[i])
            preview_text += " ".join(row_vals) + " "

        # ── Check for Tally Purchase Register ──
        is_tally = False
        if "purchase register" in preview_text:
            is_tally = True
        for i in range(min(10, len(df))):
            row_vals = _safe_str_list(df.iloc[i])
            has_particulars  = any("particulars" in v for v in row_vals)
            has_supplier_inv = any("supplier invoice" in v for v in row_vals)
            if has_particulars and has_supplier_inv:
                is_tally = True
                break
        if is_tally:
            logger.info("Format detected: Tally Purchase Register")
            return "tally_pr"

        # ── Check for GSTR-2B Excel download ──
        is_gstr2b = False
        if "gstr-2b" in preview_text or "gstr2b" in preview_text:
            is_gstr2b = True
        for i in range(min(10, len(df))):
            row_vals = _safe_str_list(df.iloc[i])
            if any("gstin of supplier" in v for v in row_vals):
                is_gstr2b = True
                break
        if is_gstr2b:
            logger.info("Format detected: GSTR-2B Excel")
            return "gstr2b_excel"

        # ── Check for standard template ──
        cols_lower = [str(c).strip().lower() for c in df.columns]
        required = ['gstin', 'trade_name', 'invoice_no', 'taxable_value', 'cgst', 'sgst', 'igst']
        if all(r in cols_lower for r in required):
            logger.info("Format detected: Standard Template")
            return "standard"

        logger.warning(f"Unknown format for file: {filename}")
        return "unknown"

    except Exception as e:
        logger.error(f"Format detection failed: {e}")
        return "unknown"

# =================== TALLY PURCHASE REGISTER PARSER =================== #

def _find_header_row(df: pd.DataFrame) -> int:
    """Find the row index where actual column headers are (has 'Particulars' and 'Date')."""
    for i in range(min(12, len(df))):
        row_vals = _safe_str_list(df.iloc[i])
        has_particulars = any("particulars" in v for v in row_vals)
        has_date        = any(v == "date" for v in row_vals)
        if has_particulars and has_date:
            return i
    return 6  # Default fallback: row index 6 (7th row)

def _get_col_idx(columns: list, *search_terms) -> Optional[int]:
    """
    Find column index by partial match (case-insensitive).
    If multiple search_terms given, finds col that contains ALL terms (AND logic).
    Falls back to finding any single term if AND match fails.
    """
    cols_lower = [str(c).strip().lower() for c in columns]
    if len(search_terms) > 1:
        # AND match: column must contain ALL search terms
        for i, col in enumerate(cols_lower):
            if all(t.strip().lower() in col for t in search_terms):
                return i
    # Single term OR fallback
    for term in search_terms:
        t = term.strip().lower()
        for i, col in enumerate(cols_lower):
            if t in col:
                return i
    return None

def parse_tally_purchase_register(raw_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Parse Tally Prime Purchase Register export (XLS/XLSX).

    File structure:
      - Rows 0..N-1 : metadata (company name, address, period, report title)
      - Row H       : actual column headers (contains 'Particulars' and 'Date')
      - Row H+1..   : data rows

    Key calculations:
      TDS columns (add back to Gross Total):
        - TDS on Profession - 194J
        - TDS on Rent 194I

      Invoice_Value  = Gross Total + TDS_Profession + TDS_Rent
      CGST           = Input CGST FY... (fallback: ITC Not Reflecting CGST)
      SGST           = Input SGST FY... (fallback: ITC Not Reflecting SGST)
      IGST           = Input IGST FY...
      Taxable_Value  = Invoice_Value - CGST - SGST - IGST
      TOTAL_TAX      = CGST + SGST + IGST + CESS (CESS = 0)

    Returns: (valid_df, no_itc_df, issues_df)
    """
    logger.info("Parsing Tally Purchase Register...")

    # ── Find header row ──
    header_row_idx = _find_header_row(raw_df)
    logger.info(f"Tally header row found at index: {header_row_idx}")

    # Rebuild DataFrame with correct headers
    headers = [("" if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v).strip()) for v in raw_df.iloc[header_row_idx].tolist()]
    data_rows = raw_df.iloc[header_row_idx + 1:].reset_index(drop=True)
    data_rows.columns = headers

    # Drop fully empty rows
    data_rows = data_rows.dropna(how='all').reset_index(drop=True)

    # ── Map column indices ──
    cols = list(data_rows.columns)

    idx_date       = _get_col_idx(cols, "date")
    idx_name       = _get_col_idx(cols, "particulars")
    idx_inv_no     = _get_col_idx(cols, "supplier invoice no", "supplier invoice")
    idx_gstin      = _get_col_idx(cols, "gstin/uin", "gstin")
    idx_gross      = _get_col_idx(cols, "gross total")

    # Tax columns — primary
    idx_cgst_main  = _get_col_idx(cols, "input cgst")
    idx_sgst_main  = _get_col_idx(cols, "input sgst")
    idx_igst_main  = _get_col_idx(cols, "input igst")

    # Tax columns — fallback (ITC Not Reflecting)
    idx_cgst_fb    = _get_col_idx(cols, "itc not reflecting", "cgst")
    idx_sgst_fb    = _get_col_idx(cols, "itc not reflecting", "sgst")

    # TDS columns
    idx_tds_prof   = _get_col_idx(cols, "tds on profession", "194j")
    idx_tds_rent   = _get_col_idx(cols, "tds on rent", "194i")

    logger.info(f"Column map → Date:{idx_date}, Name:{idx_name}, InvNo:{idx_inv_no}, "
                f"GSTIN:{idx_gstin}, Gross:{idx_gross}, CGST:{idx_cgst_main}, "
                f"SGST:{idx_sgst_main}, IGST:{idx_igst_main}, "
                f"TDS_Prof:{idx_tds_prof}, TDS_Rent:{idx_tds_rent}")

    def _get_val(row, idx):
        if idx is None:
            return 0.0
        try:
            return clean_amount(row.iloc[idx])
        except Exception:
            return 0.0

    def _get_str(row, idx):
        if idx is None:
            return ""
        try:
            v = row.iloc[idx]
            if v is None:
                return ""
            if isinstance(v, float) and np.isnan(v):
                return ""
            return str(v).strip()
        except Exception:
            return ""

    records = []

    for _, row in data_rows.iterrows():
        # Skip summary/totals rows (first cell often says "Total" or is numeric)
        first_cell = str(row.iloc[0]).strip().lower() if len(row) > 0 else ""
        if first_cell in ["total", "grand total", "totals", ""]:
            continue
        # Skip if no supplier name
        trade_name = _get_str(row, idx_name)
        if not trade_name or trade_name.lower() in ["nan", "none", "total", "grand total", ""]:
            continue

        # ── Raw financial values ──
        gross_total = _get_val(row, idx_gross)
        tds_prof    = _get_val(row, idx_tds_prof)
        tds_rent    = _get_val(row, idx_tds_rent)

        # CGST: primary column first, then fallback
        cgst_main = _get_val(row, idx_cgst_main)
        cgst_fb   = _get_val(row, idx_cgst_fb) if idx_cgst_fb != idx_cgst_main else 0.0
        cgst      = cgst_main if cgst_main != 0.0 else cgst_fb

        # SGST: same logic
        sgst_main = _get_val(row, idx_sgst_main)
        sgst_fb   = _get_val(row, idx_sgst_fb) if idx_sgst_fb != idx_sgst_main else 0.0
        sgst      = sgst_main if sgst_main != 0.0 else sgst_fb

        igst      = _get_val(row, idx_igst_main)
        cess      = 0.0

        # ── TDS add-back ──
        # Tally records Gross Total AFTER TDS deduction.
        # We add TDS back to get the actual invoice value as the supplier raised it.
        total_tds   = abs(tds_prof) + abs(tds_rent)
        invoice_val = gross_total + total_tds

        # ── Taxable value derived from invoice value ──
        taxable_val = invoice_val - cgst - sgst - igst
        if taxable_val < 0:
            taxable_val = 0.0  # Guard against data anomalies

        total_tax = cgst + sgst + igst + cess

        records.append({
            "GSTIN":         clean_string(_get_str(row, idx_gstin)),
            "Trade_Name":    trade_name,
            "Invoice_No":    clean_invoice(_get_str(row, idx_inv_no)),
            "Invoice_Date":  _get_str(row, idx_date),
            "Taxable_Value": round(taxable_val, 2),
            "CGST":          round(cgst, 2),
            "SGST":          round(sgst, 2),
            "IGST":          round(igst, 2),
            "CESS":          round(cess, 2),
            "TOTAL_TAX":     round(total_tax, 2),
            "Invoice_Value": round(invoice_val, 2),
        })

    if not records:
        raise Exception("Tally Purchase Register: No data rows could be extracted. "
                        "Check that the file has 'Particulars' and 'Date' column headers.")

    df = pd.DataFrame(records)

    # Parse dates
    df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)

    logger.info(f"Tally PR raw records extracted: {len(df)}")

    # ── Now delegate to standard validation pipeline ──
    return _validate_books_df(df)


# =================== GSTR-2B EXCEL PARSER =================== #

def _find_gstr2b_data_start(raw_df: pd.DataFrame) -> int:
    """
    Find the first row index that contains actual invoice data.
    GSTR-2B has 2 merged header rows (rows 6-7 in 1-based = indices 5-6).
    Data starts at index 7 typically, but we detect it by looking for a valid GSTIN.
    """
    gstin_re = re.compile(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}Z[0-9A-Z]{1}$')
    for i in range(min(15, len(raw_df))):
        for val in raw_df.iloc[i].tolist():
            if val and isinstance(val, str) and gstin_re.match(val.strip().upper()):
                logger.info(f"GSTR-2B data starts at row index: {i}")
                return i
    return 7  # fallback

def parse_gstr2b_excel(raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Parse official GSTR-2B Excel download from GST Portal.

    File structure:
      - Rows 0-4 : metadata (GSTIN of taxpayer, period, etc.)
      - Rows 5-6 : merged column headers (2 rows)
      - Row 7+   : data (GSTIN | Trade Name | ... | Invoice No | Date | Value | ...)

    Column positions (0-indexed, from observed GSTR-2B format):
      0  : Sr. No.
      1  : GSTIN of supplier
      2  : Trade/Legal name
      3  : ITC Availability (Yes/No)
      4  : Invoice number
      5  : Invoice type
      6  : Invoice date
      7  : Invoice value
      8  : Place of supply
      9  : Taxable value
      ...
      14 : Integrated Tax (IGST)
      15 : Central Tax (CGST)
      16 : State/UT Tax (SGST)
      17 : Cess

    We detect column positions dynamically from the header row as a safety net.
    """
    logger.info("Parsing GSTR-2B Excel download...")

    data_start = _find_gstr2b_data_start(raw_df)

    # Try to find the header row just before data_start
    header_candidates = []
    for i in range(max(0, data_start - 3), data_start):
        row_text = " ".join(_safe_str_list(raw_df.iloc[i]))
        if "gstin" in row_text or "invoice" in row_text or "taxable" in row_text:
            header_candidates.append(i)

    # Use the last header candidate row to build column name map
    col_map = {}  # field_name -> col_index (position-based fallback below)

    if header_candidates:
        hdr_row = raw_df.iloc[header_candidates[-1]].tolist()
        for ci, val in enumerate(hdr_row):
            v = ("" if (val is None or (isinstance(val, float) and np.isnan(val))) else str(val).strip().lower())
            if "gstin of supplier" in v or ("gstin" in v and "supplier" in v):
                col_map["gstin"] = ci
            elif "trade" in v and ("name" in v or "legal" in v):
                col_map["trade_name"] = ci
            elif "invoice number" in v or ("invoice" in v and "number" in v):
                col_map["invoice_no"] = ci
            elif "invoice date" in v or ("invoice" in v and "date" in v):
                col_map["invoice_date"] = ci
            elif "invoice value" in v:
                col_map["invoice_value"] = ci
            elif "taxable value" in v:
                col_map["taxable_value"] = ci
            elif "integrated tax" in v or "igst" in v:
                col_map["igst"] = ci
            elif "central tax" in v or "cgst" in v:
                col_map["cgst"] = ci
            elif "state" in v and "tax" in v or "sgst" in v:
                col_map["sgst"] = ci
            elif "cess" in v:
                col_map["cess"] = ci

    # Position-based fallback (standard GSTR-2B column order)
    FALLBACK_POS = {
        "gstin":          1,
        "trade_name":     2,
        "invoice_no":     4,
        "invoice_date":   6,
        "invoice_value":  7,
        "taxable_value":  9,
        "igst":          14,
        "cgst":          15,
        "sgst":          16,
        "cess":          17,
    }
    for field, pos in FALLBACK_POS.items():
        if field not in col_map:
            col_map[field] = pos

    logger.info(f"GSTR-2B column map: {col_map}")

    records = []
    data_section = raw_df.iloc[data_start:].reset_index(drop=True)

    gstin_re = re.compile(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}Z[0-9A-Z]{1}$')

    for _, row in data_section.iterrows():
        row_list = row.tolist()
        max_col = len(row_list) - 1

        def safe_get(field):
            idx = col_map.get(field)
            if idx is None or idx > max_col:
                return None
            return row_list[idx]

        gstin_raw = safe_get("gstin")
        if gstin_raw is None or pd.isna(gstin_raw):
            continue
        gstin_str = str(gstin_raw).strip().upper()
        if not gstin_re.match(gstin_str):
            continue  # Skip non-data rows (totals, subtitles, blanks)

        trade_name    = str(safe_get("trade_name") or "").strip()
        invoice_no    = str(safe_get("invoice_no") or "").strip()
        invoice_date  = safe_get("invoice_date")
        invoice_val   = clean_amount(safe_get("invoice_value"))
        taxable_val   = clean_amount(safe_get("taxable_value"))
        igst          = clean_amount(safe_get("igst"))
        cgst          = clean_amount(safe_get("cgst"))
        sgst          = clean_amount(safe_get("sgst"))
        cess          = clean_amount(safe_get("cess"))
        total_tax     = cgst + sgst + igst + cess

        records.append({
            "GSTIN":         gstin_str,
            "Trade_Name":    trade_name,
            "Invoice_No":    invoice_no,
            "Invoice_Date":  invoice_date,
            "Taxable_Value": round(taxable_val, 2),
            "CGST":          round(cgst, 2),
            "SGST":          round(sgst, 2),
            "IGST":          round(igst, 2),
            "CESS":          round(cess, 2),
            "TOTAL_TAX":     round(total_tax, 2),
            "Invoice_Value": round(invoice_val, 2),
        })

    if not records:
        raise Exception("GSTR-2B Excel: No valid invoice rows found. "
                        "Ensure you are uploading the correct GSTR-2B Excel file from the GST Portal.")

    df = pd.DataFrame(records)
    df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)
    df = df[df["Invoice_Date"].notna()].reset_index(drop=True)

    logger.info(f"GSTR-2B Excel parsed: {len(df)} records")
    return df


# =================== SHARED VALIDATION PIPELINE =================== #

def _validate_books_df(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Shared validation and splitting used by both parse_tally and parse_tally_purchase_register.
    Returns: (valid_df, no_itc_df, issues_df)
    """
    # ── Financial duplicate detection ──
    df_temp = df.copy()
    df_temp["GSTIN"]        = df_temp["GSTIN"].fillna("__NULL__").replace("", "__NULL__")
    df_temp["Invoice_No"]   = df_temp["Invoice_No"].fillna("__NULL__").replace("", "__NULL__")
    df_temp["Invoice_Date"] = df_temp["Invoice_Date"].fillna(pd.Timestamp("1900-01-01"))

    dup_cols = ["GSTIN", "Invoice_No", "Invoice_Date", "Taxable_Value", "CGST", "SGST", "IGST", "CESS"]
    fin_dup_mask = df_temp.duplicated(subset=dup_cols, keep=False)
    fin_dup_rows = df[fin_dup_mask].copy()

    fin_dup_issues = []
    for _, row in fin_dup_rows.iterrows():
        fin_dup_issues.append({
            "GSTIN":        row["GSTIN"],
            "Trade_Name":   row["Trade_Name"],
            "Invoice_No":   row["Invoice_No"],
            "Invoice_Date": row["Invoice_Date"],
            "Taxable_Value":row["Taxable_Value"],
            "Invoice_Value":row["Invoice_Value"],
            "TOTAL_TAX":    row["TOTAL_TAX"],
            "Issue":        "Duplicate Financial Row",
            "Source":       "books_financial",
        })

    df = df[~fin_dup_mask].copy()

    # ── Validation ──
    issues    = []
    valid_idx = []

    for idx, row in df.iterrows():
        issue = []

        gstin = row["GSTIN"]
        if pd.isna(gstin) or gstin == "":
            issue.append("No GSTIN")
        elif not validate_gstin(gstin):
            issue.append("Invalid GSTIN")

        inv_no = row["Invoice_No"]
        if pd.isna(inv_no) or str(inv_no).strip() == "" or str(inv_no).upper() in ["NAN", "NULL", "NONE"]:
            issue.append("No Invoice No")

        if pd.isna(row["Invoice_Date"]):
            issue.append("No Invoice Date")

        if issue:
            issues.append({
                "GSTIN":        row["GSTIN"],
                "Trade_Name":   row["Trade_Name"],
                "Invoice_No":   row["Invoice_No"],
                "Invoice_Date": row["Invoice_Date"],
                "Taxable_Value":row["Taxable_Value"],
                "Invoice_Value":row["Invoice_Value"],
                "TOTAL_TAX":    row["TOTAL_TAX"],
                "Issue":        ", ".join(issue),
                "Source":       "books_validation",
            })
        else:
            valid_idx.append(idx)

    all_issues   = fin_dup_issues + issues
    issues_df    = pd.DataFrame(all_issues) if all_issues else pd.DataFrame()
    valid_df     = df.loc[valid_idx].copy() if valid_idx else pd.DataFrame()

    if not valid_df.empty:
        no_itc_df = valid_df[valid_df["TOTAL_TAX"] == 0].copy()
        valid_df  = valid_df[valid_df["TOTAL_TAX"] != 0].copy()
    else:
        no_itc_df = pd.DataFrame()

    logger.info(f"Validation → Valid: {len(valid_df)}, Zero ITC: {len(no_itc_df)}, Issues: {len(issues_df)}")
    return valid_df, no_itc_df, issues_df


# =================== PARSE TALLY (STANDARD TEMPLATE) =================== #

def parse_tally(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Parse standard template Books data (9-column format)."""
    df = pre_processing_cleaner(df)
    df = df.copy()
    logger.info(f"Parsing standard Books template... Total rows: {len(df)}")

    required_cols = ["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                     "Taxable_Value", "CGST", "SGST", "IGST", "CESS"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise Exception(f"Missing columns: {missing}")

    df["GSTIN"]        = df["GSTIN"].apply(clean_string)
    df["Invoice_No"]   = df["Invoice_No"].apply(clean_invoice)
    df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)

    for col in ["Taxable_Value", "CGST", "SGST", "IGST", "CESS"]:
        df[col] = df[col].apply(clean_amount)

    df["TOTAL_TAX"]    = df["CGST"] + df["SGST"] + df["IGST"] + df["CESS"]
    df["Invoice_Value"] = df.apply(calculate_invoice_value, axis=1)

    return _validate_books_df(df)


# =================== PARSE GSTR-2B (STANDARD TEMPLATE) =================== #

def parse_gstr2b(df: pd.DataFrame) -> pd.DataFrame:
    """Parse standard template GSTR-2B data (9-column format)."""
    df = pre_processing_cleaner(df)
    df = df.copy().dropna(how="all")
    logger.info(f"Parsing standard GSTR-2B template... Total rows: {len(df)}")

    required_cols = ["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date",
                     "Taxable_Value", "CGST", "SGST", "IGST", "CESS"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise Exception(f"Missing columns: {missing}")

    df["GSTIN"]        = df["GSTIN"].apply(clean_string)
    df["Invoice_No"]   = df["Invoice_No"].apply(clean_invoice)
    df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)

    for col in ["Taxable_Value", "CGST", "SGST", "IGST", "CESS"]:
        df[col] = df[col].apply(clean_amount)

    df["TOTAL_TAX"]    = df["CGST"] + df["SGST"] + df["IGST"] + df["CESS"]
    df["Invoice_Value"] = df.apply(calculate_invoice_value, axis=1)

    df = df[df["Invoice_Date"].notna()]
    logger.info(f"GSTR-2B standard parsed: {len(df)} records")
    return df


# =================== GROUP INVOICES =================== #

def group_invoices(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    logger.info(f"Grouping {len(df)} rows by invoice...")
    sum_cols = [c for c in ["Taxable_Value", "CGST", "SGST", "IGST", "CESS", "TOTAL_TAX", "Invoice_Value"]
                if c in df.columns]
    grouped = df.groupby(["GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date"],
                         as_index=False, dropna=False)[sum_cols].sum()
    logger.info(f"Grouped into {len(grouped)} unique invoices")
    return grouped


# =================== DETECT DUPLICATES =================== #

def detect_duplicate_invoices(df: pd.DataFrame, source: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if df.empty:
        return df, pd.DataFrame(), pd.DataFrame()

    df_temp = df.copy()
    df_temp["GSTIN"]      = df_temp["GSTIN"].fillna("__NULL__").replace("", "__NULL__")
    df_temp["Invoice_No"] = df_temp["Invoice_No"].fillna("__NULL__").replace("", "__NULL__")

    duplicate_mask = df_temp.duplicated(subset=["GSTIN", "Invoice_No"], keep=False)
    duplicate_rows = df[duplicate_mask].copy()

    duplicate_issues = []
    if not duplicate_rows.empty:
        issue_text = "Duplicate Invoice in Books" if source == "books" else "Duplicate Invoice in GSTR-2B"
        for _, row in duplicate_rows.iterrows():
            duplicate_issues.append({
                "GSTIN":        row["GSTIN"],
                "Trade_Name":   row["Trade_Name"],
                "Invoice_No":   row["Invoice_No"],
                "Invoice_Date": row["Invoice_Date"],
                "Taxable_Value":row["Taxable_Value"],
                "Invoice_Value":row["Invoice_Value"],
                "TOTAL_TAX":    row["TOTAL_TAX"],
                "Issue":        issue_text,
                "Source":       source,
            })

    deduplicated_df = df[~duplicate_mask].copy()
    if not duplicate_rows.empty:
        unique_keys = duplicate_rows.drop_duplicates(subset=["GSTIN", "Invoice_No"], keep='first')
        deduplicated_df = pd.concat([deduplicated_df, unique_keys])

    duplicate_issues_df = pd.DataFrame(duplicate_issues) if duplicate_issues else pd.DataFrame()
    logger.info(f"{source}: {len(duplicate_issues)} duplicate rows, keeping {len(deduplicated_df)} unique invoices")
    return deduplicated_df, duplicate_rows, duplicate_issues_df


# =================== TRADE NAME MAPPING =================== #

def create_trade_name_mapping(gstr_df: pd.DataFrame, books_df: pd.DataFrame) -> dict:
    mapping = {}
    if not gstr_df.empty:
        for _, row in gstr_df.iterrows():
            if pd.notna(row["GSTIN"]) and row["GSTIN"] != "":
                mapping[row["GSTIN"]] = row["Trade_Name"]
    if not books_df.empty:
        for _, row in books_df.iterrows():
            if pd.notna(row["GSTIN"]) and row["GSTIN"] != "":
                if row["GSTIN"] not in mapping:
                    mapping[row["GSTIN"]] = row["Trade_Name"]
    return mapping


# =================== 3-LEVEL MATCHING =================== #

def level1_strict_match(gstr_df, books_df):
    gstr_df  = gstr_df.copy()
    books_df = books_df.copy()
    gstr_df["KEY"]  = (gstr_df["GSTIN"].fillna("").astype(str)  + "|" +
                        gstr_df["Invoice_No"].fillna("").astype(str)  + "|" +
                        gstr_df["Invoice_Date"].fillna("").astype(str))
    books_df["KEY"] = (books_df["GSTIN"].fillna("").astype(str) + "|" +
                        books_df["Invoice_No"].fillna("").astype(str) + "|" +
                        books_df["Invoice_Date"].fillna("").astype(str))
    matched_keys  = set(gstr_df["KEY"]).intersection(set(books_df["KEY"]))
    gstr_matched  = gstr_df[gstr_df["KEY"].isin(matched_keys)].copy()
    books_matched = books_df[books_df["KEY"].isin(matched_keys)].copy()
    matched       = pd.merge(gstr_matched, books_matched, on="KEY", suffixes=("_2B", "_Books"), how="inner")
    gstr_unmatched  = gstr_df[~gstr_df["KEY"].isin(matched_keys)].copy()
    books_unmatched = books_df[~books_df["KEY"].isin(matched_keys)].copy()
    logger.info(f"Level 1 Strict Match: {len(matched)} invoices matched")
    return matched, gstr_unmatched, books_unmatched, matched_keys


def level2_normalized_match(gstr_unmatched, books_unmatched, tolerance):
    if gstr_unmatched.empty or books_unmatched.empty:
        return pd.DataFrame(), gstr_unmatched, books_unmatched, set()
    gstr_df  = gstr_unmatched.copy()
    books_df = books_unmatched.copy()
    gstr_df["NORM_KEY"]  = [
        str(r['GSTIN'] or '') + "|" +
        normalize_invoice_number(str(r['Invoice_No'] or '')) + "|" +
        str(r['Invoice_Date'])
        for _, r in gstr_df.iterrows()
    ]
    books_df["NORM_KEY"] = [
        str(r['GSTIN'] or '') + "|" +
        normalize_invoice_number(str(r['Invoice_No'] or '')) + "|" +
        str(r['Invoice_Date'])
        for _, r in books_df.iterrows()
    ]
    matched_norm_keys = set(gstr_df["NORM_KEY"]).intersection(set(books_df["NORM_KEY"]))
    if not matched_norm_keys:
        return pd.DataFrame(), gstr_unmatched, books_unmatched, set()
    gstr_m  = gstr_df[gstr_df["NORM_KEY"].isin(matched_norm_keys)].copy()
    books_m = books_df[books_df["NORM_KEY"].isin(matched_norm_keys)].copy()
    matched = pd.merge(gstr_m, books_m, on="NORM_KEY", suffixes=("_2B", "_Books"), how="inner")
    matched = matched.drop(columns=["NORM_KEY"], errors="ignore")
    gstr_remaining  = gstr_df[~gstr_df["NORM_KEY"].isin(matched_norm_keys)].drop(columns=["NORM_KEY"], errors="ignore")
    books_remaining = books_df[~books_df["NORM_KEY"].isin(matched_norm_keys)].drop(columns=["NORM_KEY"], errors="ignore")
    logger.info(f"Level 2 Normalized Match: {len(matched)} invoices matched")
    matched_original_keys = set()
    for _, row in matched.iterrows():
        gstin    = row.get("GSTIN_2B", row.get("GSTIN_Books", ""))
        inv_no   = row.get("Invoice_No_2B", row.get("Invoice_No_Books", ""))
        inv_date = row.get("Invoice_Date_2B", row.get("Invoice_Date_Books", ""))
        matched_original_keys.add(f"{gstin}|{inv_no}|{inv_date}")
    return matched, gstr_remaining, books_remaining, matched_original_keys


def level3_numeric_core_match(gstr_unmatched, books_unmatched, tolerance):
    if gstr_unmatched.empty or books_unmatched.empty:
        return pd.DataFrame(), gstr_unmatched, books_unmatched, set()
    gstr_df  = gstr_unmatched.copy()
    books_df = books_unmatched.copy()

    def build_core_map(df):
        m = {}
        for idx, row in df.iterrows():
            core = extract_numeric_core(row["Invoice_No"])
            if core:
                key = f"{row['GSTIN']}|{core}|{row['Invoice_Date']}"
                m.setdefault(key, []).append(idx)
        return m

    gstr_core_map  = build_core_map(gstr_df)
    books_core_map = build_core_map(books_df)
    matched_core_keys = set(gstr_core_map).intersection(books_core_map)
    if not matched_core_keys:
        return pd.DataFrame(), gstr_unmatched, books_unmatched, set()

    matched_rows         = []
    matched_gstr_idxs    = set()
    matched_books_idxs   = set()

    for ck in matched_core_keys:
        for gi in gstr_core_map[ck]:
            gstr_row  = gstr_df.loc[gi]
            for bi in books_core_map[ck]:
                if bi in matched_books_idxs:
                    continue
                books_row = books_df.loc[bi]
                if abs(gstr_row["TOTAL_TAX"] - books_row["TOTAL_TAX"]) <= tolerance:
                    matched_gstr_idxs.add(gi)
                    matched_books_idxs.add(bi)
                    merged = {f"{c}_2B": gstr_row[c] for c in gstr_row.index}
                    merged.update({f"{c}_Books": books_row[c] for c in books_row.index})
                    merged["TAX_DIFF"]  = gstr_row["TOTAL_TAX"] - books_row["TOTAL_TAX"]
                    merged["TAX_MATCH"] = True
                    matched_rows.append(merged)
                    break

    matched         = pd.DataFrame(matched_rows) if matched_rows else pd.DataFrame()
    gstr_remaining  = gstr_df[~gstr_df.index.isin(matched_gstr_idxs)].copy()
    books_remaining = books_df[~books_df.index.isin(matched_books_idxs)].copy()
    logger.info(f"Level 3 Numeric Core Match: {len(matched)} invoices matched")
    matched_original_keys = set()
    for _, row in matched.iterrows():
        gstin    = row.get("GSTIN_2B", row.get("GSTIN_Books", ""))
        inv_no   = row.get("Invoice_No_2B", row.get("Invoice_No_Books", ""))
        inv_date = row.get("Invoice_Date_2B", row.get("Invoice_Date_Books", ""))
        matched_original_keys.add(f"{gstin}|{inv_no}|{inv_date}")
    return matched, gstr_remaining, books_remaining, matched_original_keys


# =================== RECONCILE =================== #

def reconcile(gstr_df: pd.DataFrame, books_df: pd.DataFrame, tolerance: float = 1.0) -> Dict:
    logger.info("Starting reconciliation...")
    gstr_original  = gstr_df.copy()
    books_original = books_df.copy()

    books_dedup, _, books_dup_issues = detect_duplicate_invoices(books_original, "books")
    gstr_dedup,  _, gstr_dup_issues  = detect_duplicate_invoices(gstr_original,  "gstr")

    all_dup_issues = pd.DataFrame()
    for d in [books_dup_issues, gstr_dup_issues]:
        if not d.empty:
            all_dup_issues = pd.concat([all_dup_issues, d], ignore_index=True)

    if not all_dup_issues.empty:
        logger.info(f"Duplicate issues: {len(all_dup_issues)}")

    trade_name_mapping = create_trade_name_mapping(gstr_original, books_original)

    gstr_grouped  = group_invoices(gstr_dedup)
    books_grouped = group_invoices(books_dedup)

    matched_l1, gstr_u1, books_u1, _ = level1_strict_match(gstr_grouped, books_grouped)
    matched_l2, gstr_u2, books_u2, _ = level2_normalized_match(gstr_u1, books_u1, tolerance)
    matched_l3, gstr_u3, books_u3, _ = level3_numeric_core_match(gstr_u2, books_u2, tolerance)

    all_matched = pd.concat([matched_l1, matched_l2, matched_l3], ignore_index=True) \
        if not all(d.empty for d in [matched_l1, matched_l2, matched_l3]) else pd.DataFrame()

    missing_2b_final    = books_u3.copy() if not books_u3.empty else pd.DataFrame()
    missing_books_final = gstr_u3.copy()  if not gstr_u3.empty  else pd.DataFrame()

    if not all_matched.empty:
        all_matched["TAX_DIFF"]  = all_matched["TOTAL_TAX_2B"] - all_matched["TOTAL_TAX_Books"]
        all_matched["TAX_MATCH"] = abs(all_matched["TAX_DIFF"]) <= tolerance
        matched  = all_matched[all_matched["TAX_MATCH"]].copy()
        tax_diff = all_matched[~all_matched["TAX_MATCH"]].copy()
    else:
        matched  = pd.DataFrame()
        tax_diff = pd.DataFrame()

    missing_2b_itc = missing_2b_final["TOTAL_TAX"].sum() if not missing_2b_final.empty else 0
    tax_shortage   = 0
    if not tax_diff.empty:
        shortage_rows = tax_diff[tax_diff["TAX_DIFF"] < 0]
        if not shortage_rows.empty:
            tax_shortage = abs(shortage_rows["TAX_DIFF"].sum())
    itc_at_risk = missing_2b_itc + tax_shortage

    match_pct = round((len(matched) / len(gstr_grouped) * 100), 2) if len(gstr_grouped) > 0 else 0

    summary = {
        "ITC_Books":    round(books_grouped["TOTAL_TAX"].sum(), 2),
        "ITC_GSTR":     round(gstr_grouped["TOTAL_TAX"].sum(), 2),
        "ITC_Diff":     round(gstr_grouped["TOTAL_TAX"].sum() - books_grouped["TOTAL_TAX"].sum(), 2),
        "ITC_at_Risk":  round(itc_at_risk, 2),
        "Match_%":      match_pct,
        "Total_Books":  len(books_grouped),
        "Total_GSTR":   len(gstr_grouped),
        "Matched":      len(matched),
        "Tax_Diff":     len(tax_diff),
        "Missing_2B":   len(missing_2b_final),
        "Missing_Books":len(missing_books_final),
    }

    logger.info(f"Match %: {match_pct}%, ITC at Risk: {itc_at_risk:,.2f}")

    return {
        "matched":           matched,
        "tax_diff":          tax_diff,
        "missing_2b":        missing_2b_final,
        "missing_books":     missing_books_final,
        "summary":           summary,
        "trade_name_mapping":trade_name_mapping,
        "duplicate_issues":  all_dup_issues,
    }
