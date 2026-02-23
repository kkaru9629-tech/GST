import pandas as pd
import re

# ===========================================
# COMMON CLEANERS
# ===========================================

def clean_string(val):
    if pd.isna(val):
        return ""
    return str(val).strip().upper()

def clean_invoice(inv):
    if pd.isna(inv):
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(inv).upper())

def validate_gstin(gstin):
    if pd.isna(gstin):
        return False
    gstin = str(gstin).strip()
    return len(gstin) >= 10

# ===========================================
# INTELLIGENT TAX COLUMN DETECTOR
# ===========================================

def detect_tax_columns(df):
    patterns = {
        'CGST': [
            'input cgst','cgst input','itc cgst','cgst itc',
            'central gst','central tax','cgst','c.g.s.t',
            'cgst@','cgst%','cgst amount','cgst amt'
        ],
        'SGST': [
            'input sgst','sgst input','itc sgst','sgst itc',
            'state gst','state tax','sgst','s.g.s.t',
            'sgst@','sgst%','sgst amount','sgst amt',
            'st/gst','stgst'
        ],
        'IGST': [
            'input igst','igst input','itc igst','igst itc',
            'integrated gst','integrated tax','igst','i.g.s.t',
            'igst@','igst%','igst amount','igst amt'
        ]
    }
    
    tax_cols = {'CGST': [], 'SGST': [], 'IGST': []}
    
    for col in df.columns:
        col_clean = str(col).lower().replace(" ", "").replace("_","").replace("-","")
        col_space = " " + str(col).lower() + " "
        
        for tax_type, pattern_list in patterns.items():
            for pattern in pattern_list:
                p_clean = pattern.replace(" ","").replace("_","")
                p_space = " " + pattern + " "
                
                if (p_clean in col_clean or pattern in col_space or all(word in col_space for word in pattern.split())):
                    tax_cols[tax_type].append(col)
                    break
                    
    return tax_cols

# ===========================================
# PARSE TALLY
# ===========================================

def parse_tally(df):
    df = df.copy()
    
    # Find header row
    header_idx = None
    for i in range(min(len(df), 30)):
        row = " ".join(df.iloc[i].astype(str).str.lower().values)
        if "supplier invoice no" in row and "gstin" in row:
            header_idx = i
            break
    
    if header_idx is None:
        raise Exception("Tally header not detected.")
    
    # Set header and clean data
    df.columns = df.iloc[header_idx]
    df = df.iloc[header_idx + 1:].reset_index(drop=True)
    df = df.dropna(how="all")
    
    # Rename columns
    df.rename(columns={
        "Date": "Invoice_Date",
        "Particulars": "Trade_Name",
        "Supplier Invoice No.": "Invoice_No",
        "GSTIN/UIN": "GSTIN",
        "Gross Total": "Invoice_Value"
    }, inplace=True)
    
    # Validate required columns
    required = ["GSTIN","Invoice_No","Invoice_Date"]
    for col in required:
        if col not in df.columns:
            raise Exception(f"{col} missing in Tally file.")
    
    # Clean data
    df["GSTIN"] = df["GSTIN"].apply(clean_string)
    df["Invoice_No"] = df["Invoice_No"].apply(clean_invoice)
    df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)
    
    # Detect and calculate tax columns
    tax_cols = detect_tax_columns(df)
    
    df["CGST"] = df[tax_cols['CGST']].apply(pd.to_numeric,errors='coerce').sum(axis=1) if tax_cols['CGST'] else 0
    df["SGST"] = df[tax_cols['SGST']].apply(pd.to_numeric,errors='coerce').sum(axis=1) if tax_cols['SGST'] else 0
    df["IGST"] = df[tax_cols['IGST']].apply(pd.to_numeric,errors='coerce').sum(axis=1) if tax_cols['IGST'] else 0
    
    # Detect TDS columns
    tds_cols = [c for c in df.columns if any(x in str(c).lower() for x in ['tds','t.d.s','tax deducted','tax deducted at source'])]
    df["TDS"] = df[tds_cols].apply(pd.to_numeric,errors="coerce").sum(axis=1).fillna(0) if tds_cols else 0
    
    # Calculate values
    df["Invoice_Value"] = pd.to_numeric(df["Invoice_Value"], errors="coerce")
    df["TOTAL_TAX"] = df["CGST"] + df["SGST"] + df["IGST"]
    df["Taxable_Value"] = df["Invoice_Value"] + df["TDS"] - df["TOTAL_TAX"]
    
    # Separate zero tax records
    no_itc_df = df[df["TOTAL_TAX"] == 0].copy()
    
    # Validate GSTIN
    df["GSTIN_VALID"] = df["GSTIN"].apply(validate_gstin)
    invalid_gstin_df = df[~df["GSTIN_VALID"]].copy()
    
    # Filter valid records
    df = df[df["GSTIN_VALID"]]
    df = df.drop_duplicates(subset=["GSTIN","Invoice_No"])
    df = df[df["Invoice_Date"].notna()]
    
    # Select final columns
    valid_df = df[[
        "GSTIN","Trade_Name","Invoice_No","Invoice_Date",
        "Taxable_Value","Invoice_Value","IGST","CGST","SGST","TOTAL_TAX"
    ]]
    
    return valid_df, no_itc_df, invalid_gstin_df

# ===========================================
# PARSE GSTR-2B
# ===========================================

def parse_gstr2b(df):
    df = df.copy().dropna(how="all")
    
    # Rename columns
    df.rename(columns={
        "GSTIN of supplier":"GSTIN",
        "Trade/Legal name":"Trade_Name",
        "Invoice number":"Invoice_No",
        "Invoice Date":"Invoice_Date",
        "Taxable Value (₹)":"Taxable_Value",
        "Integrated Tax(₹)":"IGST",
        "Central Tax(₹)":"CGST",
        "State/UT Tax(₹)":"SGST"
    }, inplace=True)
    
    # Validate required columns
    required = ["GSTIN","Invoice_No","Invoice_Date"]
    for col in required:
        if col not in df.columns:
            raise Exception("Incorrect GSTR-2B structured format.")
    
    # Clean data
    df["GSTIN"] = df["GSTIN"].apply(clean_string)
    df["Invoice_No"] = df["Invoice_No"].apply(clean_invoice)
    df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)
    
    # Convert numeric columns
    for col in ["Taxable_Value","IGST","CGST","SGST"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    
    # Calculate totals
    df["TOTAL_TAX"] = df["IGST"] + df["CGST"] + df["SGST"]
    df["Invoice_Value"] = df["Taxable_Value"] + df["TOTAL_TAX"]
    
    # Remove duplicates and null dates
    df = df.drop_duplicates(subset=["GSTIN","Invoice_No"])
    df = df[df["Invoice_Date"].notna()]
    
    return df[[
        "GSTIN","Trade_Name","Invoice_No","Invoice_Date",
        "Taxable_Value","Invoice_Value","IGST","CGST","SGST","TOTAL_TAX"
    ]]

# ===========================================
# RECONCILE
# ===========================================

def reconcile(gstr2b_df, tally_df):
    # Create keys for matching
    gstr2b_df = gstr2b_df.copy()
    tally_df = tally_df.copy()
    
    gstr2b_df["KEY"] = gstr2b_df["GSTIN"] + "|" + gstr2b_df["Invoice_No"]
    tally_df["KEY"] = tally_df["GSTIN"] + "|" + tally_df["Invoice_No"]
    
    # Find missing records
    missing_books = gstr2b_df[~gstr2b_df["KEY"].isin(tally_df["KEY"])].copy()
    missing_2b = tally_df[~tally_df["KEY"].isin(gstr2b_df["KEY"])].copy()
    
    # Calculate missing ITC values
    missing_books_value = missing_books["TOTAL_TAX"].sum() if not missing_books.empty else 0
    missing_2b_value = missing_2b["TOTAL_TAX"].sum() if not missing_2b.empty else 0
    
    # Merge matched records
    merged = pd.merge(
        gstr2b_df, tally_df, 
        on="KEY", 
        suffixes=("_2B","_Tally")
    )
    
    # Calculate differences
    merged["VALUE_DIFFERENCE"] = merged["Taxable_Value_2B"] - merged["Taxable_Value_Tally"]
    merged["TAX_DIFFERENCE"] = merged["TOTAL_TAX_2B"] - merged["TOTAL_TAX_Tally"]
    
    # Check matches (within 1 rupee tolerance)
    merged["VALUE_MATCH"] = abs(merged["VALUE_DIFFERENCE"]) <= 1
    merged["TAX_MATCH"] = abs(merged["TAX_DIFFERENCE"]) <= 1
    
    # Categorize matches
    fully_matched = merged[merged["VALUE_MATCH"] & merged["TAX_MATCH"]].copy()
    value_mismatch = merged[~merged["VALUE_MATCH"]].copy()
    tax_mismatch = merged[merged["VALUE_MATCH"] & ~merged["TAX_MATCH"]].copy()
    
    # Calculate summary statistics
    match_percent = round(
        (len(fully_matched) / len(gstr2b_df) * 100) if len(gstr2b_df) else 0, 
        2
    )
    
    summary = {
        "Total_Invoices_Books": len(tally_df),
        "Total_Invoices_2B": len(gstr2b_df),
        "Total_Matched": len(fully_matched),
        "Total_Missing_Books": len(missing_books),
        "Total_Missing_2B": len(missing_2b),
        "Total_Missing_Books_value": round(missing_books_value, 2),
        "Total_Missing_2B_value": round(missing_2b_value, 2),
        "Match_Percentage": match_percent,
        "Total_ITC_Books": round(tally_df["TOTAL_TAX"].sum(),2),
        "Total_ITC_2B": round(gstr2b_df["TOTAL_TAX"].sum(),2),
        "ITC_Difference": round(
            gstr2b_df["TOTAL_TAX"].sum() - tally_df["TOTAL_TAX"].sum(), 2
        )
    }
    
    return {
        "fully_matched": fully_matched,
        "missing_in_books": missing_books,
        "missing_in_2b": missing_2b,
        "value_mismatch": value_mismatch,
        "tax_mismatch": tax_mismatch,
        "summary": summary
    }
