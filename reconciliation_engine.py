import pandas as pd
import re
import streamlit as st

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

                if (p_clean in col_clean or
                    pattern in col_space or
                    all(word in col_space for word in pattern.split())):
                    tax_cols[tax_type].append(col)
                    break

    return tax_cols


# ===========================================
# PARSE TALLY (with Invalid GSTIN Handling)
# ===========================================

def parse_tally(df):
    df = df.copy()

    # Detect header safely
    header_idx = None
    for i in range(min(len(df), 30)):
        row = " ".join(df.iloc[i].astype(str).str.lower().values)
        if "supplier invoice no" in row and "gstin" in row:
            header_idx = i
            break

    if header_idx is None:
        raise Exception("Tally header not detected. Please ensure file has 'Supplier Invoice No' and 'GSTIN' columns.")

    df.columns = df.iloc[header_idx]
    df = df.iloc[header_idx + 1:].reset_index(drop=True)
    df = df.dropna(how="all")

    # Clean column names
    df.columns = [str(col).strip() for col in df.columns]
    
    # Find the correct column names (case insensitive)
    col_mapping = {}
    for col in df.columns:
        col_lower = col.lower()
        if 'date' in col_lower:
            col_mapping[col] = 'Invoice_Date'
        elif 'particulars' in col_lower or 'party' in col_lower or 'supplier' in col_lower:
            col_mapping[col] = 'Trade_Name'
        elif 'supplier invoice no' in col_lower or 'invoice no' in col_lower or 'inv no' in col_lower:
            col_mapping[col] = 'Invoice_No'
        elif 'gstin' in col_lower or ('gst' in col_lower and 'uin' in col_lower):
            col_mapping[col] = 'GSTIN'
        elif 'gross total' in col_lower or 'invoice value' in col_lower or 'total' in col_lower:
            col_mapping[col] = 'Invoice_Value'
    
    df.rename(columns=col_mapping, inplace=True)

    required = ["GSTIN","Invoice_No","Invoice_Date"]
    for col in required:
        if col not in df.columns:
            # Try to find alternative
            found = False
            for df_col in df.columns:
                if col.lower() in df_col.lower():
                    df.rename(columns={df_col: col}, inplace=True)
                    found = True
                    break
            if not found:
                raise Exception(f"{col} column not found in Tally file. Please check file format.")

    df["GSTIN"] = df["GSTIN"].apply(clean_string)
    df["Invoice_No"] = df["Invoice_No"].apply(clean_invoice)
    df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)

    # Intelligent tax detection
    tax_cols = detect_tax_columns(df)

    # Check if any tax columns were found
    has_tax_columns = any([tax_cols['CGST'], tax_cols['SGST'], tax_cols['IGST']])
    
    if has_tax_columns:
        df["CGST"] = df[tax_cols['CGST']].apply(pd.to_numeric,errors='coerce').sum(axis=1) if tax_cols['CGST'] else 0
        df["SGST"] = df[tax_cols['SGST']].apply(pd.to_numeric,errors='coerce').sum(axis=1) if tax_cols['SGST'] else 0
        df["IGST"] = df[tax_cols['IGST']].apply(pd.to_numeric,errors='coerce').sum(axis=1) if tax_cols['IGST'] else 0
    else:
        # No tax columns found - set all to 0
        df["CGST"] = 0
        df["SGST"] = 0
        df["IGST"] = 0

    # TDS detection
    tds_cols = [c for c in df.columns if any(x in str(c).lower() for x in
        ['tds','t.d.s','tax deducted','tax deducted at source'])]

    df["TDS"] = df[tds_cols].apply(pd.to_numeric,errors="coerce").sum(axis=1).fillna(0) if tds_cols else 0

    df["Invoice_Value"] = pd.to_numeric(df["Invoice_Value"], errors="coerce")
    df["TOTAL_TAX"] = df["CGST"] + df["SGST"] + df["IGST"]

    # Calculate Taxable Value
    if df["Invoice_Value"].isna().all():
        if "Taxable_Value" in df.columns:
            df["Invoice_Value"] = pd.to_numeric(df["Taxable_Value"], errors="coerce") + df["TOTAL_TAX"]
        else:
            df["Invoice_Value"] = 0
    
    df["Taxable_Value"] = df["Invoice_Value"] + df["TDS"] - df["TOTAL_TAX"]

    # ===== GSTIN VALIDATION WITH SMART HANDLING ===== #
    df["GSTIN_VALID"] = df["GSTIN"].apply(validate_gstin)
    invalid_count = (~df["GSTIN_VALID"]).sum()
    
    # Store invalid invoices in a separate DataFrame
    invalid_invoices = pd.DataFrame()
    if invalid_count > 0:
        invalid_invoices = df[~df["GSTIN_VALID"]][[
            "GSTIN", "Trade_Name", "Invoice_No", "Invoice_Date", "Invoice_Value", "TOTAL_TAX"
        ]].copy()
        
        # Keep only valid ones for reconciliation
        df_valid = df[df["GSTIN_VALID"]].copy()
        
        # Add attributes to return
        df_valid.attrs['invalid_count'] = invalid_count
        df_valid.attrs['invalid_invoices'] = invalid_invoices
        
        # Use this for reconciliation
        df = df_valid

    # Add flag for NO ITC (zero tax)
    df["HAS_ITC"] = df["TOTAL_TAX"] > 0

    # Return with attributes
    return df[[
        "GSTIN","Trade_Name","Invoice_No","Invoice_Date",
        "Taxable_Value","Invoice_Value","IGST","CGST","SGST","TOTAL_TAX", "HAS_ITC"
    ]]


# ===========================================
# PARSE GSTR-2B (Structured Format Only)
# ===========================================

def parse_gstr2b(df):
    df = df.copy().dropna(how="all")

    # Clean column names
    df.columns = [str(col).strip() for col in df.columns]
    
    # Find matching columns
    col_mapping = {}
    for col in df.columns:
        col_lower = col.lower()
        if 'gstin' in col_lower and 'supplier' in col_lower:
            col_mapping[col] = "GSTIN"
        elif 'trade' in col_lower or 'legal' in col_lower or 'supplier' in col_lower:
            col_mapping[col] = "Trade_Name"
        elif 'invoice number' in col_lower:
            col_mapping[col] = "Invoice_No"
        elif 'invoice date' in col_lower:
            col_mapping[col] = "Invoice_Date"
        elif 'taxable value' in col_lower:
            col_mapping[col] = "Taxable_Value"
        elif 'integrated tax' in col_lower or 'igst' in col_lower:
            col_mapping[col] = "IGST"
        elif 'central tax' in col_lower or 'cgst' in col_lower:
            col_mapping[col] = "CGST"
        elif ('state' in col_lower and 'tax' in col_lower) or 'sgst' in col_lower:
            col_mapping[col] = "SGST"
    
    df.rename(columns=col_mapping, inplace=True)

    required = ["GSTIN","Invoice_No","Invoice_Date"]
    for col in required:
        if col not in df.columns:
            raise Exception(f"{col} not found in GSTR-2B file. Please use the structured format.")

    df["GSTIN"] = df["GSTIN"].apply(clean_string)
    df["Invoice_No"] = df["Invoice_No"].apply(clean_invoice)
    df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)

    for col in ["Taxable_Value","IGST","CGST","SGST"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    df["TOTAL_TAX"] = df["IGST"] + df["CGST"] + df["SGST"]
    df["Invoice_Value"] = df["Taxable_Value"] + df["TOTAL_TAX"]

    df = df.drop_duplicates(subset=["GSTIN","Invoice_No"])
    df = df[df["Invoice_Date"].notna()]

    return df[[
        "GSTIN","Trade_Name","Invoice_No","Invoice_Date",
        "Taxable_Value","Invoice_Value","IGST","CGST","SGST","TOTAL_TAX"
    ]]


# ===========================================
# RECONCILE (with NO ITC handling)
# ===========================================

def reconcile(gstr2b_df, tally_df):
    # Create keys
    gstr2b_df["KEY"] = gstr2b_df["GSTIN"] + "|" + gstr2b_df["Invoice_No"]
    tally_df["KEY"] = tally_df["GSTIN"] + "|" + tally_df["Invoice_No"]

    # Separate NO ITC invoices from Tally
    no_itc_tally = tally_df[~tally_df["HAS_ITC"]].copy()
    tally_with_itc = tally_df[tally_df["HAS_ITC"]].copy()

    # Missing in books (in GSTR-2B but not in Tally WITH ITC)
    missing_books = gstr2b_df[~gstr2b_df["KEY"].isin(tally_with_itc["KEY"])].copy()
    
    # Missing in 2B (in Tally WITH ITC but not in GSTR-2B)
    missing_2b = tally_with_itc[~tally_with_itc["KEY"].isin(gstr2b_df["KEY"])].copy()

    # NO ITC category - invoices in Tally with zero tax
    no_itc_with_2b = no_itc_tally[no_itc_tally["KEY"].isin(gstr2b_df["KEY"])].copy()
    no_itc_without_2b = no_itc_tally[~no_itc_tally["KEY"].isin(gstr2b_df["KEY"])].copy()

    # Merge matched invoices
    merged = pd.merge(
        gstr2b_df,
        tally_with_itc,
        on="KEY",
        suffixes=("_2B","_Tally"),
        how="inner"
    )

    # Calculate differences
    if not merged.empty:
        merged["VALUE_DIFFERENCE"] = merged["Taxable_Value_2B"] - merged["Taxable_Value_Tally"]
        merged["TAX_DIFFERENCE"] = merged["TOTAL_TAX_2B"] - merged["TOTAL_TAX_Tally"]
        merged["VALUE_MATCH"] = abs(merged["VALUE_DIFFERENCE"]) <= 1
        merged["TAX_MATCH"] = abs(merged["TAX_DIFFERENCE"]) <= 1

        fully_matched = merged[merged["VALUE_MATCH"] & merged["TAX_MATCH"]]
        value_mismatch = merged[~merged["VALUE_MATCH"]]
        tax_mismatch = merged[merged["VALUE_MATCH"] & ~merged["TAX_MATCH"]]
    else:
        fully_matched = pd.DataFrame()
        value_mismatch = pd.DataFrame()
        tax_mismatch = pd.DataFrame()

    match_percent = round(
        (len(fully_matched) / len(gstr2b_df) * 100)
        if len(gstr2b_df) else 0,
        2
    )

    summary = {
        "Total_Invoices_Books": len(tally_df),
        "Total_Invoices_2B": len(gstr2b_df),
        "Total_Matched": len(fully_matched),
        "Total_Missing_Books": len(missing_books),
        "Total_Missing_2B": len(missing_2b),
        "Total_No_ITC": len(no_itc_tally),
        "No_ITC_with_2B": len(no_itc_with_2b),
        "No_ITC_without_2B": len(no_itc_without_2b),
        "Match_Percentage": match_percent,
        "Total_ITC_Books": round(tally_with_itc["TOTAL_TAX"].sum(), 2) if not tally_with_itc.empty else 0,
        "Total_ITC_2B": round(gstr2b_df["TOTAL_TAX"].sum(), 2),
        "ITC_Difference": round(
            gstr2b_df["TOTAL_TAX"].sum() - (tally_with_itc["TOTAL_TAX"].sum() if not tally_with_itc.empty else 0),
            2
        )
    }

    return {
        "fully_matched": fully_matched,
        "missing_in_books": missing_books,
        "missing_in_2b": missing_2b,
        "no_itc": no_itc_tally,
        "no_itc_with_2b": no_itc_with_2b,
        "no_itc_without_2b": no_itc_without_2b,
        "value_mismatch": value_mismatch,
        "tax_mismatch": tax_mismatch,
        "summary": summary
    }
