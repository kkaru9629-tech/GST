"""
GST Reconciliation Engine  v8.0 - FINAL
Fixes:
  1. detect_duplicate_invoices: uses Invoice_Date + light normalization (no false duplicates)
  2. normalize_invoice_for_grouping: light normalization only (alphanumeric, no numeric collapse)
  3. _normalize_invoice_for_level3: uses longest digit sequence (extract_numeric_core)
  4. level4_no_date_match: added for date-mismatch cases (booking date vs invoice date)
"""

import pandas as pd
import re
import logging
from typing import Tuple, Dict, Any, Optional
import numpy as np

GSTIN_PATTERN = re.compile(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}Z[0-9A-Z]{1}$')
logger = logging.getLogger(__name__)

# ── safe helpers ──────────────────────────────────────────────────────────────

def _f(val) -> float:
    if val is None: return 0.0
    if isinstance(val, float):
        return 0.0 if (np.isnan(val) or np.isinf(val)) else val
    if isinstance(val, (int, np.integer)): return float(val)
    s = str(val).strip().replace(',','').replace('₹','')
    if s in ('','-','nan','null','none','na','nil'): return 0.0
    try:
        v = float(s); return 0.0 if (np.isnan(v) or np.isinf(v)) else v
    except: return 0.0

def _s(val) -> str:
    if val is None: return ''
    if isinstance(val, float) and np.isnan(val): return ''
    return str(val).strip()

def _key(*parts) -> str:
    return '|'.join(_s(p) for p in parts)

def strict_numeric_cleaner(val) -> float: return _f(val)

def post_processing_cleaner(df):
    if df.empty: return df
    df = df.copy()
    for col in df.select_dtypes(include=[np.number]).columns:
        df[col] = df[col].apply(_f)
    return df

def pre_processing_cleaner(df):
    if df.empty: return df
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    fin = {'Taxable_Value','CGST','SGST','IGST','CESS','TOTAL_TAX','Invoice_Value'}
    for col in df.columns:
        if col in fin: df[col] = df[col].apply(_f)
        elif df[col].dtype == object:
            df[col] = df[col].apply(lambda x: str(x).strip() if pd.notna(x) else x)
    return df

def normalize_invoice_number(inv: str) -> str:
    if not inv:
        return ''
    s = str(inv).strip().upper()
    digits = re.sub(r'\D', '', s)
    if digits:
        return str(int(digits))
    return re.sub(r'[/\\\-_.\\s|,]', '', s)

def extract_numeric_core(inv: str) -> str:
    """Returns the LONGEST digit sequence from the invoice number."""
    if not inv: return ''
    seqs = re.findall(r'\d+', str(inv).strip())
    return max(seqs, key=len) if seqs else ''

def validate_gstin(g) -> bool:
    if not g: return False
    return bool(GSTIN_PATTERN.match(str(g).strip().upper()))

# ── format detection ──────────────────────────────────────────────────────────

def detect_file_format(df, filename=''):
    try:
        flat = ''
        for i in range(min(12, len(df))):
            flat += ' '.join(_s(v).lower() for v in df.iloc[i].tolist()) + ' '
        if 'purchase register' in flat: return 'tally_pr'
        for i in range(min(12, len(df))):
            vals = [_s(v).lower() for v in df.iloc[i].tolist()]
            if any('particulars' in v for v in vals) and any('supplier invoice' in v for v in vals):
                return 'tally_pr'
        if 'gstr-2b' in flat or 'gstr2b' in flat: return 'gstr2b_excel'
        for i in range(min(12, len(df))):
            vals = [_s(v).lower() for v in df.iloc[i].tolist()]
            if any('gstin of supplier' in v for v in vals): return 'gstr2b_excel'
        cols = [str(c).strip().lower() for c in df.columns]
        if all(r in cols for r in ['gstin','trade_name','invoice_no','taxable_value','cgst','sgst','igst']):
            return 'standard'
        return 'unknown'
    except Exception as e:
        logger.error(f'detect_file_format: {e}'); return 'unknown'

# ── Tally Purchase Register parser ───────────────────────────────────────────

def _find_tally_hdr(df):
    for i in range(min(12, len(df))):
        vals = [_s(v).lower() for v in df.iloc[i].tolist()]
        if any('particulars' in v for v in vals) and any('supplier invoice' in v for v in vals):
            return i
    return 5

def _map_tally_columns(header_row_values: list) -> dict:
    basic = {'date': None, 'particulars': None, 'inv_no': None, 'gstin': None, 'gross_total': None}
    tds_indices = []
    tax_indices = {'cgst': [], 'sgst': [], 'igst': [], 'cess': []}
    for idx, val in enumerate(header_row_values):
        v = _s(val).lower()
        if not v or v in ('nan', 'none', ''):
            continue
        if v == 'date':
            basic['date'] = idx
        elif 'particulars' in v:
            basic['particulars'] = idx
        elif 'supplier invoice' in v:
            basic['inv_no'] = idx
        elif 'gstin' in v and 'uin' in v:
            basic['gstin'] = idx
        elif v == 'gross total':
            basic['gross_total'] = idx
        if 'tds' in v:
            tds_indices.append(idx)
        if 'cgst' in v:
            tax_indices['cgst'].append(idx)
        if 'sgst' in v:
            tax_indices['sgst'].append(idx)
        if 'igst' in v:
            tax_indices['igst'].append(idx)
        if 'cess' in v:
            tax_indices['cess'].append(idx)
    return {**basic, 'tds_indices': tds_indices, 'tax_indices': tax_indices}

def parse_tally_purchase_register(raw_df):
    import streamlit as st
    logger.info('Parsing Tally Purchase Register...')
    hdr_row = _find_tally_hdr(raw_df)
    header_values = raw_df.iloc[hdr_row].tolist()
    cm = _map_tally_columns(header_values)
    debug_info = {'tax_indices': cm.get('tax_indices', {}), 'tds_indices': cm.get('tds_indices', [])}
    if 'tally_debug_info' not in st.session_state:
        st.session_state['tally_debug_info'] = []
    st.session_state['tally_debug_info'].append(debug_info)

    def get_basic(row, field):
        idx = cm.get(field)
        if idx is None or idx >= len(row):
            return '' if field in ('inv_no','gstin','particulars') else 0.0
        val = row[idx]
        if field in ('date','particulars','inv_no','gstin'):
            return _s(val)
        else:
            return _f(val)

    records = []
    for ri in range(hdr_row + 1, len(raw_df)):
        row = raw_df.iloc[ri].tolist()
        name = get_basic(row, 'particulars')
        if not name or name.lower() in ('', 'nan', 'none', 'grand total', 'total'):
            continue
        gross = get_basic(row, 'gross_total')
        if gross == 0.0:
            continue
        tds_total = sum(_f(row[idx]) for idx in cm.get('tds_indices', []) if idx < len(row))
        inv_val = gross + tds_total
        cgst = sum(_f(row[idx]) for idx in cm.get('tax_indices', {}).get('cgst', []) if idx < len(row) and 0 <= _f(row[idx]) <= inv_val + 0.01)
        sgst = sum(_f(row[idx]) for idx in cm.get('tax_indices', {}).get('sgst', []) if idx < len(row) and 0 <= _f(row[idx]) <= inv_val + 0.01)
        igst = sum(_f(row[idx]) for idx in cm.get('tax_indices', {}).get('igst', []) if idx < len(row) and 0 <= _f(row[idx]) <= inv_val + 0.01)
        cess = sum(_f(row[idx]) for idx in cm.get('tax_indices', {}).get('cess', []) if idx < len(row) and 0 <= _f(row[idx]) <= inv_val + 0.01)
        if cgst > 0 and sgst == 0 and igst == 0:
            sgst = cgst
        total_tax = cgst + sgst + igst + cess
        taxable = max(inv_val - total_tax, 0.0)
        gstin = get_basic(row, 'gstin').upper()
        inv_no = get_basic(row, 'inv_no')
        d_raw = row[cm.get('date', 0)] if cm.get('date', 0) is not None and cm.get('date', 0) < len(row) else None
        try:
            if hasattr(d_raw, 'year'):
                inv_date = pd.Timestamp(d_raw).normalize()
            elif d_raw is None or (isinstance(d_raw, float) and np.isnan(d_raw)):
                inv_date = pd.NaT
            else:
                inv_date = pd.to_datetime(_s(d_raw), dayfirst=True, errors='coerce').normalize()
        except:
            inv_date = pd.NaT
        records.append({
            'GSTIN': gstin, 'Trade_Name': name, 'Invoice_No': inv_no, 'Invoice_Date': inv_date,
            'Taxable_Value': round(taxable, 2), 'CGST': round(cgst, 2), 'SGST': round(sgst, 2),
            'IGST': round(igst, 2), 'CESS': round(cess, 2), 'TOTAL_TAX': round(total_tax, 2),
            'Invoice_Value': round(inv_val, 2)
        })
    if not records:
        raise ValueError('Tally Purchase Register: no data rows found.')
    return _validate_books_df(pd.DataFrame(records))

# ── GSTR-2B Excel parser ──────────────────────────────────────────────────────

def _find_gstr2b_start(df):
    for i in range(min(20, len(df))):
        for ci in [0, 1]:
            if df.shape[1] > ci:
                v = _s(df.iloc[i, ci]).upper()
                if v and GSTIN_PATTERN.match(v):
                    return i
    return 6

def _gstr2b_col_offset(raw_df, start_row):
    row = raw_df.iloc[start_row].tolist() if start_row < len(raw_df) else []
    if len(row) > 0 and GSTIN_PATTERN.match(_s(row[0]).upper()):
        return 0
    return 1

def parse_gstr2b_excel(raw_df):
    logger.info('Parsing GSTR-2B Excel...')
    start = _find_gstr2b_start(raw_df)
    off = _gstr2b_col_offset(raw_df, start)
    records = []
    for ri in range(start, len(raw_df)):
        row = raw_df.iloc[ri].tolist()
        gstin = _s(row[off+0]).upper() if len(row) > off+0 else ''
        if not gstin or not GSTIN_PATTERN.match(gstin): continue
        inv_no  = _s(row[off+2])  if len(row) > off+2  else ''
        d_raw   = row[off+4]      if len(row) > off+4  else None
        inv_val = _f(row[off+5])  if len(row) > off+5  else 0.0
        taxable = _f(row[off+8])  if len(row) > off+8  else 0.0
        igst    = _f(row[off+9])  if len(row) > off+9  else 0.0
        cgst    = _f(row[off+10]) if len(row) > off+10 else 0.0
        sgst    = _f(row[off+11]) if len(row) > off+11 else 0.0
        cess    = _f(row[off+12]) if len(row) > off+12 else 0.0
        name    = _s(row[off+1])  if len(row) > off+1  else ''
        try:
            if hasattr(d_raw, 'year'):
                inv_date = pd.Timestamp(d_raw).normalize()
            elif d_raw is None or (isinstance(d_raw, float) and np.isnan(d_raw)):
                inv_date = pd.NaT
            else:
                inv_date = pd.to_datetime(_s(d_raw), dayfirst=True, errors='coerce').normalize()
        except:
            inv_date = pd.NaT
        if pd.isna(inv_date): continue
        total_tax = cgst + sgst + igst + cess
        records.append({
            'GSTIN': gstin, 'Trade_Name': name, 'Invoice_No': inv_no,
            'Invoice_Date': inv_date, 'Taxable_Value': round(taxable, 2),
            'CGST': round(cgst, 2), 'SGST': round(sgst, 2), 'IGST': round(igst, 2),
            'CESS': round(cess, 2), 'TOTAL_TAX': round(total_tax, 2), 'Invoice_Value': round(inv_val, 2)
        })
    if not records:
        raise ValueError('GSTR-2B Excel: no valid invoice rows found.')
    return pd.DataFrame(records)

# ── standard template parsers ─────────────────────────────────────────────────

def parse_tally(df):
    df = pre_processing_cleaner(df)
    req = ['GSTIN','Trade_Name','Invoice_No','Invoice_Date','Taxable_Value','CGST','SGST','IGST','CESS']
    missing = [c for c in req if c not in df.columns]
    if missing: raise ValueError(f'Missing columns: {missing}')
    df['GSTIN'] = df['GSTIN'].apply(lambda v: _s(v).upper())
    df['Invoice_No'] = df['Invoice_No'].apply(_s)
    df['Invoice_Date'] = pd.to_datetime(df['Invoice_Date'], errors='coerce', dayfirst=True).dt.normalize()
    for col in ['Taxable_Value','CGST','SGST','IGST','CESS']:
        df[col] = df[col].apply(_f)
    df['TOTAL_TAX'] = df['CGST'] + df['SGST'] + df['IGST'] + df['CESS']
    df['Invoice_Value'] = df['Taxable_Value'] + df['TOTAL_TAX']
    return _validate_books_df(df)

def parse_gstr2b(df):
    df = pre_processing_cleaner(df.dropna(how='all').copy())
    req = ['GSTIN','Trade_Name','Invoice_No','Invoice_Date','Taxable_Value','CGST','SGST','IGST','CESS']
    missing = [c for c in req if c not in df.columns]
    if missing: raise ValueError(f'Missing columns: {missing}')
    df['GSTIN'] = df['GSTIN'].apply(lambda v: _s(v).upper())
    df['Invoice_No'] = df['Invoice_No'].apply(_s)
    df['Invoice_Date'] = pd.to_datetime(df['Invoice_Date'], errors='coerce', dayfirst=True).dt.normalize()
    for col in ['Taxable_Value','CGST','SGST','IGST','CESS']:
        df[col] = df[col].apply(_f)
    df['TOTAL_TAX'] = df['CGST'] + df['SGST'] + df['IGST'] + df['CESS']
    df['Invoice_Value'] = df['Taxable_Value'] + df['TOTAL_TAX']
    return df[df['Invoice_Date'].notna()].reset_index(drop=True)

# ── shared validation ─────────────────────────────────────────────────────────

def _row_to_issue(row):
    return {
        'GSTIN': str(row.get('GSTIN', '')), 'Trade_Name': str(row.get('Trade_Name', '')),
        'Invoice_No': str(row.get('Invoice_No', '')), 'Invoice_Date': row.get('Invoice_Date'),
        'Taxable_Value': _f(row.get('Taxable_Value', 0)), 'Invoice_Value': _f(row.get('Invoice_Value', 0)),
        'TOTAL_TAX': _f(row.get('TOTAL_TAX', 0))
    }

def _validate_books_df(df):
    df_t = df.copy()
    df_t['GSTIN'] = df_t['GSTIN'].apply(lambda v: '' if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v).strip())
    df_t['Invoice_No'] = df_t['Invoice_No'].apply(lambda v: '' if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v).strip())
    df_t['Invoice_Date'] = pd.to_datetime(df_t['Invoice_Date'], errors='coerce').dt.normalize().fillna(pd.Timestamp('1900-01-01'))
    dup_cols = ['GSTIN','Invoice_No','Invoice_Date','Taxable_Value','CGST','SGST','IGST','CESS']
    fdm = df_t.duplicated(subset=dup_cols, keep=False)
    fin_iss = [{**_row_to_issue(r), 'Issue': 'Duplicate Financial Row', 'Source': 'books_financial'} for _, r in df[fdm].iterrows()]
    df = df[~fdm].copy()
    val_iss, valid_idx = [], []
    for idx, row in df.iterrows():
        errs = []
        g = str(row['GSTIN']).strip()
        if not g or g.upper() in ('', 'NAN', 'NONE'): errs.append('No GSTIN')
        elif not validate_gstin(g): errs.append('Invalid GSTIN')
        inv = str(row['Invoice_No']).strip()
        if not inv or inv.upper() in ('', 'NAN', 'NONE'): errs.append('No Invoice No')
        if pd.isna(row['Invoice_Date']): errs.append('No Invoice Date')
        if errs: val_iss.append({**_row_to_issue(row), 'Issue': ', '.join(errs), 'Source': 'books_validation'})
        else: valid_idx.append(idx)
    all_iss = fin_iss + val_iss
    issues_df = pd.DataFrame(all_iss) if all_iss else pd.DataFrame()
    valid_df = df.loc[valid_idx].copy() if valid_idx else pd.DataFrame()
    if not valid_df.empty:
        no_itc = valid_df[valid_df['TOTAL_TAX'] == 0].copy()
        valid_df = valid_df[valid_df['TOTAL_TAX'] != 0].copy()
    else:
        no_itc = pd.DataFrame()
    logger.info(f'Valid:{len(valid_df)} NoITC:{len(no_itc)} Issues:{len(issues_df)}')
    return valid_df, no_itc, issues_df

# ── trade name mapping ────────────────────────────────────────────────────────

def create_trade_name_mapping(gstr_df, books_df):
    m = {}
    for df in [books_df, gstr_df]:
        if not df.empty:
            for _, row in df.iterrows():
                g = _s(str(row.get('GSTIN', ''))); n = _s(str(row.get('Trade_Name', '')))
                if g and n: m[g] = n
    return m

# ── normalizers ───────────────────────────────────────────────────────────────

def normalize_invoice_for_grouping(inv: str) -> str:
    """
    LIGHT normalization for grouping: remove special characters, keep alphanumeric.
    Does NOT collapse numeric values — preserves '011' as '011', not '11'.

    This is intentionally conservative so that invoices with different numeric
    identities are NOT accidentally merged during grouping.
    Numeric-level matching happens inside Level 3 and Level 4 only.

    Examples:
        SUN-INV/011  →  SUNINV011
        011          →  011
        SUN-INV/001  →  SUNINV001
        SUNINV001    →  SUNINV001   ← same as above, groups them correctly
    """
    if not inv:
        return ''
    return re.sub(r'[^A-Z0-9]', '', str(inv).strip().upper())


def _norm_l2(inv: str) -> str:
    """
    Level 2 normalization: same as grouping — remove special chars, keep alphanumeric.
    SUN-INV/001 → SUNINV001
    SUNINV001   → SUNINV001   ← L2 matches these two
    """
    if not inv:
        return ''
    return re.sub(r'[^A-Z0-9]', '', str(inv).strip().upper())


def _norm_l3(inv: str) -> str:
    """
    Level 3 normalization: extract LONGEST digit sequence, remove leading zeros.
    Uses extract_numeric_core (longest sequence) to avoid false concatenation.

    Examples:
        SUN-INV/011  →  '11'   (longest seq '011' → int 11)
        011          →  '11'
        ABC/6089/24  →  '6089' (longest of '6089','24')
        INV/001      →  '1'
    """
    if not inv:
        return ''
    core = extract_numeric_core(str(inv).strip())
    if not core:
        return ''
    try:
        return str(int(core))   # strips leading zeros
    except:
        return core

# ── group invoices ────────────────────────────────────────────────────────────

def group_invoices(df):
    """
    Group invoices within the same source (Books or GSTR-2B).
    Key: GSTIN + Invoice_Date + light-normalized Invoice_No.

    Light normalization groups 'SUN-INV/011' with 'SUNINV011' (same invoice,
    different formatting) but keeps '011' and 'SUN-INV/011' separate because
    '011' → '011' and 'SUN-INV/011' → 'SUNINV011' differ.
    Numeric-level matching ('011' ↔ 'SUN-INV/011') is handled by L3/L4.
    """
    if df.empty:
        return df
    df = df.copy()
    df['_grp'] = df['Invoice_No'].apply(lambda x: normalize_invoice_for_grouping(_s(x)))
    sc = [c for c in ['Taxable_Value','CGST','SGST','IGST','CESS','TOTAL_TAX','Invoice_Value'] if c in df.columns]
    grouped = df.groupby(['GSTIN','Invoice_Date','_grp'], as_index=False, dropna=False).agg({
        **{col: 'sum' for col in sc},
        'Trade_Name': lambda x: next((v for v in x if v and str(v).strip()), ''),
        'Invoice_No':  lambda x: next((v for v in x if v and str(v).strip()), '')
    })
    grouped = grouped.drop(columns=['_grp'])
    logger.info(f'Grouped {len(df)} → {len(grouped)}')
    return grouped

# ── detect duplicates ─────────────────────────────────────────────────────────

def detect_duplicate_invoices(df, source):
    """
    Detect genuinely duplicate invoice entries within the same source.

    FIX: duplicate key = GSTIN + light_normalized_Invoice_No + Invoice_Date + Invoice_Value.
    Invoice_Date is REQUIRED in the key to prevent false duplicate detection.

    Without date: 'SUNINV001' (02-Apr) and '01' (22-Apr) both normalize
    to same light-norm value for same supplier → incorrectly flagged as duplicates
    → one gets removed → reconciliation fails.

    With date: they have different dates → NOT flagged as duplicates → both
    reach the matching engine correctly.
    """
    if df.empty:
        return df, pd.DataFrame(), pd.DataFrame()
    dt = df.copy()
    dt['GSTIN'] = dt['GSTIN'].apply(lambda v: '' if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v).strip())
    dt['Invoice_No'] = dt['Invoice_No'].apply(lambda v: '' if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v).strip())

    # Light normalization + date + value as duplicate key
    dt['_norm_dup'] = dt['Invoice_No'].apply(lambda x: normalize_invoice_for_grouping(_s(x)))
    dt['_inv_date_str'] = pd.to_datetime(dt['Invoice_Date'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('NaT')

    mask = dt.duplicated(subset=['GSTIN', '_norm_dup', '_inv_date_str', 'Invoice_Value'], keep=False)
    dr = df[mask].copy()
    label = 'Duplicate Invoice in Books' if source == 'books' else 'Duplicate Invoice in GSTR-2B'
    di = [{**_row_to_issue(r), 'Issue': label, 'Source': source} for _, r in dr.iterrows()]
    dedup = df[~mask].copy()
    if not dr.empty:
        dt_dedup = dt[~mask].copy()
        first_dups = dt[mask].drop_duplicates(subset=['GSTIN', '_norm_dup', '_inv_date_str', 'Invoice_Value'], keep='first')
        first_dups = first_dups.drop(columns=['_norm_dup', '_inv_date_str'], errors='ignore')
        dedup = pd.concat([dt_dedup.drop(columns=['_norm_dup', '_inv_date_str'], errors='ignore'), first_dups], ignore_index=True)
    return dedup, dr, (pd.DataFrame(di) if di else pd.DataFrame())

# ── matching levels ───────────────────────────────────────────────────────────

def level1_strict_match(g, b, tol):
    """
    L1 — STRICT MATCH
    Key  : GSTIN + exact Invoice_No + Invoice_Date
    Guard: TOTAL_TAX within tolerance
    """
    g = g.copy(); b = b.copy()
    g['K'] = [_key(r['GSTIN'], r['Invoice_No'], r['Invoice_Date']) for _, r in g.iterrows()]
    b['K'] = [_key(r['GSTIN'], r['Invoice_No'], r['Invoice_Date']) for _, r in b.iterrows()]
    g['K'] = g['K'].astype(str); b['K'] = b['K'].astype(str)
    ks = set(g['K']) & set(b['K'])
    if not ks:
        return pd.DataFrame(), g.drop(columns=['K'], errors='ignore'), b.drop(columns=['K'], errors='ignore'), set()
    merged = pd.merge(g[g['K'].isin(ks)], b[b['K'].isin(ks)], on='K', suffixes=('_2B','_Books')).drop(columns=['K'], errors='ignore')
    merged['TAX_DIFF'] = merged['TOTAL_TAX_2B'].apply(_f) - merged['TOTAL_TAX_Books'].apply(_f)
    matched   = merged[merged['TAX_DIFF'].abs() <= tol].copy()
    unmatched = merged[merged['TAX_DIFF'].abs() >  tol].copy()
    failed_gk, failed_bk = set(), set()
    for _, row in unmatched.iterrows():
        failed_gk.add(_key(row.get('GSTIN_2B',       row.get('GSTIN','')),
                           row.get('Invoice_No_2B',   row.get('Invoice_No','')),
                           row.get('Invoice_Date_2B', row.get('Invoice_Date',''))))
        failed_bk.add(_key(row.get('GSTIN_Books',       row.get('GSTIN','')),
                           row.get('Invoice_No_Books',   row.get('Invoice_No','')),
                           row.get('Invoice_Date_Books', row.get('Invoice_Date',''))))
    g_rem = g[~g['K'].isin(ks) | g['K'].isin(failed_gk)].drop(columns=['K'], errors='ignore')
    b_rem = b[~b['K'].isin(ks) | b['K'].isin(failed_bk)].drop(columns=['K'], errors='ignore')
    logger.info(f'L1: {len(matched)} matched, {len(unmatched)} failed tax guard')
    return matched, g_rem, b_rem, ks


def level2_normalized_match(g, b, tol):
    """
    L2 — NORMALIZED MATCH (with date)
    Key  : GSTIN + Invoice_Date + alphanumeric-only Invoice_No
    Guard: TOTAL_TAX within tolerance

    Resolves: 'SUN-INV/001' ↔ 'SUNINV001' (same date)
              Both → 'SUNINV001' after removing special chars → match.
    """
    if g.empty or b.empty:
        return pd.DataFrame(), g, b, set()
    g = g.copy(); b = b.copy()
    g['NK'] = [_key(r['GSTIN'], r['Invoice_Date'], _norm_l2(_s(r['Invoice_No']))) for _, r in g.iterrows()]
    b['NK'] = [_key(r['GSTIN'], r['Invoice_Date'], _norm_l2(_s(r['Invoice_No']))) for _, r in b.iterrows()]
    ks = set(g['NK']) & set(b['NK'])
    if not ks:
        return pd.DataFrame(), g.drop(columns=['NK'], errors='ignore'), b.drop(columns=['NK'], errors='ignore'), set()
    merged = pd.merge(g[g['NK'].isin(ks)], b[b['NK'].isin(ks)], on='NK', suffixes=('_2B','_Books')).drop(columns=['NK'], errors='ignore')
    merged['TAX_DIFF'] = merged['TOTAL_TAX_2B'].apply(_f) - merged['TOTAL_TAX_Books'].apply(_f)
    matched = merged[merged['TAX_DIFF'].abs() <= tol].copy()
    consumed = set()
    for _, row in matched.iterrows():
        gstin    = _s(row.get('GSTIN_2B',       row.get('GSTIN', '')))
        inv_date =    row.get('Invoice_Date_2B', row.get('Invoice_Date', ''))
        inv      = _s(row.get('Invoice_No_2B',   row.get('Invoice_No', '')))
        consumed.add(_key(gstin, inv_date, _norm_l2(inv)))
    g_rem = g[~g['NK'].isin(consumed)].drop(columns=['NK'], errors='ignore')
    b_rem = b[~b['NK'].isin(consumed)].drop(columns=['NK'], errors='ignore')
    logger.info(f'L2: {len(matched)} matched')
    return matched, g_rem, b_rem, consumed


def level3_numeric_core_match(g, b, tol):
    """
    L3 — NUMERIC CORE MATCH (with date)
    Key  : GSTIN + Invoice_Date + longest digit sequence (leading zeros stripped)
    Guard: TOTAL_TAX within tolerance

    Resolves: 'SUN-INV/011' ↔ '011' (same date)
              Both → '11' after numeric core extraction → match.

    Uses extract_numeric_core (longest sequence) NOT re.sub(r'\\D','')
    to avoid false concatenation on multi-part invoice numbers like 'ABC/6089/24'.
    """
    if g.empty or b.empty:
        return pd.DataFrame(), g, b, set()
    g = g.copy(); b = b.copy()

    def _ck(row):
        core = _norm_l3(_s(row['Invoice_No']))
        return _key(row['GSTIN'], row['Invoice_Date'], core) if core else ''

    g['CK'] = [_ck(r) for _, r in g.iterrows()]
    b['CK'] = [_ck(r) for _, r in b.iterrows()]
    g_nc = g[g['CK'] == ''].drop(columns=['CK'], errors='ignore')
    b_nc = b[b['CK'] == ''].drop(columns=['CK'], errors='ignore')
    g = g[g['CK'] != ''].copy(); b = b[b['CK'] != ''].copy()
    ck = set(g['CK']) & set(b['CK'])
    if not ck:
        return pd.DataFrame(), \
               pd.concat([g.drop(columns=['CK'], errors='ignore'), g_nc], ignore_index=True), \
               pd.concat([b.drop(columns=['CK'], errors='ignore'), b_nc], ignore_index=True), set()
    gmap = {}
    for idx, row in g[g['CK'].isin(ck)].iterrows(): gmap.setdefault(row['CK'], []).append(idx)
    bmap = {}
    for idx, row in b[b['CK'].isin(ck)].iterrows(): bmap.setdefault(row['CK'], []).append(idx)
    rows, gu, bu = [], set(), set()
    for k in ck:
        for gi in gmap.get(k, []):
            if gi in gu: continue
            gr = g.loc[gi]
            for bi in bmap.get(k, []):
                if bi in bu: continue
                br = b.loc[bi]
                if abs(_f(gr['TOTAL_TAX']) - _f(br['TOTAL_TAX'])) <= tol:
                    gu.add(gi); bu.add(bi)
                    m = {f'{c}_2B': gr[c] for c in gr.index if c != 'CK'}
                    m.update({f'{c}_Books': br[c] for c in br.index if c != 'CK'})
                    m['TAX_DIFF'] = _f(gr['TOTAL_TAX']) - _f(br['TOTAL_TAX'])
                    rows.append(m); break
    matched = pd.DataFrame(rows) if rows else pd.DataFrame()
    g_rem = pd.concat([g[~g.index.isin(gu)].drop(columns=['CK'], errors='ignore'), g_nc], ignore_index=True)
    b_rem = pd.concat([b[~b.index.isin(bu)].drop(columns=['CK'], errors='ignore'), b_nc], ignore_index=True)
    logger.info(f'L3: {len(matched)} matched')
    return matched, g_rem, b_rem, set()


def level4_no_date_match(g, b, tol):
    """
    L4 — NUMERIC CORE MATCH (NO date)
    Key  : GSTIN + longest digit sequence only (date intentionally excluded)
    Guard: TOTAL_TAX within tolerance
    Pair : strictly one-to-one

    Resolves the real-world GST scenario where supplier's invoice date in
    GSTR-2B differs from accountant's booking date in Tally:
        SUN-INV/007 (06-Apr, GSTR) ↔ 007 (20-Apr, Books) → core '7' → match
        SUN-INV/001 (02-Apr, GSTR) ↔ 01  (22-Apr, Books) → core '1' → match

    Only reached after L1, L2, L3 all failed — so date mismatch is confirmed.
    Tax tolerance guard prevents false positives.
    """
    if g.empty or b.empty:
        return pd.DataFrame(), g, b, set()
    g = g.copy(); b = b.copy()

    def _l4k(row):
        core = _norm_l3(_s(row['Invoice_No']))
        return _key(row['GSTIN'], core) if core else ''

    g['L4K'] = [_l4k(r) for _, r in g.iterrows()]
    b['L4K'] = [_l4k(r) for _, r in b.iterrows()]
    g_nc = g[g['L4K'] == ''].drop(columns=['L4K'], errors='ignore')
    b_nc = b[b['L4K'] == ''].drop(columns=['L4K'], errors='ignore')
    g = g[g['L4K'] != ''].copy(); b = b[b['L4K'] != ''].copy()
    ck = set(g['L4K']) & set(b['L4K'])
    if not ck:
        return pd.DataFrame(), \
               pd.concat([g.drop(columns=['L4K'], errors='ignore'), g_nc], ignore_index=True), \
               pd.concat([b.drop(columns=['L4K'], errors='ignore'), b_nc], ignore_index=True), set()
    gmap = {}
    for idx, row in g[g['L4K'].isin(ck)].iterrows(): gmap.setdefault(row['L4K'], []).append(idx)
    bmap = {}
    for idx, row in b[b['L4K'].isin(ck)].iterrows(): bmap.setdefault(row['L4K'], []).append(idx)
    rows, gu, bu = [], set(), set()
    for k in ck:
        for gi in gmap.get(k, []):
            if gi in gu: continue
            gr = g.loc[gi]
            for bi in bmap.get(k, []):
                if bi in bu: continue
                br = b.loc[bi]
                if abs(_f(gr['TOTAL_TAX']) - _f(br['TOTAL_TAX'])) <= tol:
                    gu.add(gi); bu.add(bi)
                    m = {f'{c}_2B': gr[c] for c in gr.index if c != 'L4K'}
                    m.update({f'{c}_Books': br[c] for c in br.index if c != 'L4K'})
                    m['TAX_DIFF'] = _f(gr['TOTAL_TAX']) - _f(br['TOTAL_TAX'])
                    rows.append(m); break
    matched = pd.DataFrame(rows) if rows else pd.DataFrame()
    g_rem = pd.concat([g[~g.index.isin(gu)].drop(columns=['L4K'], errors='ignore'), g_nc], ignore_index=True)
    b_rem = pd.concat([b[~b.index.isin(bu)].drop(columns=['L4K'], errors='ignore'), b_nc], ignore_index=True)
    logger.info(f'L4 (no-date): {len(matched)} matched')
    return matched, g_rem, b_rem, set()

# ── reconcile ─────────────────────────────────────────────────────────────────

def reconcile(gstr_df, books_df, tolerance=1.0):
    logger.info('Starting reconciliation...')
    go = gstr_df.copy(); bo = books_df.copy()
    bd, _, bdup = detect_duplicate_invoices(bo, 'books')
    gd, _, gdup = detect_duplicate_invoices(go, 'gstr')
    dup_iss = pd.DataFrame()
    for d in [bdup, gdup]:
        if not d.empty: dup_iss = pd.concat([dup_iss, d], ignore_index=True)
    tmap = create_trade_name_mapping(go, bo)
    gg = group_invoices(gd); bg = group_invoices(bd)
    m1, gu1, bu1, _ = level1_strict_match(gg, bg, tolerance)
    m2, gu2, bu2, _ = level2_normalized_match(gu1, bu1, tolerance)
    m3, gu3, bu3, _ = level3_numeric_core_match(gu2, bu2, tolerance)
    m4, gu4, bu4, _ = level4_no_date_match(gu3, bu3, tolerance)
    am = pd.concat([m1, m2, m3, m4], ignore_index=True) if any(not x.empty for x in [m1, m2, m3, m4]) else pd.DataFrame()
    m2b = bu4.copy() if not bu4.empty else pd.DataFrame()
    mb  = gu4.copy() if not gu4.empty else pd.DataFrame()
    if not am.empty:
        am['TAX_DIFF']  = am['TOTAL_TAX_2B'].apply(_f) - am['TOTAL_TAX_Books'].apply(_f)
        am['TAX_MATCH'] = am['TAX_DIFF'].abs() <= tolerance
        matched = am[am['TAX_MATCH']].copy()
        tdiff   = am[~am['TAX_MATCH']].copy()
    else:
        matched = pd.DataFrame(); tdiff = pd.DataFrame()
    r2b = m2b['TOTAL_TAX'].apply(_f).sum() if not m2b.empty else 0
    ts  = abs(tdiff[tdiff['TAX_DIFF'] < 0]['TAX_DIFF'].apply(_f).sum()) if not tdiff.empty and (tdiff['TAX_DIFF'] < 0).any() else 0
    ng  = len(gg)
    summary = {
        'ITC_Books':   round(bg['TOTAL_TAX'].apply(_f).sum(), 2),
        'ITC_GSTR':    round(gg['TOTAL_TAX'].apply(_f).sum(), 2),
        'ITC_Diff':    round(gg['TOTAL_TAX'].apply(_f).sum() - bg['TOTAL_TAX'].apply(_f).sum(), 2),
        'ITC_at_Risk': round(r2b + ts, 2),
        'Match_%':     round(len(matched) / ng * 100, 2) if ng else 0,
        'Total_Books': len(bg), 'Total_GSTR': ng, 'Matched': len(matched),
        'Tax_Diff':    len(tdiff), 'Missing_2B': len(m2b), 'Missing_Books': len(mb)
    }
    logger.info(f"Match {summary['Match_%']}%, Risk {summary['ITC_at_Risk']:,.2f}")
    return {
        'matched': matched, 'tax_diff': tdiff, 'missing_2b': m2b, 'missing_books': mb,
        'summary': summary, 'trade_name_mapping': tmap, 'duplicate_issues': dup_iss
    }
