"""
GST Reconciliation Engine  v9.5 - FINAL IMPROVEMENTS
- Preserve original invoice numbers; add Normalized_Invoice_No for matching
- Level 3 numeric core selection uses longest meaningful sequence
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
    s = s.replace('₹', '')
    if s.lower() in ('','-','nan','null','none','na','nil'): return 0.0
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

def validate_gstin(g)->bool:
    if not g: return False
    return bool(GSTIN_PATTERN.match(str(g).strip().upper()))

# ── DATE NORMALIZATION HELPER ─────────────────────────────────────────────────

def normalize_date(dt):
    if pd.isna(dt):
        return ''
    parsed = parse_invoice_date(dt)
    if pd.isna(parsed):
        return ''
    return parsed.strftime('%Y-%m-%d')

def parse_invoice_date(dt):
    if pd.isna(dt):
        return pd.NaT
    try:
        if hasattr(dt, 'year'):
            return pd.Timestamp(dt).normalize()
        if isinstance(dt, (int, float, np.integer, np.floating)) and not pd.isna(dt):
            if 1 <= float(dt) <= 80000:
                return (pd.Timestamp('1899-12-30') + pd.to_timedelta(float(dt), unit='D')).normalize()
        return pd.to_datetime(_s(dt), dayfirst=True, errors='coerce').normalize()
    except Exception:
        return pd.NaT

# ── NORMALIZATION FUNCTIONS ───────────────────────────────────────────────────

def light_normalize(inv: str) -> str:
    if not inv:
        return ''
    return re.sub(r'[^A-Z0-9]', '', str(inv).strip().upper())

def extract_numeric_core(inv: str) -> str:
    """
    Extract the most relevant numeric sequence from invoice number.
    Ignores year-like numbers (1900-2100) and selects the longest remaining sequence.
    Returns as string with leading zeros removed via lstrip.
    """
    if not inv:
        return ''
    s = str(inv).strip().upper()
    nums = re.findall(r'\d+', s)
    if not nums:
        return ''
    
    filtered = []
    for n in nums:
        try:
            v = int(n)
            if 1900 <= v <= 2100:
                continue
            filtered.append(n)
        except:
            pass
    
    if filtered:
        core = max(filtered, key=len)
    else:
        core = max(nums, key=len)
    
    return core.lstrip('0') or '0'

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
        if any(x in v for x in ['cgst', 'central tax', 'input cgst', 'central gst']):
            tax_indices['cgst'].append(idx)
        if any(x in v for x in ['sgst', 'state tax', 'input sgst', 'state gst']):
            tax_indices['sgst'].append(idx)
        if any(x in v for x in ['igst', 'integrated tax', 'input igst', 'integrated gst']):
            tax_indices['igst'].append(idx)
        if any(x in v for x in ['cess', 'compensation cess']):
            tax_indices['cess'].append(idx)
    return {**basic, 'tds_indices': tds_indices, 'tax_indices': tax_indices}

def parse_tally_purchase_register(raw_df):
    import streamlit as st
    logger.info('Parsing Tally Purchase Register...')
    hdr_row = _find_tally_hdr(raw_df)
    header_values = raw_df.iloc[hdr_row].tolist()
    cm = _map_tally_columns(header_values)
    logger.info(f'Tally column mapping: {cm}')
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
        gstin = get_basic(row, 'gstin').upper()
        inv_no = get_basic(row, 'inv_no')

        cgst = sum(_f(row[idx]) for idx in cm.get('tax_indices', {}).get('cgst', []) if idx < len(row) and 0 <= _f(row[idx]) <= inv_val + 0.01)
        sgst = sum(_f(row[idx]) for idx in cm.get('tax_indices', {}).get('sgst', []) if idx < len(row) and 0 <= _f(row[idx]) <= inv_val + 0.01)
        igst = sum(_f(row[idx]) for idx in cm.get('tax_indices', {}).get('igst', []) if idx < len(row) and 0 <= _f(row[idx]) <= inv_val + 0.01)
        cess = sum(_f(row[idx]) for idx in cm.get('tax_indices', {}).get('cess', []) if idx < len(row) and 0 <= _f(row[idx]) <= inv_val + 0.01)
        if cgst > 0 and sgst == 0 and igst == 0:
            logger.warning(f'Possible SGST mapping issue for invoice {inv_no}: CGST found but SGST/IGST are zero')
        total_tax = cgst + sgst + igst + cess
        taxable = max(inv_val - total_tax, 0.0)

        d_raw = row[cm.get('date', 0)] if cm.get('date', 0) is not None and cm.get('date', 0) < len(row) else None
        inv_date = parse_invoice_date(d_raw)

        if len(records) < 20:
            logger.info(
                f"Invoice debug | Invoice:{inv_no} | Party:{name} | Gross:{gross} | "
                f"CGST:{cgst} | SGST:{sgst} | IGST:{igst} | CESS:{cess} | "
                f"TOTAL_TAX:{total_tax} | Invoice_Value:{inv_val} | Date_Raw:{d_raw} | Parsed_Date:{inv_date}"
            )

        record = {
            'GSTIN': gstin, 'Trade_Name': name, 'Invoice_No': inv_no, 'Invoice_Date': inv_date,
            'Taxable_Value': round(taxable, 2), 'CGST': round(cgst, 2), 'SGST': round(sgst, 2),
            'IGST': round(igst, 2), 'CESS': round(cess, 2), 'TOTAL_TAX': round(total_tax, 2),
            'Invoice_Value': round(inv_val, 2)
        }
        records.append(record)
    if not records:
        raise ValueError('Tally Purchase Register: no data rows found.')
    df = pd.DataFrame(records)
    return _validate_books_df(df)

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
        gstin = _s(row[off+0]).upper() if len(row)>off+0 else ''
        if not gstin or not GSTIN_PATTERN.match(gstin): continue
        inv_no = _s(row[off+2]) if len(row)>off+2 else ''
        d_raw = row[off+4] if len(row)>off+4 else None
        inv_val = _f(row[off+5]) if len(row)>off+5 else 0.0
        taxable = _f(row[off+8]) if len(row)>off+8 else 0.0
        igst = _f(row[off+9]) if len(row)>off+9 else 0.0
        cgst = _f(row[off+10]) if len(row)>off+10 else 0.0
        sgst = _f(row[off+11]) if len(row)>off+11 else 0.0
        cess = _f(row[off+12]) if len(row)>off+12 else 0.0
        name = _s(row[off+1]) if len(row)>off+1 else ''
        inv_date = parse_invoice_date(d_raw)
        total_tax = cgst+sgst+igst+cess
        records.append({'GSTIN':gstin,'Trade_Name':name,'Invoice_No':inv_no,
            'Invoice_Date':inv_date,'Taxable_Value':round(taxable,2),
            'CGST':round(cgst,2),'SGST':round(sgst,2),'IGST':round(igst,2),
            'CESS':round(cess,2),'TOTAL_TAX':round(total_tax,2),'Invoice_Value':round(inv_val,2)})
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
    df['Invoice_Date'] = df['Invoice_Date'].apply(parse_invoice_date)
    for col in ['Taxable_Value','CGST','SGST','IGST','CESS']:
        df[col] = df[col].apply(_f)
    df['TOTAL_TAX'] = df['CGST']+df['SGST']+df['IGST']+df['CESS']
    df['Invoice_Value'] = df['Taxable_Value']+df['TOTAL_TAX']
    return _validate_books_df(df)

def parse_gstr2b(df):
    df = pre_processing_cleaner(df.dropna(how='all').copy())
    req = ['GSTIN','Trade_Name','Invoice_No','Invoice_Date','Taxable_Value','CGST','SGST','IGST','CESS']
    missing = [c for c in req if c not in df.columns]
    if missing: raise ValueError(f'Missing columns: {missing}')
    df['GSTIN'] = df['GSTIN'].apply(lambda v: _s(v).upper())
    df['Invoice_No'] = df['Invoice_No'].apply(_s)
    df['Invoice_Date'] = df['Invoice_Date'].apply(parse_invoice_date)
    for col in ['Taxable_Value','CGST','SGST','IGST','CESS']:
        df[col] = df[col].apply(_f)
    df['TOTAL_TAX'] = df['CGST']+df['SGST']+df['IGST']+df['CESS']
    df['Invoice_Value'] = df['Taxable_Value']+df['TOTAL_TAX']
    return df.reset_index(drop=True)

# ── shared validation ─────────────────────────────────────────────────────────

def _row_to_issue(row):
    return {'GSTIN':str(row.get('GSTIN','')), 'Trade_Name':str(row.get('Trade_Name','')),
            'Invoice_No':str(row.get('Invoice_No','')), 'Invoice_Date':row.get('Invoice_Date'),
            'Taxable_Value':_f(row.get('Taxable_Value',0)), 'Invoice_Value':_f(row.get('Invoice_Value',0)),
            'TOTAL_TAX':_f(row.get('TOTAL_TAX',0))}

def _validate_books_df(df):
    df_t = df.copy()
    df_t['GSTIN'] = df_t['GSTIN'].apply(lambda v: '' if (v is None or (isinstance(v,float) and np.isnan(v))) else str(v).strip())
    df_t['Invoice_No'] = df_t['Invoice_No'].apply(lambda v: '' if (v is None or (isinstance(v,float) and np.isnan(v))) else str(v).strip())
    df_t['Invoice_Date'] = pd.to_datetime(df_t['Invoice_Date'], errors='coerce').dt.normalize().fillna(pd.Timestamp('1900-01-01'))
    dup_cols = ['GSTIN','Invoice_No','Invoice_Date','Taxable_Value','CGST','SGST','IGST','CESS']
    fdm = df_t.duplicated(subset=dup_cols, keep=False)
    fin_iss = [{**_row_to_issue(r),'Issue':'Duplicate Financial Row','Source':'books_financial'} for _,r in df[fdm].iterrows()]
    df = df[~fdm].copy()
    val_iss, valid_idx = [], []
    for idx, row in df.iterrows():
        errs = []
        g = str(row['GSTIN']).strip()
        if not g or g.upper() in ('','NAN','NONE'): errs.append('No GSTIN')
        elif not validate_gstin(g): errs.append('Invalid GSTIN')
        inv = str(row['Invoice_No']).strip()
        if not inv or inv.upper() in ('','NAN','NONE'): errs.append('No Invoice No')
        if pd.isna(row['Invoice_Date']): errs.append('No Invoice Date')
        if errs: val_iss.append({**_row_to_issue(row),'Issue':', '.join(errs),'Source':'books_validation'})
        else: valid_idx.append(idx)
    all_iss = fin_iss + val_iss
    issues_df = pd.DataFrame(all_iss) if all_iss else pd.DataFrame()
    valid_df = df.loc[valid_idx].copy() if valid_idx else pd.DataFrame()
    if not valid_df.empty:
        no_itc = valid_df[valid_df['TOTAL_TAX']==0].copy()
        valid_df = valid_df[valid_df['TOTAL_TAX']!=0].copy()
    else: no_itc = pd.DataFrame()
    logger.info(f'Valid:{len(valid_df)} NoITC:{len(no_itc)} Issues:{len(issues_df)}')
    return valid_df, no_itc, issues_df

# ── trade name mapping ────────────────────────────────────────────────────────

def create_trade_name_mapping(gstr_df, books_df):
    m = {}
    for df in [books_df, gstr_df]:
        if not df.empty:
            for _,row in df.iterrows():
                g=_s(str(row.get('GSTIN',''))).upper(); n=_s(str(row.get('Trade_Name','')))
                if g and n: m[g]=n
    return m

# ── GROUPING (PRESERVE ORIGINAL INVOICE, ADD NORMALIZED COLUMN) ──────────────

def group_invoices(df):
    if df.empty:
        return df
    
    df = df.copy()
    
    df['_normalized_invoice'] = df['Invoice_No'].apply(
        lambda x: light_normalize(_s(x))
    )
    
    df['_group_key'] = df.apply(
        lambda r: _key(
            r['GSTIN'], 
            normalize_date(r['Invoice_Date']),
            r['_normalized_invoice']
        ), axis=1
    )
    
    sc = [c for c in ['Taxable_Value', 'CGST', 'SGST', 'IGST', 'CESS', 'TOTAL_TAX', 'Invoice_Value'] 
          if c in df.columns]
    
    grouped = df.groupby('_group_key', as_index=False).agg({
        **{col: 'sum' for col in sc},
        'GSTIN': 'first',
        'Invoice_Date': 'first',
        'Trade_Name': lambda x: next((v for v in x if v and str(v).strip()), ''),
        'Invoice_No': 'first',
        '_normalized_invoice': 'first'
    })
    
    grouped['Normalized_Invoice_No'] = grouped['_normalized_invoice']
    grouped = grouped.drop(columns=['_group_key', '_normalized_invoice'])
    
    logger.info(f'Grouped {len(df)} → {len(grouped)}')
    return grouped

# ── DUPLICATE DETECTION (INCLUDES DATE) ──────────────────────────────────────

def detect_duplicate_invoices(df, source):
    if df.empty:
        return df, pd.DataFrame(), pd.DataFrame()
    
    dt = df.copy()
    dt['GSTIN'] = dt['GSTIN'].apply(lambda v: '' if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v).strip())
    dt['Invoice_No'] = dt['Invoice_No'].apply(lambda v: '' if (v is None or (isinstance(v, float) and np.isnan(v))) else str(v).strip())
    
    dt['_norm_dup'] = dt['Invoice_No'].apply(lambda x: light_normalize(_s(x)))
    dt['_inv_date_str'] = dt['Invoice_Date'].apply(normalize_date)
    
    mask = dt.duplicated(subset=['GSTIN', '_norm_dup', '_inv_date_str', 'Invoice_Value'], keep=False)
    dr = df[mask].copy()
    
    label = 'Duplicate Invoice in Books' if source == 'books' else 'Duplicate Invoice in GSTR-2B'
    di = [{**_row_to_issue(r), 'Issue': label, 'Source': source} for _, r in dr.iterrows()]
    
    dedup = df[~mask].copy()
    if not dr.empty:
        dt_dedup = dt[~mask].copy()
        first_dups = dt[mask].drop_duplicates(subset=['GSTIN', '_norm_dup', '_inv_date_str', 'Invoice_Value'], keep='first')
        first_dups = first_dups.drop(columns=['_norm_dup', '_inv_date_str'])
        dedup = pd.concat([dt_dedup.drop(columns=['_norm_dup', '_inv_date_str']), first_dups], ignore_index=True)
    
    return dedup, dr, (pd.DataFrame(di) if di else pd.DataFrame())

# ── 3-LEVEL MATCHING ──────────────────────────────────────────────────────────

def level1_strict_match(g, b, tol):
    g = g.copy()
    b = b.copy()

    g['K'] = [_key(r['GSTIN'], r['Invoice_No'], normalize_date(r['Invoice_Date'])) for _, r in g.iterrows()]
    b['K'] = [_key(r['GSTIN'], r['Invoice_No'], normalize_date(r['Invoice_Date'])) for _, r in b.iterrows()]

    ks = set(g['K']) & set(b['K'])
    if not ks:
        return pd.DataFrame(), g, b, set()

    gm = g[g['K'].isin(ks)].copy()
    bm = b[b['K'].isin(ks)].copy()
    merged = pd.merge(gm, bm, on='K', suffixes=('_2B', '_Books')).drop(columns=['K'], errors='ignore')

    merged['TAX_DIFF'] = merged['TOTAL_TAX_2B'].apply(_f) - merged['TOTAL_TAX_Books'].apply(_f)
    matched = merged[merged['TAX_DIFF'].abs() <= tol].copy()
    unmatched = merged[merged['TAX_DIFF'].abs() > tol].copy()

    failed_gkeys = set()
    failed_bkeys = set()
    if not unmatched.empty:
        for _, row in unmatched.iterrows():
            failed_gkeys.add(_key(row.get('GSTIN_2B'), row.get('Invoice_No_2B'), normalize_date(row.get('Invoice_Date_2B'))))
            failed_bkeys.add(_key(row.get('GSTIN_Books'), row.get('Invoice_No_Books'), normalize_date(row.get('Invoice_Date_Books'))))

    g_remaining = g[~g['K'].isin(ks) | g['K'].isin(failed_gkeys)].drop(columns=['K'], errors='ignore')
    b_remaining = b[~b['K'].isin(ks) | b['K'].isin(failed_bkeys)].drop(columns=['K'], errors='ignore')

    logger.info(f'L1: {len(matched)} matched')
    return matched, g_remaining, b_remaining, ks


def level2_normalized_match(g, b, tol):
    if g.empty or b.empty:
        return pd.DataFrame(), g, b, set()

    g = g.copy()
    b = b.copy()

    if 'Normalized_Invoice_No' in g.columns:
        g['NK'] = g['Normalized_Invoice_No']
    else:
        g['NK'] = [light_normalize(_s(r['Invoice_No'])) for _, r in g.iterrows()]
    
    if 'Normalized_Invoice_No' in b.columns:
        b['NK'] = b['Normalized_Invoice_No']
    else:
        b['NK'] = [light_normalize(_s(r['Invoice_No'])) for _, r in b.iterrows()]

    g['NK'] = [_key(r['GSTIN'], normalize_date(r['Invoice_Date']), r['NK']) for _, r in g.iterrows()]
    b['NK'] = [_key(r['GSTIN'], normalize_date(r['Invoice_Date']), r['NK']) for _, r in b.iterrows()]

    ks = set(g['NK']) & set(b['NK'])
    
    if not ks:
        g_remaining = g.drop(columns=['NK'], errors='ignore')
        b_remaining = b.drop(columns=['NK'], errors='ignore')
        return pd.DataFrame(), g_remaining, b_remaining, set()

    used_g = set()
    used_b = set()
    rows = []

    for nk in ks:
        g_rows = g[g['NK'] == nk]
        b_rows = b[b['NK'] == nk]
        
        for gi, gr in g_rows.iterrows():
            if gi in used_g:
                continue
            
            for bi, br in b_rows.iterrows():
                if bi in used_b:
                    continue
                
                tax_diff = abs(_f(gr['TOTAL_TAX']) - _f(br['TOTAL_TAX']))
                
                if tax_diff <= tol:
                    used_g.add(gi)
                    used_b.add(bi)
                    
                    merged_row = {
                        **{f'{c}_2B': gr[c] for c in gr.index if c != 'NK'},
                        **{f'{c}_Books': br[c] for c in br.index if c != 'NK'},
                        'TAX_DIFF': _f(gr['TOTAL_TAX']) - _f(br['TOTAL_TAX'])
                    }
                    rows.append(merged_row)
                    break

    matched = pd.DataFrame(rows) if rows else pd.DataFrame()
    g_remaining = g[~g.index.isin(used_g)].drop(columns=['NK'], errors='ignore')
    b_remaining = b[~b.index.isin(used_b)].drop(columns=['NK'], errors='ignore')

    logger.info(f'L2: {len(matched)} matched')
    return matched, g_remaining, b_remaining, set()


def level3_numeric_core_match(g, b, tol):
    if g.empty or b.empty:
        return pd.DataFrame(), g, b, set()

    g = g.copy()
    b = b.copy()

    g['CORE'] = g['Invoice_No'].apply(lambda x: extract_numeric_core(_s(x)))
    b['CORE'] = b['Invoice_No'].apply(lambda x: extract_numeric_core(_s(x)))

    g['CK'] = [_key(r['GSTIN'], normalize_date(r['Invoice_Date']), r['CORE']) for _, r in g.iterrows()]
    b['CK'] = [_key(r['GSTIN'], normalize_date(r['Invoice_Date']), r['CORE']) for _, r in b.iterrows()]

    g_valid = g[g['CORE'] != ''].copy()
    b_valid = b[b['CORE'] != ''].copy()
    g_no_key = g[g['CORE'] == ''].drop(columns=['CK', 'CORE'], errors='ignore')
    b_no_key = b[b['CORE'] == ''].drop(columns=['CK', 'CORE'], errors='ignore')

    ck = set(g_valid['CK']) & set(b_valid['CK'])
    if not ck:
        g_remaining = pd.concat([g_valid.drop(columns=['CK', 'CORE'], errors='ignore'), g_no_key], ignore_index=True)
        b_remaining = pd.concat([b_valid.drop(columns=['CK', 'CORE'], errors='ignore'), b_no_key], ignore_index=True)
        return pd.DataFrame(), g_remaining, b_remaining, set()

    gmap = {k: g_valid[g_valid['CK'] == k].index.tolist() for k in ck}
    bmap = {k: b_valid[b_valid['CK'] == k].index.tolist() for k in ck}

    rows, gu, bu = [], set(), set()
    for k in ck:
        for gi in gmap.get(k, []):
            if gi in gu:
                continue
            gr = g_valid.loc[gi]
            for bi in bmap.get(k, []):
                if bi in bu:
                    continue
                br = b_valid.loc[bi]
                if abs(_f(gr['TOTAL_TAX']) - _f(br['TOTAL_TAX'])) <= tol:
                    gu.add(gi)
                    bu.add(bi)
                    merged = {f'{c}_2B': gr[c] for c in gr.index if c not in ['CK', 'CORE']}
                    merged.update({f'{c}_Books': br[c] for c in br.index if c not in ['CK', 'CORE']})
                    merged['TAX_DIFF'] = _f(gr['TOTAL_TAX']) - _f(br['TOTAL_TAX'])
                    rows.append(merged)
                    break

    matched = pd.DataFrame(rows) if rows else pd.DataFrame()

    g_remaining = pd.concat([
        g_valid[~g_valid.index.isin(gu)].drop(columns=['CK', 'CORE'], errors='ignore'),
        g_no_key
    ], ignore_index=True)
    b_remaining = pd.concat([
        b_valid[~b_valid.index.isin(bu)].drop(columns=['CK', 'CORE'], errors='ignore'),
        b_no_key
    ], ignore_index=True)

    logger.info(f'L3: {len(matched)} matched')
    return matched, g_remaining, b_remaining, set()


def level4_date_mismatch_match(g, b, tol):
    if g.empty or b.empty:
        return pd.DataFrame(), g, b, set()

    g = g.copy()
    b = b.copy()

    g['DK_NORM'] = g['Normalized_Invoice_No'] if 'Normalized_Invoice_No' in g.columns else g['Invoice_No'].apply(lambda x: light_normalize(_s(x)))
    b['DK_NORM'] = b['Normalized_Invoice_No'] if 'Normalized_Invoice_No' in b.columns else b['Invoice_No'].apply(lambda x: light_normalize(_s(x)))
    g['DK_CORE'] = g['Invoice_No'].apply(lambda x: extract_numeric_core(_s(x)))
    b['DK_CORE'] = b['Invoice_No'].apply(lambda x: extract_numeric_core(_s(x)))

    g['DK'] = [
        [_key(r['GSTIN'], 'N', r['DK_NORM'])] +
        ([_key(r['GSTIN'], 'C', r['DK_CORE'])] if r['DK_CORE'] else [])
        for _, r in g.iterrows()
    ]
    b['DK'] = [
        [_key(r['GSTIN'], 'N', r['DK_NORM'])] +
        ([_key(r['GSTIN'], 'C', r['DK_CORE'])] if r['DK_CORE'] else [])
        for _, r in b.iterrows()
    ]

    common_keys = set(k for keys in g['DK'] for k in keys) & set(k for keys in b['DK'] for k in keys)
    rows, used_g, used_b = [], set(), set()

    for key in common_keys:
        g_rows = g[g['DK'].apply(lambda keys: key in keys)]
        b_rows = b[b['DK'].apply(lambda keys: key in keys)]

        for gi, gr in g_rows.iterrows():
            if gi in used_g:
                continue

            best_bi = None
            best_diff = None
            for bi, br in b_rows.iterrows():
                if bi in used_b:
                    continue
                tax_diff_abs = abs(_f(gr.get('TOTAL_TAX')) - _f(br.get('TOTAL_TAX')))
                if best_diff is None or tax_diff_abs < best_diff:
                    best_bi = bi
                    best_diff = tax_diff_abs

            if best_bi is not None:
                br = b.loc[best_bi]
                used_g.add(gi)
                used_b.add(best_bi)
                merged = {f'{c}_2B': gr[c] for c in gr.index if c not in ['DK', 'DK_NORM', 'DK_CORE']}
                merged.update({f'{c}_Books': br[c] for c in br.index if c not in ['DK', 'DK_NORM', 'DK_CORE']})
                merged['TAX_DIFF'] = _f(gr.get('TOTAL_TAX')) - _f(br.get('TOTAL_TAX'))
                g_date = normalize_date(gr.get('Invoice_Date'))
                b_date = normalize_date(br.get('Invoice_Date'))
                merged['DATE_MISMATCH'] = bool(g_date and b_date and g_date != b_date)
                rows.append(merged)

    matched = pd.DataFrame(rows) if rows else pd.DataFrame()
    g_remaining = g[~g.index.isin(used_g)].drop(columns=['DK', 'DK_NORM', 'DK_CORE'], errors='ignore')
    b_remaining = b[~b.index.isin(used_b)].drop(columns=['DK', 'DK_NORM', 'DK_CORE'], errors='ignore')

    logger.info(f'L4 date mismatch: {len(matched)} matched')
    return matched, g_remaining, b_remaining, set()

# ── RECONCILE ─────────────────────────────────────────────────────────────────

def reconcile(gstr_df, books_df, tolerance=1.0):
    logger.info('Starting reconciliation...')
    go = gstr_df.copy()
    bo = books_df.copy()
    
    bd, _, bdup = detect_duplicate_invoices(bo, 'books')
    gd, _, gdup = detect_duplicate_invoices(go, 'gstr')
    
    dup_iss = pd.DataFrame()
    for d in [bdup, gdup]:
        if not d.empty:
            dup_iss = pd.concat([dup_iss, d], ignore_index=True)
    
    tmap = create_trade_name_mapping(go, bo)
    gg = group_invoices(gd)
    bg = group_invoices(bd)
    
    m1, gu1, bu1, _ = level1_strict_match(gg, bg, tolerance)
    m2, gu2, bu2, _ = level2_normalized_match(gu1, bu1, tolerance)
    m3, gu3, bu3, _ = level3_numeric_core_match(gu2, bu2, tolerance)
    m4, gu4, bu4, _ = level4_date_mismatch_match(gu3, bu3, tolerance)
    
    am = pd.concat([m1, m2, m3, m4], ignore_index=True) if any(not x.empty for x in [m1, m2, m3, m4]) else pd.DataFrame()
    m2b = bu4.copy() if not bu4.empty else pd.DataFrame()
    mb = gu4.copy() if not gu4.empty else pd.DataFrame()
    
    if not am.empty:
        am['TAX_DIFF'] = am['TOTAL_TAX_2B'].apply(_f) - am['TOTAL_TAX_Books'].apply(_f)
        am['TAX_MATCH'] = am['TAX_DIFF'].abs() <= tolerance
        matched = am[am['TAX_MATCH']].copy()
        tdiff = am[~am['TAX_MATCH']].copy()
    else:
        matched = pd.DataFrame()
        tdiff = pd.DataFrame()
    
    r2b = m2b['TOTAL_TAX'].apply(_f).sum() if not m2b.empty else 0
    ts = abs(tdiff[tdiff['TAX_DIFF'] < 0]['TAX_DIFF'].apply(_f).sum()) if not tdiff.empty and (tdiff['TAX_DIFF'] < 0).any() else 0
    ng = len(gg)
    
    summary = {
        'ITC_Books': round(bg['TOTAL_TAX'].apply(_f).sum(), 2),
        'ITC_GSTR': round(gg['TOTAL_TAX'].apply(_f).sum(), 2),
        'ITC_Diff': round(gg['TOTAL_TAX'].apply(_f).sum() - bg['TOTAL_TAX'].apply(_f).sum(), 2),
        'ITC_at_Risk': round(r2b + ts, 2),
        'Match_%': round(len(matched) / ng * 100, 2) if ng else 0,
        'Total_Books': len(bg),
        'Total_GSTR': ng,
        'Matched': len(matched),
        'Tax_Diff': len(tdiff),
        'Missing_2B': len(m2b),
        'Missing_Books': len(mb)
    }
    
    logger.info(f"Match {summary['Match_%']}%, Risk {summary['ITC_at_Risk']:,.2f}")
    
    return {
        'matched': matched,
        'tax_diff': tdiff,
        'missing_2b': m2b,
        'missing_books': mb,
        'summary': summary,
        'trade_name_mapping': tmap,
        'duplicate_issues': dup_iss
    }
