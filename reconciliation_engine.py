"""
GST Reconciliation Engine  v7.0
Position-based parsers verified against Karu.xls + Book1.xlsx
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

# ===== MODIFIED: normalize_invoice_number =====
def normalize_invoice_number(inv: str) -> str:
    if not inv:
        return ''
    s = str(inv).strip().upper()
    nums = re.findall(r'\d+', s)
    if nums:
        return ''.join(nums)
    return re.sub(r'[/\\\-_.\\s|,]', '', s)
# =============================================

def extract_numeric_core(inv:str)->str:
    if not inv: return ''
    seqs = re.findall(r'\d+', str(inv).strip())
    return max(seqs, key=len) if seqs else ''

def validate_gstin(g)->bool:
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
    """
    Returns a dict:
      - basic fields: date, particulars, inv_no, gstin, gross_total
      - tds_indices: list of column indices that contain 'tds' (anywhere in header)
      - tax_indices: dict with keys 'cgst','sgst','igst','cess' -> list of column indices
    """
    basic = {'date': None, 'particulars': None, 'inv_no': None, 'gstin': None, 'gross_total': None}
    tds_indices = []
    tax_indices = {'cgst': [], 'sgst': [], 'igst': [], 'cess': []}

    for idx, val in enumerate(header_row_values):
        v = _s(val).lower()
        if not v or v in ('nan', 'none', ''):
            continue

        # Basic fields
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

        # TDS columns (any column with 'tds' in header)
        if 'tds' in v:
            tds_indices.append(idx)

        # Tax columns - substring match for cgst, sgst, igst, cess
        if 'cgst' in v:
            tax_indices['cgst'].append(idx)
        if 'sgst' in v:
            tax_indices['sgst'].append(idx)
        if 'igst' in v:
            tax_indices['igst'].append(idx)
        if 'cess' in v:
            tax_indices['cess'].append(idx)

    return {
        **basic,
        'tds_indices': tds_indices,
        'tax_indices': tax_indices,
    }

def parse_tally_purchase_register(raw_df):
    import streamlit as st
    
    logger.info('Parsing Tally Purchase Register...')
    hdr_row = _find_tally_hdr(raw_df)
    logger.info(f'Header at row {hdr_row}')

    header_values = raw_df.iloc[hdr_row].tolist()
    cm = _map_tally_columns(header_values)
    logger.info(f'Tally column map: {cm}')

    # ===== DEBUG EXPORT =====
    debug_info = {
        'detected_columns': {},
        'header_row_index': hdr_row,
        'header_values': header_values[:25] if len(header_values) > 25 else header_values,
    }
    for field, idx in cm.items():
        if field in ('date','particulars','inv_no','gstin','gross_total') and idx is not None:
            if 0 <= idx < len(header_values):
                debug_info['detected_columns'][field] = {
                    'index': idx,
                    'value': header_values[idx]
                }
    debug_info['tax_indices'] = cm.get('tax_indices', {})
    debug_info['tds_indices'] = cm.get('tds_indices', [])
    
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
    parsed_rows_for_debug = []
    
    for ri in range(hdr_row + 1, len(raw_df)):
        row = raw_df.iloc[ri].tolist()

        name = get_basic(row, 'particulars')
        if not name or name.lower() in ('', 'nan', 'none', 'grand total', 'total'):
            continue
        
        gross = get_basic(row, 'gross_total')
        if gross == 0.0:
            continue

        tds_total = 0.0
        for tds_idx in cm.get('tds_indices', []):
            if tds_idx < len(row):
                tds_total += _f(row[tds_idx])
        
        inv_val = gross + tds_total

        tax_indices = cm.get('tax_indices', {})
        cgst = 0.0
        for idx in tax_indices.get('cgst', []):
            if idx < len(row):
                val = _f(row[idx])
                if 0 <= val <= inv_val + 0.01:
                    cgst += val
        sgst = 0.0
        for idx in tax_indices.get('sgst', []):
            if idx < len(row):
                val = _f(row[idx])
                if 0 <= val <= inv_val + 0.01:
                    sgst += val
        igst = 0.0
        for idx in tax_indices.get('igst', []):
            if idx < len(row):
                val = _f(row[idx])
                if 0 <= val <= inv_val + 0.01:
                    igst += val
        cess = 0.0
        for idx in tax_indices.get('cess', []):
            if idx < len(row):
                val = _f(row[idx])
                if 0 <= val <= inv_val + 0.01:
                    cess += val
        
        if cgst > 0 and sgst == 0 and igst == 0:
            sgst = cgst
        
        total_tax = cgst + sgst + igst + cess
        taxable = max(inv_val - total_tax, 0.0)

        validation_flags = []
        if total_tax > inv_val + 0.01:
            validation_flags.append(f"TAX_EXCEEDS_INVOICE: total_tax={total_tax:.2f} > inv_val={inv_val:.2f}")
        if taxable > 0 and total_tax > 0:
            if total_tax / taxable > 0.30:
                validation_flags.append(f"HIGH_TAX_RATE: tax/taxable={total_tax/taxable:.1%}")
        if cgst > 0 and sgst > 0 and abs(cgst - sgst) > 0.01 and igst == 0:
            validation_flags.append(f"CGST_SGST_MISMATCH: cgst={cgst:.2f}, sgst={sgst:.2f}")

        gstin = get_basic(row, 'gstin').upper()
        inv_no = get_basic(row, 'inv_no')
        d_raw = row[cm.get('date', 0)] if cm.get('date', 0) is not None and cm.get('date', 0) < len(row) else None

        # ===== MODIFIED: date parsing with .normalize() =====
        try:
            if hasattr(d_raw, 'year'):
                inv_date = pd.Timestamp(d_raw).normalize()
            elif d_raw is None or (isinstance(d_raw, float) and np.isnan(d_raw)):
                inv_date = pd.NaT
            else:
                inv_date = pd.to_datetime(_s(d_raw), dayfirst=True, errors='coerce').normalize()
        except Exception:
            inv_date = pd.NaT
        # ====================================================

        record = {
            'GSTIN': gstin,
            'Trade_Name': name,
            'Invoice_No': inv_no,
            'Invoice_Date': inv_date,
            'Taxable_Value': round(taxable, 2),
            'CGST': round(cgst, 2),
            'SGST': round(sgst, 2),
            'IGST': round(igst, 2),
            'CESS': round(cess, 2),
            'TOTAL_TAX': round(total_tax, 2),
            'Invoice_Value': round(inv_val, 2),
            '_validation_flags': validation_flags if validation_flags else None,
        }
        
        if len(parsed_rows_for_debug) < 10:
            parsed_rows_for_debug.append({
                'row_index': ri,
                'name': name,
                'gross': gross,
                'tds': tds_total,
                'inv_val': inv_val,
                'cgst': cgst,
                'sgst': sgst,
                'igst': igst,
                'total_tax': total_tax,
                'taxable': taxable,
                'validation_flags': validation_flags,
                'raw_values': {
                    'date': d_raw,
                    'gstin': gstin,
                    'inv_no': inv_no,
                }
            })
        
        records.append(record)
    
    st.session_state['tally_debug_parsed_rows'] = parsed_rows_for_debug

    if not records:
        raise ValueError(
            'Tally Purchase Register: no data rows found. '
            'Ensure the file contains Purchase Register data with a Gross Total column.'
        )
    
    df = pd.DataFrame(records)
    df['_is_suspicious'] = df['_validation_flags'].notna() & (df['_validation_flags'].astype(str) != 'None')
    suspicious_rows = df[df['_is_suspicious']].copy() if df['_is_suspicious'].any() else pd.DataFrame()
    df_clean = df.drop(columns=['_validation_flags', '_is_suspicious'], errors='ignore')
    
    logger.info(f'Tally PR: {len(df_clean)} rows extracted, {len(suspicious_rows)} suspicious rows flagged')
    st.session_state['tally_suspicious_rows'] = suspicious_rows
    
    return _validate_books_df(df_clean)

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
    off   = _gstr2b_col_offset(raw_df, start)
    logger.info(f'Data starts at row {start}, GSTIN column offset={off}')
    records = []
    for ri in range(start, len(raw_df)):
        row = raw_df.iloc[ri].tolist()
        gstin = _s(row[off+0]).upper() if len(row)>off+0 else ''
        if not gstin or not GSTIN_PATTERN.match(gstin): continue

        inv_no  = _s(row[off+2])  if len(row)>off+2  else ''
        d_raw   = row[off+4]      if len(row)>off+4  else None
        inv_val = _f(row[off+5])  if len(row)>off+5  else 0.0
        taxable = _f(row[off+8])  if len(row)>off+8  else 0.0
        igst    = _f(row[off+9])  if len(row)>off+9  else 0.0
        cgst    = _f(row[off+10]) if len(row)>off+10 else 0.0
        sgst    = _f(row[off+11]) if len(row)>off+11 else 0.0
        cess    = _f(row[off+12]) if len(row)>off+12 else 0.0
        name    = _s(row[off+1])  if len(row)>off+1  else ''

        # ===== MODIFIED: date parsing with .normalize() =====
        try:
            if hasattr(d_raw,'year'):
                inv_date = pd.Timestamp(d_raw).normalize()
            elif d_raw is None or (isinstance(d_raw,float) and np.isnan(d_raw)):
                inv_date = pd.NaT
            else:
                inv_date = pd.to_datetime(_s(d_raw), dayfirst=True, errors='coerce').normalize()
        except:
            inv_date = pd.NaT
        # ====================================================
        if pd.isna(inv_date): continue

        total_tax = cgst+sgst+igst+cess
        records.append({'GSTIN':gstin,'Trade_Name':name,'Invoice_No':inv_no,
            'Invoice_Date':inv_date,'Taxable_Value':round(taxable,2),
            'CGST':round(cgst,2),'SGST':round(sgst,2),'IGST':round(igst,2),
            'CESS':round(cess,2),'TOTAL_TAX':round(total_tax,2),'Invoice_Value':round(inv_val,2)})

    if not records:
        raise ValueError('GSTR-2B Excel: no valid invoice rows found.')
    df = pd.DataFrame(records)
    logger.info(f'GSTR-2B Excel: {len(df)} records')
    return df

# ── standard template parsers ─────────────────────────────────────────────────

def parse_tally(df):
    df = pre_processing_cleaner(df)
    req = ['GSTIN','Trade_Name','Invoice_No','Invoice_Date','Taxable_Value','CGST','SGST','IGST','CESS']
    missing = [c for c in req if c not in df.columns]
    if missing: raise ValueError(f'Missing columns: {missing}')
    df['GSTIN']        = df['GSTIN'].apply(lambda v: _s(v).upper())
    df['Invoice_No']   = df['Invoice_No'].apply(_s)
    # ===== MODIFIED: date parsing with .normalize() =====
    df['Invoice_Date'] = pd.to_datetime(df['Invoice_Date'], errors='coerce', dayfirst=True).dt.normalize()
    # ====================================================
    for col in ['Taxable_Value','CGST','SGST','IGST','CESS']:
        df[col] = df[col].apply(_f)
    df['TOTAL_TAX']    = df['CGST']+df['SGST']+df['IGST']+df['CESS']
    df['Invoice_Value']= df['Taxable_Value']+df['TOTAL_TAX']
    return _validate_books_df(df)

def parse_gstr2b(df):
    df = pre_processing_cleaner(df.dropna(how='all').copy())
    req = ['GSTIN','Trade_Name','Invoice_No','Invoice_Date','Taxable_Value','CGST','SGST','IGST','CESS']
    missing = [c for c in req if c not in df.columns]
    if missing: raise ValueError(f'Missing columns: {missing}')
    df['GSTIN']        = df['GSTIN'].apply(lambda v: _s(v).upper())
    df['Invoice_No']   = df['Invoice_No'].apply(_s)
    # ===== MODIFIED: date parsing with .normalize() =====
    df['Invoice_Date'] = pd.to_datetime(df['Invoice_Date'], errors='coerce', dayfirst=True).dt.normalize()
    # ====================================================
    for col in ['Taxable_Value','CGST','SGST','IGST','CESS']:
        df[col] = df[col].apply(_f)
    df['TOTAL_TAX']    = df['CGST']+df['SGST']+df['IGST']+df['CESS']
    df['Invoice_Value']= df['Taxable_Value']+df['TOTAL_TAX']
    return df[df['Invoice_Date'].notna()].reset_index(drop=True)

# ── shared validation ─────────────────────────────────────────────────────────

def _row_to_issue(row):
    return {'GSTIN':str(row.get('GSTIN','')), 'Trade_Name':str(row.get('Trade_Name','')),
            'Invoice_No':str(row.get('Invoice_No','')), 'Invoice_Date':row.get('Invoice_Date'),
            'Taxable_Value':_f(row.get('Taxable_Value',0)), 'Invoice_Value':_f(row.get('Invoice_Value',0)),
            'TOTAL_TAX':_f(row.get('TOTAL_TAX',0))}

def _validate_books_df(df):
    df_t = df.copy()
    df_t['GSTIN']        = df_t['GSTIN'].apply(lambda v: '' if (v is None or (isinstance(v,float) and np.isnan(v))) else str(v).strip())
    df_t['Invoice_No']   = df_t['Invoice_No'].apply(lambda v: '' if (v is None or (isinstance(v,float) and np.isnan(v))) else str(v).strip())
    # ===== MODIFIED: normalize dates, fill NaT =====
    df_t['Invoice_Date'] = pd.to_datetime(df_t['Invoice_Date'], errors='coerce').dt.normalize().fillna(pd.Timestamp('1900-01-01'))
    # ================================================
    dup_cols = ['GSTIN','Invoice_No','Invoice_Date','Taxable_Value','CGST','SGST','IGST','CESS']
    fdm = df_t.duplicated(subset=dup_cols, keep=False)
    fin_iss = [{**_row_to_issue(r),'Issue':'Duplicate Financial Row','Source':'books_financial'}
               for _,r in df[fdm].iterrows()]
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
    valid_df  = df.loc[valid_idx].copy() if valid_idx else pd.DataFrame()
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
                g=_s(str(row.get('GSTIN',''))); n=_s(str(row.get('Trade_Name','')))
                if g and n: m[g]=n
    return m

# ── group invoices ────────────────────────────────────────────────────────────

def group_invoices(df):
    if df.empty: return df
    sc = [c for c in ['Taxable_Value','CGST','SGST','IGST','CESS','TOTAL_TAX','Invoice_Value'] if c in df.columns]
    g = df.groupby(['GSTIN','Trade_Name','Invoice_No','Invoice_Date'], as_index=False, dropna=False)[sc].sum()
    logger.info(f'Grouped {len(df)}→{len(g)}')
    return g

# ── detect duplicates ─────────────────────────────────────────────────────────

def detect_duplicate_invoices(df, source):
    if df.empty: return df, pd.DataFrame(), pd.DataFrame()
    dt = df.copy()
    dt['GSTIN']      = dt['GSTIN'].apply(lambda v: '' if (v is None or (isinstance(v,float) and np.isnan(v))) else str(v).strip())
    dt['Invoice_No'] = dt['Invoice_No'].apply(lambda v: '' if (v is None or (isinstance(v,float) and np.isnan(v))) else str(v).strip())
    mask = dt.duplicated(subset=['GSTIN','Invoice_No'], keep=False)
    dr = df[mask].copy()
    label = 'Duplicate Invoice in Books' if source=='books' else 'Duplicate Invoice in GSTR-2B'
    di = [{**_row_to_issue(r),'Issue':label,'Source':source} for _,r in dr.iterrows()]
    dedup = df[~mask].copy()
    if not dr.empty: dedup = pd.concat([dedup, dr.drop_duplicates(subset=['GSTIN','Invoice_No'], keep='first')])
    return dedup, dr, (pd.DataFrame(di) if di else pd.DataFrame())

# ── 3-level matching ──────────────────────────────────────────────────────────

def level1_strict_match(g, b):
    g=g.copy(); b=b.copy()
    g['K']=[_key(r['GSTIN'],r['Invoice_No'],r['Invoice_Date']) for _,r in g.iterrows()]
    b['K']=[_key(r['GSTIN'],r['Invoice_No'],r['Invoice_Date']) for _,r in b.iterrows()]
    g['K']=g['K'].astype(str); b['K']=b['K'].astype(str)
    ks=set(g['K'])&set(b['K'])
    m=pd.merge(g[g['K'].isin(ks)],b[b['K'].isin(ks)],on='K',suffixes=('_2B','_Books'))
    logger.info(f'L1:{len(m)} matched')
    return m, g[~g['K'].isin(ks)].copy(), b[~b['K'].isin(ks)].copy(), ks

def level2_normalized_match(g, b, tol):
    if g.empty or b.empty: return pd.DataFrame(),g,b,set()
    g=g.copy(); b=b.copy()
    g['NK']=[_key(r['GSTIN'],normalize_invoice_number(_s(r['Invoice_No'])),r['Invoice_Date']) for _,r in g.iterrows()]
    b['NK']=[_key(r['GSTIN'],normalize_invoice_number(_s(r['Invoice_No'])),r['Invoice_Date']) for _,r in b.iterrows()]
    g['NK']=g['NK'].astype(str); b['NK']=b['NK'].astype(str)
    ks=set(g['NK'])&set(b['NK'])
    if not ks: return pd.DataFrame(),g.drop(columns=['NK'],errors='ignore'),b.drop(columns=['NK'],errors='ignore'),set()
    m=pd.merge(g[g['NK'].isin(ks)],b[b['NK'].isin(ks)],on='NK',suffixes=('_2B','_Books')).drop(columns=['NK'],errors='ignore')
    logger.info(f'L2:{len(m)} matched')
    return m, g[~g['NK'].isin(ks)].drop(columns=['NK'],errors='ignore'), b[~b['NK'].isin(ks)].drop(columns=['NK'],errors='ignore'), ks

def level3_numeric_core_match(g, b, tol):
    if g.empty or b.empty: return pd.DataFrame(),g,b,set()
    g=g.copy(); b=b.copy()
    def bmap(df):
        m={}
        for idx,row in df.iterrows():
            c=extract_numeric_core(_s(row['Invoice_No']))
            if c: m.setdefault(_key(row['GSTIN'],c,row['Invoice_Date']),[]).append(idx)
        return m
    gm=bmap(g); bm=bmap(b); ck=set(gm)&set(bm)
    if not ck: return pd.DataFrame(),g,b,set()
    rows,gu,bu=[],set(),set()
    for k in ck:
        for gi in gm[k]:
            gr=g.loc[gi]
            for bi in bm[k]:
                if bi in bu: continue
                br=b.loc[bi]
                if abs(_f(gr['TOTAL_TAX'])-_f(br['TOTAL_TAX']))<=tol:
                    gu.add(gi); bu.add(bi)
                    merged={f'{c}_2B':gr[c] for c in gr.index}
                    merged.update({f'{c}_Books':br[c] for c in br.index})
                    merged['TAX_DIFF']=_f(gr['TOTAL_TAX'])-_f(br['TOTAL_TAX'])
                    rows.append(merged); break
    m=pd.DataFrame(rows) if rows else pd.DataFrame()
    logger.info(f'L3:{len(m)} matched')
    return m, g[~g.index.isin(gu)].copy(), b[~b.index.isin(bu)].copy(), set()

# ── reconcile ─────────────────────────────────────────────────────────────────

def reconcile(gstr_df, books_df, tolerance=1.0):
    logger.info('Starting reconciliation...')
    go=gstr_df.copy(); bo=books_df.copy()
    bd,_,bdup = detect_duplicate_invoices(bo,'books')
    gd,_,gdup = detect_duplicate_invoices(go,'gstr')
    dup_iss = pd.DataFrame()
    for d in [bdup,gdup]:
        if not d.empty: dup_iss=pd.concat([dup_iss,d],ignore_index=True)
    tmap = create_trade_name_mapping(go,bo)
    gg = group_invoices(gd); bg = group_invoices(bd)
    m1,gu1,bu1,_ = level1_strict_match(gg,bg)
    m2,gu2,bu2,_ = level2_normalized_match(gu1,bu1,tolerance)
    m3,gu3,bu3,_ = level3_numeric_core_match(gu2,bu2,tolerance)
    am = pd.concat([m1,m2,m3],ignore_index=True) if any(not x.empty for x in [m1,m2,m3]) else pd.DataFrame()
    m2b = bu3.copy() if not bu3.empty else pd.DataFrame()
    mb  = gu3.copy() if not gu3.empty else pd.DataFrame()
    if not am.empty:
        am['TAX_DIFF']  = am['TOTAL_TAX_2B'].apply(_f)-am['TOTAL_TAX_Books'].apply(_f)
        am['TAX_MATCH'] = am['TAX_DIFF'].abs()<=tolerance
        matched=am[am['TAX_MATCH']].copy(); tdiff=am[~am['TAX_MATCH']].copy()
    else: matched=pd.DataFrame(); tdiff=pd.DataFrame()
    r2b = m2b['TOTAL_TAX'].apply(_f).sum() if not m2b.empty else 0
    ts  = abs(tdiff[tdiff['TAX_DIFF']<0]['TAX_DIFF'].apply(_f).sum()) if not tdiff.empty and (tdiff['TAX_DIFF']<0).any() else 0
    ng  = len(gg)
    summary={'ITC_Books':round(bg['TOTAL_TAX'].apply(_f).sum(),2),
             'ITC_GSTR':round(gg['TOTAL_TAX'].apply(_f).sum(),2),
             'ITC_Diff':round(gg['TOTAL_TAX'].apply(_f).sum()-bg['TOTAL_TAX'].apply(_f).sum(),2),
             'ITC_at_Risk':round(r2b+ts,2),
             'Match_%':round(len(matched)/ng*100,2) if ng else 0,
             'Total_Books':len(bg),'Total_GSTR':ng,'Matched':len(matched),
             'Tax_Diff':len(tdiff),'Missing_2B':len(m2b),'Missing_Books':len(mb)}
    logger.info(f"Match {summary['Match_%']}%, Risk {summary['ITC_at_Risk']:,.2f}")
    return {'matched':matched,'tax_diff':tdiff,'missing_2b':m2b,'missing_books':mb,
            'summary':summary,'trade_name_mapping':tmap,'duplicate_issues':dup_iss}
