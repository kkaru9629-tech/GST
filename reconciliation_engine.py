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

def normalize_invoice_number(inv:str)->str:
    if not inv: return ''
    return re.sub(r'[/\-_\.\s\\|,]','',str(inv).strip().upper())

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
# Verified column positions from Karu.xls (header at row-index 5, data from row 6):
#  0=Date  1=Particulars  3=Supplier Inv No  4=GSTIN/UIN  8=Gross Total
#  11=ITC NR CGST  12=ITC NR SGST  13=Purchases
#  14=Input CGST  15=Input SGST  18=Input IGST
#  21=TDS Prof 194J  23=TDS Rent 194I
#
# KEY RULES:
#  CGST = col14 else col11
#  SGST = col15 else col12; if still 0 & CGST>0 & IGST=0 → SGST = CGST (intra-state)
#  IGST = col18
#  TDS  = col21 + col23
#  Invoice_Value = Gross Total + TDS
#  Taxable_Value = Invoice_Value - CGST - SGST - IGST

def _find_tally_hdr(df):
    for i in range(min(12, len(df))):
        vals = [_s(v).lower() for v in df.iloc[i].tolist()]
        if any('particulars' in v for v in vals) and any('supplier invoice' in v for v in vals):
            return i
    return 5

def _map_tally_columns(header_row_values: list) -> dict:
    """Scan header row, return field→col_index mapping. Falls back to verified defaults."""
    defaults = {
        'date': 0, 'particulars': 1, 'inv_no': 3, 'gstin': 4,
        'gross_total': 8,
        'itc_nr_cgst': 11, 'itc_nr_sgst': 12,
        'purchases': 13,
        'input_cgst': 14, 'input_sgst': 15,
        'input_igst': 18,
        'tds_prof': 21, 'tds_rent': 23,
    }
    found = {}
    for idx, val in enumerate(header_row_values):
        v = _s(val).lower()
        if not v or v in ('nan', 'none', ''):
            continue
        if v == 'date':
            found['date'] = idx
        elif 'particulars' in v:
            found['particulars'] = idx
        elif 'supplier invoice' in v:
            found['inv_no'] = idx
        elif 'gstin' in v and 'uin' in v:
            found['gstin'] = idx
        elif v == 'gross total':
            found['gross_total'] = idx
        elif 'itc not reflecting' in v and 'cgst' in v:
            found['itc_nr_cgst'] = idx
        elif 'itc not reflecting' in v and 'sgst' in v:
            found['itc_nr_sgst'] = idx
        elif v == 'purchases':
            found['purchases'] = idx
        elif 'input cgst' in v:
            found['input_cgst'] = idx
        elif 'input sgst' in v:
            found['input_sgst'] = idx
        elif 'input igst' in v:
            found['input_igst'] = idx
        elif 'tds on profession' in v or ('tds' in v and '194j' in v):
            found['tds_prof'] = idx
        elif 'tds on rent' in v or ('tds' in v and '194i' in v):
            found['tds_rent'] = idx
    return {**defaults, **found}

def parse_tally_purchase_register(raw_df):
    logger.info('Parsing Tally Purchase Register...')
    hdr_row = _find_tally_hdr(raw_df)
    logger.info(f'Header at row {hdr_row}')

    cm = _map_tally_columns(raw_df.iloc[hdr_row].tolist())
    logger.info(f'Tally column map: {cm}')

    records = []
    for ri in range(hdr_row + 1, len(raw_df)):
        row = raw_df.iloc[ri].tolist()

        def gc(field):
            idx = cm.get(field, -1)
            return _f(row[idx]) if 0 <= idx < len(row) else 0.0

        def gs(field):
            idx = cm.get(field, -1)
            return _s(row[idx]) if 0 <= idx < len(row) else ''

        name = gs('particulars')
        if not name or name.lower() in ('', 'nan', 'none', 'grand total', 'total'):
            continue
        gross = gc('gross_total')
        if gross == 0.0:
            continue

        cgst      = gc('input_cgst') or gc('itc_nr_cgst')
        sgst_main = gc('input_sgst')
        sgst_nr   = gc('itc_nr_sgst')
        sgst      = sgst_main if sgst_main != 0.0 else sgst_nr
        igst      = gc('input_igst')

        # Intra-state rule: infer SGST = CGST ONLY when BOTH SGST source columns
        # were genuinely empty (not zero-entered) and IGST is also absent.
        if cgst > 0.0 and igst == 0.0 and sgst_main == 0.0 and sgst_nr == 0.0:
            sgst = cgst

        tds       = gc('tds_prof') + gc('tds_rent')
        inv_val   = gross + tds
        taxable   = max(inv_val - cgst - sgst - igst, 0.0)
        total_tax = cgst + sgst + igst

        gstin  = gs('gstin').upper()
        inv_no = gs('inv_no')
        d_raw  = row[cm.get('date', 0)] if cm.get('date', 0) < len(row) else None

        # Parse date
        try:
            if hasattr(d_raw, 'year'):
                inv_date = pd.Timestamp(d_raw)
            elif d_raw is None or (isinstance(d_raw, float) and np.isnan(d_raw)):
                inv_date = pd.NaT
            else:
                inv_date = pd.to_datetime(_s(d_raw), dayfirst=True, errors='coerce')
        except Exception:
            inv_date = pd.NaT

        records.append({
            'GSTIN':         gstin,
            'Trade_Name':    name,
            'Invoice_No':    inv_no,
            'Invoice_Date':  inv_date,
            'Taxable_Value': round(taxable, 2),
            'CGST':          round(cgst, 2),
            'SGST':          round(sgst, 2),
            'IGST':          round(igst, 2),
            'CESS':          0.0,
            'TOTAL_TAX':     round(total_tax, 2),
            'Invoice_Value': round(inv_val, 2),
        })

    if not records:
        raise ValueError(
            'Tally Purchase Register: no data rows found. '
            'Ensure the file contains Purchase Register data with a Gross Total column.'
        )
    df = pd.DataFrame(records)
    logger.info(f'Tally PR: {len(df)} rows extracted')
    return _validate_books_df(df)

# ── GSTR-2B Excel parser ──────────────────────────────────────────────────────
# Verified column positions from Book1.xlsx (data starts row 6):
#  0=GSTIN  1=Trade Name  2=Invoice No  4=Invoice Date  5=Invoice Value
#  8=Taxable  9=IGST  10=CGST  11=SGST  12=CESS

def _find_gstr2b_start(df):
    """Find first row containing a valid GSTIN in col 0 or col 1."""
    for i in range(min(20, len(df))):
        for ci in [0, 1]:
            if df.shape[1] > ci:
                v = _s(df.iloc[i, ci]).upper()
                if v and GSTIN_PATTERN.match(v):
                    return i
    return 6

def _gstr2b_col_offset(raw_df, start_row):
    """Detect whether GSTIN is in column 0 or column 1 (some portals add a blank col A)."""
    row = raw_df.iloc[start_row].tolist() if start_row < len(raw_df) else []
    if len(row) > 0 and GSTIN_PATTERN.match(_s(row[0]).upper()):
        return 0
    return 1  # GSTIN is in col 1 (B)

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

        try:
            if hasattr(d_raw,'year'): inv_date = pd.Timestamp(d_raw)
            elif d_raw is None or (isinstance(d_raw,float) and np.isnan(d_raw)): inv_date = pd.NaT
            else: inv_date = pd.to_datetime(_s(d_raw), dayfirst=True, errors='coerce')
        except: inv_date = pd.NaT
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
    df['Invoice_Date'] = pd.to_datetime(df['Invoice_Date'], errors='coerce', dayfirst=True)
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
    df['Invoice_Date'] = pd.to_datetime(df['Invoice_Date'], errors='coerce', dayfirst=True)
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
    df_t['Invoice_Date'] = pd.to_datetime(df_t['Invoice_Date'],errors='coerce').fillna(pd.Timestamp('1900-01-01'))
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
