# streamlit_app.py
import io
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime

# ---------- Criteria ----------
colors = ['D','E','F','G','H','I','J','K','L','M']
clarity = ['IF','VVS1','VVS2','VS1','VS2','SI1','SI2']
cut     = ['3EX','EX','VG']
flour   = ['NON','FNT','MED']
key_symbol_block = ['Cavity','Chip','Cavity, Chip','Chip, Cavity']
gia_comment_block = "Clarity grade is based on internal graining that is not shown."

# ---------- Column mapping (target : aliases in input) ----------
columns_mapping = {
    'SRNO': [],
    'LOTNO': ['Certi No.'],
    'CARAT': ['Cts'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symm'],
    'FLRC': ['Fls'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table(%)'],
    'DEPTH': ['Depth(%)'],
    'LAB': ['Lab'],
    'DISC': ['Disc(%)'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['Table Black'],
    'Natts2':['Crown Black'],
    "Stone id":['Stock Id']
}

# ---------- Helpers ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    for c in out.columns:
        if out[c].dtype == "object":
            out[c] = out[c].astype(str).str.strip()
    return out

def _to_num(series: pd.Series):
    return pd.to_numeric(series, errors="coerce")

def apply_filters(data: pd.DataFrame) -> pd.DataFrame:
    df = data.copy()

    # Normalize for robust comparisons
    if 'Lab' in df.columns:
        df['Lab'] = df['Lab'].astype(str).str.strip().str.upper()
    if 'Shape' in df.columns:
        df['Shape'] = df['Shape'].astype(str).str.strip().str.upper()
    if 'Location' in df.columns:
        df['Location'] = df['Location'].astype(str).str.strip().str.upper()

    # 1) Lab == GIA
    if 'Lab' in df.columns:
        df = df[df['Lab'] == 'GIA']

    # 2) Shape == ROUND
    if 'Shape' in df.columns:
        df = df[df['Shape'] == 'ROUND']

    # 3) Cts 0.30–3.99
    if 'Cts' in df.columns:
        cts = _to_num(df['Cts'])
        df = df[cts.between(0.30, 3.99, inclusive='both')]

    # 4) Location == INDIA
    if 'Location' in df.columns:
        df = df[df['Location'] == 'INDIA']

    # 5) Disc(%) <= -10
    if 'Disc(%)' in df.columns:
        disc = _to_num(df['Disc(%)'])
        df = df[disc <= -10]

    # 6) Attribute lists
    if 'Color' in df.columns:
        df = df[df['Color'].isin(colors)]
    if 'Clarity' in df.columns:
        df = df[df['Clarity'].isin(clarity)]
    for c in ['Cut', 'Polish', 'Symm']:
        if c in df.columns:
            df = df[df[c].isin(cut)]
    if 'Fls' in df.columns:
        df = df[df['Fls'].isin(flour)]

    # 7) Key To Symbol not in blocked list (exact match, same as your script)
    if 'Key To Symbol' in df.columns:
        df = df[~df['Key To Symbol'].isin(key_symbol_block)]

    # 8) GIA Comments should NOT start with specific clarity sentence
    if 'GIA Comments' in df.columns:
        s = df['GIA Comments'].astype(str)
        df = df[~s.str.startswith(gia_comment_block, na=False)]

    return df

def build_output(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)

    # Fill mapped columns
    for tgt_col, aliases in columns_mapping.items():
        if tgt_col in ['LENGTH','WIDTH','HEIGHT','RATIO','RAP']:
            out[tgt_col] = 0
        elif tgt_col == 'SRNO':
            out[tgt_col] = range(1, len(df) + 1)
        elif tgt_col in ['TINCH','LUSTER','T OPENS','B OPENS']:
            out[tgt_col] = 'N/A'
        elif tgt_col == 'NATTS':
            out[tgt_col] = 'CENTER - L'
        elif aliases:
            # pick first available alias
            src = next((a for a in aliases if a in df.columns), None)
            out[tgt_col] = df[src] if src else ""
        else:
            out[tgt_col] = ""

    # Reorder as per mapping
    out = out[list(columns_mapping.keys())]

    # Normalize labels where possible
    if 'Shape' in df.columns:
        out['SHAPE'] = 'ROUND'
    if 'Lab' in df.columns:
        out['LAB'] = 'GIA'
    if 'Cts' in df.columns:
        out['CARAT'] = _to_num(df['Cts']).fillna(0)
    if 'Disc(%)' in df.columns:
        out['DISC'] = _to_num(df['Disc(%)']).fillna("")

    return out

# ---------- Streamlit App ----------
def run_shairu_mod():
    st.title("Shairu")

    # Sidebar inputs + downloads live here
    with st.sidebar:
        uploaded = st.file_uploader("Select Excel file", type=["xlsx","xls"])
        skiprows = st.number_input("Skip rows", min_value=0, max_value=10000, value=5, step=1)

    if not uploaded:
        st.info("Upload an .xlsx/.xls file to begin.")
        st.stop()

    try:
        raw = pd.read_excel(uploaded, skiprows=skiprows)
        raw = _sanitize_df(raw)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw (sanitized)")
    st.dataframe(raw,  height=320)

    filt = apply_filters(raw)
   
    out = build_output(filt)
    st.subheader("Final Output")
    st.dataframe(out,  height=420)

    # ---------- Downloads in sidebar ----------
    today = datetime.now().strftime("%d.%m.%Y")
    csv_name  = f"shairu {today}.csv"
    
    csv_bytes = out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=csv_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_shairu_mod()
