# streamlit_app.py
import base64
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime

# ---------- Criteria (Sheetal rules) ----------
colors  = ['D','E','F','G','H','I','J','K','L','M']
clarity = ['IF','VVS1','VVS2','VS1','VS2','SI1','SI2']
cut     = ['EX','VG']
flour   = ['NON','FNT','MED']
key_symbol_block = ['Cavity','Cavity','Chip','Cavity, Chip','Chip, Cavity']  # keep exact

# ---------- Column mapping (target : input aliases) ----------
columns_mapping = { 
    'SRNO': [],
    'LOTNO': ['Certificate'],
    'CARAT': ['Carat'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symmetry'],
    'FLRC': ['Fluorescence'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table%'],
    'DEPTH': ['Depth%'],
    'LAB': ['Lab'],
    'DISC': ['Disc%'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['BIT'],
    'Natts2':['BIC'],
    "Stone id":['StoneID']
}

# ---------- Helpers ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    for c in out.columns:
        if out[c].dtype == "object":
            out[c] = out[c].astype(str).str.strip()
    return out

def _to_num(s: pd.Series):
    return pd.to_numeric(s, errors="coerce")

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()

    # normalize text columns used
    for c in ['Lab','Shape','Location','Hue','Milky','BIT','BIC','BOC','Keytosymbols']:
        if c in d.columns:
            d[c] = d[c].astype(str).str.strip()

    # Shape: RBC -> ROUND
    if 'Shape' in d.columns:
        d['Shape'] = d['Shape'].str.upper().str.replace('RBC','ROUND', regex=False)

    # Lab == GIA
    if 'Lab' in d.columns:
        d = d[d['Lab'].str.upper() == 'GIA']

    # Shape == ROUND
    if 'Shape' in d.columns:
        d = d[d['Shape'].str.upper() == 'ROUND']

    # Carat 0.30—3.99
    if 'Carat' in d.columns:
        car = _to_num(d['Carat'])
        d = d[car.between(0.30, 3.99, inclusive='both')]

    # Location == INDIA
    if 'Location' in d.columns:
        d = d[d['Location'].str.upper() == 'INDIA']

    # Disc% <= -10
    if 'Disc%' in d.columns:
        disc = _to_num(d['Disc%'])
        d = d[disc <= -10]

    # Hue == WH
    if 'Hue' in d.columns:
        d = d[d['Hue'].str.upper() == 'WH']

    # Milky == NO
    if 'Milky' in d.columns:
        d = d[d['Milky'].str.upper() == 'NO']

    # Exclude BIT2 / BIC2 / BOC2
    if 'BIT' in d.columns:
        d = d[d['BIT'].str.upper() != 'BIT2']
    if 'BIC' in d.columns:
        d = d[d['BIC'].str.upper() != 'BIC2']
    if 'BOC' in d.columns:
        d = d[d['BOC'].str.upper() != 'BOC2']

    # Attribute lists
    if 'Color' in d.columns:
        d = d[d['Color'].isin(colors)]
    if 'Clarity' in d.columns:
        d = d[d['Clarity'].isin(clarity)]
    if 'Cut' in d.columns:
        d = d[d['Cut'].isin(cut)]
    if 'Polish' in d.columns:
        d = d[d['Polish'].isin(cut)]
    if 'Symmetry' in d.columns:
        d = d[d['Symmetry'].isin(cut)]
    if 'Fluorescence' in d.columns:
        d = d[d['Fluorescence'].isin(flour)]

    # Keytosymbols not in blocked list
    if 'Keytosymbols' in d.columns:
        d = d[~d['Keytosymbols'].isin(key_symbol_block)]

    return d

def build_output(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)

    # map into target schema
    for tgt_col, src_cols in columns_mapping.items():
        if tgt_col in ['LENGTH','WIDTH','HEIGHT','RATIO','RAP']:
            out[tgt_col] = 0
        elif tgt_col == 'SRNO':
            out[tgt_col] = range(1, len(df) + 1)
        elif tgt_col in ['TINCH','LUSTER','T OPENS','B OPENS']:
            out[tgt_col] = 'N/A'
        elif tgt_col == 'NATTS':
            out[tgt_col] = 'CENTER - L'
        elif src_cols:
            src = next((c for c in src_cols if c in df.columns), None)
            out[tgt_col] = df[src] if src else np.nan
        else:
            out[tgt_col] = np.nan

    # reorder
    out = out[list(columns_mapping.keys())]

    # normalize outputs
    if 'Shape' in df.columns:
        out['SHAPE'] = 'ROUND'
    if 'Lab' in df.columns:
        out['LAB']   = 'GIA'
    if 'Carat' in df.columns:
        out['CARAT'] = _to_num(df['Carat']).fillna(0)
    if 'Disc%' in df.columns:
        out['DISC']  = _to_num(df['Disc%'])

    # drop rows with any NaN to match your original behavior
    out = out.dropna(how='any')

    return out

def _find_sample(filename: str):
    """Locate a file inside a 'Sample Files' folder, searching upward from this
    script. Works on Linux (Streamlit Cloud) and Windows regardless of the
    current working directory or where this page sits in a multipage app."""
    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        candidate = parent / "Sample Files" / filename
        if candidate.exists():
            return candidate
    return None

# ---------- Streamlit App (Shairu-style layout) ----------
def run_sheetal_mod():
    st.title("Sheetal")

    # Sidebar: upload + options (same pattern as Shairu)
    sample_path = _find_sample("Sheetal.xlsx")
    
    with st.sidebar:
    
        if sample_path is not None:
            with open(sample_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            href = (
                f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" '
                f'download="{sample_path.name}">Sample File</a>'
            )
            st.markdown(href, unsafe_allow_html=True)
        else:
            st.warning("Sample file not found on server.")
        uploaded = st.file_uploader("Select Excel file", type=["xlsx","xls"])
        skiprows = st.number_input("Skip rows", min_value=0, max_value=10000, value=8, step=1)

    if not uploaded:
        st.info("Upload an .xlsx/.xls file to begin.")
        st.stop()

    # Read & sanitize (Sheet1, skiprows default 8 like your original)
    try:
        raw = pd.read_excel(uploaded, skiprows=skiprows)
        raw = _sanitize_df(raw)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("File ")
    st.dataframe(raw,  height=320)

    # Apply filters and build final output
    filt = apply_filters(raw)
    out  = build_output(filt)

    st.subheader("Final Output")
    st.dataframe(out,  height=420)

    # Downloads in sidebar (same as Shairu flow)
    today = datetime.now().strftime("%d.%m.%Y")
    csv_name  = f"sheetal {today}.csv"
    csv_bytes = out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=csv_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_sheetal_mod()
