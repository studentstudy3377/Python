import streamlit as st
import pandas as pd
import numpy as np

# ---------- Config ----------
columns_mapping = {
    'SRNO': [],
    'LOTNO': ['Report No'],
    'CARAT': ['Carats'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Sym'],
    'FLRC': ['Flour'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table %'],
    'DEPTH': ['Depth %'],
    'LAB': ['Lab'],
    'DISC': ['Disc %'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': []
}

# Lists (same as your script)
colors  = ['D','E','F','G','H','I','J','K','L','M']
clarity = ['IF','VVS1','VVS2','VS1','VS2','SI1','SI2']
cut     = ['EX','VG']
flour   = ['NON','FNT','MED']

# ---------- Utils ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore") if isinstance(x, (bytes, bytearray)) else x
            )
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df

def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)

def numeric(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df

def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df = numeric(df, col)
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df

def negation(df: pd.DataFrame, col: str, allowed_values, negate: bool = False, keep_na: bool = True) -> pd.DataFrame:
    if col in df.columns:
        mark = df[col].isin(allowed_values)
        if negate:
            mark = ~mark
        if keep_na:
            mark = mark | df[col].isna()
        return df[mark]
    return df

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""
    return df[list(columns_mapping)]

def default_value(df: pd.DataFrame) -> pd.DataFrame:
    for col in ['LENGTH','WIDTH','HEIGHT','RATIO','RAP']:
        if col in df.columns:
            df[col] = 0
    for num_col in ['CARAT','TABLE','DEPTH','DISC','RATIO','RAP','LENGTH','WIDTH','HEIGHT']:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors='coerce')
    for col in ['TINCH','LUSTER','T OPENS','B OPENS']:
        if col in df.columns:
            df[col] = 'N/A'
    if 'SRNO' in df.columns:
        df['SRNO'] = range(1, len(df) + 1)
    if 'NATTS' in df.columns:
        df['NATTS'] = 'CENTER - L'
    if 'DISC' in df.columns:
        df['DISC'] = df['DISC'].round(2)
    # For output, ensure SHAPE reads ROUND (mirrors your BR->ROUND replacement)
    if 'SHAPE' in df.columns:
        df['SHAPE'] = df['SHAPE'].replace({'BR': 'ROUND'})
    return df

# ---------- Filters (mirrors your script) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = equal(df, 'Lab', 'GIA')
    df = equal(df, 'Shape', 'BR')             # original filter first…
    df = equal(df, 'Location', 'MUMBAI')
    # map BR -> ROUND (like your script) for downstream display/consistency
    if 'Shape' in df.columns:
        df['Shape'] = df['Shape'].replace({'BR': 'ROUND'})

    # Disc % <= -10
    if 'Disc %' in df.columns:
        df = numeric(df,'Disc %')
        df = df[df['Disc %'] <= -10.00]

    df = between(df, 'Carats', 0.30, 3.99, inclusive='both')
    df = negation(df, 'Color', colors)              # keep only in lists
    df = negation(df, 'Clarity', clarity)
    # Cut/Polish/Sym as EX/VG
    if all(c in df.columns for c in ['Cut','Polish','Sym']):
        df = df[df['Cut'].isin(cut) & df['Polish'].isin(cut) & df['Sym'].isin(cut)]
    else:
        df = negation(df, 'Cut', cut)
        df = negation(df, 'Polish', cut)
        df = negation(df, 'Sym', cut)
    df = negation(df, 'Flour', flour)
    df = equal(df, 'Culet Size', 'NON')
    return df

# ---------- App ----------
def run_hrv_mod() -> None:
    st.header('HRV')

    with st.sidebar:
        upload = st.file_uploader('Select HRV File', type=['xlsx','xls'])
        skip = st.number_input('Skip Rows', 0, 10000, 0, 1)

    if upload is None:
        st.stop()

    try:
        df = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f'Failed to read file: {e}')
        st.stop()

    st.subheader('Raw Data')
    st.dataframe(df)

    # Apply filters and normalize
    df = apply_filters(df)
    df = fix_columns(df)
    df = default_value(df)

    st.subheader('Processed Data')
    st.dataframe(df)

    csv_bytes = df.to_csv(index=False)
    file_name = f"hrv {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label='Generate CSV',
            data=csv_bytes,
            file_name=file_name,
            mime='csv',  
        )

if __name__ == '__main__':
    run_hrv_mod()
