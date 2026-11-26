import streamlit as st
import pandas as pd
import numpy as np

# ---------- Configuration ----------
columns_mapping = {
    'SRNO': [],
    'LOTNO': ['Report No'],
    
    'CARAT': ['Size'],
    'SHAPE': [''],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Pol'],
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
    'B OPENS': [],
    'Natts_1':[],
    'Natts_2':[],
    "Stone id":['Stock Id']
}

colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
key_symbol = ['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']

# ---------- Utilities ----------
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
    """
    If negate=False  -> keep rows where df[col] is in allowed_values (optionally keeping NAs)
    If negate=True   -> keep rows where df[col] is NOT in allowed_values (optionally keeping NAs)
    """
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
    for col in ['LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP']:
        if col in df.columns:
            df[col] = 0
    for num_col in ['CARAT', 'TABLE', 'DEPTH', 'DISC', 'RATIO', 'RAP', 'LENGTH', 'WIDTH', 'HEIGHT']:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors='coerce')
    for col in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
        if col in df.columns:
            df[col] = 'N/A'
    if 'SRNO' in df.columns:
        df['SRNO'] = range(1, len(df) + 1)
    if 'NATTS' in df.columns:
        df['NATTS'] = 'CENTER - L'
    if 'DISC' in df.columns:
        df['DISC'] = df['DISC'].round(2)
    if 'SHAPE' in df.columns:
        df['SHAPE'] = 'ROUND'
    return df

# ---------- Filters (your correct format) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Follow original script logic but be resilient to missing columns
    df = equal(df, 'Lab', 'GIA')
    df = equal(df, 'Shape', 'ROUND')
    df = between(df, 'Size', 0.30, 3.99, inclusive='both')
    df = negation(df, "Color", colors)
    df = negation(df, "Clarity", clarity)
    df = negation(df, "Cut", cut)
    df = negation(df, "Pol", cut)
    df = negation(df, "Sym", cut)
    df = negation(df, "Flour", flour)
    df = negation(df, 'Key to Symbols', key_symbol,negate=True)

    if "Disc %" in df.columns:
        
        bad = df["Disc %"].between(-10, -0, inclusive="both")
        df = df[~bad & df["Disc %"].notna()]

    return df

# ---------- App ----------
def run_hiradias_mod() -> None:
    st.header('Hiradias')

    with st.sidebar:
        upload = st.file_uploader('Select Excel File', type=['xlsx', 'xls'])
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

    # Apply filters and normalize columns
    df = apply_filters(df)
    df = fix_columns(df)
    df = default_value(df)

    st.subheader('Processed Data')
    st.dataframe(df)

    csv_file = df.to_csv(index=False)
    file_name = f'hiradias {pd.Timestamp.now():%d.%m.%Y}.csv'

    with st.sidebar:
        st.download_button(
            label='Generate CSV',
            data=csv_file,
            file_name=file_name,
            mime='csv',
           
        )

if __name__ == '__main__':
    run_hiradias_mod()
