import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

columns_mapping = {
    'SRNO': [],
    'LOTNO': ['Report No'],
    'CARAT': ['Weight'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Pol'],
    'SYM': ['Sym'],
    'FLRC': ['Flu'],
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
    'Natts1':[],
    'Natts2':[],
    "Stone id":['Stock ID']
}

colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
flour = ['NON', 'FNT', 'MED']
cut = ['EX', 'VG']
culet = ['NON']
key_symbol = ['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']

def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names, drop 'Unnamed' residuals, and coerce mixed object cols to str."""
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(lambda x: x.decode("utf-8", "ignore") if isinstance(x, (bytes, bytearray)) else x)
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df

def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""
    return df[list(columns_mapping)]

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df[df['Shape'] == 'RBC']
    df['Shape'] = df['Shape'].str.replace('RBC', 'ROUND')
    df = df[(df['Weight'] >= 0.30) & (df['Weight'] <= 3.99)]
    df = df[df['Disc%'] <= -10]
    df = df[df['Color'].isin(colors)]
    df = df[df['Clarity'].isin(clarity)]
    df = df[df['Flu'].isin(flour)]
    df = df[df['Culet Size'].isin(culet)]
    df = df[df['Cut'].isin(cut) & df['Pol'].isin(cut) & df['Sym'].isin(cut)]
    df = df[~df['Key to Symbol'].isin(key_symbol)]
    return df

def default_value(df: pd.DataFrame) -> pd.DataFrame:
    for col in ['LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP']:
        if col in df.columns:
            df[col] = 0
    for num_col in ['CARAT', 'TABLE', 'DEPTH', 'DISC', 'RATIO', 'RAP', 'LENGTH', 'WIDTH', 'HEIGHT']:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")
    for col in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
        if col in df.columns:
            df[col] = 'N/A'
    if 'SRNO' in df.columns:
        df['SRNO'] = range(1, len(df) + 1)
    if 'NATTS' in df.columns:
        df['NATTS'] = 'CENTER - L'
    return df

def run_komal_mod() -> None:
    st.header("Komal Export")

    with st.sidebar:
        upload = st.file_uploader("Select File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)

    if upload is None:
        st.stop()

    try:
        df = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    # Show the raw (sanitized) data first
    st.dataframe(df)

    # Filters and normalization
    df = apply_filters(df)
    df = fix_columns(df)
    df = default_value(df)

    st.dataframe(df)

    csv_file = df.to_csv(index=False)
    file_name = f"komal export {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_file,
            file_name=file_name,
            mime="csv"
           
        )

if __name__ == "__main__":
    run_komal_mod()
