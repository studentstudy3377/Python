import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import Workbook

# Column mapping (similar to 0.py)
columns_mapping = {
    'SRNO': [],
    'LOTNO': ['CertNo'],
    'CARAT': ['SIZE'],
    'SHAPE': ['Shape'],
    'COLOR': ['COL'],
    'CLARITY': ['CLR'],
    'CUT': ['CUT'],
    'POL': ['POL'],
    'SYM': ['SYM'],
    'FLRC': ['FLR'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['TA%'],
    'DEPTH': ['TD%'],
    'LAB': ['Lab'],
    'DISC': ['DISC'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': []
}

# Lists for filtering
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
culet = ['NON']
key_symbol = ['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']

# Sanitize data
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

# Read uploaded file
def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)

# Fix column names
def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""  # for missing columns
    return df[list(columns_mapping)]

# Apply filters to the dataframe
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df[df['Lab'] == 'GIA']
    df = df[df['Shape'] == 'BR']
    df['Shape'] = df['Shape'].str.replace('BR', 'ROUND')
    df = df[df['DISC'] <= -10]
    df = df[df['LOC'] == 'INDIA']
    df = df[(df['SIZE'] >= 0.30) & (df['SIZE'] <= 3.99)]
    df = df[df['COL'].isin(colors)]
    df = df[df['CLR'].isin(clarity)]
    df = df[df['CUT'].isin(cut) & df['POL'].isin(cut) & df['SYM'].isin(cut)]
    df = df[df['FLR'].isin(flour)]
    df = df[df['BIC'] != 'Heavy']
    df = df[df['BIS'] != 'Heavy']
    df = df[~df['KTS'].isin(key_symbol)]
    return df

# Default values for missing fields
def default_value(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        if col in df.columns:
            df[col] = 0
    for num_col in ["CARAT", "TABLE", "DEPTH", "DISC", "RATIO", "RAP", "LENGTH", "WIDTH", "HEIGHT"]:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")
    for col in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
        if col in df.columns:
            df[col] = "N/A"
    if "SRNO" in df.columns:
        df["SRNO"] = range(1, len(df) + 1)
    if "NATTS" in df.columns:
        df["NATTS"] = "CENTER - L"
    if "DISC" in df.columns:
        df["DISC"] = df["DISC"].round(2)
    if "SHAPE" in df.columns:
        df["SHAPE"] = "ROUND"
    return df

# Main Streamlit app
def run_kbs_mod():
    st.header("KBS Diamond")
    with st.sidebar:
        upload = st.file_uploader("Select Excel file", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)
    
    if upload is None:
        st.info("Upload an Excel file to begin.")
        st.stop()

    try:
        df = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw Data (Sanitized)")
    st.dataframe(df)

    # Filtering and normalization (ported from original 0.py)
    df = apply_filters(df)

    # Column normalization
    df = fix_columns(df)

    # Default/typing
    df = default_value(df)

    st.subheader("Transformed Output")
    st.dataframe(df)

    # Prepare CSV for download
    csv_bytes = df.to_csv(index=False)
    today_str = pd.Timestamp.now().strftime("%d.%m.%Y")
    file_name = f"kbs diamonds {today_str}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_kbs_mod()
