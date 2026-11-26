# app_sagar_list.py
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["CertNo"],
    "CARAT": ["Carats"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Sym"],
    "FLRC": ["FL"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["TAB%"],
    "DEPTH": ["T.Dep%"],
    "LAB": ["Lab"],
    "DISC": ["Back%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':[],
    'Natts2':[],
    "Stone id":['StockID']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
location_allow = ["Mumbai", "MUMBAI"]
flour = ["NON", "FNT", "MED"]
key_symbol_block = ["Cavity", "Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

# ---------- Helpers ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore")
                if isinstance(x, (bytes, bytearray)) else x
            )
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df

def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)

def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df

def isin(df: pd.DataFrame, col: str, values) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col].isin(values)]
    return df

def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df

# ---------- Filters ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = isin(df, "Location", location_allow)

    # Shape RD -> ROUND
    if "Shape" in df.columns:
        df = df[df["Shape"] == "RD"]
        df["Shape"] = df["Shape"].astype(str).str.replace("RD", "ROUND", regex=False)

    # Back%: force negative; keep <= -10
    if "Back%" in df.columns:
        df["Back%"] = pd.to_numeric(df["Back%"], errors="coerce")
        df["Back%"] = df["Back%"].apply(lambda x: -x if pd.notna(x) and x >= 0 else x)
        df = df[df["Back%"] <= -10]

    # Carat range
    df = between(df, "Carats", 0.30, 3.99, inclusive="both")

    # Attributes
    df = isin(df, "Color", colors)
    df = isin(df, "Clarity", clarity)
    for c in ["Cut", "Polish", "Sym"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]
    df = isin(df, "FL", flour)

    # Blocks
    if "KeytoSymbol" in df.columns:
        df = df[~df["KeytoSymbol"].isin(key_symbol_block)]
    if "Comments" in df.columns:
        s = df["Comments"].astype(str)
        df = df[~s.str.startswith("Clarity", na=False)]

    return df

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Rename to standard headers + ensure expected columns exist
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""

    # SRNO
    if "SRNO" in df.columns:
        df["SRNO"] = range(1, len(df) + 1)

    # Defaults
    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        if col in df.columns:
            df[col] = 0
    for col in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
        if col in df.columns:
            df[col] = "N/A"
    if "NATTS" in df.columns:
        df["NATTS"] = "CENTER - L"

    # Column order
    return df[list(columns_mapping)]

# ---------- App ----------
def run_sagar_list_mod() -> None:
    st.header("Sagar List")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)  # your desktop script used header row

    if upload is None:
        st.stop()

    try:
        df_raw = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw Input (sanitized)")
    st.dataframe(df_raw)

    df = apply_filters(df_raw)
    df = fix_columns(df)

    st.subheader("Final Output")
    st.dataframe(df)

    # Download CSV
    file_name = f"sagar list {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_sagar_list_mod()
