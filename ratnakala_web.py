# app_ratnakala_web.py
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Report No"],
    "CARAT": ["Carat"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symmetry"],
    "FLRC": ["Fluorescent"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table%"],
    "DEPTH": ["Depth%"],
    "LAB": ["Lab"],
    "DISC": ["Disc %"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['TB'],
    'Natts2':['SB'],
    "Stone id":['Stone Id']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
milky_allow = ["NONE", "NON"]
cut_allow = ["EX", "VG"]
flour_allow = ["NON", "FNT", "MED"]
table_black_allow = ["NONE", "PINPOINT"]
culet_allow = ["NONE"]
key_symbol_block = ["Cavity", "Chip"]

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

def read_file(uploaded_file, skiprows: int = 4) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    df = _sanitize_df(df)
    # Trim to non-null 'SR No' count (replicates your desktop script)
    if "SR No" in df.columns:
        length = len(df["SR No"].dropna())
        df = df.iloc[:length]
    return df

def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df

def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df

# ---------- Filters (Ratnakala Web) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    df = between(df, "Carat", 0.30, 3.99, inclusive="both")
    df = equal(df, "Country", "INDIA")

    # Disc % <= -10
    if "Disc %" in df.columns:
        df["Disc %"] = pd.to_numeric(df["Disc %"], errors="coerce")
        df = df[df["Disc %"] <= -10]

    # Attribute filters
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    if "Milky" in df.columns:
        df = df[df["Milky"].isin(milky_allow)]
    if "Shade" in df.columns:
        df = df[df["Shade"] == "NN"]
    for c in ["Cut", "Polish", "Symmetry"]:
        if c in df.columns:
            df = df[df[c].isin(cut_allow)]
    if "Fluorescent" in df.columns:
        df = df[df["Fluorescent"].isin(flour_allow)]
    if "TB" in df.columns:
        df = df[df["TB"].isin(table_black_allow)]
    if "SB" in df.columns:
        df = df[df["SB"].isin(table_black_allow)]
    if "Table Open" in df.columns:
        df = df[df["Table Open"].isin(culet_allow)]
    if "Key To Symbol" in df.columns:
        df = df[~df["Key To Symbol"].isin(key_symbol_block)]

    return df

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Rename aliases + ensure all expected columns exist
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

    # Order
    return df[list(columns_mapping)]

# ---------- App ----------
def run_ratnakala_web_mod() -> None:
    st.header("Ratnakala Web")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 4, 1)  # default 4 like your script

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

    # Drop rows with any blank values (your desktop script behavior)
    df = df.dropna(how="any")

    st.subheader("Final Output")
    st.dataframe(df)

    # Download CSV
    file_name = f"ratnakal {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_ratnakala_web_mod()
