# app_red_exim.py
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Report#"],
    "CARAT": ["Weight"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["CutGrade"],
    "POL": ["Polish"],
    "SYM": ["Symmetry"],
    "FLRC": ["Flour"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table"],
    "DEPTH": ["Depth"],
    "LAB": ["Lab"],
    "DISC": ["Disc"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':[],
    'Natts2':[],
    "Stone id":['Stock#']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut_allow = ["EX", "VG"]
flour_allow = ["NON", "FNT", "MED"]
culet_allow = ["NONE", "POINTED"]
no_bgm_variants = {"NO BGM", "NO BGM ", "NO BGM  "}

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

def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df

# ---------- Filters (Red Exim) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Force negative Disc
    if "Disc" in df.columns:
        df["Disc"] = pd.to_numeric(df["Disc"], errors="coerce")
        df["Disc"] = df["Disc"].apply(lambda x: -x if pd.notna(x) and x >= 0 else x)

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    if "Disc" in df.columns:
        df = df[df["Disc"] <= -10]
    df = between(df, "Weight", 0.30, 3.99, inclusive="both")
    df = equal(df, "Country", "INDIA")

    # Attribute filters
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    for c in ["CutGrade", "Polish", "Symmetry"]:
        if c in df.columns:
            df = df[df[c].isin(cut_allow)]
    if "Flour" in df.columns:
        df = df[df["Flour"].isin(flour_allow)]
    if "CuletSize" in df.columns:
        df = df[df["CuletSize"].isin(culet_allow)]

    # Milky/Shade allowed: null or 'NONE'
    if "Milky" in df.columns:
        df = df[df["Milky"].isna() | (df["Milky"] == "NONE")]
    if "Shade" in df.columns:
        df = df[df["Shade"].isna() | (df["Shade"] == "NONE")]

    # Block key symbols
    if "Key To Sym" in df.columns:
        df = df[~df["Key To Sym"].isin(["Cavity", "Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"])]

    # Comment must be one of NO BGM variants
    if "Comment" in df.columns:
        s = df["Comment"].astype(str).str.strip()
        df = df[s.isin({x.strip() for x in no_bgm_variants})]

    return df

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Rename to standard headers + ensure all expected columns exist
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
def run_red_exim_mod() -> None:
    st.header("Red Exim")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)  # default 0 like your script

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
    file_name = f"red exim {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",  
        )

if __name__ == "__main__":
    run_red_exim_mod()
