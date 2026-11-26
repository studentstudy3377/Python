# app_rays.py
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Cert.No."],
    "CARAT": ["Carat"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symmetry"],
    "FLRC": ["FL"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table%"],
    "DEPTH": ["T.Dep."],
    "LAB": ["Lab"],
    "DISC": ["Disc%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Natts'],
    'Natts2':[],
    "Stone id":['Packet No']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut_allow = ["EX", "VG"]
flour_allow = ["NONE", "FNT", "MED"]  # input FL; later FLUOR is normalized NON/FNT/MED
shade_allow = ["NONE"]
milky_allow = ["NONE"]  # you require Milky == 'NONE'
table_open_allow = ["NONE"]
key_symbol_block = ["Cavity", "Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]
natts_block = ["C2", "T2", "T2C1", "T2C2", "T3C3"]

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

def read_file(uploaded_file, skiprows: int = 1) -> pd.DataFrame:
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

# ---------- Filters (Rays) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    df = between(df, "Carat", 0.30, 3.99, inclusive="both")
    df = equal(df, "Location", "INDIA")

    # Disc% <= -10
    if "Disc%" in df.columns:
        df["Disc%"] = pd.to_numeric(df["Disc%"], errors="coerce")
        df = df[df["Disc%"] <= -10]

    # Attribute filters
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    for c in ["Cut", "Polish", "Symmetry"]:
        if c in df.columns:
            df = df[df[c].isin(cut_allow)]
    if "FL" in df.columns:
        df = df[df["FL"].isin(flour_allow)]
        # Normalize into separate FLUOR, as in your script
        df.loc[:, "FLUOR"] = df["FL"].replace(
            {"FNT": "FNT", "MED": "MED", "NONE": "NON"}
        )

    # Shade & Milky strict equals 'NONE'
    if "Shade" in df.columns:
        df = df[df["Shade"].isin(shade_allow)]
    if "Milky" in df.columns:
        df = df[df["Milky"].isin(milky_allow)]

    # T.Open == 'NONE'
    if "T.Open" in df.columns:
        df = df[df["T.Open"].isin(table_open_allow)]

    # Exclusions
    if "Natts" in df.columns:
        df = df[~df["Natts"].isin(natts_block)]
    if "Key To Sym" in df.columns:
        df = df[~df["Key To Sym"].isin(key_symbol_block)]

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
def run_rays_web_mod() -> None:
    st.header("Rays Web")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 1, 1)  # default 1 like your script

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

    # Drop rows with any blanks, like your desktop pipeline
    df = df.dropna(how="any")

    st.subheader("Final Output")
    st.dataframe(df)

    # Download CSV
    file_name = f"rays {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_rays_web_mod()
