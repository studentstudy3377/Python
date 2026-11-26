# app_pragati.py
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Report No"],
    "CARAT": ["Size"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symm"],
    "FLRC": ["Fls"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table %"],
    "DEPTH": ["Depth %"],
    "LAB": ["Lab"],
    "DISC": ["Disc %"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Table Black'],
    'Natts2':['Side Black'],
    "Stone id":['Ref No']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
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
            # normalize mixed types to string to avoid issues
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


# ---------- Filters (Pragati) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")  # your input already uses ROUND

    # Disc % <= -10
    if "Disc %" in df.columns:
        df["Disc %"] = pd.to_numeric(df["Disc %"], errors="coerce")
        df = df[df["Disc %"] <= -10]

    # Carat/Size range
    df = between(df, "Size", 0.30, 3.99, inclusive="both")

    # Location == IN
    df = equal(df, "Loc", "IN")

    # Attribute lists
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    if "Cut" in df.columns:
        df = df[df["Cut"].isin(cut)]
    if "Polish" in df.columns:
        df = df[df["Polish"].isin(cut)]
    if "Symm" in df.columns:
        df = df[df["Symm"].isin(cut)]
    if "Fls" in df.columns:
        df = df[df["Fls"].isin(flour)]

    # Culet not VSM
    if "Culet" in df.columns:
        df = df[df["Culet"] != "VSM"]

    # Table/Side Black exclusions (treat as strings)
    if "Table Black" in df.columns:
        s = df["Table Black"].astype(str)
        df = df[s != "TB-2"]
    if "Side Black" in df.columns:
        s = df["Side Black"].astype(str)
        df = df[s != "SB-2"]

    # Key To Symbols blocklist
    if "Key To Symbols" in df.columns:
        df = df[~df["Key To Symbols"].isin(key_symbol_block)]

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
def run_pragati_mod() -> None:
    st.header("Pragati")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)  # your sheet uses header at row 0

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
    file_name = f"pragati {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_pragati_mod()
