import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Report"],
    "CARAT": ["Cts"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Flor."],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table"],
    "DEPTH": ["Depth"],
    "LAB": ["Lab"],
    "DISC": ["Dis"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':[],
    'Natts2':[],
    "Stone id":[]
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


# ---------- Filters (Pansurya) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Lab == GIA
    df = equal(df, "Lab", "GIA")

    # Shape == RD, then replace to ROUND (so SHAPE shows ROUND later)
    if "Shape" in df.columns:
        df = df[df["Shape"] == "RD"]
        df["Shape"] = df["Shape"].astype(str).str.replace("RD", "ROUND", regex=False)

    # Location == MUMBAI
    df = equal(df, "Location", "MUMBAI")

    # Discount <= -10
    if "Dis" in df.columns:
        df["Dis"] = pd.to_numeric(df["Dis"], errors="coerce")
        df = df[df["Dis"] <= -10]

    # Carat range (0.30–3.99)
    df = between(df, "Cts", 0.30, 3.99, inclusive="both")

    # Attribute lists
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    for c in ["Cut", "Pol", "Sym"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]
    if "Flor." in df.columns:
        df = df[df["Flor."].isin(flour)]

    # Block Key To Sym
    if "Key To Sym" in df.columns:
        df = df[~df["Key To Sym"].isin(key_symbol_block)]

    # Exclude Comment starting with 'Clarity'
    if "Comment" in df.columns:
        s = df["Comment"].astype(str)
        df = df[~s.str.startswith("Clarity", na=False)]

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
def run_pansurya_mod() -> None:
    st.header("Pansurya Impex")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)  # default 0 for your sheets

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
    file_name = f"pansurya impex {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_pansurya_mod()