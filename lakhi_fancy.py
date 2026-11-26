import re
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Report #"],
    "CARAT": ["Weight"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symmetry"],
    "FLRC": ["Fluorescence Intensity"],
    "LENGTH": ['M1'],
    "WIDTH": ['M2'],
    "HEIGHT": ['M3'],
    "TABLE": ["Table %"],
    "DEPTH": ["Depth %"],
    "LAB": ["Lab"],
    "DISC": ["Rapnet  Discount %"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": ['Ratio'],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Black Inclusion   '],
    'Natts2':[],
    "Stone id":['Stock #']
}

# ---------- Lists (criteria) ---------- 
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
shape = ["PRINCESS","EMERALD","OVAL","PEAR"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
flour = ["NON", "FNT", "MED"]

# ---------- Helpers ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names, drop 'Unnamed' residuals, and coerce mixed object cols to str."""
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore")
                if isinstance(x, (bytes, bytearray))
                else x
            )
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


def filter_none(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in df.columns:
        return df.iloc[0:0]
    s = df[col].astype(str).str.strip()
    mask_missing = df[col].isna() | s.eq("") | s.str.casefold().isin({"nan", "none", "null", "n/a", "na"})
    return df[mask_missing]

def get_col(df: pd.DataFrame, target: str):
    """Return real column name matching target (case/space-insensitive), or None."""
    norm = lambda s: re.sub(r"\s+", " ", str(s).strip()).casefold()
    t = norm(target)
    for c in df.columns:
        if norm(c) == t:
            return c
    return None

# ---------- Core Filters (Lakhi) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Rapnet Discount handling: x*100 → force negative → format 2 decimals → back to numeric
    if "Rapnet  Discount %" in df.columns:
        col = pd.to_numeric(df["Rapnet  Discount %"], errors="coerce")
        col = col * 100  # as in your original
        col = col.apply(lambda x: f"{x if x < 0 else -x:.2f}")
        df["Rapnet  Discount %"] = pd.to_numeric(col, errors="coerce")

    # Base filters
    df = equal(df, "Lab", "GIA")
    if "Shape" in df.columns:
        df["Shape"] = df["Shape"].replace({"Square Modified Brilliant":"PRINCESS","Emerald Cut":"EMERALD","Oval Brilliant":"OVAL","Pear Brilliant":"PEAR"})
    df = between(df, "Weight", 0.30, 3.99, inclusive="both")
    # if "Country" in df.columns:
    #     df = equal(df, "Country", "INDIA")

    if "Rapnet  Discount %" in df.columns:
        df = df[df["Rapnet  Discount %"] <= -10]

    if "Shape" in df.columns:
        df = df[df["Shape"].isin(shape)]
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]

    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    if "Fluorescence Intensity" in df.columns:
        df = df[df["Fluorescence Intensity"].isin(flour)]


    # Cut/Polish/Symmetry all EX/VG
    for c in ["Polish", "Symmetry"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]

    # Black Inclusion (case/space tolerant)
    col_bi = get_col(df, "Black Inclusion")
    if col_bi:
        s = df[col_bi].astype(str).str.strip().str.casefold()
        df = df[~s.str.startswith("n3,", na=False)]
        
    if "Cert comment" in df.columns:
        bad = "Clarity grade is based on internal graining that is not shown."
        df = df[~df["Cert comment"].astype(str).str.startswith(bad, na=False)]

    return df


def default_value(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["RAP"]:
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
    if "RATIO" in df.columns:
        df["RATIO"] = df["RATIO"].abs().round(2)

    if "NATTS" in df.columns:
        df["NATTS"] = "CENTER - L"
    if "CUT" in df.columns:
        df["CUT"] = "VG"

    if "DISC" in df.columns:
        df["DISC"] = df["DISC"].round(2)

    return df

# ---------- App ----------
def run_lakhi_fancy_mod() -> None:
    st.header("Lakhi")

    with st.sidebar:
        upload = st.file_uploader("Select Lakhi File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)

    if upload is None:
        st.stop()

    with st.spinner("Processing..."):

        try:
            # Initialize progress bar
            progress_bar = st.progress(0)

            df = read_file(upload, skiprows=skip)
        except Exception as e:
            st.error(f"Failed to read file: {e}")
            st.stop()

        # Show raw (sanitized)
        progress_bar.progress(20)
        st.dataframe(df)

        # Process
        progress_bar.progress(40)
        df = apply_filters(df)
        progress_bar.progress(60)
        df = fix_columns(df)
        progress_bar.progress(80)
        df = default_value(df)
        progress_bar.progress(100)
        st.header("Output:")
        st.dataframe(df)
        
        # Download
        csv_file = df.to_csv(index=False)
        file_name = f"lakhi fancy{pd.Timestamp.now():%d.%m.%Y}.csv"

        with st.sidebar:
            st.download_button(
                label="Generate CSV",
                data=csv_file,
                file_name=file_name,
                mime="csv" 
            )
if __name__ == "__main__":
    run_lakhi_fancy_mod()
