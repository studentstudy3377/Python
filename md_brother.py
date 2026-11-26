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
    "Stone id":['Stock ID']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
flour = ["NON", "FNT", "MED"]
culet_block = ["Very Small", "VSM"]
black_block = ["SPPB2", "CPPB2"]
key_symbol_block = ["Cavity", "Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

# ---------- Helpers ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Trim column names, drop Unnamed cols, make mixed object cols strings."""
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


def read_file(uploaded_file, skiprows: int = 1) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)


def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    # rename aliases and ensure all target columns present
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""
    # reorder
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

def keep_none_or_empty(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in df.columns:
        return df
    s = df[col].astype(str).str.strip().str.casefold()
    none_like = {"none"}                 # literal NONE (any case)
    empty_like = {"", "nan", "null", "n/a", "na"}  # blank / NA-ish text
    mask = df[col].isna() | s.isin(none_like | empty_like)
    return df[mask]

# ---------- Filters ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Base filters from your script
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    df = equal(df, "Loc", "MUMBAI")

    if "Disc %" in df.columns:
        df["Disc %"] = pd.to_numeric(df["Disc %"], errors="coerce")
        # force everything to negative
        df["Disc %"] = df["Disc %"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        # now filter
        df = df[df["Disc %"] <= -10]

    df = between(df, "Size", 0.30, 3.99, inclusive="both")

    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]

    # Cut/Polish/Symm must be EX/VG
    for c in ["Cut", "Polish", "Symm"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]

    if "Fls" in df.columns:
        df = df[df["Fls"].isin(flour)]

    df = keep_none_or_empty(df, "Shade")
    
    df = keep_none_or_empty(df, "Milky")

    if "Culet" in df.columns:
        df = df[~df["Culet"].isin(culet_block)]

    if "Black" in df.columns:
        df = df[~df["Black"].isin(black_block)]

    if "Key To Symbols" in df.columns:
        df = df[~df["Key To Symbols"].isin(key_symbol_block)]

    if "Comments" in df.columns:
        s = df["Comments"].astype(str)
        df = df[
            ~s.str.startswith("Clarity", na=False)
            | ~s.str.startswith("* This diamond has been", na=False)
        ]

    if "Grading Comments" in df.columns:
        s = df["Grading Comments"].astype(str)
        df = df[s.str.startswith("NO BGM", na=False)]

    return df

def default_value(df: pd.DataFrame) -> pd.DataFrame:
    # Defaults & numeric coercions similar to your earlier apps
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


# ---------- App ----------
def run_md_brothers_mod() -> None:
    st.header("M D Brothers")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 1, 1)  # your script used skiprows=1

    if upload is None:
        st.stop()

    try:
        df = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    # Show raw (sanitized)
    st.dataframe(df)

    # Process
    df = apply_filters(df)
    df = fix_columns(df)
    df = default_value(df)

    st.dataframe(df)

    # Download CSV
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    file_name = f"m d brothers {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_md_brothers_mod()
