import streamlit as st
import pandas as pd
import numpy as np

columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Certificate No"],
    "CARAT": ["Carat"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Flour"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table"],
    "DEPTH": ["Depth"],
    "LAB": ["Lab"],
    "DISC": ["Dis%"],
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

ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR = ["NON", "FNT", "MED"]
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names, drop Unnamed residuals, coerce mixed objects to str."""
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore") if isinstance(x, (bytes, bytearray)) else x
            )
            # if mixed types remain, coerce to string
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df


def read_excel(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)


def numeric(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df


def isin(df: pd.DataFrame, col: str, values) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col].isin(values)]
    return df


def notin(df: pd.DataFrame, col: str, values) -> pd.DataFrame:
    if col in df.columns:
        return df[~df[col].isin(values)]
    return df


def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df = numeric(df, col)
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df


def empty_or_na(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in df.columns:
        return df.iloc[0:0]
    s = df[col].astype(str).str.strip()
    mask = df[col].isna() | s.eq("") | s.str.casefold().isin({"nan", "none", "null", "n/a", "na"})
    return df[mask]

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Exact matches
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Location", "MUMBAI")

    # Shape (user file uses 'ROUND')
    df = equal(df, "Shape", "ROUND")

    # Carat range
    df = between(df, "Carat", 0.30, 3.99, inclusive="both")

    # Lists
    df = isin(df, "Color", ALLOWED_COLOR)
    df = isin(df, "Clarity", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Pol", ALLOWED_CUT)
    df = isin(df, "Sym", ALLOWED_CUT)
    df = isin(df, "Flour", ALLOWED_FLOUR)

    # Dis%: force negative and require <= -10
    if "Dis%" in df.columns:
        df["Dis%"] = pd.to_numeric(df["Dis%"], errors="coerce")
        df["Dis%"] = df["Dis%"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        df = df[df["Dis%"] <= -10]

    # Empties only
    if "Milky" in df.columns:
        df = df[df.index.isin(empty_or_na(df, "Milky").index)]
    if "Shade" in df.columns:
        df = df[df.index.isin(empty_or_na(df, "Shade").index)]
    if "SGP Comment" in df.columns:
        # First keep NA/empty rows
        mask_empty = df.index.isin(empty_or_na(df, "SGP Comment").index)
        # Then exclude those starting with 'Clarity' (case-insensitive)
        s = df["SGP Comment"].astype(str).str.strip()
        mask_not_clarity = ~s.str.startswith("Clarity", na=False)
        df = df[mask_empty & mask_not_clarity]

    # Culet == 'NON'
    if "Culet" in df.columns:
        df = equal(df, "Culet", "NON")

    # Key To Symbol: exclude specified
    if "Key To Symbol" in df.columns:
        df = notin(df, "Key To Symbol", EXCLUDED_KEY_SYMBOLS)

    return df

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Rename to our output headers when source headers exist
    rename_map = {src: tgt for tgt, sources in columns_mapping.items() for src in (sources + [tgt])}
    df = df.rename(columns=rename_map)

    # Ensure all expected columns exist
    for col in columns_mapping:
        if col not in df.columns:
            df[col] = np.nan

    # Defaults
    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        df[col] = 0

    if "SRNO" in df.columns:
        df["SRNO"] = range(1, len(df) + 1)

    for col in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
        df[col] = "N/A"

    if "NATTS" in df.columns:
        df["NATTS"] = "CENTER - L"

    # SHAPE forced to ROUND in output
    if "SHAPE" in df.columns:
        df["SHAPE"] = "ROUND"

    # Numeric coercion for usual suspects
    for num_col in ["CARAT", "TABLE", "DEPTH", "DISC", "RATIO", "RAP", "LENGTH", "WIDTH", "HEIGHT"]:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")

    # Round DISC to 2 decimals if present
    if "DISC" in df.columns:
        df["DISC"] = df["DISC"].round(2)

    return df[list(columns_mapping)]


def run_shivan_mod() -> None:
    st.header("Shivan")

    with st.sidebar:
        upload = st.file_uploader("Select Excel file", type=["xlsx", "xls"])
        skip = st.number_input("Skip rows", 0, 10000, 0, 1)

    if not upload:
        st.info("Upload a file to begin.")
        st.stop()

    try:
        df = read_excel(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw data (sanitized)")
    st.dataframe(df)

    # Apply filters and shape output
    df_filtered = apply_filters(df)
    df_out = fix_columns(df_filtered.copy())

    st.subheader("Filtered & Mapped Output")
    st.dataframe(df_out)

    # Download
    file_name = f"shivan list {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_shivan_mod()
