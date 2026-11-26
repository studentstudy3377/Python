import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (from your VD Global script)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR = ["NON", "FNT", "MED"]
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Cavity, Chip", "Cavity, Chip", "Chip, Cavity"]
EXCLUDED_H_FLAG = ["H"]  # for BT, BCR, WT, etc. where applied

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["ReportNo"],
    "CARAT": ["CTS"],
    "SHAPE": ["Shape"],
    "COLOR": ["Col"],
    "CLARITY": ["Cla"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Flo"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Tab%"],
    "DEPTH": ["TD%"],
    "LAB": ["LAB"],
    "DISC": ["Disc%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['BT'],
    'Natts2':[],
    "Stone id":['Stock No']
}

# =====================
# Utility helpers
# =====================
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    for col in df.columns:
        if df[col].dtype == "object":
            def _fix(x):
                if isinstance(x, (bytes, bytearray)):
                    try:
                        x = x.decode("utf-8", "ignore")
                    except Exception:
                        x = str(x)
                if isinstance(x, str):
                    return x.strip()
                return x
            df[col] = df[col].map(_fix)

    # Coerce numeric columns used later
    for num in ["Disc%", "CTS", "Tab%", "TD%"]:
        if num in df.columns:
            df[num] = pd.to_numeric(df[num], errors="coerce")

    return df


def read_excel(uploaded_file, skiprows: int = 4) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    df = _sanitize_df(df)

    # Force Disc% negative like your script
    if "Disc%" in df.columns:
        df["Disc%"] = df["Disc%"].apply(lambda x: -abs(x) if pd.notna(x) else x)

    return df


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

# =====================
# Core filter pipeline (mirrors your script)
# =====================
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Base columns
    df = equal(df, "LAB", "GIA")

    # Shape RD then standardize to ROUND
    df = equal(df, "Shape", "RD")
    if "Shape" in df.columns:
        df.loc[:, "Shape"] = df["Shape"].str.replace("RD", "ROUND", regex=False)

    # Ranges and mandatory filters
    df = between(df, "CTS", 0.30, 3.99, inclusive="both")
    df = equal(df, "Location", "India")
    if "Disc%" in df.columns:
        df = df[df["Disc%"] <= -10]

    # Lists
    df = isin(df, "Col", ALLOWED_COLOR)
    df = isin(df, "Cla", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Pol", ALLOWED_CUT)
    df = isin(df, "Sym", ALLOWED_CUT)
    df = isin(df, "Flo", ALLOWED_FLOUR)

    # Empties/NA required
    for blank_col in ["Color Shade", "MILKY"]:
        if blank_col in df.columns:
            df = df[df[blank_col].isna()]

    # Culet not VSM
    if "CULET" in df.columns:
        df = df[df["CULET"] != "VSM"]

    # Exclude 'H' flags
    for col in ["BT", "BCR", "WT"]:
        if col in df.columns:
            df = notin(df, col, EXCLUDED_H_FLAG)

    # Key to Symbol exclusions
    for key_col in ["Key To Symbol", "Key To Symbols", "Key To Sym"]:
        if key_col in df.columns:
            df = notin(df, key_col, EXCLUDED_KEY_SYMBOLS)

    return df


# =====================
# Output shaping
# =====================
def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {src: tgt for tgt, sources in COLUMNS_MAPPING.items() for src in (sources + [tgt])}
    # Trim headers before rename to catch space variants
    df = df.rename(columns={c: c.strip() for c in df.columns})
    df = df.rename(columns=rename_map)

    # Ensure all expected columns exist
    for col in COLUMNS_MAPPING:
        if col not in df.columns:
            df[col] = np.nan

    # Defaults & formatting
    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        df[col] = 0

    df["SRNO"] = range(1, len(df) + 1)
    for col in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
        df[col] = "N/A"
    df["NATTS"] = "CENTER - L"
    df["SHAPE"] = "ROUND"

    # Numeric coercion & rounding
    for num_col in ["CARAT", "TABLE", "DEPTH", "DISC", "RATIO", "RAP", "LENGTH", "WIDTH", "HEIGHT"]:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")
    if "DISC" in df.columns:
        df["DISC"] = df["DISC"].round(2)

    return df[list(COLUMNS_MAPPING)]

def _find_sample(filename: str):
    """Locate a file inside a 'Sample Files' folder, searching upward from this
    script. Works on Linux (Streamlit Cloud) and Windows regardless of the
    current working directory or where this page sits in a multipage app."""
    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        candidate = parent / "Sample Files" / filename
        if candidate.exists():
            return candidate
    return None

# =====================
# App
# =====================
def run_vd_global_mod() -> None:
    st.header("VD Global")

    sample_path = _find_sample("VD Global.xlsx")
    
    with st.sidebar:
    
        if sample_path is not None:
            with open(sample_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            href = (
                f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" '
                f'download="{sample_path.name}">Sample File</a>'
            )
            st.markdown(href, unsafe_allow_html=True)
        else:
            st.warning("Sample file not found on server.")
            
        upload = st.file_uploader("Select Excel file", type=["xlsx", "xls"])
        skip = st.number_input("Skip rows", 0, 10000, 4, 1)

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

    df_filtered = apply_filters(df)
    df_out = fix_columns(df_filtered.copy())

    st.subheader("Filtered & Mapped Output")
    st.dataframe(df_out)

    file_name = f"vd global {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_vd_global_mod()
