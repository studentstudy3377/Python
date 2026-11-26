import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (from your Unique script)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR = ["NON", "FNT", "MED"]
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Report No"],
    "CARAT": ["Carat"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut ", "Cut"],
    "POL": [" Pol", "Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Fls"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table%"],
    "DEPTH": ["Depth%"],
    "LAB": ["Lab"],
    "DISC": [" Disc %", "Disc %"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['TB'],
    'Natts2':['SB'],
    "Stone id":['Packet No']
}

# =====================
# Utility helpers
# =====================
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    # Trim headers and drop Excel "Unnamed"
    df.columns = [str(c) for c in df.columns]
    trimmed = {c: c.strip() for c in df.columns}
    df = df.rename(columns=trimmed)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    # Clean objects
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
    return df


def read_excel(uploaded_file, skiprows: int = 4) -> pd.DataFrame:
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

# =====================
# Core filter pipeline
# =====================
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")

    # Disc % <= -10
    if "Disc %" in df.columns:
        df = numeric(df, "Disc %")
        df = df[df["Disc %"] <= -10]
    if "Disc%" in df.columns:  # tolerate alternate header
        df = numeric(df, "Disc%")
        df = df[df["Disc%"] <= -10]
    if "Disc" in df.columns:
        df = numeric(df, "Disc")
        df = df[df["Disc"] <= -10]

    # Carat range
    df = between(df, "Carat", 0.30, 3.99, inclusive="both")

    # Lists
    df = isin(df, "Color", ALLOWED_COLOR)
    df = isin(df, "Clarity", ALLOWED_CLARITY)
    df = isin(df, "Cut ", ALLOWED_CUT) if "Cut " in df.columns else isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, " Pol", ALLOWED_CUT) if " Pol" in df.columns else isin(df, "Pol", ALLOWED_CUT)
    df = isin(df, "Sym", ALLOWED_CUT)
    df = isin(df, "Fls", ALLOWED_FLOUR)

    # TI must be NON
    if "TI" in df.columns:
        df = equal(df, "TI", "NON")

    # Exclude Key To Symbols
    for key_col in ["Key To Symbols", "Key To Symbol", "Key To Sym"]:
        if key_col in df.columns:
            df = notin(df, key_col, EXCLUDED_KEY_SYMBOLS)

    return df


# =====================
# Output shaping
# =====================
def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {src: tgt for tgt, sources in COLUMNS_MAPPING.items() for src in (sources + [tgt])}
    df = df.rename(columns=rename_map)

    for col in COLUMNS_MAPPING:
        if col not in df.columns:
            df[col] = np.nan

    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        df[col] = 0

    df["SRNO"] = range(1, len(df) + 1)

    for col in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
        df[col] = "N/A"

    df["NATTS"] = "CENTER - L"
    df["SHAPE"] = "ROUND"

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
def run_unique_mod() -> None:
    st.header("Unique")

    sample_path = _find_sample("Unique.xlsx")
    
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

    file_name = f"unique {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_unique_mod()
