import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (from your Starrays script)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR_SRC = ["N", "F", "M"]  # normalized below
FLOUR_NORMALIZE = {"F": "FNT", "M": "MED", "N": "NON"}
EXCLUDED_BTI = ["N2", "N3"]
EXCLUDED_LIGHT_GRADE = ["L1", "L2"]
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]
EXCLUDED_COMMENT_PREFIXES = [
    "Clarity grade is based on internal graining that is not shown.",
    "Clarity grade is based on clouds that are not shown.",
]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Cert.No"],
    "CARAT": ["Cts."],
    "SHAPE": ["Shape"],
    "COLOR": ["Col."],
    "CLARITY": ["Cla."],
    "CUT": ["Cut"],
    "POL": ["Pol."],
    "SYM": ["Sym."],
    "FLRC": ["FL"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table %"],
    "DEPTH": ["Depth %"],
    "LAB": ["Cert."],
    "DISC": ["Rap.%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['BIC'],
    'Natts2':['BIT'],
    "Stone id":[]
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
    return df


def read_excel(uploaded_file, skiprows: int = 3) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    df = _sanitize_df(df)
    if "Rap.%" in df.columns:
        df["Rap.%"] = pd.to_numeric(df["Rap.%"], errors="coerce")
        df["Rap.%"] = df["Rap.%"].apply(lambda x: -abs(x) if pd.notna(x) else x)
    if "Cla." in df.columns:
        df["Cla."] = df["Cla."].astype(str).str.replace("+", "", regex=False).str.strip()
    if "Col." in df.columns:
        df["Col."] = df["Col."].astype(str).str.replace("+", "", regex=False).str.strip()
    if "Light Grade" in df.columns:
        df["Light Grade"] = df["Light Grade"].astype(str).str.strip()
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


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = equal(df, "Cert.", "GIA")
    df = equal(df, "Shape", "BR")
    df = between(df, "Cts.", 0.30, 3.99)
    df = equal(df, "Location", "India")

    if "Rap.%" in df.columns:
        df = df[df["Rap.%"] <= -10]

    # optional filters – comment out if too strict
    if "Col.Shade" in df.columns:
        df = df[df["Col.Shade"].isna()]
    if "Col.Desc." in df.columns:
        df = df[df["Col.Desc."].isna()]

    df = isin(df, "Col.", ALLOWED_COLOR)
    df = isin(df, "Cla.", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Pol.", ALLOWED_CUT)
    df = isin(df, "Sym.", ALLOWED_CUT)
    df = isin(df, "FL", ALLOWED_FLOUR_SRC)
    if "FL" in df.columns:
        df["FL"] = df["FL"].replace(FLOUR_NORMALIZE)
    for col in ["BIT", "BIC"]:
        if col in df.columns:
            df = notin(df, col, EXCLUDED_BTI)
    if "Light Grade" in df.columns:
        df = notin(df, "Light Grade", EXCLUDED_LIGHT_GRADE)
    if "Luster" in df.columns:
        df = df[df["Luster"].isin(["EX", ""]) | df["Luster"].isna()]
    for key_col in ["KeyToSymbols", "Key To Symbols", "Key To Symbol", "Key To Sym"]:
        if key_col in df.columns:
            df = notin(df, key_col, EXCLUDED_KEY_SYMBOLS)
    if "Comments" in df.columns:
        s = df["Comments"].astype(str).str.strip()
        mask = pd.Series(True, index=df.index)
        for pref in EXCLUDED_COMMENT_PREFIXES:
            mask &= ~s.str.startswith(pref)
        df = df[mask]
    return df



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

def run_starrays_mod() -> None:
    st.header("Starrays")

    sample_path = _find_sample("Star Rays.xlsx")
    
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
        skip = st.number_input("Skip rows", 0, 10000, 3, 1)

    if not upload:
        st.info("Upload a file to begin.")
        st.stop()

    try:
        df = read_excel(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw data")
    st.dataframe(df)
    df_filtered = apply_filters(df)
    df_out = fix_columns(df_filtered.copy())
  
    st.subheader("Filtered & Mapped Output")
    st.dataframe(df_out)

    file_name = f"starrays {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_starrays_mod()
