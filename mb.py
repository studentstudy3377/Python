import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Cert No."],
    "CARAT": ["Weight"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Fluo"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table %"],
    "DEPTH": ["Depth %"],
    "LAB": ["Cert Lab"],
    "DISC": ["Disc"],
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
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2", "SI2"]
cut = ["EX", "VG"]
flour = ["NON", "FNT", "MED"]
key_symbol_block = ["Cavity", "Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]
table_black_allow = ["NN", "BT1", "BT1+"]
side_black_allow = ["NN", "BC1", "BC1+"]
shade_allow = ["NN"]
open_block = ["OT1"]
cullet_block = ["VERY SMALL"]  # (kept commented, like your script)

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


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Base filters
    df = equal(df, "Cert Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    df = between(df, "Weight", 0.30, 3.99, inclusive="both")

    # Disc: force all values negative, then keep <= -10
    if "Disc" in df.columns:
        df["Disc"] = pd.to_numeric(df["Disc"], errors="coerce")
        df["Disc"] = df["Disc"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        df = df[df["Disc"] <= -10]

    # Attribute filters
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    for c in ["Cut", "Pol", "Sym"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]
    if "Fluo" in df.columns:
        df = df[df["Fluo"].isin(flour)]
    if "Shade" in df.columns:
        df = df[df["Shade"].isin(shade_allow)]
    if "Table Black" in df.columns:
        df = df[df["Table Black"].isin(table_black_allow)]
    if "Side Black" in df.columns:
        df = df[df["Side Black"].isin(side_black_allow)]
    if "Open" in df.columns:
        df = df[~df["Open"].isin(open_block)]
    if "Culet" in df.columns:
        df = df[~df["Culet"].isin(cullet_block)]

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

# ---------- App ----------
def run_mb_mod() -> None:
    st.header("MB")

    sample_path = _find_sample("MB.xlsx")
    
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
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)

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
    file_name = f"mb {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",   
        )

if __name__ == "__main__":
    run_mb_mod()
