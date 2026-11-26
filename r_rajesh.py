# app_r_rajesh.py
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Report No"],
    "CARAT": ["Carat"],
    "SHAPE": ["Shape"],
    "COLOR": ["Col"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Fls"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table %"],
    "DEPTH": ["TD%"],
    "LAB": ["Lab"],
    "DISC": ["Disc %"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['TB'],
    'Natts2':['SB'],
    "Stone id":['Stone ID']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
flour = ["NONE", "FNT", "MED"]
tb_block = ["TBL2", "TBL3"]
sb_block = ["SBL2", "SBL3"]
tw_block = ["CW3"]
sw_block = ["SW3"]
culet_block = ["VERY SMALL"]
milky_allow = ["NONE", ""]
open_table_block = ["MEDIUM TABLE OPEN", "SMALL TABLE OPEN"]
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


def read_file(uploaded_file, skiprows: int = 5) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    df = _sanitize_df(df)
    # Trim rows to non-null 'Col' length (replicates your desktop script)
    if "Col" in df.columns:
        length = len(df["Col"].dropna())
        df = df.iloc[:length]
    return df


def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df


def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df


# ---------- Filters (R Rajesh) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Force negative Disc %
    if "Disc %" in df.columns:
        df["Disc %"] = pd.to_numeric(df["Disc %"], errors="coerce")
        df["Disc %"] = df["Disc %"].apply(lambda x: -x if pd.notna(x) and x >= 0 else x)

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    df = between(df, "Carat", 0.30, 3.99, inclusive="both")
    df = equal(df, "Location", "INDIA")

    if "Disc %" in df.columns:
        df = df[df["Disc %"] <= -10]

    # Attribute filters
    if "Col" in df.columns:
        df = df[df["Col"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    for c in ["Cut", "Pol", "Sym"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]
    if "Fls" in df.columns:
        df = df[df["Fls"].isin(flour)]

    # Additional exclusions/inclusions
    if "Tinge" in df.columns:
        df = df[df["Tinge"].isin(milky_allow)]
    if "TB" in df.columns:
        df = df[~df["TB"].isin(tb_block)]
    if "SB" in df.columns:
        df = df[~df["SB"].isin(sb_block)]
    if "TW" in df.columns:
        df = df[~df["TW"].isin(tw_block)]
    if "SW" in df.columns:
        df = df[~df["SW"].isin(sw_block)]
    if "Culet" in df.columns:
        df = df[~df["Culet"].isin(culet_block)]
    if "Milky" in df.columns:
        df = df[df["Milky"].isin(milky_allow)]
    if "Open On Table" in df.columns:
        df = df[~df["Open On Table"].isin(open_table_block)]
    if "Report Comments" in df.columns:
        s = df["Report Comments"].astype(str)
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
def run_r_rajesh_mod() -> None:
    st.header("R Rajesh")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 5, 1)  # default 5 like your script

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
    file_name = f"r rajesh {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_r_rajesh_mod()
