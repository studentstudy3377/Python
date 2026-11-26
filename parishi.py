# app_parishi.py
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Lot no."],
    "CARAT": ["CT"],
    "SHAPE": ["SHAPE"],
    "COLOR": ["COL"],
    "CLARITY": ["CLA"],
    "CUT": ["CUT"],
    "POL": ["POL"],
    "SYM": ["SYM"],
    "FLRC": ["FL"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["TABLE"],
    "DEPTH": ["T.DEPTH"],
    "LAB": ["LAB"],
    "DISC": ["DISC"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['BIT'],
    'Natts2':['BIC'],
    "Stone id":[]
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
flour = ["NONE", "FNT", "MED"]  # 'NONE' will be normalized to 'NON' later (as per your script)
key_symbol_block = ["Cavity", "Cavity, Chip", "Cavity, Chip", "Chip, Cavity"]
culet_allow = ["N", "NN", "PO"]
bt_block = ["BT2", "BT3"]
bic_allow = ["BC0", "BC1", "NV"]

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


# ---------- Filters (Parishi) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # LAB == GIA
    df = equal(df, "LAB", "GIA")

    # SHAPE == RD, then RD -> ROUND
    if "SHAPE" in df.columns:
        df = df[df["SHAPE"] == "RD"]
        df["SHAPE"] = df["SHAPE"].astype(str).str.replace("RD", "ROUND", regex=False)

    # 0.30 <= CT <= 3.99
    df = between(df, "CT", 0.30, 3.99, inclusive="both")

    # DISC <= -10
    if "DISC" in df.columns:
        df["DISC"] = pd.to_numeric(df["DISC"], errors="coerce")
        df = df[df["DISC"] <= -10]

    # Attribute lists
    if "COL" in df.columns:
        df = df[df["COL"].isin(colors)]
    if "CLA" in df.columns:
        df = df[df["CLA"].isin(clarity)]
    for c in ["CUT", "POL", "SYM"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]
    if "FL" in df.columns:
        df = df[df["FL"].isin(flour)]
        # Normalize 'NONE' -> 'NON' after the filter (keeps your original behavior)
        df.loc[:, "FL"] = df["FL"].replace({"NONE": "NON"})

    # CULET allowed
    if "CULET" in df.columns:
        df = df[df["CULET"].isin(culet_allow)]

    # Exclude BIT in ['BT2','BT3']
    if "BIT" in df.columns:
        df = df[~df["BIT"].isin(bt_block)]

    # Keep BIC in ['BC0','BC1','NV']
    if "BIC" in df.columns:
        df = df[df["BIC"].isin(bic_allow)]

    # Blocked key symbols
    if "KEY TO SYMBOLS" in df.columns:
        df = df[~df["KEY TO SYMBOLS"].isin(key_symbol_block)]

    # Exclude COMMENTS starting with 'Clarity'
    if "COMMENTS" in df.columns:
        s = df["COMMENTS"].astype(str)
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
def run_parishi_mod() -> None:
    st.header("Parishi")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)  # default 0 as per your script

    if upload is None:
        st.stop()

    try:
        df_raw = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw Input (sanitized)")
    st.dataframe(df_raw,)

    df = apply_filters(df_raw)
    df = fix_columns(df)

    st.subheader("Final Output")
    st.dataframe(df)

    # Download CSV
    file_name = f"parishi {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_parishi_mod()
