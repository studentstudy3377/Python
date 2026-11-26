# app_pramukh.py
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
    "SYM": ["Sy"],
    "FLRC": ["Flo"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Tabl"],
    "DEPTH": ["Depth"],
    "LAB": ["Lab"],
    "DISC": ["Disc%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['T.BLK'],
    'Natts2':['S.BLK'],
    "Stone id":['Stone ID']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
# Input may have Ex/Vg/N/Fnt/Med; we filter on those first, then normalize to EX/VG and NON/FNT/MED
cut_raw = ["Ex", "Vg"]
flour_raw = ["N", "Fnt", "Med"]
key_symbol_block = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

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

# ---------- Filters (Pramukh) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Lab == GIA
    df = equal(df, "Lab", "GIA")

    # Shape == RBC, then normalize to ROUND
    if "Shape" in df.columns:
        df = df[df["Shape"] == "RBC"]
        df["Shape"] = df["Shape"].astype(str).str.replace("RBC", "ROUND", regex=False)

    # Discount <= -10
    if "Disc%" in df.columns:
        df["Disc%"] = pd.to_numeric(df["Disc%"], errors="coerce")
        df = df[df["Disc%"] <= -10]

    # Carat range
    df = between(df, "Carat", 0.30, 3.99, inclusive="both")

    # Attribute lists
    if "Col" in df.columns:
        df = df[df["Col"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]

    # Cut/Pol/Sy initially in Ex/Vg -> filter, then normalize to EX/VG
    for c in ["Cut", "Pol", "Sy"]:
        if c in df.columns:
            df = df[df[c].isin(cut_raw)]
            df.loc[:, c] = df[c].replace({"Ex": "EX", "Vg": "VG"})

    # Fluorescence initially N/Fnt/Med -> filter, then normalize to NON/FNT/MED
    if "Flo" in df.columns:
        df = df[df["Flo"].isin(flour_raw)]
        df.loc[:, "Flo"] = df["Flo"].replace({"N": "NON", "Fnt": "FNT", "Med": "MED"})

    # Exclusions
    if "Lust" in df.columns:
        df = df[df["Lust"] != "VG"]
    if "T.BLK" in df.columns:
        df = df[df["T.BLK"] != "B2"]
    if "T.OP" in df.columns:
        df = df[df["T.OP"] != "HL"]
    if "KeyToSym" in df.columns:
        df = df[~df["KeyToSym"].isin(key_symbol_block)]
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
def run_pramukh_mod() -> None:
    st.header("Pramukh")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)  # default 0 like your script

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
    file_name = f"pramukh {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",     
        )

if __name__ == "__main__":
    run_pramukh_mod()