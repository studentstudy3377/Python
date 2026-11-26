import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Stock No"],
    
    "CARAT": ["Carat","Size"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Flo"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Tab%"],
    "DEPTH": ["TD%"],
    "LAB": ["Lab","Cert"],
    "DISC": ["Rap%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts_1':['BT'],
    'Natts_2':['BC'],
    "Stone id":[]
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
flour = ["NON", "FNT", "MED"]
culet_allow = ["NONE", "POINTED"]
milky_block = ["BC0","BC1"]
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


def read_file(uploaded_file, skiprows: int = 4) -> pd.DataFrame:
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

def keep_empty_like(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in df.columns:
        return df
    s = df[col].astype("string")
    mask = s.isna() | s.str.strip().str.casefold().isin({"", "none", "na", "n/a", "nan"})
    return df[mask]

# ---------- Filters ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")

    # Disc % -> force negative, then keep <= -10
    if "Disc %" in df.columns:
        df["Disc %"] = pd.to_numeric(df["Disc %"], errors="coerce")
        df["Disc %"] = df["Disc %"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        df = df[df["Disc %"] <= -10]

    df = between(df, "Carat", 0.30, 3.99, inclusive="both")

    # Attribute lists
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    for c in ["Cut", "Pol", "Sym"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]
    df['Flo'] = df['Flo'].replace({'F': 'FNT', 'N': 'NON', 'MED': 'MED'})

    if "Flo" in df.columns:
        df = df[df["Flo"].isin(flour)]
    
    if "BT" in df.columns:
        df = df[df["BT"] == "BT0"]
    if "BC" in df.columns:
        df = df[df["BC"] .isin(milky_block)]
    df = keep_empty_like(df, "Shade")
    df = keep_empty_like(df, "Table Open")
    
    if "Comments" in df.columns:
        s = df["Comments"].astype(str)
        df = df[(s == "NO BGM") | s.isna()]
    if "Key To Symbol" in df.columns:
        df = df[~df["Key To Symbol"].isin(key_symbol_block)]
    if "Culet" in df.columns:
        df = df[df["Culet"].isin(culet_allow)]

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
def run_harekrishna_mod() -> None:
    st.header("HAREKRISHNA")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)  # your script used skiprows=4

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
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    file_name = f"harekrishna {pd.Timestamp.now():%d.%m.%Y}.csv"
    
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_harekrishna_mod()
