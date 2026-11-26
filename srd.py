import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (from your SRD script)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR_SRC = ["N", "F", "M"]  # will be normalized to NON/FNT/MED
FLOUR_NORMALIZE = {"F": "FNT", "M": "MED", "N": "NON"}
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Cert No"],
    "CARAT": ["Cts"],
    "SHAPE": ["Shape"],
    "COLOR": ["Col"],
    "CLARITY": ["Cla"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["FL"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Tab%"],
    "DEPTH": ["TD%"],
    "LAB": ["Lab"],
    "DISC": ["Rap %"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Natts'],
    'Natts2':[],
    "Stone id":['Stone No']
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
    # Rap %: force negative
    if "Rap %" in df.columns:
        df["Rap %"] = pd.to_numeric(df["Rap %"], errors="coerce")
        df["Rap %"] = df["Rap %"].apply(lambda x: -abs(x) if pd.notna(x) else x)

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    if "Rap %" in df.columns:
        df = df[df["Rap %"] <= -10]
    df = between(df, "Cts", 0.30, 3.99, inclusive="both")
    df = equal(df, "Loc", "MUMBAI")

    # Lists
    df = isin(df, "Col", ALLOWED_COLOR)
    df = isin(df, "Cla", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Pol", ALLOWED_CUT)
    df = isin(df, "Sym", ALLOWED_CUT)

    # Fluorescence normalize
    df = isin(df, "FL", ALLOWED_FLOUR_SRC)
    if "FL" in df.columns:
        df["FL"] = df["FL"].replace(FLOUR_NORMALIZE)

    # Additional flags
    df = equal(df, "CS", "WH") if "CS" in df.columns else df
    if "BT" in df.columns:
        df = df[df["BT"] != "BT2"]
    if "BC" in df.columns:
        df = df[df["BC"] != "BC2"]
    df = equal(df, "TO", "TO0") if "TO" in df.columns else df
    if "FlShd" in df.columns:
        df = df[df["FlShd"] != "Yellow"]
    df = equal(df, "Milky", "M0") if "Milky" in df.columns else df
    df = equal(df, "Luster", "EX") if "Luster" in df.columns else df

    # Fancy Color Description must be NA
    if "Fancy Color Description" in df.columns:
        df = df[df["Fancy Color Description"].isna()]

    # Exclude Key To Sym
    if "Key To Sym" in df.columns:
        df = notin(df, "Key To Sym", EXCLUDED_KEY_SYMBOLS)

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


# =====================
# App
# =====================
def run_srd_mod() -> None:
    st.header("SRD ")

    with st.sidebar:
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

    file_name = f"srd {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_srd_mod()
