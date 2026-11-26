import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (from your Veer Gems script)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "3EX", "VG"]
ALLOWED_FLOUR_SRC = ["NON", "FNT", "MED"]  
FLOUR_NORMALIZE = {"FN": "FNT", "NN": "NON"}
EXCLUDED_KEY_SYMBOLS = ["Chip", "Cavity, Chip", "Chip, Cavity", "Cavity"]
EXCLUDED_VALS_23 = [2, 3]
EXCLUDED_VALS_3 = [3]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Cert#"],
    "CARAT": ["Cts"],
    "SHAPE": ["Shape"],
    "COLOR": ["Col"],
    "CLARITY": ["Clrt"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Flr"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Tbl"],
    "DEPTH": ["T.Dpt"],
    "LAB": ["Cert"],
    "DISC": ["Disc"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['CTB'],
    'Natts2':['STB'],
    "Stone id":['Ref#']
}

# =====================
# Utility helpers
# =====================
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    # Trim headers, drop Excel Unnamed
    df.columns = [str(c).strip() for c in df.columns]
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

    # Numeric coercions commonly used
    for num in ["Disc", "Cts", "Tbl", "T.Dpt"]:
        if num in df.columns:
            df[num] = pd.to_numeric(df[num], errors="coerce")

    return df


def read_excel(uploaded_file, skiprows: int = 3) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    df = _sanitize_df(df)

    # Force Disc negative like your script
    if "Disc" in df.columns:
        df["Disc"] = df["Disc"].apply(lambda x: -abs(x) if pd.notna(x) else x)

    # Normalize fluorescence
    if "Flr" in df.columns:
        df["Flr"] = df["Flr"].replace(FLOUR_NORMALIZE)

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
# Core filter pipeline
# =====================
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Base filters
    df = equal(df, "Cert", "GIA")
    df = equal(df, "Loc", "Mum")
    df = equal(df, "Shape", "ROUND")

    if "Disc" in df.columns:
        df = df[df["Disc"] <= -10]

    df = between(df, "Cts", 0.30, 3.99, inclusive="both")

    # # Lists
    df = isin(df, "Col", ALLOWED_COLOR)
    df = isin(df, "Clrt", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Pol", ALLOWED_CUT)
    df = isin(df, "Sym", ALLOWED_CUT)

    # # Fluorescence intake and normalization
    df = isin(df, "Flr", ALLOWED_FLOUR_SRC)

    # # # Exclude table/side/center black/white flags
    for col in ["CTB", "STB", "CRB"]:
        if col in df.columns:
            df = notin(df, col, EXCLUDED_VALS_23)
    for col in ["CTW", "STW", "CRW"]:
        if col in df.columns:
            df = notin(df, col, EXCLUDED_VALS_3)

    # Culet must not be OTHER
    if "Culet" in df.columns:
        df = df[df["Culet"] != "OTHER"]

    # Exclude key symbols
    for key_col in ["Key To Symbol", "Key To Symbols", "Key To Sym"]:
        if key_col in df.columns:
            df = notin(df, key_col, EXCLUDED_KEY_SYMBOLS)

    return df


# =====================
# Output shaping
# =====================
def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {src: tgt for tgt, sources in COLUMNS_MAPPING.items() for src in (sources + [tgt])}
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


# =====================
# App
# =====================
def run_veer_gems_mod() -> None:
    st.header("Veer Gems")

    with st.sidebar:
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

    st.subheader("Raw data (sanitized)")
    st.dataframe(df)

    df_filtered = apply_filters(df)
    df_out = fix_columns(df_filtered.copy())

    st.subheader("Filtered & Mapped Output")
    st.dataframe(df_out)

    file_name = f"veer gems {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_veer_gems_mod()
