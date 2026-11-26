import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (from your Sunrays script)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR = ["NON", "FNT", "MED"]
EXCLUDED_CULET = ["VSM", "Very Small"]
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Report No"],
    "CARAT": ["Size"],
    "SHAPE": ["Shape"],
    "COLOR": ["Col"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symm"],
    "FLRC": ["Fls"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table %"],
    "DEPTH": ["Depth %"],
    "LAB": ["Lab"],
    "DISC": ["Disc %"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['B I'],
    'Natts2':['T I'],
    "Stone id":['Packet No']
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


def read_excel(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    df = _sanitize_df(df)

    # Coerce Disc % to negative as per your script
    if "Disc %" in df.columns:
        df["Disc %"] = pd.to_numeric(df["Disc %"], errors="coerce")
        df["Disc %"] = df["Disc %"].apply(lambda x: -abs(x) if pd.notna(x) else x)
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
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    if "Disc %" in df.columns:
        df = df[df["Disc %"] <= -10]
    df = between(df, "Size", 0.30, 3.99, inclusive="both")
    df = equal(df, "Location", "IN")

    # Lists
    df = isin(df, "Col", ALLOWED_COLOR)
    df = isin(df, "Clarity", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Polish", ALLOWED_CUT)
    df = isin(df, "Symm", ALLOWED_CUT)

    # Flour list
    df = isin(df, "Fls", ALLOWED_FLOUR)

    # Additional domain rules
    df = equal(df, "Shade", "WHT") if "Shade" in df.columns else df
    if "Milky" in df.columns:
        df = df[df["Milky"] != "M01"]
    if "B I" in df.columns:
        df = df[df["B I"] != "B02"]
    if "T I" in df.columns:
        df = df[df["T I"] != "T02"]

    if "Culet" in df.columns:
        df = notin(df, "Culet", EXCLUDED_CULET)

    # Exclude key symbols
    for key_col in ["Key To Symbols", "Key To Symbol", "Key To Sym"]:
        if key_col in df.columns:
            df = notin(df, key_col, EXCLUDED_KEY_SYMBOLS)

    # Exclude report comments starting with 'Clarity'
    if "Report Comments" in df.columns:
        s = df["Report Comments"].astype(str).str.strip()
        df = df[~s.str.startswith("Clarity")]

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

    # Defaults
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
def run_sunrays_mod() -> None:
    st.header("Sunrays")

    with st.sidebar:
        upload = st.file_uploader("Select Excel file", type=["xlsx", "xls"])
        skip = st.number_input("Skip rows", 0, 10000, 0, 1)

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

    file_name = f"sunrays {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_sunrays_mod()
