import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (derived from your spec)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR_SRC = ["NONE", "FAINT", "MEDIUM"]
FLOUR_NORMALIZE = {"FAINT": "FNT", "MEDIUM": "MED", "NONE": "NON"}
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Report #"],
    "CARAT": ["Weight"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symmetry"],
    "FLRC": ["Fl.Int."],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table%"],
    "DEPTH": ["Depth%"],
    "LAB": ["Lab"],
    "DISC": ["Disc.", "Disc. "],  # tolerate trailing space variant
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':[],
    'Natts2':[],
    "Stone id":[]
}

# =====================
# Utility helpers
# =====================
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Trim headers, drop Unnamed, unify whitespace, coerce byte objects, etc."""
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    # strip whitespace within object cells and decode bytes
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


def empty_or_na_mask(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    return series.isna() | s.eq("") | s.str.casefold().isin({"nan", "none", "null", "n/a", "na"})


# =====================
# Core filter pipeline (as per your script)
# =====================
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Exact matches
    df = equal(df, "Lab", "GIA")
    df = equal(df, "City", "MUMBAI")
    df = equal(df, "Shape", "ROUND")

    # Weight 0.30 to 3.99
    df = between(df, "Weight", 0.30, 3.99, inclusive="both")

    # Lists
    df = isin(df, "Color", ALLOWED_COLOR)
    df = isin(df, "Clarity", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Polish", ALLOWED_CUT)
    df = isin(df, "Symmetry", ALLOWED_CUT)

    # Fluorescence intake list, then normalize to NON/FNT/MED
    df = isin(df, "Fl.Int.", ALLOWED_FLOUR_SRC)
    if "Fl.Int." in df.columns:
        df["Fl.Int."] = df["Fl.Int."].replace(FLOUR_NORMALIZE)

    # Disc.: force negative and require <= -10
    disc_cols = [c for c in df.columns if c.strip().startswith("Disc.")]
    if disc_cols:
        dcol = disc_cols[0]
        df[dcol] = pd.to_numeric(df[dcol], errors="coerce")
        df[dcol] = df[dcol].apply(lambda x: -abs(x) if pd.notna(x) else x)
        df = df[df[dcol] <= -10]

    # Key to Symbols: exclude specified
    if "Key to Symbols" in df.columns:
        df = notin(df, "Key to Symbols", EXCLUDED_KEY_SYMBOLS)

    # Cert comment: keep rows starting with NO BGM / NO BMG (allow leading space)
    if "Cert comment" in df.columns:
        s = df["Cert comment"].astype(str).str.strip()
        mask = s.str.startswith("NO BGM") | s.str.startswith("NO BMG")
        df = df[mask | df["Cert comment"].isna()]

    return df


# =====================
# Output shaping
# =====================
def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Build a rename map allowing both exact and tolerant headers
    rename_map = {}
    for tgt, sources in COLUMNS_MAPPING.items():
        for src in sources:
            rename_map[src.strip()] = tgt
            rename_map[src] = tgt  # also keep original in case of trailing spaces
        # also allow same-named target from source
        rename_map[tgt] = tgt

    # First, create a trimmed header version to catch variants
    trimmed = {c: c.strip() for c in df.columns}
    df = df.rename(columns=trimmed)
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

    # Force SHAPE in output
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
def run_shyam_mod() -> None:
    st.header("Shyam Corporation")

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

    # Filters & mapping
    df_filtered = apply_filters(df)
    df_out = fix_columns(df_filtered.copy())

    st.subheader("Filtered & Mapped Output")
    st.dataframe(df_out)

    # Download button
    file_name = f"shyam coperation {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_shyam_mod()
