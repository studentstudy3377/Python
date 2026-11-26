import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (from your Starlight script)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR = ["NONE", "FNT", "MB"]  # will normalize below
FLOUR_NORMALIZE = {"MB": "MED", "NONE": "NON"}  # FNT stays FNT
EXCLUDED_LUSTER = ["L1", "L2", "L3"]
ALLOWED_CULET = ["NON", "POINTED", "NONE", "N", "PO"]
EXCLUDED_BLACK_TABLE = ["#N/A", "TN3", "TN4", "TN5", "TN6"]
ALLOWED_TABLE_INCLUSION = ["TW0", "TW1", "TW2"]
ALLOWED_SIDE_BLACK = ["CN0", "CN1"]  # or NA allowed
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Cert.No."],
    "CARAT": ["Carat"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symmetry"],
    "FLRC": ["FL"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table"],
    "DEPTH": ["T.Dep."],
    "LAB": ["Lab"],
    "DISC": ["Back"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Black In Table'],
    'Natts2':['Side Black'],
    "Stone id":['StockID']
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
    # Lab
    df = equal(df, "Lab", "GIA")

    # Shape: RD -> keep, then standardize to ROUND
    df = equal(df, "Shape", "RD")
    if "Shape" in df.columns:
        df.loc[:, "Shape"] = df["Shape"].replace("RD", "ROUND")

    # Discount <= -10 (Back column)
    if "Back" in df.columns:
        df["Back"] = pd.to_numeric(df["Back"], errors="coerce")
        df = df[df["Back"] <= -10]

    # Carat range
    df = between(df, "Carat", 0.30, 3.99, inclusive="both")

    # Lists
    df = isin(df, "Color", ALLOWED_COLOR)
    df = isin(df, "Clarity", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Polish", ALLOWED_CUT)
    df = isin(df, "Symmetry", ALLOWED_CUT)

    # FL list and normalize
    df = isin(df, "FL", ALLOWED_FLOUR)
    if "FL" in df.columns:
        df["FL"] = df["FL"].replace(FLOUR_NORMALIZE)

    # Luster exclusions
    if "Luster" in df.columns:
        df = notin(df, "Luster", EXCLUDED_LUSTER)

    # Culet allowed list
    if "Culet" in df.columns:
        df = isin(df, "Culet", ALLOWED_CULET)

    # Black In Table exclusions (treat '#N/A' as string value if present)
    if "Black In Table" in df.columns:
        df = notin(df, "Black In Table", EXCLUDED_BLACK_TABLE)

    # Table Inclusion allowed or NA
    if "Table Inclusion" in df.columns:
        df = df[df["Table Inclusion"].isin(ALLOWED_TABLE_INCLUSION) | df["Table Inclusion"].isna()]

    # Shade NOBGM or NA
    if "Shade" in df.columns:
        df = df[(df["Shade"] == "NOBGM") | (df["Shade"].isna())]

    # Brown / Milky NONE or NA
    for col in ["Brown", "Milky"]:
        if col in df.columns:
            df = df[(df[col] == "NONE") | (df[col].isna())]

    # Side Black allowed or NA
    if "Side Black" in df.columns:
        df = df[df["Side Black"].isin(ALLOWED_SIDE_BLACK) | df["Side Black"].isna()]

    # Exclude Key To Sym
    for key_col in ["Key To Sym", "Key To Symbol", "Key To Symbols"]:
        if key_col in df.columns:
            df = notin(df, key_col, EXCLUDED_KEY_SYMBOLS)

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

    # Force final output SHAPE to ROUND
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
def run_starlight_mod() -> None:
    st.header("Starlight Diamonds")

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

    file_name = f"starlight diamonds {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_starlight_mod()
