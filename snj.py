import streamlit as st
import pandas as pd
import numpy as np

ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR = ["NON", "FNT", "MED"]
EXCLUDED_TABLE_BLACK = ["TN2", "TN3"]
EXCLUDED_SIDE_BLACK = ["SN2", "SN3"]
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Certificate No"],
    "CARAT": ["Size"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Sym"],
    "FLRC": ["Flo"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table"],
    "DEPTH": ["Depth"],
    "LAB": ["Lab"],
    "DISC": ["Final Disc%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Table Black'],
    'Natts2':['Side Black'],
    "Stone id":['Stone Id']
}

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


def read_excel(uploaded_file, skiprows: int = 1) -> pd.DataFrame:
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


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Lab, Shape, Location
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    df = equal(df, "Location", "INDIA")

    # Final Disc% <= -10
    if "Final Disc%" in df.columns:
        df["Final Disc%"] = pd.to_numeric(df["Final Disc%"], errors="coerce")
        df = df[df["Final Disc%"] <= -10]

    # Size 0.30 to 3.99
    df = between(df, "Size", 0.30, 3.99, inclusive="both")

    # Lists
    df = isin(df, "Color", ALLOWED_COLOR)
    df = isin(df, "Clarity", ALLOWED_CLARITY)
    df = isin(df, "Cut", ALLOWED_CUT)
    df = isin(df, "Polish", ALLOWED_CUT)
    df = isin(df, "Sym", ALLOWED_CUT)

    # Exclusions
    if "Table Black" in df.columns:
        df = notin(df, "Table Black", EXCLUDED_TABLE_BLACK)
    if "Side Black" in df.columns:
        df = notin(df, "Side Black", EXCLUDED_SIDE_BLACK)
    if "Luster" in df.columns:
        df = df[df["Luster"] != "VG"]
    if "Flo" in df.columns:
        df = isin(df, "Flo", ALLOWED_FLOUR)

    # Table Open: keep rows with 'N' or NA; handle potential newline header variant
    col_opts = [c for c in df.columns if c.replace("\n", " ").strip() == "Table  Open"]
    if "Table \nOpen" in df.columns:
        df = df[(df["Table \nOpen"] == "N") | (df["Table \nOpen"].isna())]
    elif "Table Open" in df.columns:
        df = df[(df["Table Open"] == "N") | (df["Table Open"].isna())]

    # Culet Size: NONE or NA
    if "Culet Size" in df.columns:
        df = df[(df["Culet Size"] == "NONE") | (df["Culet Size"].isna())]

    # Key To Symbol exclusion
    for key_col in ["Key To Symbol", "Key To Symbols"]:
        if key_col in df.columns:
            df = notin(df, key_col, EXCLUDED_KEY_SYMBOLS)

    return df


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
def run_snj_mod() -> None:
    st.header("SNJ")

    with st.sidebar:
        upload = st.file_uploader("Select Excel file", type=["xlsx", "xls"])
        skip = st.number_input("Skip rows", 0, 10000, 5, 1)

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

    file_name = f"snj list {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_snj_mod()
