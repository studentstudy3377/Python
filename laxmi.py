import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Certificate"],
    "CARAT": ["Carats"],
    "SHAPE": ["Shape"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symm"],
    "FLRC": ["Fluo Int"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table(%)"],
    "DEPTH": ["Depth(%)"],
    "LAB": ["Lab"],
    "DISC": ["Disc(%)"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Black Inc'],
    'Natts2':[],
    "Stone id":['Stone ID']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
flour = ["NON", "FNT", "MED"]
black_inc = ["TB2", "SB2", "SB2,TB2"]
key_symbol = [
    "Cavity",
    "Cavity",
    "Chip",
    "Cavity, Chip",
    "Chip, Cavity",
    "Bruise, Crystal",
    "Brown patch of color, Pinpoint",
]
tinge = ["NO", ""]
milky = ["M0", ""]


# ---------- Helpers ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names, drop 'Unnamed' residuals, and coerce mixed object cols to str."""
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in df.columns:
        if df[col].dtype == "object":
            # decode bytes, make types consistent
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore")
                if isinstance(x, (bytes, bytearray))
                else x
            )
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df


def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)


def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Rename known aliases + ensure all expected columns exist
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""
    # Reorder
    return df[list(columns_mapping)]


def numeric(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df


def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df = numeric(df, col)
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Force Disc(%) negative
    if "Disc(%)" in df.columns:
        df["Disc(%)"] = pd.to_numeric(df["Disc(%)"], errors="coerce")
        df["Disc(%)"] = df["Disc(%)"].apply(lambda x: -x if pd.notna(x) and +x >= 0 else x)

    # Remove + / - from Color, Clarity, Cut (if present)
    for c in ["Color", "Clarity", "Cut"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.replace("+", "", regex=False).str.replace("-", "", regex=False)

    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "RBC")
    if "Shape" in df.columns:
        df["Shape"] = df["Shape"].replace("RBC", "ROUND")

    df = between(df, "Carats", 0.30, 3.99, inclusive="both")

    if "Country" in df.columns:
        # Case-insensitive match to 'India'
        df = df[df["Country"].astype(str).str.casefold() == "india"]

    if "Disc(%)" in df.columns:
        df = df[df["Disc(%)"] <= -10]

    # Inclusion/Exclusion attributes
    if "Black Inc" in df.columns:
        df = df[~df["Black Inc"].isin(black_inc)]
    if "Tinge" in df.columns:
        df = df[df["Tinge"].isin(tinge)]
    if "Milky" in df.columns:
        df = df[df["Milky"].isin(milky)]
    if "Luster" in df.columns:
        df = equal(df, "Luster", "EX")
    if "Culet" in df.columns:
        df = equal(df, "Culet", "NON")

    # Quality lists
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]
    # Cut/Polish/Symm all EX/VG
    for c in ["Cut", "Polish", "Symm"]:
        if c in df.columns:
            df = df[df[c].isin(cut)]

    if "Fluo Int" in df.columns:
        df = df[df["Fluo Int"].isin(flour)]
    if "Key To Symbols" in df.columns:
        df = df[~df["Key To Symbols"].isin(key_symbol)]

    return df


def default_value(df: pd.DataFrame) -> pd.DataFrame:
    # Set defaults
    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        if col in df.columns:
            df[col] = 0

    for num_col in ["CARAT", "TABLE", "DEPTH", "DISC", "RATIO", "RAP", "LENGTH", "WIDTH", "HEIGHT"]:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")

    for col in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
        if col in df.columns:
            df[col] = "N/A"

    if "SRNO" in df.columns:
        df["SRNO"] = range(1, len(df) + 1)

    if "NATTS" in df.columns:
        df["NATTS"] = "CENTER - L"

    if "DISC" in df.columns:
        df["DISC"] = df["DISC"].round(2)

    if "SHAPE" in df.columns:
        df["SHAPE"] = "ROUND"

    return df


# ---------- App ----------
def run_laxmi_mod() -> None:
    st.header("Laxmi")

    with st.sidebar:
        upload = st.file_uploader("Select Laxmi File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 5, 1)  # your script had skiprows=5

    if upload is None:
        st.stop()

    try:
        df = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    # Show the raw (sanitized) data first
    st.dataframe(df)

    # Filters and normalization
    df = apply_filters(df)
    df = fix_columns(df)
    df = default_value(df)

    st.dataframe(df)

    # Download CSV
    csv_file = df.to_csv(index=False)  # BOM helps Excel
    file_name = f"laxmi {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_file,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_laxmi_mod()
