import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np

columns_mapping = {
    'SRNO': [],
    'LOTNO': ['Report No'],
    
    'CARAT': ['Ct.'],
    'SHAPE': ['Shape'],
    'COLOR': ['Col.'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Pol.'],
    'SYM': ['Sym.'],
    'FLRC': ['Fl.'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table %'],
    'DEPTH': ['Depth %'],
    'LAB': ['Lab'],
    'DISC': ['Disc. (%)'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['Centre Black'],
    'Natts2':['Crown Black'],
    "Stone id":['Packet No']
}

colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX','VG']
flour = ['NON', 'FNT', 'MED']
key_symbol = ['Cavity', 'Chip','Brown patch of color, Pinpoint, Needle','Cavity, Chip','Chip, Cavity'] 
black= ['B2','B3']
crown = ['B2','B3','CrownCentre']

def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names, drop 'Unnamed' residuals, and coerce mixed object cols to str."""
    # Normalize column names to strings
    df.columns = [str(c).strip() for c in df.columns]

    # Drop Excel residual index columns like "Unnamed: 0", "Unnamed: 0.1", etc.
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    # Decode any bytes in object columns; then make types consistent
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore")
                if isinstance(x, (bytes, bytearray))
                else x
            )
            # If mixed types remain, coerce to string to keep Arrow happy
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)

    return df

def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""
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

def negation(df: pd.DataFrame,col: str,allowed_values,negate: bool = False, keep_na: bool = True,) -> pd.DataFrame:
    if col in df.columns:
        mark = df[col].isin(allowed_values)
        if negate:
            mark = ~mark
        if keep_na:
            mark = mark | df[col].isna()
        return df[mark]
    return df

def filter_none(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in df.columns:
        return df.iloc[0:0]  
    s = df[col].astype(str).str.strip()
    mask_missing = ( df[col].isna() | s.eq("") | s.str.casefold().isin({"nan", "none", "null", "n/a", "na"}))
       
    return df[mask_missing]

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Base filters
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "ROUND")
    df = equal(df, "Location", "INDIA")
    df = between(df, "Ct.", 0.30, 3.99, inclusive="both")
    df = negation(df, "Col.", colors)
    df = negation(df, "Clarity", clarity)
    df = negation(df, "Cut", cut)
    df = negation(df, "Pol.", cut)
    df = negation(df, "Sym.", cut)
    df = negation(df, "Fl.", flour)
    df = negation(df, "Centre Black", black, negate=True)
    df = negation(df, "Crown Black", crown, negate=True)
    df = negation(df, "Key To Symbols", key_symbol, negate=True)
    if 'Report Comments' in df.columns:
        df = df[~df['Report Comments'].astype(str).str.startswith(
            "Clarity grade is based on internal graining that is not shown", na=False)]
    # DISC filter
    if "Disc. (%)" in df.columns:
        df["Disc. (%)"] = df["Disc. (%)"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        bad = df["Disc. (%)"].between(-10, -0, inclusive="both")
        df = df[~bad & df["Disc. (%)"].notna()]

    # Shade: keep only None/empty-like (if that's what you want)
    if "BGM" in df.columns:
        df = df[df.index.isin(filter_none(df, "BGM").index)]

    return df

def default_value(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        if col in df.columns:
            df[col] = 0

    # Keep numeric-looking fields typed as numeric if possible
    for num_col in [ "CARAT","TABLE", "DEPTH","DISC","RATIO", "RAP","LENGTH", "WIDTH","HEIGHT", ]:
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

def _find_sample(filename: str):
    """Locate a file inside a 'Sample Files' folder, searching upward from this
    script. Works on Linux (Streamlit Cloud) and Windows regardless of the
    current working directory or where this page sits in a multipage app."""
    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        candidate = parent / "Sample Files" / filename
        if candidate.exists():
            return candidate
    return None

def run_d_navin_mod() -> None:
    st.header("D Navin")

    sample_path = _find_sample("D Navin.xlsx")
    
    with st.sidebar:
    
        if sample_path is not None:
            with open(sample_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            href = (
                f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" '
                f'download="{sample_path.name}">Sample File</a>'
            )
            st.markdown(href, unsafe_allow_html=True)
        else:
            st.warning("Sample file not found on server.")
        upload = st.file_uploader("Select D Navin File",type=["xlsx","xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 1, 1)

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

    csv_file = df.to_csv(index=False)  # BOM helps Excel
    file_name = f"d navin {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_file,
            file_name=file_name,
            mime="csv",
        )

if __name__ == "__main__":
    run_d_navin_mod()
