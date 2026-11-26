import streamlit as st
import pandas as pd
import numpy as np

columns_mapping = {
    'SRNO': [],
    'LOTNO': ['Packet No'],
    
    'CARAT': ['Carat'],
    'SHAPE': [''],
    'COLOR': ['Col'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Pol'],
    'SYM': ['Sym'],
    'FLRC': ['Fls'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table %'],
    'DEPTH': ['Depth %'],
    'LAB': ['Lab'],
    'DISC': ['Disc %'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['CN'],
    'Natts2':['SN'],
    "Stone id":['Packet No']
}
# Lists
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
key_symbol = ['Cavity', 'Cavity, Chip','Cavity, Chip','Chip, Cavity'] 
black = ['CN3','SN3','CW3','SW3','Y']

def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names, drop 'Unnamed' residuals, and coerce mixed object cols to str."""
    # Work on a copy to avoid chained-assignment warnings
    df = df.copy()

    # Normalize column names to strings
    df.columns = [str(c).strip() for c in df.columns]

    # Drop Excel residual index columns like "Unnamed: 0", "Unnamed: 0.1", etc.
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()

    # Decode any bytes in object columns; then make types consistent
    for col in df.columns:
        if df[col].dtype == "object":
            decoded = df[col].apply(
                lambda x: x.decode("utf-8", "ignore") if isinstance(x, (bytes, bytearray)) else x
            )
            # If mixed types remain, coerce to string to keep Arrow happy
            if decoded.map(type).nunique() > 1:
                decoded = decoded.astype(str)
            df[col] = decoded

    return df

def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""
    return df[list(columns_mapping)].copy()

def numeric(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col in df.columns:
        df = df.copy()
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        mask = df[col] == value
        return df.loc[mask].copy()
    return df

def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        s = pd.to_numeric(df[col], errors="coerce")
        mask = s.between(low, high, inclusive=inclusive)
        return df.loc[mask].copy()
    return df

def negation(
    df: pd.DataFrame,
    col: str,
    allowed_values,
    negate: bool = False,
    keep_na: bool = True,
) -> pd.DataFrame:
    if col in df.columns:
        mark = df[col].isin(allowed_values)
        if negate:
            mark = ~mark
        if keep_na:
            mark = mark | df[col].isna()
        return df.loc[mark].copy()
    return df

def filter_none(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in df.columns:
        return df.iloc[0:0].copy()
    s = df[col].astype(str).str.strip()
    mask_missing = (df[col].isna() | s.eq("") | s.str.casefold().isin({"nan", "none", "null", "n/a", "na"}))
    return df.loc[mask_missing].copy()

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Base filters (each returns a fresh copy)
    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "RD")
    df = equal(df, "Location", "IND")
    # df = equal(df, "Shade", "WHT")
    df = negation(df, "Col", colors)              
    df = negation(df, "Clarity", clarity)           
    df = negation(df, "Cut", cut)                  
    df = negation(df, "Pol", cut)                  
    df = negation(df, "Sym", cut)                  
    df = negation(df, "Fls", flour)                 
    df = negation(df,'CN',black,negate=True)                 
    df = negation(df,'SN',black,negate=True)                 
    df = negation(df,'CW',black,negate=True)                 
    df = negation(df,'SW',black,negate=True)                 
    df = negation(df,'To',black,negate=True)                 
    df = between(df, "Carat", 0.30, 3.99, inclusive="both")    
    df = negation(df, "Key To Symbols", key_symbol, negate=True)

    if "Disc %" in df.columns:
        df["Disc %"] = df["Disc %"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        bad = df["Disc %"].between(-10, -0, inclusive="both")
        df = df[~bad & df["Disc %"].notna()]

    if 'Report Comments' in df.columns:
        rc = df['Report Comments'].astype(str)
        df = df.loc[~rc.str.startswith("Clarity grade is based on internal graining that is not shown", na=False)].copy()

    if "Culet" in df.columns:
        s = df["Culet"].astype(str).str.strip().str.upper()

        allowed_culet = ["NONE", "POINTED", "NON"]
        mask_culet = s.isin(allowed_culet)

        mask_blank = df.index.isin(filter_none(df, "Culet").index)

        df = df[mask_culet | mask_blank]
    
    if "Brl." in df.columns:
        s = df["Brl."].astype(str).str.strip().str.casefold()
        mask_rb = s.eq("ex")
        mask_none = df.index.isin(filter_none(df, "Brl.").index)
        df = df[mask_rb | mask_none]

    return df.reset_index(drop=True)

def default_value(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

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

def run_finestar_mod() -> None:
    st.header("Finestar")

    with st.sidebar:
        upload = st.file_uploader("Select Finestar File",type=["xlsx","xls"])
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
    file_name = f"finestar {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_file,
            file_name=file_name,
            mime="csv"
            )
        
if __name__ == "__main__":
    run_finestar_mod()
