import streamlit as st
import pandas as pd
import numpy as np

columns_mapping = {
    'SRNO': [],
    'LOTNO': ['Report No'],
    
    'CARAT': ['Cts'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Sym'],
    'FLRC': ['Fluo'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table%'],
    'DEPTH': ['Depth%'],
    'LAB': ['Lab'],
    'DISC': ['Net %'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts_1':[],
    'Natts_2':[],
    "Stock Id":['Stone ID']
}

# Lists
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
culet = ['NON']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'] 

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
    df = equal(df, "Location", "MUMBAI")
    df = negation(df, "Color", colors)
    df = negation(df, "Clarity", clarity)
    df = negation(df, "Cut", cut)
    df = negation(df, "Polish", cut)
    df = negation(df, "Sym", cut)
    df = negation(df, "Fluo", flour)
    df = between(df, "Cts", 0.30, 3.99, inclusive="both")
    df = negation(df, "Key To Symbols", key_symbol, negate=True)

    # DISC filter
    if "Net %" in df.columns:
        df["Net %"] = df["Net %"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        bad = df["Net %"].between(-10, -0, inclusive="both")
        df = df[~bad & df["Net %"].notna()]
    
    if 'Comments' in df.columns:
        df = df[~df['Comments'].astype(str).str.startswith(
            "Clarity", na=False)]
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

def run_gyanimpex_mod() -> None:
    st.header("Gyam Impex")

    with st.sidebar:
        upload = st.file_uploader("Select Amora File",type=["xlsx","xls"])
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
    file_name = f"gyamimpex {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_file,
            file_name=file_name,
            mime="csv"
           
        )

if __name__ == "__main__":
    run_gyanimpex_mod()