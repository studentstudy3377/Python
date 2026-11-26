import streamlit as st
import pandas as pd
import numpy as np

columns_mapping = {
    'SRNO': [],
    'LOTNO': ['Stock No'],
    'CARAT': ['Carat'],
    'SHAPE': [''],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symmetry'],
    'FLRC': ['Flo'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Tab%'],
    'DEPTH': ['TD%'],
    'LAB': ['Lab'],
    'DISC': ['Rap%'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['BT'],
    'Natts2':['BC'],
    "Stone id":[]
}

# Lists
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX','VG']
flour = ['N', 'F', 'M']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'] 
bt = ['BT2','BT3']
bc = ['BC2','BC3']
wt = ['WT3']
wc = ['WC3']
milky = ['M1','M2','M3','NV']
oopv= ['PO2','PO3']
opcr = ['CO2','CO3']
opta = ['NV','TO1','TO2']
culet = ['L','VSM','']
flo_color = ['YL','GR']
luster  =  ['GD','VG']


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
    df = equal(df, 'Lab', 'GIA')
    df = equal(df, 'Shape', 'ROUND')
    df = equal(df, 'Location', 'MUMBAI')
    df = equal(df, 'Color Shade', 'WH')
    df = negation(df, 'Color', colors)            
    df = negation(df, 'Clarity', clarity)         
    df = negation(df, 'Flo', flour)           
    df = negation(df, 'Cut', cut)           
    df = negation(df, 'Polish', cut)           
    df = negation(df, 'Symmetry', cut)           
    df = negation(df, 'BT', bt, negate=True)
    df = negation(df, 'BC', bc, negate=True)
    df = negation(df, 'WT', wt, negate=True)
    df = negation(df, 'WC', wc, negate=True)
    df = negation(df, 'Milky', milky, negate=True)
    df = negation(df, 'OPPV', oopv, negate=True)
    df = negation(df, 'OPCR', opcr, negate=True)
    df = negation(df, 'OPTA', opta, negate=True)
    df = negation(df, 'Culet', culet, negate=True)
    df = negation(df, 'Flo Color', flo_color, negate=True)
    df = negation(df, 'Luster', luster, negate=True)
    df = negation(df, 'Key To Symbol', key_symbol, negate=True)
    df = between(df, 'Carat', 0.30, 3.99, inclusive='both')
    
    # DISC filter
    if "Rap%" in df.columns:
        bad = df["Rap%"].between(-10, -0, inclusive="both")
        df = df[~bad & df["Rap%"].notna()]

    if 'GIA Comments' in df.columns:
        df = df[~df['GIA Comments'].astype(str).str.startswith("Clarity grade is based on internal graining that is not shown", na=False)]

    # Shade: keep only None/empty-like (if that's what you want)
    if "Color Desc" in df.columns:
        df = df[df.index.isin(filter_none(df, "Color Desc").index)]

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
    
    if "FLRC" in df.columns:
        df["FLRC"] = df["FLRC"].replace({'F': 'FNT', 'M': 'MED', 'N': 'NON'})
    
    return df


def run_hk_mod() -> None:
    st.header("HK")

    with st.sidebar:
        upload = st.file_uploader("Select HK File",type=["xlsx","xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)

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
    file_name = f"hk {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_file,
            file_name=file_name,
            mime="csv"
            
        )

if __name__ == "__main__":
    run_hk_mod()
