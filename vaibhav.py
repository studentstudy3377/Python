import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np

# ==========================================
# Config & Constants (from your Vaibhav script)
# ==========================================
ALLOWED_COLOR = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
ALLOWED_CLARITY = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
ALLOWED_CUT = ["EX", "VG"]
ALLOWED_FLOUR = ["NON", "FNT", "MED"]
ALLOWED_CULET = ["NONE", "Pointed", "NON", "Blank"]
SHADE_ALLOWED = ["NONE", "Blank"]
EXCLUDED_CB = [3, 4]
EXCLUDED_SW = [3, 4, 5]
EXCLUDED_KEY_SYMBOLS = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]
EXCLUDED_COMMENT_PREFIXES = [
    "Clarity grade is based on internal graining that is not shown.",
    "CLARITY GRADE IS BASED ON INTERNAL GRAINING THAT IS NOT SHOWN.",
]

COLUMNS_MAPPING = {
    "SRNO": [],
    "LOTNO": ["Report No"],
    "CARAT": ["Wgt"],
    "SHAPE": ["Shape"],
    "COLOR": ["Col"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
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
    'Natts1':['CB'],
    'Natts2':['SB'],
    "Stone id":['Packet No']
}

# =====================
# Utility helpers
# =====================
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    # trim headers, drop excel unnamed
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    # fill blanks like your script
    df = df.fillna('Blank')

    # normalize numeric columns you use
    if 'Wgt' in df.columns:
        df['Wgt'] = pd.to_numeric(df['Wgt'], errors='coerce')
    if 'Disc %' in df.columns:
        df['Disc %'] = pd.to_numeric(df['Disc %'], errors='coerce')

    # strip object cells, decode bytes
    for col in df.columns:
        if df[col].dtype == 'object':
            def _fix(x):
                if isinstance(x, (bytes, bytearray)):
                    try:
                        x = x.decode('utf-8', 'ignore')
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
        df[col] = pd.to_numeric(df[col], errors='coerce')
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


def between(df: pd.DataFrame, col: str, low, high, inclusive: str = 'both') -> pd.DataFrame:
    if col in df.columns:
        df = numeric(df, col)
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df

# =====================
# Core filter pipeline (mirrors your script)
# =====================
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = equal(df, 'Lab', 'GIA')
    df = equal(df, 'Shape', 'ROUND')
    df = between(df, 'Wgt', 0.30, 3.99, inclusive='both')
    df = equal(df, 'Location', 'INDIA')

    if 'Disc %' in df.columns:
        df = df[df['Disc %'] <= -10]

    df = isin(df, 'Col', ALLOWED_COLOR)
    df = isin(df, 'Clarity', ALLOWED_CLARITY)
    df = isin(df, 'Cut', ALLOWED_CUT)
    df = isin(df, 'Pol', ALLOWED_CUT)
    df = isin(df, 'Sym', ALLOWED_CUT)
    df = isin(df, 'Fls', ALLOWED_FLOUR)

    if 'Culet' in df.columns:
        df = isin(df, 'Culet', ALLOWED_CULET)

    if 'Comments 1' in df.columns:
        df = df[df['Comments 1'] != 'BROWN']

    if 'Milky' in df.columns:
        df = equal(df, 'Milky', 'ML-0')

    if 'Shade' in df.columns:
        df = isin(df, 'Shade', SHADE_ALLOWED)

    for col, bad in [('CB', EXCLUDED_CB), ('SB', EXCLUDED_CB), ('CW', EXCLUDED_CB), ('SW', EXCLUDED_SW)]:
        if col in df.columns:
            df = notin(df, col, bad)

    if 'Fls Col' in df.columns:
        df = df[df['Fls Col'] != 'Yellow']

    for key_col in ['Key To Symbols', 'Key To Symbol', 'Key To Sym']:
        if key_col in df.columns:
            df = notin(df, key_col, EXCLUDED_KEY_SYMBOLS)

    if 'Report Comments' in df.columns:
        s = df['Report Comments'].astype(str).str.strip()
        for pref in EXCLUDED_COMMENT_PREFIXES:
            s = s[~s.str.startswith(pref)]
        df = df.loc[s.index]

    return df


# =====================
# Output shaping
# =====================
def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {src: tgt for tgt, sources in COLUMNS_MAPPING.items() for src in (sources + [tgt])}
    # Trim headers again to catch space-variants before rename
    df = df.rename(columns={c: c.strip() for c in df.columns})
    df = df.rename(columns=rename_map)

    for col in COLUMNS_MAPPING:
        if col not in df.columns:
            df[col] = np.nan

    for col in ['LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP']:
        df[col] = 0

    df['SRNO'] = range(1, len(df) + 1)

    for col in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
        df[col] = 'N/A'

    df['NATTS'] = 'CENTER - L'
    df['SHAPE'] = 'ROUND'

    for num_col in ['CARAT', 'TABLE', 'DEPTH', 'DISC', 'RATIO', 'RAP', 'LENGTH', 'WIDTH', 'HEIGHT']:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors='coerce')
    if 'DISC' in df.columns:
        df['DISC'] = df['DISC'].round(2)

    return df[list(COLUMNS_MAPPING)]

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

# =====================
# App
# =====================
def run_vaibhav_mod() -> None:
    st.header('Vaibhav')

    sample_path = _find_sample("Vaibhav.xlsx")
    
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
            
        upload = st.file_uploader('Select Excel file', type=['xlsx', 'xls'])
        skip = st.number_input('Skip rows', 0, 10000, 1, 1)

    if not upload:
        st.info('Upload a file to begin.')
        st.stop()

    try:
        df = read_excel(upload, skiprows=skip)
    except Exception as e:
        st.error(f'Failed to read file: {e}')
        st.stop()

    st.subheader('Raw data (sanitized)')
    st.dataframe(df)

    df_filtered = apply_filters(df)
    df_out = fix_columns(df_filtered.copy())

    st.subheader('Filtered & Mapped Output')
    st.dataframe(df_out)

    file_name = f"vaibhav {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_out.to_csv(index=False)

    with st.sidebar:
        st.download_button(
            label='Download CSV',
            data=csv_bytes,
            file_name=file_name,
            mime='csv',
            
        )


if __name__ == '__main__':
    run_vaibhav_mod()
