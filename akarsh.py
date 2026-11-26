import base64
import io
import os
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np
import re

OUTPUT_MAP = {
    'SRNO': [],
    'LOTNO': ['Report No'],
    
    'CARAT': ['Cts'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Purity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symm'],
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
    'Natts1':['Black Incl.'],
    'Natts2':['Table Incl.'],
    "Stone id":['Packet No']
}

# Lists
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
milky = ['NONE']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
culet = ['NON']
key_symbol = ['Cavity', 'Chip']
green = ['NONE']
black_in = ['B2', 'B3']


def read_file(uploaded_file, skiprows: int = 0):
    return pd.read_excel(uploaded_file, skiprows=skiprows)


def fix_columns(df):
    df = df.rename(columns={a: k for k, v in OUTPUT_MAP.items() for a in v + [k]})
    for k in OUTPUT_MAP:
        if k not in df.columns:
            df[k] = ""
    return df[list(OUTPUT_MAP)]


def numeric(df, col):
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df


def equal(df, col, value):
    if col in df.columns:
        return df[df[col] == value]
    return df


def between(df, col, low, high, inclusive='both'):
    if col in df.columns:
        df = numeric(df, col)
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df


def negation(df, col, allowed_values, negate=False):
    if col in df.columns:
        mark = df[col].isin(allowed_values)
        if negate:
            mark = ~mark
        return df[mark]
    return df


def apply_filters(df: pd.DataFrame):
    df = equal(df, 'Lab', 'GIA')
    df = equal(df, 'Shape', 'ROUND')
    df = equal(df, 'Location', 'INDIA')
    df = between(df, 'Cts', 0.30, 3.99, inclusive='both')
    df = numeric(df, 'Disc %')
    if "Disc %" in df.columns:
        df["Disc %"] = df["Disc %"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        bad = df["Disc %"].between(-10, -0, inclusive="both")
        df = df[~bad & df["Disc %"].notna()]
    df = negation(df, 'Color', colors)
    df = negation(df, 'Purity', clarity)
    df = negation(df, 'Milky', milky)
    df = negation(df, 'Tinge', green)
    df = negation(df, 'Brown', green)
    df = negation(df, 'Cut', cut)
    df = negation(df, 'Polish', cut)
    df = negation(df, 'Symm', cut)
    df = negation(df, 'Fls', flour)
    df = negation(df, 'Culet', culet)
    df = negation(df, 'Green', green)
    df = negation(df, 'Black Incl.', black_in, negate=True)
    df = negation(df, 'Key to symbols', key_symbol, negate=True)
    return df


def default_value(df: pd.DataFrame):
    for col in ['LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP']:
        if col in df.columns:
            df[col] = 0

    # Keep numeric-looking fields typed as numeric if possible
    for num_col in ['CARAT', 'TABLE', 'DEPTH', 'DISC', 'RATIO', 'RAP', 'LENGTH', 'WIDTH', 'HEIGHT']:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors='coerce')

    for col in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
        if col in df.columns:
            df[col] = 'N/A'

    if 'SRNO' in df.columns:
        df['SRNO'] = range(1, len(df) + 1)

    if 'NATTS' in df.columns:
        df['NATTS'] = 'CENTER - L'

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


def run_akarsh_mod():
    """Streamlit UI entrypoint."""

    st.subheader("Akarsh File")

    sample_path = _find_sample("Akarsh.xlsx")

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

        uploaded = st.file_uploader("Select File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 1000, 1, 1)

    if uploaded is None:
        st.info("Upload a file to continue.")
        return

    try:
        df = read_file(uploaded, skiprows=skip)
    except Exception as e:
        st.error(f"Couldn’t read the file: {e}")
        return

    st.dataframe(df)
    df = apply_filters(df)
    df = fix_columns(df)
    df = default_value(df)
    st.dataframe(df)

    # Create a BytesIO buffer to store the CSV content
    csv_buffer = io.BytesIO()

    # Write the DataFrame to the buffer as CSV
    df.to_csv(csv_buffer, index=False)

    # Move the cursor to the start of the buffer to prepare for download
    csv_buffer.seek(0)

    file_name = f"akarsh {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_buffer,
            file_name=file_name,
            mime="csv",
        )


# If running as a page:
if __name__ == "__main__":
    run_akarsh_mod()