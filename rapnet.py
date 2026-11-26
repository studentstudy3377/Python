import io
import streamlit as st
import pandas as pd
import numpy as np
import re

OUTPUT_MAP = {
    'SRNO': [],
    'LOTNO': ['Stone ID'],
    'CARAT': ['Size'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symmetry'],
    'FLRC': ['Fluorescence'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table'],
    'DEPTH': ['Depth'],
    'LAB': ['Lab'],
    'DISC': ['%Rap'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':[],
    'Natts2':[],
    "Stone id":[]
}

def read_file(uploaded_file, skiprows: int = 0):
    
    return pd.read_excel(uploaded_file, skiprows=skiprows,keep_default_na=False)

def fix_columns(df):
    df = df.rename(columns={a: k for k, v in OUTPUT_MAP.items() for a in v + [k]})
    for k in OUTPUT_MAP:
        if k not in df.columns:
            df[k] = ""
    return df[list(OUTPUT_MAP)]

def numeric (df,col):
    if col in df.columns:
        df[col] = pd.to_numeric(df[col],errors='coerce')
    return df

def equal(df,col,value):
    if col in df.columns:
        return df[df[col] == value]
    return df

def between(df,col, low, high,inclusive= 'both'):
    if col in df.columns:
        df = numeric(df,col)
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df 

def negation(df, col, allowed_values , negate = False):
    if col in df.columns:
        mark  = df[col].isin(allowed_values)
        if negate:
            mark = ~mark
        return df[mark]
    return df

# Lists
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
milky = ['NONE']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
culet = ['NON']
key_symbol = ['Cavity', 'Chip']
green = ['NONE']
black_in = ['N', 'BPP']

def apply_filters (df:pd.DataFrame):
    
    df = equal(df,'Lab','GIA')
    df = equal(df,'Shape','Round')
    df = between(df,'Size',0.30,3.99,inclusive='both')
    df = df[df['%Rap'] <= 10]
    df = negation(df,'Color',colors)
    df = negation(df,'Clarity',clarity)
    df['Fluorescence'] = df['Fluorescence'].replace({'Faint': 'FNT', 'None': 'NON', 'Medium': 'MED'})
    df = df[df['Fluorescence'].isin(flour)]
    df["Cut"] = df["Cut"].replace({"Very Good":"VG","Excellent":"EX"})
    df = df[df['Cut'].isin(cut)]
    df["Polish"] = df["Polish"].replace({"Very Good":"VG","Excellent":"EX"})
    df = df[df['Polish'].isin(cut)]
    df["Symmetry"] = df["Symmetry"].replace({"Very Good":"VG","Excellent":"EX"})
    df = df[df['Symmetry'].isin(cut)]
    df = negation(df,'Key To Symbols',key_symbol,negate=True)
    return df

def default_value (df:pd.DataFrame):
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

def run_rapnet_mod():
    st.subheader("Rapnet")

    with st.sidebar:
        uploaded = st.file_uploader("Select File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 1000, 0, 1)

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

    file_name = f"rapnet {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_buffer,
            file_name=file_name,
            mime="csv",  
        )
# If running as a page:
if __name__ == "__main__":
    run_rapnet_mod()
