import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX', 'VG']
flour = ['NONE', 'FAINT', 'MEDIUM']
culet = ['NON']
key_symbol = ['Cavity', 'Cavity', 'Chip', 'Cavity/Chip', 'Chip/Cavity']
luster = ['EX', '']
shade = ['NONE', '']
side_black = ['MEDIUM', 'CLOUD', 'HEAVY']
side_spot = ['RED MEDIUM', 'RED MINOR']
table_spot = ['RED MINOR', 'GREEN MINOR', 'RED MEDIUM']
flour_color = ['BLUE,VERY SLIGHT', 'NONE,VERY SLIGHT', 'NONE,YELLOW', 'YELLOW', '']

# Function to sanitize data
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore") if isinstance(x, (bytes, bytearray)) else x
            )
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df

# Function to read uploaded file
def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)

# Function to apply filters to the data
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df['DISC'] = df['DISC'].apply(lambda x: -x if x >= 0 else x)
    df = df[df['LAB'] == 'GIA']
    df = df[df['SHAPE'] == 'ROUND']
    df = df[(df['CTS'] >= 0.30) & (df['CTS'] <= 3.99)]
    df = df[df['LOC.'] == 'Mumbai']
    df = df[df['DISC'] <= -10]
    df = df[df['COLOR'].isin(colors)]
    df = df[df['CLARITY'].isin(clarity)]
    df = df[df['FLUOR'].isin(flour)]
    df['FLUOR'] = df['FLUOR'].replace({'FAINT': 'FNT', 'MEDIUM': 'MED', 'NONE': 'NON'})
    df = df[df['CUT'].isin(cut) & df['POL'].isin(cut) & df['SYMM'].isin(cut)]
    df = df[df['SHADE'].isin(shade)]
    df = df[df['TABLE OPEN INCLUSION'].isin(shade)]
    df = df[df['LUSTER'].isin(luster)]
    df = df[~df['FLUORESCENCE COLOR'].isin(flour_color)]
    df = df[~df['SIDE BLACK INCLUSION'].isin(side_black)]
    df = df[~df['TABLE SPOT INCLUSION'].isin(table_spot)]
    df = df[df['TABLE BLACK INCLUSION'].isin(side_black)]
    df = df[~df['KEY TO SYMBOL'].isin(key_symbol)]
    df = df[~df['REPORT COMMENTS'].str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)]
    return df

# Function to concatenate dataframes based on column names
def concatenate_dataframes(df, column_names):
    concatenated = []
    for sub_df in df.values():
        for column_name in column_names:
            if sub_df is not None and column_name in sub_df.columns:
                concatenated.append(sub_df[column_name])
                break
    concatenated = [series for series in concatenated if not series.empty]
    return pd.concat(concatenated, ignore_index=True) if concatenated else pd.Series(dtype='str')

# Column mapping from original code
columns_mapping = {
    'SRNO': [],
    'LOTNO': ['REPORT NO.'],
    'CARAT': ['CTS'],
    'SHAPE': ['SHAPE'],
    'COLOR': ['COLOR'],
    'CLARITY': ['CLARITY'],
    'CUT': ['CUT'],
    'POL': ['POL'],
    'SYM': ['SYMM'],
    'FLRC': ['FLUOR'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['TABLE %'],
    'DEPTH': ['DEPTH %'],
    'LAB': ['LAB'],
    'DISC': ['DISC'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['Natts'],
    'Natts2':[],
    "Stone id":['Stone No']   
}

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

# Streamlit app to load, process, and download data
def run_kg_mod():
    st.header("KG")

    # Sidebar for file upload and skipping rows
    sample_path = _find_sample("KG.xlsx")
    
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
            
        upload = st.file_uploader("Select Excel file", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)

    if upload is None:
        st.info("Upload an Excel file to begin.")
        st.stop()

    try:
        # Read and sanitize the file
        df = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw Data (Sanitized)")
    st.dataframe(df)

    # Apply filters and normalization
    df = apply_filters(df)

    # Create output DataFrame with mapped columns
    output_data = pd.DataFrame()
    for column, keys in columns_mapping.items():
        if column in ['LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP']:
            output_data[column] = 0
        elif column == 'SRNO':
            output_data[column] = range(1, len(df) + 1)
        elif column in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
            output_data[column] = 'N/A'
        elif column == 'NATTS':
            output_data[column] = 'CENTER - L'
        elif not keys:
            output_data[column] = np.nan
        else:
            output_data[column] = concatenate_dataframes({0: df}, keys)

    st.subheader("Transformed Data")
    st.dataframe(output_data)

    # Prepare CSV for download
    csv_bytes = output_data.to_csv(index=False)
    today_str = pd.Timestamp.now().strftime("%d.%m.%Y")
    file_name = f"kg {today_str}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_kg_mod()
