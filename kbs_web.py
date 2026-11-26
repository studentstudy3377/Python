import base64
from pathlib import Path

import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime

# Function to load and clean the data
def load_and_process_data(uploaded_file):
    # Load data
   
    data = pd.read_excel(uploaded_file, skiprows=2)

    # Define filter criteria
    colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
    clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
    cut = ['EX', 'VG']
    flour = ['N', 'F', 'M']
    culet = ['NON']
    bic = ['Heavy', 'Major']
    mix_tinch = ['MT1', 'MT2']
    key_symbol = ['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']

    # Apply the filtering
    data = data[data['LAB'] == 'GIA']
    data = data[data['SHAPE'] == 'Round']
    data['SHAPE'] = data['SHAPE'].str.replace('Round', 'ROUND')
    data = data[data['DISC'] <= -10]
    data = data[data['LOC'] == 'INDIA']
    data = data[(data['WEIGHT'] >= 0.30) & (data['WEIGHT'] <= 3.99)]
    data = data[data['COL'].isin(colors)]
    data = data[data['CLR'].isin(clarity)]
    data = data[data['CUT'].isin(cut) & data['POL'].isin(cut) & data['SYM'].isin(cut)]
    data = data[data['FLR'].isin(flour)]
    data['FLR'] = data['FLR'].replace({'F': 'FNT', 'M': 'MED', 'N': 'NON'})
    data = data[(data['Brown'] == 'BR0') | (data['Brown'].isna())]
    data = data[data['MILKY'].isna()]
    data = data[~data['Mix Tinge'].isin(mix_tinch)]
    data = data[~data['BIC'].isin(bic)]
    data = data[~data['BIS'].isin(bic)]
    data = data[~data['KTS'].isin(key_symbol)]

    return data

# Function to concatenate dataframes
def concatenate_dataframes(df, column_names):
    concatenated = []
    for sub_df in df.values():
        for column_name in column_names:
            if sub_df is not None and column_name in sub_df.columns:
                concatenated.append(sub_df[column_name])
                break
    concatenated = [series for series in concatenated if not series.empty]
    return pd.concat(concatenated, ignore_index=True) if concatenated else pd.Series(dtype='str')

# Column mapping
columns_mapping = {
    'SRNO': [],
    'LOTNO': ['CertNo'],
    'CARAT': ['WEIGHT'],
    'SHAPE': ['SHAPE'],
    'COLOR': ['COL'],
    'CLARITY': ['CLR'],
    'CUT': ['CUT'],
    'POL': ['POL'],
    'SYM': ['SYM'],
    'FLRC': ['FLR'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['TA%'],
    'DEPTH': ['TD%'],
    'LAB': ['LAB'],
    'DISC': ['DISC'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['BIC'],
    'Natts2':['BIS'],
    "Stone id":['StoneID']
}

# Function to process the data
def process_data(data):
    output_data = pd.DataFrame()

    # Populate output_data
    for column, keys in columns_mapping.items():
        if column in ['LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP']:
            output_data[column] = 0
        elif column == 'SRNO':
            output_data[column] = range(1, len(data) + 1)
        elif column in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
            output_data[column] = 'N/A'
        elif column == 'NATTS':
            output_data[column] = 'CENTER - L'
        elif not keys:
            output_data[column] = np.nan
        else:
            output_data[column] = concatenate_dataframes({0: data}, keys)

    return output_data

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

# Streamlit app
def run_kbs_web_mod():
    st.title("Diamond Data Processing")
    sample_path = _find_sample("KBS Diamond Web.xlsx")
        
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
    # File uploader
        uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

    if uploaded_file:
        # Process data
        data = load_and_process_data(uploaded_file)

        # Show the raw data before processing
        st.subheader("Raw Data")
        st.dataframe(data)

        # Process the data
        processed_data = process_data(data)

        # Show processed data
        st.subheader("Processed Data")
        st.dataframe(processed_data)

        # Prepare CSV download
        csv = processed_data.to_csv(index=False)

        # Save file path
        current_datetime = datetime.now().strftime("%d.%m.%Y")
        file_name = f"kbs_diamonds_{current_datetime}.csv"
        with st.sidebar:
            st.download_button(
                label="Generate CSV",
                data=csv,
                file_name=file_name,
                mime="csv",
                
            )

if __name__ == "__main__":
    run_kbs_web_mod()