import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime

# Function to load and clean the data
def load_and_process_data(uploaded_file, skip):
    # Load data
    data = pd.read_excel(uploaded_file, skiprows=skip)

    # Define filter criteria
    colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
    clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI2']
    cut = ['EX', 'VG']
    flour = ['NONE', 'FA', 'MD']
    key_symbol = ['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']
    milky = ['EX']
    bgm = ['Yes', '']
    side_black = ['B2-S', 'B3-S']
    table_black = ['B2-T']

    # Apply filters
    data = data[data['Lab'] == 'GIA']
    data = data[data['Shape'] == 'Round']
    data['Shape'] = data['Shape'].str.replace('Round', 'ROUND')
    data = data[(data['Ct'] >= 0.30) & (data['Ct'] <= 3.99)]
    data = data[data['Location'] == 'Mumbai']
    if 'Rapnet Back' in data.columns:
        data = data[data['Rapnet Back'] <= -10]
    elif 'Online Back' in data.columns:
        data = data[data['Online Back'] <= -10]
        data = data[data['Col'].isin(colors)]
        data = data[data['Clr'].isin(clarity)]
        data = data[data['Fls'].isin(flour)]
        data['Fls'] = data['Fls'].replace({'FA': 'FNT', 'MD': 'MED', 'NONE': 'NON'})
        data = data[~data['BGM'].isin(bgm)]
        data = data[data['brown'].isna()]
        data = data[~data['Tab_BLK'].isin(table_black)]
        data = data[data['Op_Tab'].isna()]
        data = data[data['Culet'] == 'NONE']
        data = data[data['Cut'].isin(cut) & data['Pol'].isin(cut) & data['Sym'].isin(cut)]
        data = data[data['Milky/Luster'].isin(milky)]
        data = data[~data['Side_BLK'].isin(side_black)]
        data = data[~data['Key To Symbols'].isin(key_symbol)]
        data = data[~data['Report Comments'].str.startswith(
            "Clarity grade is based on internal graining that is not shown.", na=False
        )]

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
    'LOTNO': ['Report No'],
    'CARAT': ['Ct'],
    'SHAPE': ['Shape'],
    'COLOR': ['Col'],
    'CLARITY': ['Clr'],
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
    'DISC': ['Rapnet Back','Online Back'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['Tab_BLK'],
    'Natts2':['Side_BLK'],
    "Stone id":['Packet No']
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

# Streamlit app
def run_m_suresh_mod():
    st.title("M Suresh Diamond Data Processing")

    with st.sidebar:
        # File uploader
        uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

        # NEW: Skip rows input
        skip = st.number_input("Skip Rows", 0, 10000, 1, 1)  # default 5

    if uploaded_file:
        # Process data (skip controlled by input)
        data = load_and_process_data(uploaded_file, skip)

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
        file_name = f"m_suresh_{current_datetime}.csv"
        with st.sidebar:
            st.download_button(
                label="Generate CSV",
                data=csv,
                file_name=file_name,
                mime="text/csv",
            )

if __name__ == "__main__":
    run_m_suresh_mod()
