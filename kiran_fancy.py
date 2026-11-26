import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
flour = ['Non', 'Fnt', 'Med']
key_symbol = ['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']
milky = ['NONE']
shade = ['NONE']
cut = ['EX', 'VG']
culet = ['NON']
shape = ["PR","OL","EM"]
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
    df['Disc%'] = df['Disc%'].apply(lambda x: -x if x >= 0 else x)
    df['L:W'] = df['L:W'].apply(lambda x: f"{abs(x):.2f}")
    df = df[df['Lab'] == 'GIA']
    df = df[df['Shp'].isin(shape)]
    df['Shp'] = df['Shp'].replace({'RB': 'ROUND','PR': 'PRINCESS','OL': 'OVAL','EM': 'EMERALD'})
    df = df[(df['Cts'] >= 0.30) & (df['Cts'] <= 3.99)]
    df = df[df['Loc'] == 'INDIA']
    df = df[df['Disc%'] <= 10]
    df = df[df['Col'].isin(colors)]
    df = df[df['Clr'].isin(clarity)]
    df = df[df['Flr'].isin(flour)]
    df['Flr'] = df['Flr'].replace({'Fnt': 'FNT', 'Med': 'MED', 'Non': 'NON'})
    df = df[df['Pol'].isin(cut) & df['Sym'].isin(cut)]
    # df = df[~df['K T S'].isin(key_symbol)]
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
    'LOTNO': ['Rep No'],
    'CARAT': ['Cts'],
    'SHAPE': ['Shp'],
    'COLOR': ['Col'],
    'CLARITY': ['Clr'],
    'CUT': ['Cut'],
    'POL': ['Pol'],
    'SYM': ['Sym'],
    'FLRC': ['Flr'],
    'LENGTH': ['Max'],
    'WIDTH': ['Min'],
    'HEIGHT': ['Hgt'],
    'TABLE': ['Tbl'],
    'DEPTH': ['TD%'],
    'LAB': ['Lab'],
    'DISC': ['Disc%'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': ['L:W'],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': [],
    'Natts1':['BIC'],
    'Natts2':['BIS'],
    "Stone id":['Stone No']
}

# Streamlit app to load, process, and download data
def run_kiran_fancy_mod():
    st.header("KIRAN")

    # Sidebar for file upload and skipping rows
    with st.sidebar:
        upload = st.file_uploader("Select Excel file", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)


    if upload is None:
            st.stop()

    with st.spinner("Processing..."):

        try:
            # Initialize progress bar
            progress_bar = st.progress(0)

            df = read_file(upload, skiprows=skip)
        except Exception as e:
            st.error(f"Failed to read file: {e}")
            st.stop()

        # Show raw (sanitized)
        progress_bar.progress(20)
        st.dataframe(df)

        # Create output DataFrame with mapped columns
        progress_bar.progress(40)
        output_data = pd.DataFrame()
        for column, keys in columns_mapping.items():
            if column in ['RAP']:
                output_data[column] = 0
            elif column == 'SRNO':
                output_data[column] = range(1, len(df) + 1)
            elif column in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
                output_data[column] = 'N/A'
            elif column == 'NATTS':
                output_data[column] = 'CENTER - L'
            elif column == 'CUT':
                output_data[column] = 'VG'
            elif not keys:
                output_data[column] = np.nan
            else:
                output_data[column] = concatenate_dataframes({0: df}, keys)

        st.subheader("Transformed Data")
        progress_bar.progress(60)
        st.dataframe(output_data)
        progress_bar.progress(80)
        # Prepare CSV for download
        csv_bytes = output_data.to_csv(index=False)
        today_str = pd.Timestamp.now().strftime("%d.%m.%Y")
        file_name = f"kiran fancy {today_str}.csv"
        progress_bar.progress(100)
        with st.sidebar:
            st.download_button(
                label="Generate CSV",
                data=csv_bytes,
                file_name=file_name,
                mime="csv",
                
            )
if __name__ == "__main__":
    run_kiran_fancy_mod()


