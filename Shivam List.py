from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df)

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'] 
# Additional filtering
data = data[data['Lab'] == 'GIA']
data = data[data['Location'] == 'MUMBAI']
data = data[data['Shape'] == 'ROUND']
data['Dis%'] = data['Dis%'].apply(lambda x: -x if x >= 0 else x)
data = data[data['Dis%'] <= -10]
data = data[(data['Carat'] >= 0.30) & (data['Carat'] <= 3.99)]
data = data[data['Color'].isin(colors)]
data = data[data['Clarity'].isin(clarity)]
data = data[data['Cut'].isin(cut) & data['Pol'].isin(cut) & data['Sym'].isin(cut)]
data = data[data['Flour'].isin(flour)]
data = data[data['Milky'].isna()]
data = data[data['Shade'].isna()]
data = data[data['SGP Comment'].isna()]
data = data[data['Culet'] == 'NON']
data = data[~data['Key To Symbol'].isin(key_symbol)]
data = data[~data['SGP Comment'].str.startswith('Clarity',na=False)]
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
    'LOTNO': ['Certificate No'],
    'CARAT': ['Carat'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Pol'],
    'SYM': ['Sym'],
    'FLRC': ['Flour'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table'],
    'DEPTH': ['Depth'],
    'LAB': ['Lab'],
    'DISC': ['Dis%'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': []
}

output_data = pd.DataFrame()

# Populate output_data
# Create output DataFrame
output_data = pd.DataFrame()

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

# Save output data as a CSV file
output_dir = r"G:\Certificate\Python Output\DEAL Output"
current_datetime = datetime.now().strftime("%d.%m.%Y")
file_path = os.path.join(output_dir, f"shivan list {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
