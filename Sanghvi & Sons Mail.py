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
culet = ['VSM','LNG']
luster = ['EX','N']
flourcolor = ['Green', 'Yellow']
patterns = [' CRB2', ' TB2', 'TB3', 'TW3']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'] 
# Additional filtering
data = data[data['LAB'] == 'GIA']
data = data[data['SHAPE'] == 'ROUND']
data = data[data['RAP DISC%'] <= -10]
data = data[(data['CTS'] >= 0.30) & (data['CTS'] <= 3.99)]
data = data[data['COLOR'].isin(colors)]
data = data[data['CLARITY'].isin(clarity)]
data = data[data['CUT'].isin(cut) & data['POL'].isin(cut) & data['SYM'].isin(cut)]
data = data[data['FLOUR'].isin(flour)]
data = data[data['TINGE'] == 'NBG']
data = data[data['MILKY'].isna() | (data['MILKY'] == 'NONE')]
data = data[(data['LUSTER'].isin(luster)) | (data['LUSTER'].isna())]
data = data[~data['CULET'].isin(culet)]
data = data[~data['FL COLOR'].isin(flourcolor)]
data = data[~data['KEY TO SYMBOL'].isin(key_symbol)]
data = data[~data['REPORT COMMENT'].str.startswith('Clarity',na=False)]
regex_pattern = r'\b(' + '|'.join(map(pd.api.types.infer_dtype, patterns)) + r')\b'
data['INCLUSION'] = data['INCLUSION'].str.replace(regex_pattern, '', regex=True).str.strip()

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
    'LOTNO': ['CERTNO'],
    'CARAT': ['CTS'],
    'SHAPE': ['SHAPE'],
    'COLOR': ['COL'],
    'CLARITY': ['CLAR'],
    'CUT': ['CUT'],
    'POL': ['POL'],
    'SYM': ['SYM'],
    'FLRC': ['FL'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['TAB%'],
    'DEPTH': ['T.D.%'],
    'LAB': ['LAB'],
    'DISC': ['RAP DISC%'],
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
file_path = os.path.join(output_dir, f"sanghvi & sons {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
