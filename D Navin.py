from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="D Navin", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df, skiprows=1)

# Apply filters

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX','VG']
flour = ['NON', 'FNT', 'MED']
bgm = ['']
key_symbol = ['Cavity', 'Chip','Brown patch of color, Pinpoint, Needle','Cavity, Chip','Chip, Cavity'] 
black= ['B2','B3']
crown = ['B2','B3','CrownCentre']

# Additional filtering
data = data[data['Lab'] == 'GIA']
data = data[data['Shape'] == 'ROUND']
data = data[(data['Ct.'] >= 0.30) & (data['Ct.'] <= 3.99)]
data = data[data['Location'] == 'INDIA']
data = data[data['Disc. (%)'] <= -10]
data = data[data['Col.'].isin(colors)]
data = data[data['Clarity'].isin(clarity)]
data = data[data['Fl.'].isin(flour)]
data = data[data['Cut'].isin(cut) & data['Pol.'].isin(cut) & data['Sym.'].isin(cut)]
data = data[data['Fl.'].isin(flour)]
data = data[data['BGM'].isnull()]
data = data[~data['Key To Symbols'].isin(key_symbol)]
data = data[~data['Centre Black'].isin(black)]
data = data[~data['Crown Black'].isin(crown)]
data = data[~data['Report Comments'].str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)]
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
    'CARAT': ['Ct.'],
    'SHAPE': ['Shape'],
    'COLOR': ['Col.'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Pol.'],
    'SYM': ['Sym.'],
    'FLRC': ['Fl.'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table %'],
    'DEPTH': ['Depth %'],
    'LAB': ['Lab'],
    'DISC': ['Disc. (%)'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': []
}

output_data = pd.DataFrame()

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
file_path = os.path.join(output_dir, f"d navin {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")

