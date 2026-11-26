from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="HK", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df)

# Apply filters

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX','VG']
flour = ['N', 'F', 'M']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'] 
# data['Stock No'] = data['Stock No'].str.replace('-', '')

bt = ['BT2','BT3']
bc = ['BC2','BC3']
wt = ['WT3']
wc = ['WC3']
milky = ['M1','M2','M3','NV']
oopv= ['PO2','PO3']
opcr = ['CO2','CO3']
opta = ['NV','TO1','TO2']
culet = ['L','VSM','']
flo_color = ['YL','GR']
luster  =  ['GD','VG']
# Additional filtering
data = data[data['Lab'] == 'GIA']
data = data[data['Shape'] == 'ROUND']
data = data[(data['Carat'] >= 0.30) & (data['Carat'] <= 3.99)]
data = data[data['Location'] == 'MUMBAI']
data = data[data['Rap%'] <= -10]
data = data[data['Color Shade'] == 'WH']
data = data[data['Color'].isin(colors)]
data = data[data['Clarity'].isin(clarity)]
data = data[data['Flo'].isin(flour)]
data['Flo'] = data['Flo'].replace({'F': 'FNT', 'M': 'MED', 'N': 'NON'})
data = data[data['Cut'].isin(cut) & data['Polish'].isin(cut) & data['Symmetry'].isin(cut)]
data = data[data['Color Desc'].isna()]
data = data[~data['BT'].isin(bt)]
data = data[~data['BC'].isin(bc)]
data = data[~data['WT'].isin(wt)]
data = data[~data['WC'].isin(wc)]
data = data[~data['Milky'].isin(milky)]
data = data[~data['OPPV'].isin(oopv)]
data = data[~data['OPCR'].isin(opcr)]
data = data[~data['OPTA'].isin(opta)]
data = data[~data['Culet'].isin(culet)]
data = data[~data['Flo Color'].isin(flo_color)]
data = data[~data['Luster'].isin(luster)]
data = data[~data['Key To Symbol'].isin(key_symbol)]
data = data[~data['GIA Comments'].str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)]
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
    'LOTNO': ['Stock No'],
    'CARAT': ['Carat'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symmetry'],
    'FLRC': ['Flo'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Tab%'],
    'DEPTH': ['TD%'],
    'LAB': ['Lab'],
    'DISC': ['Rap%'],
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

# Remove rows with any blank (NaN) values
output_data = output_data.dropna(how='any')

# Save output data as a CSV file
output_dir = r"G:\Certificate\Python Output\DEAL Output"
current_datetime = datetime.now().strftime("%d.%m.%Y")
file_path = os.path.join(output_dir, f"hk {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
