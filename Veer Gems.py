from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df, skiprows=3)

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX','3EX', 'VG']
flour = ['NN', 'FN', 'MED']
key_symbol = ['Chip','Cavity, Chip','Chip, Cavity','Cavity'] 
cbt =[2,3]
ctw = [3]
# Additional filtering
data = data[data['Cert'] == 'GIA']
data = data[data['Loc'] == 'Mum']
data = data[data['Shape'] == 'ROUND']
data['Disc'] = data['Disc'].apply(lambda x: -x if x >= 0 else x)
data = data[data['Disc'] <= -10]
data = data[(data['Cts'] >= 0.30) & (data['Cts'] <= 3.99)]
data = data[data['Col'].isin(colors)]
data = data[data['Clrt'].isin(clarity)]
data = data[data['Cut'].isin(cut) & data['Pol'].isin(cut) & data['Sym'].isin(cut)]
data = data[data['Flr'].isin(flour)]
data['Flr'] = data['Flr'].replace({'FN': 'FNT', 'NN': 'NON'})
data = data[~data['CTB'].isin(cbt)]
data = data[~data['CTW'].isin(ctw)]
data = data[~data['STB'].isin(cbt)]
data = data[~data['STW'].isin(ctw)]
data = data[~data['CRB'].isin(cbt)]
data = data[~data['CRW'].isin(ctw)]
data = data[data['Culet']!= 'OTHER']
data = data[~data['Key To Symbol'].isin(key_symbol)]
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
    'LOTNO': ['Cert#'],
    'CARAT': ['Cts'],
    'SHAPE': ['Shape'],
    'COLOR': ['Col'],
    'CLARITY': ['Clrt'],
    'CUT': ['Cut'],
    'POL': ['Pol'],
    'SYM': ['Sym'],
    'FLRC': ['Flr'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Tbl'],
    'DEPTH': ['T.Dpt'],
    'LAB': ['Cert'],
    'DISC': ['Disc'],
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
file_path = os.path.join(output_dir, f"veer gems {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
