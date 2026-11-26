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
cut = ['Ex', 'Vg']
flour = ['N', 'Fnt', 'Med']
key_symbol = ['Cavity','Chip','Cavity, Chip','Chip, Cavity'] 

# Additional filtering
data = data[data['Lab'] == 'GIA']
data = data[data['Shape'] == 'RBC']
data['Shape'] = data['Shape'].str.replace('RBC','ROUND')
data = data[data['Disc%'] <= -10]
data = data[(data['Carat'] >= 0.30) & (data['Carat'] <= 3.99)]
data = data[data['Col'].isin(colors)]
data = data[data['Lust']!='VG']
data = data[data['T.BLK']!='B2']
data = data[data['T.OP']!='HL']
data = data[data['Clarity'].isin(clarity)]
data = data[data['Cut'].isin(cut) & data['Pol'].isin(cut) & data['Sy'].isin(cut)]
data['Cut'] = data['Cut'].replace({'Ex': 'EX', 'Vg': 'VG'})
data['Pol'] = data['Pol'].replace({'Ex': 'EX', 'Vg': 'VG'})
data['Sy'] = data['Sy'].replace({'Ex': 'EX', 'Vg': 'VG'})
data = data[data['Flo'].isin(flour)]
data['Flo'] = data['Flo'].replace({'Fnt': 'FNT', 'Med': 'MED', 'N': 'NON'})
data = data[~data['KeyToSym'].isin(key_symbol)]
data = data[~data['Comment'].str.startswith('Clarity', na=False)]
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
    'CARAT': ['Carat'],
    'SHAPE': ['Shape'],
    'COLOR': ['Col'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Pol'],
    'SYM': ['Sy'],
    'FLRC': ['Flo'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Tabl'],
    'DEPTH': ['Depth'],
    'LAB': ['Lab'],
    'DISC': ['Disc%'],
    'TINCH': [],
    'NATTS': [],
    'RATIO': [],
    'RAP': [],
    'LUSTER': [],
    'T OPENS': [],
    'B OPENS': []
}

output_data = pd.DataFrame()

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
file_path = os.path.join(output_dir, f"pramukh {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
