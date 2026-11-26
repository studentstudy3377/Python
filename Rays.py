from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="Rays", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df,skiprows=1)

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
milky = ['NONE', 'NON']
cut = ['EX', 'VG']
flour = ['NONE', 'FNT', 'MED']
table_black = ['NONE', 'PINPOINT']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'] 
shade = ['NONE']
natts = ['C2','T2','T2C1','T2C2','T3C3']
# Apply filters
data = data[data['Lab'] == 'GIA']
data = data[data['Shape'] == 'ROUND']
data = data[(data['Carat'] >= 0.30) & (data['Carat'] <= 3.99)]
data = data[data['Location'] == 'INDIA']
data = data[data['Disc%'] <= -10]
# Additional filtering
data = data[data['Color'].isin(colors)]
data = data[data['Clarity'].isin(clarity)]
data = data[data['Cut'].isin(cut) & data['Polish'].isin(cut) & data['Symmetry'].isin(cut)]
data = data[data['FL'].isin(flour)]
data['FLUOR'] = data['FL'].replace({'FNT': 'FNT', 'MED': 'MED', 'NONE': 'NON'})
data = data[data['Shade'].isin(shade)]
data = data[data['Milky'].isin(shade)]
data = data[data['T.Open']=='NONE']
data = data[~data['Natts'].isin(natts)]
data = data[~data['Key To Sym'].isin(key_symbol)]

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
    'LOTNO': ['Cert.No.'],
    'CARAT': ['Carat'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symmetry'],
    'FLRC': ['FL'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table%'],
    'DEPTH': ['T.Dep.'],
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
file_path = os.path.join(output_dir, f"rays {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
