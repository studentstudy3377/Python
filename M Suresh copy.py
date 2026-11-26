from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="M Suresh", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df,skiprows=1)

# Apply filters

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2','SI2']
cut = ['EX','VG']
flour = ['NONE', 'FA', 'MD']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'] 
milky = ['EX']
bgm = ['Yes','']
side_black = ['B2-S','B3-S']
table_black = [ 'B2-T']
# Additional filtering
data = data[data['Lab'] == 'GIA']
data = data[data['Shape'] == 'Round']
data['Shape'] = data['Shape'].str.replace('Round', 'ROUND')
data = data[(data['Ct'] >= 0.30) & (data['Ct'] <= 3.99)]
data = data[data['Location'] == 'Mumbai']
data = data[data['Online Back'] <= -10.00]
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
    'DISC': ['Online Back'],
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

# Remove rows with any blank (NaN) values
output_data = output_data.dropna(how='any')


# Save output data as a CSV file
output_dir = r"G:\Certificate\Python Output\DEAL Output"
current_datetime = datetime.now().strftime("%d.%m.%Y")
file_path = os.path.join(output_dir, f"m suresh {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
