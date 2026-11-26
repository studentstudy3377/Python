from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df, skiprows=10)
data.fillna('')
# Apply filters
data = data[data['Lab'] == 'GIA']
data = data[data['Shape'] == 'ROUND']
data = data[(data['CRTWT'] >= 0.30) & (data['CRTWT'] <= 3.99)]
data = data[data['LOCATION'] == 'MUMBAI']
data['RAP DIS'] = data['RAP DIS'].apply(lambda x: -x if x >= 0 else x).apply(lambda x: f"{x if x < 0 else -x:.2f}")
data['RAP DIS'] = pd.to_numeric(data['RAP DIS'], errors='coerce')
data = data[data['RAP DIS'] <= -10]
# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
milky = ['NO','']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
table_black = ['NONE', 'PINPOINT']
culet = ['NON','PTD','']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'] 
natts= ['YT2']

# Additional filtering
data = data[~data['Comments'].str.startswith("Clarity grade is based on internal graining that is not shown", na=False)]

data = data[data['COL'].isin(colors)]
data = data[data['CLR'].isin(clarity)]
data = data[data['Milky'].isin(milky)]
data = data[data['Shade'] == 'NB']
data = data[data['CUT'].isin(cut) & data['POL'].isin(cut) & data['SYM'].isin(cut)]
data = data[data['Flur'].isin(flour)]
data = data[data['Culet'].isin(culet)]
data = data[data['OPTA'].isin(milky)]
data = data[~data['Natts'].isin(natts)]
data = data[~data['KTSVIEW'].isin(key_symbol)]

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
    'CARAT': ['CRTWT'],
    'SHAPE': ['Shape'],
    'COLOR': ['COL'],
    'CLARITY': ['CLR'],
    'CUT': ['CUT'],
    'POL': ['POL'],
    'SYM': ['SYM'],
    'FLRC': ['Flur'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['TBL'],
    'DEPTH': ['DPL'],
    'LAB': ['Lab'],
    'DISC': ['RAP DIS'],
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
file_path = os.path.join(output_dir, f"ankit gems {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
