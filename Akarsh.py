from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="Akarsh", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df, skiprows=1)
length = len(data['Report No'].dropna())
data = data.iloc[:length]

# Apply filters
data = data[data['Lab'] == 'GIA']
data = data[data['Shape'] == 'ROUND']
data = data[(data['Cts'] >= 0.30) & (data['Cts'] <= 3.99)]
data = data[data['Location'] == 'INDIA']

# Convert 'Disc %' column to numeric, coercing errors to NaN
data['Disc %'] = pd.to_numeric(data['Disc %'], errors='coerce')
data = data[~data['Disc %'].between(-10, -1, inclusive='both')]
# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
milky = ['NONE']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
culet = ['NON']
key_symbol = ['Cavity', 'Chip']
green = ['NONE']
black_in = ['B2', 'B3']

# Additional filtering
data = data[data['Color'].isin(colors)]
data = data[data['Purity'].isin(clarity)]
data = data[data['Milky'].isin(milky)]
data = data[data['Tinge'].isin(green)]
data = data[data['Brown'].isin(green)]
data = data[data['Cut'].isin(cut) & data['Polish'].isin(cut) & data['Symm'].isin(cut)]
data = data[data['Fls'].isin(flour)]
data = data[data['Culet'].isin(culet)]
data = data[data['Green'].isin(green)]
data = data[~data['Key to symbols'].isin(key_symbol) & ~data['Black Incl.'].isin(black_in)]

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
    'CARAT': ['Cts'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Purity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symm'],
    'FLRC': ['Fls'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table %'],
    'DEPTH': ['Depth %'],
    'LAB': ['Lab'],
    'DISC': ['Disc %'],
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
file_path = os.path.join(output_dir, f"akarsh export {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
