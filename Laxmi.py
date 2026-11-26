from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime

# Load data
df = askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df, skiprows=5)
data['Disc(%)'] = data['Disc(%)'].apply(lambda x: -x if +x >= 0 else x)
data['Color'] = data['Color'].str.replace('+', '')
data['Color'] = data['Color'].str.replace('-', '')
data['Clarity'] = data['Clarity'].str.replace('+', '')
data['Cut'] = data['Cut'].str.replace('+', '')


# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX', 'VG']
flour = ['NON', 'FNT', 'MED']
black_inc =['TB2','SB2','SB2,TB2']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity','Bruise, Crystal','Brown patch of color, Pinpoint'] 
tinge = ['NO','']
milky = ['M0','']
# Apply filters
data = data[data['Lab'] == 'GIA']
data = data[data['Shape'] == 'RBC']
data['Shape'] = data['Shape'].str.replace('RBC', 'ROUND')
data = data[(data['Carats'] >= 0.30) & (data['Carats'] <= 3.99)]
data = data[data['Country'] == 'India']
data = data[data['Disc(%)'] <= -10]
data = data[~data['Black Inc'].isin(black_inc)]
data = data[data['Tinge'].isin(tinge)]
data = data[data['Milky'].isin(milky)]
data = data[data['Luster'] == 'EX']
data = data[data['Culet'] == 'NON']
data = data[data['Color'].isin(colors)]
data = data[data['Clarity'].isin(clarity)]
data = data[data['Cut'].isin(cut) & data['Polish'].isin(cut) & data['Symm'].isin(cut)]
data = data[data['Fluo Int'].isin(flour)]
data = data[~data['Key To Symbols'].isin(key_symbol)]

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
    'LOTNO': ['Certificate'],
    'CARAT': ['Carats'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symm'],
    'FLRC': ['Fluo Int'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['Table(%)'],
    'DEPTH': ['Depth(%)'],
    'LAB': ['Lab'],
    'DISC': ['Disc(%)'],
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
file_path = os.path.join(output_dir, f"laxmi {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
