from datetime import datetime
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# Load data
df = askopenfilename(title="KG", filetypes=[("Excel files", "*.xlsx *.xls")])
data = pd.read_excel(df,skiprows=0)

# Define filter criteria
colors = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
clarity = ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2']
cut = ['EX','VG']
flour = ['NONE', 'FAINT', 'MEDIUM']
key_symbol = ['Cavity', 'Cavity', 'Chip','Cavity/Chip','Chip/Cavity'] 
# data['REPORT NO.'] = data['REPORT NO.'].str.replace('-', '')
data['DISC'] = data['DISC'].apply(lambda x: -x if x >= 0 else x)
luster  =  ['EX','']
shade = ['NONE','']
side_black = ['MEDIUM','CLOUD','HEAVY']
side_spot = ['RED MEDIUM','RED MINOR']
table_spot = ['RED MINOR','GREEN MINOR','RED MEDIUM']
flour_color= ['BLUE,VERY SLIGHT','NONE,VERY SLIGHT','NONE,YELLOW','YELLOW','']

# Additional filtering
data = data[data['LAB'] == 'GIA']
data = data[data['SHAPE'] == 'ROUND']
data = data[(data['CTS'] >= 0.30) & (data['CTS'] <= 3.99)]
data = data[data['LOC.'] == 'Mumbai']
data = data[data['DISC'] <= -10]
data = data[data['COLOR'].isin(colors)]
data = data[data['CLARITY'].isin(clarity)]
data = data[data['FLUOR'].isin(flour)]
data['FLUOR'] = data['FLUOR'].replace({'FAINT': 'FNT', 'MEDIUM': 'MED', 'NONE': 'NON'})
data = data[data['CUT'].isin(cut) & data['POL'].isin(cut) & data['SYMM'].isin(cut)]
data = data[data['SHADE'].isin(shade)]
data = data[data['TABLE OPEN INCLUSION'].isin(shade)]
data = data[data['LUSTER'].isin(luster)]
data = data[~data['FLUORESCENCE COLOR'].isin(flour_color)]
data = data[~data['SIDE BLACK INCLUSION '].isin(side_black)]
data = data[~data['TABLE SPOT INCLUSION'].isin(table_spot)]
data = data[data['TABLE BLACK INCLUSION'].isin(side_black)]
data = data[~data['KEY TO SYMBOL'].isin(key_symbol)]
data = data[~data['REPORT COMMENTS'].str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)]

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
    'LOTNO': ['REPORT NO.'],
    'CARAT': ['CTS'],
    'SHAPE': ['SHAPE'],
    'COLOR': ['COLOR'],
    'CLARITY': ['CLARITY'],
    'CUT': ['CUT'],
    'POL': ['POL'],
    'SYM': ['SYMM'],
    'FLRC': ['FLUOR'],
    'LENGTH': [],
    'WIDTH': [],
    'HEIGHT': [],
    'TABLE': ['TABLE %'],
    'DEPTH': ['DEPTH %'],
    'LAB': ['LAB'],
    'DISC': ['DISC'],
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
file_path = os.path.join(output_dir, f"kg {current_datetime}.csv")

try:
    # Save the DataFrame to a CSV file
    output_data.to_csv(file_path, index=False)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)
except Exception as e:
    print(f"An error occurred: {e}")
