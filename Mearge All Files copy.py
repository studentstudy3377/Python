from datetime import datetime
from tkinter import filedialog
from tkinter.filedialog import askdirectory
import pandas as pd
import numpy as np
import os
import glob

def load_data(file_path, skip_rows=0):
    data = pd.read_excel(file_path, skiprows=skip_rows)
    return data

def apply_filters(data, filters):
    for column, criteria in filters.items():
        if isinstance(criteria, list):
            data = data[data[column].isin(criteria)]
        elif callable(criteria):
            data = data[criteria(data[column])]
    return data

def concatenate_dataframes(df, column_names):
    concatenated = []
    for sub_df in df.values():
        for column_name in column_names:
            if sub_df is not None and column_name in sub_df.columns:
                concatenated.append(sub_df[column_name])
                break
    concatenated = [series for series in concatenated if not series.empty]
    return pd.concat(concatenated, ignore_index=True) if concatenated else pd.Series(dtype='str')

def map_columns(data, column_mapping):
    output_data = pd.DataFrame()
    for column, keys in column_mapping.items():
        if column in ['LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP']:
            output_data[column] = 0
        elif column == 'SRNO':
            output_data[column] = range(1, len(data) + 1)
        elif column in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
            output_data[column] = 'N/A'
        elif column == 'SHAPE':
            output_data[column] = 'ROUND'
        elif column == 'NATTS':
            output_data[column] = 'CENTER - L'
        elif not keys:
            output_data[column] = np.nan
        else:
            if column == 'DISC':
                output_data[column] = concatenate_dataframes({0: data}, keys).round(2).apply(lambda x: -x if x > 0 else x)
            elif column == 'FLRC':
                output_data[column] = concatenate_dataframes({0: data}, keys).replace(
                    {'F': 'FNT', 'M': 'MED', 'N': 'NON', 'FAINT': 'FNT', 'NONE': 'NON', 'MEDIUM': 'MED',
                     'Fnt': 'FNT', 'Med': 'MED', ' F': 'FNT', ' N': 'NON', ' M': 'MED', 'FA': 'FNT', 'MD': 'MED', 'NN': 'NON', 'FN': 'FNT'})
            elif column in ['CUT', 'POL', 'SYM']:
                output_data[column] = concatenate_dataframes({0: data}, keys).replace({'Vg': 'VG', 'Ex': 'EX'})
            else:
                output_data[column] = concatenate_dataframes({0: data}, keys)
    return output_data

def save_output(output_data, output_dir, prefix):
    current_datetime = datetime.now().strftime("%d.%m.%Y")
    file_path = os.path.join(output_dir, f"{prefix.lower()} {current_datetime}.csv")

    try:
        output_data.to_csv(file_path, index=False)
        print(f"File successfully saved at: {file_path}")
        os.startfile(file_path)
    except Exception as e:
        print(f"An error occurred while saving CSV: {e}")

def process_data(file_path, filters, column_mapping, output_dir, prefix, skip_rows=0):
    print(f"Processing {prefix} data for file: {file_path}")
    data = load_data(file_path, skip_rows=skip_rows)
    if 'Stock No' in data.columns:
        data['Stock No'] = data['Stock No'].str.replace('-', '')
    if 'Rapnet  Discount %' in data.columns:
        data['Rapnet  Discount %'] = (data['Rapnet  Discount %'] * 100).apply(lambda x: f"{x if x < 0 else -x:.2f}")
        data['Rapnet  Discount %'] = pd.to_numeric(data['Rapnet  Discount %'], errors='coerce')
    filtered_data = apply_filters(data, filters)
    if prefix == 'KGK' and 'Size' in filtered_data.columns:
        filtered_data.rename(columns={'Size': 'Updated_Size'}, inplace=True)

    output_data = map_columns(filtered_data, column_mapping)
    save_output(output_data, output_dir, prefix)

def main():
    files = filedialog.askopenfilenames(
        title="Select Excel or CSV Files",
        filetypes=[("Excel/CSV Files", "*.xlsx *.xls *.csv"), ("All Files", "*.*")]
    )

    excel_files = list(files)

    if not excel_files:
        print("No files selected. Exiting.")
        return

    print(f"Detected {len(excel_files)} files. Processing...")

    filters_akarsh = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['INDIA'],
        'Disc %': lambda x: x <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Purity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Milky': ['NONE'],
        'Tinge': ['NONE'],
        'Brown': ['NONE'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symm': ['EX', 'VG'],
        'Fls': ['NON', 'FNT', 'MED'],
        'Culet': ['NON'],
        'Green': ['NONE'],
        'Key to symbols': lambda x: ~x.isin(['Cavity', 'Chip']),
        'Black Incl.': lambda x: ~x.isin(['B2', 'B3'])
    }

    filters_ankit = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'CRTWT': lambda x: (x >= 0.30) & (x <= 3.99),
        'LOCATION': ['MUMBAI'],
        'RAP DIS': lambda x: x <= -10,
        'COL': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLR': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Milky': ['NO', ''],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'Flur': ['NON', 'FNT', 'MED'],
        'Culet': ['NON', 'PTD', ''],
        'OPTA': ['NO',''],
        'KTSVIEW': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity']),
        'Natts': lambda x: ~x.isin(['YT2']),
        'Shade': ['NB'],
        'Comments': lambda x: ~x.str.startswith(
            "Clarity grade is based on internal graining that is not shown", na=False
        )
    }

    filters_amora = {
        'LAB': ['GIA'],
        'Shape': ['RBC'],
        'DISC.': lambda x: pd.to_numeric(x, errors='coerce').apply(lambda val: val <= -10),
        'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clar': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fluor': ['NON', 'FNT', 'MED'],
        'Culet Size': lambda x: x != 'VSM',
        'Milkey': lambda x: x.isna() | (x == 'None '),
        'Shade': lambda x: x.isna(),
        'Shape Description': lambda x: x.isna() | (x == 'Round Brilliant'),
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity']),
        'Report Comments': lambda x: ~x.str.startswith('Clarity', na=False)
    }

    filters_asian_star = {
        'Lab': ['GIA'],
        'Shape': ['RD'],
        'Disc': lambda x: x <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['IND'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Final Cut': ['EX', 'VG', '3EX'],
        'Polish': ['EX', 'VG', '3EX'],
        'Sym.': ['EX', 'VG', '3EX'],
        'Fluor.': ['NON', 'FNT', 'MED'],
        'Culet': ['NON', 'NONE', 'POINTED'],
        'Tinge': lambda x: x.isna() | (x == 'NN'),
        'Milky': ['M0'],
        'BIT': lambda x: x != 'B2',
        'BIS': lambda x: x != 'B2',
        'TI': lambda x: x != 'W3',
        'Origin': lambda x: x != 'RUSSIA',
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
    }
    filters_asian_star_web = {
        'Lab': ['GIA'],
        'Shape': ['RD'],
        'Disc': lambda x: x <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['IND'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Final Cut': ['EX', 'VG', '3EX'],
        'Polish': ['EX', 'VG', '3EX'],
        'Sym.': ['EX', 'VG', '3EX'],
        'Fluor.': ['NON', 'FNT', 'MED'],
        'Culet': ['NON', 'NONE', 'POINTED'],
        'Tinge': lambda x: x.isna() | (x == 'NN'),
        'Milky': ['M0'],
        'BIT': lambda x: x != 'B2',
        'BIS': lambda x: x != 'B2',
        'TI': lambda x: x != 'W3',
        'Origin': lambda x: x != 'RUSSIA',
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
    }

    filters_d_navin = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Ct.': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['INDIA'],
        'Disc. (%)': lambda x: x <= -10,
        'Col.': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Fl.': ['NON', 'FNT', 'MED'],
        'Cut': ['EX', 'VG'],
        'Pol.': ['EX', 'VG'],
        'Sym.': ['EX', 'VG'],
        'BGM': lambda x: x.isnull(),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip','Brown patch of color, Pinpoint, Needle','Cavity, Chip','Chip, Cavity']),
        'Centre Black': lambda x: ~x.isin(['B2', 'B3']),
        'Crown Black': lambda x: ~x.isin(['B2', 'B3','CrownCentre']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)
    }

    filters_d_navin_web = {
        'LAB': ['GIA'],
        'Shape': ['ROUND'],
        'Ct.': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['INDIA'],
        'Disc. (%)': lambda x: x <= -10,
        'Col.': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Fl.': ['NON', 'FNT', 'MED'],
        'Cut': ['EX', 'VG'],
        'Pol.': ['EX', 'VG'],
        'Sym.': ['EX', 'VG'],
        'BGM': lambda x: x.isnull(),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip','Brown patch of color, Pinpoint, Needle','Cavity, Chip','Chip, Cavity']),
        'Centre Black': lambda x: ~x.isin(['B2', 'B3']),
        'Crown Black': lambda x: ~x.isin(['B2', 'B3','CrownCentre']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)
    }

    filters_finstar = {
        'Lab': ['GIA'],
        'Shape': ['RD'],
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['IND'],
        'Disc %': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fls': ['NON', 'FNT', 'MED'],
        'Culet': ['NONE', 'POINTED', ''],
        'Shade': ['WHT'],
        'Brl.': ['EX', ''],
        'CN': lambda x: x != 'CN3',
        'SN': lambda x: x != 'SN3',
        'CW': lambda x: x != 'CW3',
        'SW': lambda x: x != 'SW3',
        'To': lambda x: x != 'Y',
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity, Chip', 'Chip, Cavity']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity", na=False)
    }

    filters_glowstar = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['India'],
        'Disc %': lambda x: pd.to_numeric(x, errors='coerce') <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Natts': lambda x: ~x.isin(['BC0 - BT2','BC0 - BT3','BC1 - BT3','BC2 - BT0','BC2 - BT1','BC2 - BT2','BC2 - BT3','BC3 - BT0','BC3 - BT2','BC3 - BT3']),
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'CULET': ['NON','Pointed'],
        'Shade': lambda x: x == 'NO',
        'BGM Comment': lambda x: x == 'No BROWN NO MILKY',
        'Comment': lambda x: ~x.str.startswith("Clar", na=False),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'])
    }

    filters_gyanimpex = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Net %': lambda x: -x if x >= 0 else x,
        'Net %': lambda x: x <= -10,
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['MUMBAI'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fluo': ['NON', 'FNT', 'MED'],
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comments': lambda x: ~x.str.startswith('Clarity', na=False)
    }

    filters_haridas = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Disc %': lambda x: x <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
    }

    filters_hk = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['MUMBAI'],
        'Rap%': lambda x: x <= -10,
        'Color Shade': ['WH'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Flo': ['N', 'F', 'M'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symmetry': ['EX', 'VG'],
        'Color Desc': lambda x: x.isna(),
        'BT': lambda x: ~x.isin(['BT2', 'BT3']),
        'BC': lambda x: ~x.isin(['BC2', 'BC3']),
        'WT': lambda x: ~x.isin(['WT3']),
        'WC': lambda x: ~x.isin(['WC3']),
        'Milky': lambda x: ~x.isin(['M1', 'M2', 'M3', 'NV']),
        'OPPV': lambda x: ~x.isin(['PO2', 'PO3']),
        'OPCR': lambda x: ~x.isin(['CO2', 'CO3']),
        'OPTA': lambda x: ~x.isin(['NV', 'TO1', 'TO2']),
        'Culet': lambda x: ~x.isin(['L', 'VSM', '']),
        'Flo Color': lambda x: ~x.isin(['YL', 'GR']),
        'Luster': lambda x: ~x.isin(['GD', 'VG']),
        'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'GIA Comments': lambda x: ~x.str.startswith("Clarity", na=False)
    }
    
    filters_hk_web = {
        'LAB': ['GIA'],
        'Shape': ['ROUND'],
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['MUMBAI'],
        'Rap%': lambda x: x <= -10,
        'Color Shade': ['WH'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Flo': ['N', 'F', 'M'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symmetry': ['EX', 'VG'],
        # 'Color Desc': lambda x: x.isna(),
        'BT': lambda x: ~x.isin(['BT2', 'BT3']),
        'BC': lambda x: ~x.isin(['BC2', 'BC3']),
        'WT': lambda x: ~x.isin(['WT3']),
        'WC': lambda x: ~x.isin(['WC3']),
        'MILKY': lambda x: ~x.isin(['M1', 'M2', 'M3', 'NV']),
        'OPPV': lambda x: ~x.isin(['PO2', 'PO3']),
        'OPCR': lambda x: ~x.isin(['CO2', 'CO3']),
        'OPTA': lambda x: ~x.isin(['NV', 'TO1', 'TO2']),
        'CULET': lambda x: ~x.isin(['L', 'VSM', '']),
        # 'Flo Color': lambda x: ~x.isin(['YL', 'GR']),
        # 'Luster': lambda x: ~x.isin(['GD', 'VG']),
        'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'GIA Comments': lambda x: ~x.str.startswith("Clarity", na=False)
    }

    filters_hrv = {
        'Lab': ['GIA'],
        'Shape': ['BR'],
        'Location': ['MUMBAI'],
        'Disc %': lambda x: x <= -10,   
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'Culet Size': ['NON']
    }
    
    filters_hrv_mail = {
        'Lab': ['GIA'],
        'Shape': ['BR'],
        'Location': ['MUMBAI'],
        'Disc %': lambda x: x <= -10,   
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'Culet Size': ['NON']
    }

    filters_jewelex = {
        'Lab': ['GIA'],
        'Shape': ['RBC'],
        'Disc %': lambda x: x <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }
    
    filters_jodhani = {
        'Lab': ['GIA'],
        'Location': ['India'],
        'Shape': ['Round'],
        # 'Disc(%)': lambda x: -x if x >= 0 else x,
        # 'Disc(%)': lambda x: x <= -10,
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symmetry': ['EX', 'VG'],
        'Fluorescence': ['Faint', 'Medium','NON'],
        'Brown':['N'],
        'Culet':['N'],
        'Luster':['EX'],
        'Table Open':['N',''],
        'Table Black': lambda x: ~x.isin(['B2','HVY']),
        'Side Black': lambda x: ~x.isin(['B2','HVY']),
        'Red Inclusion': lambda x: ~x.isin(['T-PP']),
        'Key Symbol': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity']),
    }

    filters_kbs = {
        'Lab': ['GIA'],
        'Shape': ['BR'],
        'DISC': lambda x: x <= -10,
        'LOC': ['INDIA'],
        'SIZE': lambda x: (x >= 0.30) & (x <= 3.99),
        'COL': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLR': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'FLR': ['NON', 'FNT', 'MED'],
        'BIC': lambda x: x != 'Heavy',
        'BIS': lambda x: x != 'Heavy',
        'KTS': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_kbs_web = {
        'LAB': ['GIA'],
        'SHAPE': ['Round'],
        'DISC': lambda x: x <= -10,
        'LOC': ['INDIA'],
        'WEIGHT': lambda x: (x >= 0.30) & (x <= 3.99),
        'COL': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLR': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'FLR': ['N', 'F', 'M'],
        'BIC': lambda x: x != 'Heavy',
        'BIS': lambda x: x != 'Heavy',
        'KTS': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_kc = {
        'Cert': ['GIA'],
        'Shape': ['Round'],
        'Dis %': lambda x: x <= -10,
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Carats': lambda x: (x >= 0.3) & (x <= 3.9),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fluo': ['N', 'F', 'MED'],
        'Culet': lambda x: x.isna(),
        'Comments': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity", na=False),
    }
    filters_kc_web = {
        'Cert': ['GIA'],
        'Shape': ['Round'],
        'Dis %': lambda x: x <= -10,
        'Carats': lambda x: (x >= 0.3) & (x <= 3.9),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fluo': ['N', 'F', 'MED'],
    }

    filters_kg = {
        'LAB': ['GIA'],
        'SHAPE': ['ROUND'],
        'CTS': lambda x: (x >= 0.30) & (x <= 3.99),
        'LOC.': ['Mumbai'],
        'DISC':lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'COLOR': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLARITY': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'FLUOR': ['NONE', 'FAINT', 'MEDIUM'],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYMM': ['EX', 'VG'],
        'SHADE': ['NONE', ''],
        'TABLE OPEN INCLUSION': ['NONE', ''],
        'LUSTER': ['EX', ''],
        'FLUORESCENCE COLOR': lambda x: ~x.isin(['BLUE,VERY SLIGHT', 'NONE,VERY SLIGHT', 'NONE,YELLOW', 'YELLOW', '']),
        'SIDE BLACK INCLUSION ': lambda x: ~x.isin(['MEDIUM', 'CLOUD', 'HEAVY']),
        'TABLE SPOT INCLUSION': lambda x: ~x.isin(['RED MINOR', 'GREEN MINOR', 'RED MEDIUM']),
        'TABLE BLACK INCLUSION': lambda x: x.isin(['MEDIUM', 'CLOUD', 'HEAVY']),
        'KEY TO SYMBOL': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity/Chip', 'Chip/Cavity']),
        'REPORT COMMENTS': lambda x: ~x.str.startswith("Clarity", na=False),
    }

    filters_kgk = {
        'Lab': ['GIA'],
        'Shape': ['BR'],
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'BACK': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Milky': ['M0'],
        'FLUO': ['NON', 'FNT', 'MED'],
        'Culet': lambda x: ~x.isin(['Long', 'Very Small']),
        'Color Shade': ['White'],
        'Luster': ['EX'],
        'Table Inclusion': lambda x: x != 'T3',
        'Side Inclusion': lambda x: x != 'S3',
        'Side Black': lambda x: ~x.isin(['SB2', 'SB3']),
        'Table Black': lambda x: ~x.isin(['B2', 'B3']),
        'CUT': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_lakhi = {
        'Lab': ['GIA'],
        'Shape': ['RBC'],
        'Country': ['INDIA'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Fluorescence Intensity': ['NON', 'FNT', 'MED'],
        'Culet Size': ['NON'],
        'Shade ': ['NONE'],
        'Milky': ['NONE'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symmetry': ['EX', 'VG'],
        'Rapnet  Discount %': lambda x: x <= -10,
        'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
        'Key to symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Black Inclusion   ': lambda x: ~x.str.startswith("N3,", na=False),
        'Cert comment': lambda x: ~x.str.startswith("Clarity", na=False)
    }
    
    filters_lakhi_mail = {
        'Lab': ['GIA'],
        'Shape': ['RBC'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Fluorescence Intensity': ['NON', 'FNT', 'MED'],
        'Culet Size': ['NON'],
        'Shade ': ['NONE'],
        'Milky': ['NONE'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symmetry': ['EX', 'VG'],
        'Rapnet  Discount %': lambda x: x <= -10,
        'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
        'Key to symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Black Inclusion   ': lambda x: ~x.str.startswith("N3,", na=False),
        'Cert comment': lambda x: ~x.str.startswith("Clarity", na=False)
   }

    filters_m_d_brother= {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Loc': ['MUMBAI'],
        'Disc %': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symm': ['EX', 'VG'],
        'Fls': ['NON', 'FNT', 'MED'],
        'Shade': lambda x: (x == 'NONE') | (x.isna()),
        'Milky': lambda x: (x == 'NONE') | (x.isna()),
        'Culet': lambda x: ~x.isin(['Very Small', 'VSM']),
        'Black': lambda x: ~x.isin(['SPPB2', 'CPPB2']),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comments': lambda x: ~x.str.startswith(('Clarity', '* This diamond has been'), na=False),
        'Grading Comments': lambda x: x.str.startswith('NO BGM', na=False)
    }

    filters_m_suresh = {
        'Lab': ['GIA'],
        'Shape': ['Round'],
        'Ct': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['Mumbai'],
        'Online Back': lambda x: x <= -10.00,
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clr': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fls': ['NONE', 'FA', 'MD'],
        'BGM': lambda x: ~x.isin(['Yes', '']),
        'brown': lambda x: x.isna(),
        'Tab_BLK': lambda x: ~x.isin(['B2-T']),
        'Op_Tab': lambda x: x.isna(),
        'Culet': lambda x: x == 'NONE',
        'Milky/Luster': ['EX'],
        'Side_BLK': lambda x: ~x.isin(['B2-S', 'B3-S']),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)
    }

    filters_mb = {
            'Cert Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['MUM'],
            'Disc': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Fluo': ['NON', 'FNT', 'MED'],
            'Shade': ['NN'],
            'Table Black': ['NN', 'BT1', 'BT1+'],
            'Side Black': ['NN', 'BC1', 'BC1+'],
            'Open': lambda x: ~x.isin(['OT1']),
            'Culet': lambda x: ~x.isin(['VERY SMALL']),
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_narola = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Location': ['INDIA'],
            'Disc%': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Flo': ['NON', 'FNT', 'MED'],
            'TableOpen': lambda x: x != 'SMALL-OPEN',
            'KeyToSym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
            'NGS Comment': lambda x: ~x.str.startswith('CULET SMALL', na=False)
        }

    filters_pansurya_impex = {
        'Lab': ['GIA'],
        'Shape': ['RD'],
        'Location': ['MUMBAI'],
        'Dis': lambda x: x <= -10,
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flor.': ['NON', 'FNT', 'MED'],
        'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comment': lambda x: ~x.str.startswith('Clarity', na=False)
    }

    filters_pragati = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Disc %': lambda x: x <= -10,
            'Size': lambda x: (x >= 0.30) & (x <= 3.99),
            'Loc': ['IN'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symm': ['EX', 'VG'],
            'Fls': ['NON', 'FNT', 'MED'],
            'Culet': lambda x: x != 'VSM',
            'Table Black': lambda x: x != 'TB-2',
            'Side Black': lambda x: x != 'SB-2',
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_pramukh = {
        'Lab': ['GIA'],
        'Shape': ['RBC'],
        'Disc%': lambda x: x <= -10,
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Lust': lambda x: x != 'VG',
        'T.BLK': lambda x: x != 'B2',
        'T.OP': lambda x: x != 'HL',
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['Ex', 'Vg'],
        'Pol': ['Ex', 'Vg'],
        'Sy': ['Ex', 'Vg'],
        'Flo': ['N', 'Fnt', 'Med'],
        'KeyToSym': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comment': lambda x: ~x.str.startswith('Clarity', na=False)
    }

    filters_rn_export = {
        'LAB': ['GIA'],
        'SHAPE': ['RBC'],
        'DISC': lambda x: (x * 100).apply(lambda y: f"{y if y < 0 else -y:.2f}").astype(float) <= -10,
        'LOC': ['IND'],
        'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', ' EX', 'EX ', 'VG', ' VG', 'VG '],
        'Polish': ['EX', ' EX', 'EX ', 'VG', ' VG', 'VG '],
        'Symm': ['EX', ' EX', 'EX ', 'VG', ' VG', 'VG '],
        'Fluor': ['N', 'F', ' F', 'M', ' M', ' N'],
        'Key to Sym': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
    }

    filters_r_rajesh = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['INDIA'],
            'Disc %': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Fls': ['NONE', 'FNT', 'MED'],
            'Tinge': ['NONE', ''],
            'TB': lambda x: ~x.isin(['TBL2', 'TBL3']),
            'SB': lambda x: ~x.isin(['SBL2', 'SBL3']),
            'TW': lambda x: ~x.isin(['CW3']),
            'SW': lambda x: ~x.isin(['SW3']),
            'Culet': lambda x: ~x.isin(['VERY SMALL']),
            'Milky': ['NONE', ''],
            'Open On Table': lambda x: ~x.isin(['MEDIUM TABLE OPEN', 'SMALL TABLE OPEN']),
            'Report Comments': lambda x: ~x.str.startswith('Clarity', na=False)
        }

    filters_rays = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['INDIA'],
            'Disc%': lambda x: x <= -10,
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'FL': ['NONE', 'FNT', 'MED'],
            'Shade': ['NONE'],
            'Milky': ['NONE', 'NON'],
            'T.Open': ['NONE'],
            'Natts': lambda x: ~x.isin(['C2', 'T2', 'T2C1', 'T2C2', 'T3C3']),
            'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }
    
    filters_rays_mails = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Wgt': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['MUMBAI'],
            'Disc %': lambda x: x <= -10,
            'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Fls': ['NONE', 'FNT', 'MED'],
            'BLACK': ['T1','T2','C1','C2','C2','T2','T2C1','T2C2','T3C3','NO BLACK'],
            'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_red_exim = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Disc': lambda x: x.apply(lambda y: -y if y >= 0 else y),
            'Disc': lambda x: x <= -10,
            'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
            'Country': ['INDIA'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'CutGrade': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'Flour': ['NON', 'FNT', 'MED'],
            'CuletSize': ['NONE', 'POINTED'],
            'Milky': lambda x: x.isna() | (x == 'NONE'),
            'Shade': lambda x: x.isna() | (x == 'NONE'),
            'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
            'Comment': lambda x: x.isin(['NO BGM ', 'NO BGM', 'NO BGM  '])
        }

    filters_sagar_list = {
        'Lab': ['GIA'],
        'Location': ['Mumbai', 'MUMBAI'],
        'Shape': ['RD'],
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Back%':lambda x: x <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'FL': ['NON', 'FNT', 'MED'],
        'KeytoSymbol': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comments': lambda x: ~x.str.startswith('Clarity', na=False),
        'Back%': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10
    }

    filters_shairu = {
        'Lab': ['GIA '],
        'Shape': ['ROUND'],
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['India'],
        'Disc(%)': lambda x: x <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['3EX', 'EX', 'VG'],
        'Polish': ['3EX', 'EX', 'VG'],
        'Symm': ['3EX', 'EX', 'VG'],
        'Fls': ['NON', 'FNT', 'MED'],
        'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'GIA Comments': lambda x: ~x.str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)
    }

    filters_sheetal = {
            'Lab': ['GIA'],
            'Shape': ['RBC'],
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['INDIA'],
            'Disc%': lambda x: x <= -10,
            'Hue': ['WH'],
            'Milky': ['NO'],
            'BIT': lambda x: x != 'BIT2',
            'BIC': lambda x: x != 'BIC2',
            'BOC': lambda x: x != 'BOC2',
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'Fluorescence': ['NON', 'FNT', 'MED'],
            'Keytosymbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_shyam = {
            'Lab': ['GIA'],
            'City': ['MUMBAI'],
            'Shape': ['ROUND'],
            'Disc.': lambda x: -x if x >= 0 else x,
            'Disc.': lambda x: x <= -10,
            'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'Fl.Int.': ['NONE', 'FAINT', 'MEDIUM'],
            'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
            'Cert comment': lambda x: x.str.startswith('NO BGM', na=True) | x.str.startswith(' NO BGM', na=True) | x.str.startswith('NO BMG', na=True) | x.str.startswith(' NO BMG', na=True)
        }

    filters_ske = {
            'Lab': ['GIA'],
            'Shape': ['RBC'],
            'Disc%': lambda x: x <= -10,
            'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['MUMBAI'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symm': ['EX', 'VG'],
            'Flour.': ['NON', 'FNT', 'MED'],
            'Luster': lambda x: x != 'VG',
            'Table Open': lambda x: x == 'N',
            'Table Black': lambda x: x != 'TB2',
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_snj_list = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Final Disc%': lambda x: x <= -10,
            'Size': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['INDIA'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Table Black': lambda x: ~x.isin(['TN2', 'TN3']),
            'Side Black': lambda x: ~x.isin(['SN2', 'SN3']),
            'Luster': lambda x: x != 'VG',
            'Flo': ['NON', 'FNT', 'MED'],
            'Flo Color': lambda x: x != 'Yellow',
            'Table \nOpen': lambda x: (x == 'N') | (x.isna()),
            'Culet Size': lambda x: (x == 'NONE') | (x.isna()),
            'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_srd = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        # 'Rap %': lambda x: x <= 10,
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Loc': ['MUMBAI'],
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Cla': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'FL': ['N', 'F', 'M'],
        'CS': ['WH'],
        'BT': lambda x: x != 'BT2',
        'BC': lambda x: x != 'BC2',
        'TO': ['TO0'],
        'FlShd': lambda x: x != 'Yellow',
        'Milky': ['M0'],
        'Luster': ['EX'],
        'Fancy Color Description': lambda x: x.isna(),
        'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
    }

    filters_sunrays = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Disc %': lambda x: -x if +x >= 0 else x,
            'Disc %': lambda x: x <= -10,
            'Size': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['IN'],
            'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symm': ['EX', 'VG'],
            'Fls': ['NON', 'FNT', 'MED'],
            'Shade': ['WHT'],
            'Milky': lambda x: x != 'M01',
            'B I': lambda x: x != 'B02',
            'T I': lambda x: x != 'T02',
            'Culet': lambda x: ~x.isin(['VSM', 'Very Small']),
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
            'Report Comments': lambda x: ~x.str.startswith('Clarity', na=False)
        }

    filters_vd_global = {
        'LAB': ['GIA'],
        'Shape': ['RD'],
        'CTS': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['India'],
        'Disc%': lambda x: x <= -10,
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Cla': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flo': ['NON', 'FNT', 'MED'],
        'Color Shade': lambda x: x.isna(),
        'MILKY': lambda x: x.isna(),
        'Fancy Color': lambda x: x.isna(),
        'CULET': lambda x: x != 'VSM',
        'BT': lambda x: ~x.isin(['H']),
        'BC': lambda x: ~x.isin(['H']),
        'BCR': lambda x: ~x.isin(['H']),
        'BP': lambda x: ~x.isin(['H']),
        'WT': lambda x: ~x.isin(['H']),
        'WC': lambda x: ~x.isin(['H']),
        'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Cavity, Chip', 'Chip, Cavity']),
        'GIA Comments': lambda x: ~x.str.startswith("Clarity", na=False)
    }

    filters_veer_gems = {
            'Cert': ['GIA'],
            'Loc': ['Mum'],
            'Shape': ['ROUND'],
            'Disc': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
            'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clrt': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', '3EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Flr': ['NN', 'FN', 'MED'],
            'CTB': lambda x: ~x.isin([2, 3]),
            'CTW': lambda x: ~x.isin([3]),
            'STB': lambda x: ~x.isin([2, 3]),
            'STW': lambda x: ~x.isin([3]),
            'CRB': lambda x: ~x.isin([2, 3]),
            'CRW': lambda x: ~x.isin([3]),
            'Culet': lambda x: x != 'OTHER',
            'Key To Symbol': lambda x: ~x.isin(['Chip', 'Cavity, Chip', 'Chip, Cavity', 'Cavity'])
        }
    
    filters_unique = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            ' Disc %': lambda x: x <= -10,
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut ': ['EX', 'VG'],
            ' Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Fls': ['NON', 'FNT', 'MED'],
            'TI': lambda x: x == 'NON',
            'TB': ['MINOR BLACK','NON','PINPOINT '],
            'SB': ['MINOR BLACK','NON','PINPOINT '],
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    column_mapping = {
        'SRNO': [],
        'LOTNO': ['Stock No', 'Cert#', 'Report No', 'Cert No', 'Certificate No', 'Report No.', 'Stock No', 'Packet No', 'CertNo', 'Cert No.', 'REPORT NO.', 'Report #', 'Control No', 'ReportNo', 'Report', 'Certificate', 'Cert.No.', 'Report#', 'Certi No.'],
        'CARAT': ['Ct.','Cts', 'Wgt', 'CRTWT', 'Weight', 'Size', 'Carats', 'Size', 'Carat', 'SIZE', 'CTS', 'Ct'],
        'SHAPE': [],
        'COLOR': ['Col.','Color', 'COL', 'Col', 'COLOR'],
        'CLARITY': ['Purity', 'Cla', 'CLR', 'Clar', 'Clarity', 'CLARITY', 'Clr', 'Cts','WEIGHT'],
        'CUT': ['Cut ','Cut', 'CUT', 'Final Cut', 'CutGrade'],
        'POL': ['Pol.',' Pol','Polish', 'POL', 'Pol'],
        'SYM': ['Symm', 'SYM', 'Sym', 'Sym.', 'Symmetry', 'SYMM', 'Sy'],
        'FLRC': ['Fl.','Fls', 'Flur', 'Fluor', 'Fluor.', 'Flour', 'Fluo', 'Flo', 'FLR', 'FLUOR', 'FLUO', 'Fluorescence Intensity', 'Fluo Int', 'Flor.', 'FL', 'Fluorescence', 'Fl.Int.', 'Flour.', 'Flr'],
        'LENGTH': [],
        'WIDTH': [],
        'HEIGHT': [],
        'TABLE': ['Table %', 'TBL', 'Table%', 'Tab%', 'TA%', 'T.Size', 'TABLE %', 'Table', 'Table(%)', 'Tabl', 'TAB%', 'Tbl'],
        'DEPTH': ['Depth %', 'DPL', 'Depth%', 'TD%', 'T Depth', 'DEPTH %', 'Depth Percentage', 'Depth(%)', 'Depth', 'DEPTH', 'T.Dep.', 'T.Dpt'],
        'LAB': ['Lab', 'LAB', 'Cert', 'Cert Lab'],
        'DISC': ['Disc. (%)',' Disc %','Disc %', 'Rap %', 'RAP DIS', 'DISC.', 'Disc', 'Net %', 'Rap%', 'DISC', 'Dis %', 'BACK', 'Rapnet  Discount %', 'Disc(%)', 'Disc%', 'Dis', 'Back%', 'Dis%', 'Disc. ', 'Final Disc%', 'Online Back'],
        'TINCH': [],
        'NATTS': [],
        'RATIO': [],
        'RAP': [],
        'LUSTER': [],
        'T OPENS': [],
        'B OPENS': []
    }

    prefix_info = {
        
        
        
        "AKARSH": {"filters": filters_akarsh, "skip_rows": 1},
        "ANKIT":  {"filters": filters_ankit,  "skip_rows": 10},   
        "AMORA":  {"filters": filters_amora,  "skip_rows": 0},
        "ASIAN STAR":[ {"filters": filters_asian_star, "skip_rows": 7},
                       {"filters": filters_asian_star_web, "skip_rows": 0}],
        "D NAVIN": [
                {"filters": filters_d_navin, "skip_rows": 1},
                {"filters": filters_d_navin_web, "skip_rows":0 }],
        "FINSTAR": {"filters": filters_finstar, "skip_rows": 1},
        "GLOWSTAR": {"filters": filters_glowstar, "skip_rows": 4},
        "GNYAN IMPEX": {"filters": filters_gyanimpex, "skip_rows": 1},
        "HIRADIAS LIST": {"filters": filters_haridas, "skip_rows": 0},
        "HK": [{"filters": filters_hk, "skip_rows": 0},
                {"filters": filters_hk_web, "skip_rows": 1}],
        "HRV": [{"filters": filters_hrv, "skip_rows": 0},
                {"filters": filters_hrv_mail, "skip_rows": 1}],
        "JEWELEX": {"filters": filters_jewelex, "skip_rows": 1},
        "JODHANI": {"filters": filters_jodhani, "skip_rows": 0},
        "KBS":[{"filters": filters_kbs, "skip_rows": 2},
                {"filters": filters_kbs_web, "skip_rows": 0}],
        "KC": [{"filters": filters_kc, "skip_rows": 0},
               {"filters": filters_kc_web, "skip_rows": 5}],
        "KG": {"filters": filters_kg, "skip_rows": 1}, 
        "KGK": {"filters": filters_kgk, "skip_rows": 1},
        "LAKHI": [ {"filters": filters_lakhi, "skip_rows": 0},
                   {"filters": filters_lakhi_mail, "skip_rows": 5}],
        "M D": {"filters": filters_m_d_brother, "skip_rows": 1},
        "M SURESH": {"filters": filters_m_suresh, "skip_rows": 1},
        "MB": {"filters": filters_mb, "skip_rows": 1},
        "NAROLA": {"filters": filters_narola, "skip_rows": 0},
        "PANSURYA": {"filters": filters_pansurya_impex, "skip_rows": 0},
        "PRAGATI": {"filters": filters_pragati, "skip_rows": 0},
        "PRAMUKH": {"filters": filters_pramukh, "skip_rows": 0},
        "R N EXPORT": {"filters": filters_rn_export, "skip_rows": 0},
        "R.RAJESh": {"filters": filters_r_rajesh, "skip_rows": 5},
        "RAYS": [{"filters": filters_rays, "skip_rows": 0},
                 {"filters": filters_rays, "skip_rows": 1} ],
        'RED EXIM': {"filters": filters_red_exim, "skip_rows": 0},
        'SAGRA LIST':{"filters": filters_sagar_list, "skip_rows": 0},
        'SHAIRU': {"filters": filters_shairu, "skip_rows": 5},
        'SHEETAL': {"filters": filters_sheetal, "skip_rows": 8},
        'SHYAM': {"filters": filters_shyam, "skip_rows": 0},
        'SKE': {"filters": filters_ske, "skip_rows": 0},
        'SNJ LIST': {"filters": filters_snj_list, "skip_rows": 5},
        'SRD': {"filters": filters_srd, "skip_rows": 4},
        'SUNRAYS': {"filters": filters_sunrays, "skip_rows": 0},    
        "VD  GLOBAl": {"filters": filters_vd_global, "skip_rows": 4},
        'VEER GEMS': {"filters": filters_veer_gems, "skip_rows": 3},
        'UNQUE': {"filters": filters_unique, "skip_rows": 0},
       
    }
    for file_path in excel_files:
        file_name = os.path.basename(file_path).upper()
        for prefix, configurations in prefix_info.items():
            if prefix in file_name:  # Check if the file name contains the prefix
                # Check if configurations is a list or a dictionary
                if isinstance(configurations, list):  # For 'D NAVIN' which has multiple configurations
                    for config in configurations:
                        try:
                            process_data(
                                file_path=file_path,
                                filters=config["filters"],  # Access filter settings correctly
                                column_mapping=column_mapping,
                                output_dir=r"G:\Certificate\Python Output\DEAL Output",
                                prefix=prefix,
                                skip_rows=config["skip_rows"],  # Access skip_rows correctly
                            )
                        except Exception as e:
                            print(f"Error processing file {file_name} with prefix '{prefix}': {e}")
                            print("Skipping this file and moving to the next.\n")
                elif isinstance(configurations, dict):  # For 'HK' and 'UNQUE' where we have single config
                    
                    try:
                        process_data(
                            file_path=file_path,
                            filters=configurations["filters"],  # Access filter settings directly
                            column_mapping=column_mapping,
                            output_dir=r"G:\Certificate\Python Output\DEAL Output",
                            prefix=prefix,
                            skip_rows=configurations["skip_rows"],  # Access skip_rows directly
                        )
                    except Exception as e:
                        print(f"Error processing file {file_name} with prefix '{prefix}': {e}")
                        print("Skipping this file and moving to the next.\n")
                break
        else:
            print(f"No matching prefix found in file: {file_name}. Skipping.")

if __name__ == "__main__":
    main()
