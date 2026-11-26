import datetime as dt
import os
import platform
import tkinter as tk
from tkinter import filedialog, Tk
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

def select_and_load_excel_files():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Select Excel/CSV Files to Merge",
        filetypes=[("All Files", "*.*")]
    )
    root.destroy()
    return file_paths

def read_file(path, skiprows=0):
    filename = os.path.basename(path)
    ext = os.path.splitext(filename)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'
            df = pd.read_excel(path, engine=engine, skiprows=skiprows)
        elif ext == '.csv':
            df = pd.read_csv(path, skiprows=skiprows)
        else:
            return (filename, None)
        return (filename, df)
    except Exception:
        print(f"❌ Failed to read file: {filename}")
        return (filename, None)

def load_files_parallel(file_paths):
    dfs = {}
    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(executor.map(read_file, file_paths))
    for filename, df in results:
        if df is not None:
            dfs[filename] = df
    return dfs

# Use dict comprehensions for mapping
def map_column(df, old_names, new_name):
    for col in old_names:
        if col in df.columns:
            return df[col]
    return pd.Series([''] * len(df))

def process_column(df, column, func):
    df[column] = df[column].map(func)

def cost(v):
    return ',' if str(v).strip() == '' else v

def process_cut_polish_symmetry(v):
    v = str(v).strip().upper()
    return (
        'Good' if v in {'G', 'GOOD'} else
        'Fear' if v == 'FR' else
        '' if v == {'-','NAN'} else
        'EX' if v in {'EXCELLENT', '3EX', 'EXCL', 'EXC', 'EX','3EXCELLENT'} else
        'VG' if v in {'VERY GOOD', 'VGOOD', 'VERY'} else v
    )

def fluorescence(v):
    lookup = {
        'NN': 'NONE', 'NON': 'NONE', 'N': 'NONE', 'NIL': 'NONE','None':'NONE',
        'F': 'FAINT', 'FNT': 'FAINT', 'FA': 'FAINT',
        'MED. BLUE':'MED', 'MEDIUM': 'MED', 'MD': 'MED','M':'MED',
        'SLIGHT': 'SLIGHT', 'SL': 'SLIGHT', 'STG': 'STRONG',
        'V.STRONG': 'V.STRONG', 'VERY SLIGHT': 'VERY SLIGHT'
    }
    return lookup.get(str(v).strip().upper(), str(v).strip().upper() if pd.notna(v) else 'None')

def culet(v):
    lookup = {
        'NN': 'None', 'NON': 'None', 'N': 'None', 'NIL': 'None'
    }
    return lookup.get(str(v).strip().upper(), str(v).strip().upper() if pd.notna(v) else 'None')

def location(v):
    lookup = {
        'MUMBAI': 'Mumbai', 'NEWYORK': 'New York'
    }
    return lookup.get(str(v).strip().upper(), str(v).strip().upper() if pd.notna(v) else 'None')

def open_file(file_path):
    if platform.system() == "Windows":
        os.startfile(file_path)

def combine_cols(df, col1, col2):
    """
    Combine two columns with ' - ' separator, handling missing columns.
    Returns a Series with cleaned values.
    """
    if col1 in df.columns:
        val1 = df[col1].astype(str).str.strip()
    else:
        val1 = pd.Series([''] * len(df))

    if col2 in df.columns:
        val2 = df[col2].astype(str).str.strip()
    else:
        val2 = pd.Series([''] * len(df))

    combined = val1 + ' - ' + val2
    combined = combined.replace(' - ', '')  # Optional: clean placeholder-only
    return combined


def build_inclusion_series(dfs, col1a='T.B', col1b='S.B', col2a='BT', col2b='BC'):
    
    result_series = []

    for df in dfs.values():
        s1 = combine_cols(df, col1a, col1b)
        s2 = combine_cols(df, col2a, col2b)

        # Prefer s1 unless it’s empty
        combined = s1.combine(s2, lambda x, y: x if x else y)
        result_series.append(combined)

    return pd.concat(result_series, ignore_index=True)


def main():

    lab_columns = [
        'CERTI', 'Lab', 'LABNAME', 'Lab', 'Cert', 'LAB', 'Cert Lab',
        'Crt', 'Cert.'
    ]
    measurement_columns = [
        'Diam.', 'Dia', 'MM', 'L*W/Diam', 'Measurements(L*W*D)', 'MEASUREMENTS',
        'Measurements', 'Measurement', 'L/W Or Dia', 'Meas MM', 'Meas',
        'Measurement'
    ]
    report_no_columns = ['Stone ID','Stock No','Vendor Stock number']

    columns_mapping = {
        'Stone ID': report_no_columns,
        'Shape': ['SHP','shape', 'ShapeName', 'Shp.', 'SHAPE', 'Shape', 'Shp', 'Shape', 'shp', 'Shape and Cutting Style'],
        'Size': ['Carat','Size'],
        'Color': ['COL','Colour Grade', 'color', 'Color Grade', 'ColorName', 'col', 'Colour', 'COLOR', 'Color', 'DispColor', 'Color', 'Col', 'Color', 'COL'],
        'Clarity': ['CLA','Clrt','Quality', 'purity', 'CTY', 'CLAR', 'Clarity Grade', 'ClarityName', 'Clr', 'Clarity', 'CLARITY', 'Purity', 'CLR', 'DispClarity', 'Clr', 'Cl.', 'Clarity', 'Cla', 'Clar'],
        'Cut': ['C','Sh.Ap. / Cut','Proportions', 'Cut Grade', 'Cut', 'Cut Grade', 'Final Cut', 'CUT', ],
        'Polish': ['P','POL','polish', 'Polish', 'POLISH', 'PolishName', 'Pol', 'POL.', 'POL', 'Pl', 'Po', 'Pol.'],
        'Symmetry': ['S','SYM','Symm.','sym', 'Symmetry', 'SymmetryName', 'Sym', 'SYM.', 'SYM', 'Symm', 'Sym.', 'Sy'],
        'Fluorescence': ['FLUO','FLS','LW-fluo', 'fl', 'FLR', 'Fluorescence', 'Flu', 'FLIntName', 'Flor', 'FLO.', 'Fluorescence ', 'FlrIntens', 'Fluorescence Intensity', 'Fluor', 'Fluor.', 'Flur', 'Flour', 'Fluor', 'Flo', 'Fl', 'Flor', 'FLUOR', 'Flr', 'Fluo Int', 'Fluo', 'Fluorescent', 'Fluorescence', 'Fls', 'FL'],
        'Lab': lab_columns,
        'Ratio':['Ratio','D/R'],
        'Depth': ['Total Depth','DEPTH','Depth','TD','T.Dpt','Total Depth Per.','Depth(%)','Dep%','Depth%','Tot.Depth %','Tot Depth', 'T.Dep.', 'Dep', 'Total depth%', 'TD%', 'TD', 'Depth %', 'T.D.%', 'Depth%'],
        'Table': ['TABLE','TA','Tbl','Table Size Per.','Table width %', 'TAB%', 'TA%', 'Tableprct', 'TBL', 'Tab', 'Table', 'Tab%', 'Tab%', 'Table Depth', 'Tbl', 'Table%', 'TA', 'Table(%)', 'Table %'],
        'Measurement':measurement_columns,
        'Shade':['Shade','Shd.','Color Shade'],
        'Inclusion':['Inclusion'],
        'Key To Symbols':['Key To Symbol','Key To Symbol','Key To Symbols'],
        'Vendor Stock number':[],
        'Rap':['Rap ','Rap Rate','Rap $/CT'],
        '$/ct':['Pr/Ct','Base Rate','$/ct'],
        '%Rap':['Rap%','Base Off %','%Rap'],
        'Total':['Amount','Total'],
        'State':['State','Location'],
        'Milky':['Lust','MILKY'] 
    }
    file_paths = select_and_load_excel_files()
    if not file_paths:
        raise ValueError("No files selected.")
    
    dfs = load_files_parallel(file_paths)

    data = pd.DataFrame()
    # Party name generation vectorized
    party_names = []
    for name, df in dfs.items():
        party_names.extend([name.split('.')[0]] * len(df))
    data['Seller'] = pd.Series(party_names)
    
    bt_bc_series = []
    for df in dfs.values():
        if 'BT' in df.columns and 'BC' in df.columns:
            s = df['BT'].astype(str).str.strip() + ' - ' + df['BC'].astype(str).str.strip()
        elif 'BT' in df.columns:
            s = df['BT'].astype(str).str.strip() + ' - '
        elif 'BC' in df.columns:
            s = ' - ' + df['BC'].astype(str).str.strip()
        else:
            s = pd.Series([' - '] * len(df))
        bt_bc_series.append(s.reset_index(drop=True))
    
    tb_sb_series = []
    for df in dfs.values():
        if 'T.B' in df.columns and 'S.B' in df.columns:
            s = df['T.B'].astype(str).str.strip() + ' - ' + df['S.B'].astype(str).str.strip()
        elif 'T.B' in df.columns:
            s = df['T.B'].astype(str).str.strip() + ' - '
        elif 'S.B' in df.columns:
            s = ' - ' + df['S.B'].astype(str).str.strip()
        else:
            s = pd.Series([' - '] * len(df))
        tb_sb_series.append(s.reset_index(drop=True))

    for column, keys in columns_mapping.items():
        series_list = [map_column(df, keys, column) for df in dfs.values()]
        data[column] = pd.concat(series_list, ignore_index=True)
        # Vectorized mapping (replace .apply with .map where possible)
        if column in ['Cut', 'Polish', 'Symmetry']:
            data[column] = data[column].map(process_cut_polish_symmetry)
        elif column == 'Fluorescence':
            data[column] = data[column].map(fluorescence)
        elif column == 'Culet':
            data[column] = data[column].map(culet)
        elif column == 'Location':
            data[column] = data[column].map(location)
        # elif column == 'Inclusion':
        #     data[column] = pd.concat(bt_bc_series, ignore_index=True)
        elif column == 'Inclusion':
            data[column] = build_inclusion_series(dfs)
        
    output_dir = r"G:\Certificate\Python Output"
    os.makedirs(output_dir, exist_ok=True)
    filename = f"Rapnet Format {dt.datetime.now().strftime('%d-%m-%Y')}.xlsx"
    output_path = os.path.join(output_dir, filename)
    data.to_excel(output_path, index=False, engine='openpyxl')
    open_file(output_path)

if __name__ == "__main__":
    main()
    