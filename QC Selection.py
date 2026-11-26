import datetime as dt
import os
import platform
import re
import tkinter as tk
from tkinter import filedialog, Tk
import numpy as np
import pandas as pd

def select_and_load_excel_files():
    """Opens a file dialog to select multiple Excel or CSV files. Returns {filename: DataFrame} with custom skiprows logic."""
    root = Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        title="Select Excel/CSV Files to Merge", filetypes=[("All Files", "*.*")]
    )

    skiprows_map = { }
    # skiprows_map = {'ASIAN STAR': 7, "RATNAKALA":4,"HARDIKMEHTA":6,
    #                 'GLOW STAR': 4}
    dfs = {}

    for path in file_paths:
        filename = os.path.basename(path)
        ext = os.path.splitext(filename)[1].lower()
        skiprows = next((v for k, v in skiprows_map.items() if k in filename.upper()), 0)

        try:
            if ext in ['.xlsx', '.xls']:
                engine = 'openpyxl' if ext == '.xlsx' else 'xlrd'
                dfs[filename] = pd.read_excel(path, engine=engine, skiprows=skiprows)
            elif ext == '.csv':
                dfs[filename] = pd.read_csv(path, skiprows=skiprows)
        except Exception:
            print(f"❌ Failed to read file: {filename}")

    return dfs


def open_file(file_path):
    if platform.system() == "Windows":
        os.startfile(file_path)


def concatenate_dataframes(dfs, column_names):
    concatenated = []
    for df in dfs.values():
        for column_name in column_names:
            if column_name in df.columns:
                concatenated.append(df[column_name])
                break
        else:
            concatenated.append(pd.Series([''] * len(df)))
    return pd.concat(concatenated, ignore_index=True) if concatenated else pd.Series(dtype='str')

def cost(v): return ',' if str(v).strip() == '' else v

def main():
    lab_columns = ['CERTI','Lab', 'LABNAME', 'Lab', 'Cert', 'LAB', 'Cert Lab', 'Crt', 'Cert.']
    measurement_columns = ['Diam.', 'Dia', 'MM', 'L*W/Diam', 'Measurements(L*W*D)', 'MEASUREMENTS', 'Measurements', 'Measurement', 'L/W Or Dia', 'Meas MM', 'Meas', 'Measurement']
    report_no_columns = ['Stock No','Stone ID','Diamond ID']
    dfs = select_and_load_excel_files()
    if not dfs:
        raise ValueError("No files selected.")
    data = pd.DataFrame()
    # data['Files Name'] = pd.concat(
    #     [pd.Series([name.split('.')[0]] * len(df)) for name, df in dfs.items()],
    #     ignore_index=True)

    columns_mapping = {
        'Diamond ID': report_no_columns,
        'Shape': ['SHP','shape', 'ShapeName', 'Shp.', 'SHAPE', 'Shape', 'Shp', 'Shape', 'shp', 'Shape and Cutting Style'],
        'Size': ['Carat','Carat','Size'],      
        'Color': ['COL','Colour Grade', 'color', 'Color Grade', 'ColorName', 'col', 'Colour', 'COLOR', 'Color', 'DispColor', 'Color', 'Col', 'Color', 'COL'],
        'Clarity': ['CLA','Clrt','Quality', 'purity', 'CTY', 'CLAR', 'Clarity Grade', 'ClarityName', 'Clr', 'Clarity', 'CLARITY', 'Purity', 'CLR', 'DispClarity', 'Clr', 'Cl.', 'Clarity', 'Cla', 'Clar'],
        'Cut': ['C','Sh.Ap. / Cut','Proportions', 'Cut Grade', 'Cut', 'Cut Grade', 'Final Cut', 'CUT', ],
        'Polish': ['P','POL','polish', 'Polish', 'POLISH', 'PolishName', 'Pol', 'POL.', 'POL', 'Pl', 'Po', 'Pol.'],
        'Symmetry': ['S','SYM','Symm.','sym', 'Symmetry', 'SymmetryName', 'Sym', 'SYM.', 'SYM', 'Symm', 'Sym.', 'Sy'],
        'Fluorescence Intensity': ['Fluorescence','Flo','Flour'],
        'Lab': lab_columns,
        'Depth': ['Total Depth','DEPTH','Depth','TD','T.Dpt','Total Depth Per.','Depth(%)','Dep%','Depth%','Tot.Depth %','Tot Depth', 'T.Dep.', 'Dep', 'Total depth%', 'TD%', 'TD', 'Depth %', 'T.D.%', 'Depth%'],
        'Table': ['TABLE','TA','Tbl','Table Size Per.','Table width %', 'TAB%', 'TA%', 'Tableprct', 'TBL', 'Tab', 'Table', 'Tab%', 'Tab%', 'Table Depth', 'Tbl', 'Table%', 'TA', 'Table(%)', 'Table %'],
        'Measurement':measurement_columns,
        'Key To Symbols':['Key To Symbol','Key To Symbol','Key To Symbols'],
        'Vendor Stock number':['Vendor Stock number'],
        '$/ct':['Pr/Ct','Rap $/CT','$/ct'],
        '%Rap':['Rap%','Base Off %','%Rap'],
        'Total':['Amount','Total'],
        'Seller':['Seller']
    }

    for column, keys in columns_mapping.items():
        data[column] = concatenate_dataframes(dfs, keys)
        # if column in ['Search']:
        #     data[column] = data.apply(lambda row: f"{row['Diamond ID']},{row[column]}", axis=1)
        if column in ['Fluorescence Intensity'] :
            data['Fluorescence Intensity'] = data['Fluorescence Intensity'].fillna('None')

    output_dir = r"G:\Certificate\Python Output"
    os.makedirs(output_dir, exist_ok=True)
    filename = f"QC Selection {dt.datetime.now().strftime('%d-%m-%Y')}.csv"
    output_path = os.path.join(output_dir, filename)
    data.to_csv(output_path, index=False)
    print(f"✅ File saved: {output_path}")
    open_file(output_path)

if __name__ == "__main__":
    main()
