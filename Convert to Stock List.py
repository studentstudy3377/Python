import os
import platform
import pandas as pd
import tkinter as tk
from tkinter.filedialog import askopenfilenames
import datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill

def select_files():
    root = tk.Tk()
    root.withdraw()
    file_paths = askopenfilenames(
        title='Select Excel Files',
        filetypes=[('Excel Files', '*.xlsx *.xls')]
    )
    return file_paths

def open_file(file_path):
    if platform.system() == "Windows":
        os.startfile(file_path)

def process_cut_polish_symmetry(value):
    if pd.isna(value):
        return value
    value = str(value).strip().upper()
    mappings = {'G': 'GD', 'GOOD': 'GD', 'FR': 'FAIR',
                'EXCELLENT': 'EX', '3EX': 'EX', 'EXCL': 'EX', 'EXC': 'EX', 'V. GOOD':'VG',
                'VERY GOOD': 'VG', 'VGOOD': 'VG', 'VERY': 'VG'}
    return mappings.get(value, value)

def shape(value):
    if pd.isna(value):
        return value
    value = str(value).strip().upper()
    mappings = {'BR': 'ROUND', 'RBC': 'ROUND', 'RD': 'ROUND', 'ROUND BRILLIANT': 'ROUND',
                'BRILLIANT': 'ROUND', 'PR': 'PEAR', 'OL': 'OVAL', 'OV': 'OVAL','MARQUIES':'MARQUISE',
                'MQ': 'MARQUIES', 'EM': 'EMERALD', 'EMD': 'EMERALD',
                'CUS CUT': 'CUSHION CUT', 'CUSH': 'CUSHION CUT'}
    return mappings.get(value, value)

def load_excel_files(file_paths):
    return [pd.read_excel(file) for file in file_paths]

def concatenate_dataframes(dfs, keys):
    for key in keys:
        for df in dfs:
            if key in df.columns:
                return pd.concat([d[key] for d in dfs if key in d.columns], ignore_index=True)
    return pd.Series()

def process_tinch(value):
    if pd.isna(value):
        return 'NONE'
    value = str(value).strip().upper()
    mappings = {'LIGHT BROWN': 'B-1', 'BROWN': 'B-2'}
    return mappings.get(value, value)

def process_fluor(value):
    value = str(value).strip().upper()
    mappings = {'FAINT': 'FNT', 'MED. BLUE': 'MED','STRONG':'STG','V.STRONG':'VST'}
    return mappings.get(value, value)

def process_luster(value):
    value = str(value).strip().upper()
    mappings = {'NORMAL': 'EX', 'LIGHT MILKY': 'M-1','LOW':'LOW'}
    return mappings.get(value, 'M-2')

def format_measurement(length, width, depth):
    if pd.isna(length) or pd.isna(width) or pd.isna(depth):
        return ''
    return f"{length:.2f} - {width:.2f} X {depth:.2f}"

def categorize_rapsize(carat):
    if 0.70 <= carat < 0.80:
        return '70UP'
    elif 0.80 <= carat < 0.90:
        return '80UP'
    return None  # Will be ignored during update


def apply_style(ws, row, col, fill_color=None):
    cell = ws.cell(row=row, column=col)
    if fill_color:
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        cell.fill = fill

def apply_font_colors(data):
    wb = openpyxl.load_workbook(data)
    ws = wb.active

    font = Font(color="ff0000")
    columns_to_color = ['DISC', 'PRICE', 'VALUE','RAPPRICE','RATIO']
    for column_name in columns_to_color:
        column_index = None
        for col_num, cell in enumerate(ws[1], start=1):
            if cell.value == column_name:
                column_index = col_num
                break
        if column_index:
            for row in ws.iter_rows(min_col=column_index, max_col=column_index, min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.font = font

    # New conditional coloring for SHAPE
    header = [cell.value for cell in ws[1]]
    try:
        shape_idx = header.index("SHAPE") + 1
        tinch_idx = header.index("TINCH") + 1
        luster_idx = header.index("LUSTER") + 1
    except ValueError:
        wb.save(data)
        return

    for excel_row in range(2, ws.max_row + 1):
        shape_cell = ws.cell(row=excel_row, column=shape_idx)
        tinch_cell = ws.cell(row=excel_row, column=tinch_idx)
        luster_cell = ws.cell(row=excel_row, column=luster_idx)

        if (
            shape_cell.value is not None
            and str(tinch_cell.value).strip().upper() != "NONE"
            or str(luster_cell.value).strip().upper() != "EX"
        ):
            apply_style(ws, excel_row, shape_idx, fill_color="BDD7EE")

    wb.save(data)


def main():
    file_paths = select_files()
    if not file_paths:
        raise ValueError("No files selected.")

    dfs = load_excel_files(file_paths)

    data = pd.DataFrame()

    columns_mapping = {
        'STOCKID': ['Sr. No'],
        'SHAPE': ['shape', 'ShapeName', 'Shp.', 'SHAPE', 'Shape', 'Shp', 'Shape', 'shp', 'Shape and Cutting Style'],
        'RAPSIZE': ['Size'],
        'CARAT': ['Weight R', 'weight', 'SIZE', 'Carat Weight', 'CARAT', 'carat', 'WEIGHT', 'Cts.', 'Weight', 'cts', 'Cts', 'CRTWT', 'Size', 'CTS', 'Cts', 'Carats', 'Weight', 'Cts', 'Carat', 'Wgt'],
        'Color': ['Colour Grade', 'color', 'Color Grade', 'ColorName', 'col', 'Colour', 'COLOR', 'Color', 'DispColor', 'Color', 'Col', 'Color', 'COL'],
        'Clarity': ['Quality', 'purity', 'CTY', 'CLAR', 'Clarity Grade', 'ClarityName', 'Clr', 'Clarity', 'CLARITY', 'Purity', 'CLR', 'DispClarity', 'Clr', 'Cl.', 'Clarity', 'Cla', 'Clar'],
        'CERT': ['Cert'],
        'CUT': ['Proportions', 'Cut Grade', 'Cut', 'Cut Grade', 'Final Cut', 'CUT', 'Ct'],
        'POL': ['polish', 'Polish', 'POLISH', 'PolishName', 'Pol', 'POL.', 'POL', 'Pl', 'Po', 'Pol.'],
        'SYM': ['Symetry', 'Symmetry', 'SymmetryName', 'Sym', 'SYM.', 'SYM', 'Symm', 'Sym.', 'Sy'],
        'FLUOR': ['FLRC'],
        'MEASUREMENT': ['MEASUREMENT'],
        'RAPPRICE': ['RAPAPORT'],
        'RAP V': [''],
        'DISC': ['DISC'],
        'PRICE': ['PRICE'],
        'VALUE': [''],
        'CERTNO': ['Job No'],
        'DEPTH': ['Depth'],
        'TABLE': ['Table'],
        'RATIO': ['Ratio'],
        'BLACK': ['Natts No'],
        'TINCH': ['Tinch'],
        'LUSTER': ['Luster'],
        'CA': ['C Angle'],
        'CH': ['C Height'],
        'PA': ['P Angle'],
        'PH': ['P Height'],
        'GIRDLE': ['Girdle'],
        'GIRDLE%': ['Girdle %'],
        'CULET': ['Culet'],
        'KEY TO SYMBOLS': ['Key To Symbol'],
        'REPORT COMM': ['Report Comment'],
        'Length': ['Length', 'Length (mm)', 'Len'],
        'Width': ['Width', 'Width (mm)', 'Wid'],
        'Depth MM': ['Depth MM', 'Depth (mm)', 'Dep'],
        'Rec Date':['Rec Date'],
        'Lock':['Lock']
    }

    for column, keys in columns_mapping.items():
        data[column] = concatenate_dataframes(dfs, keys)

        if column in ['CUT', 'POL', 'SYM']:
            data[column] = data[column].apply(process_cut_polish_symmetry)
        elif column == 'TINCH':
            data[column] = data[column].apply(process_tinch)
        elif column == 'LUSTER':
            data[column] = data[column].apply(process_luster)
        elif column == 'FLUOR':
            data[column] = data[column].apply(process_fluor)
        elif column == 'SHAPE':
            data[column] = data[column].apply(shape)
        elif column == 'Rec Date':
            data[column] = pd.to_datetime(data[column], errors='coerce').dt.strftime('%d/%m/%Y')
    # Replace 'CUT' with '3EX' if 'CUT', 'POL', and 'SYM' are all 'EX'
    if all(col in data.columns for col in ['CUT', 'POL', 'SYM']):
        data.loc[(data['CUT'] == 'EX') & (data['POL'] == 'EX') & (data['SYM'] == 'EX'), 'CUT'] = '3EX'

    if all(col in data.columns for col in ['Length', 'Width', 'Depth MM']):
        data['MEASUREMENT'] = data.apply(
            lambda row: format_measurement(row['Length'], row['Width'], row['Depth MM']), axis=1
        )
    if all(col in data.columns for col in ['RAPPRICE', 'CARAT']):
            data['RAP V'] = data['RAPPRICE'] * data['CARAT']

    if 'DISC' in data.columns:
        data['DISC'] = data['DISC'] + 20
        data = data[data['DISC'] != 20]

    if all(col in data.columns for col in ['RAPPRICE', 'DISC']):
        data['PRICE'] = data['RAPPRICE'] + data['RAPPRICE'] * data['DISC'] / 100

    if all(col in data.columns for col in ['PRICE', 'CARAT']):
        data['VALUE'] = data['PRICE'] * data['CARAT']

     # Overwrite RAPSIZE based on CARAT range
    if 'CARAT' in data.columns and 'RAPSIZE' in data.columns:
        updated_sizes = data['CARAT'].apply(categorize_rapsize)
        data.loc[updated_sizes.notna(), 'RAPSIZE'] = updated_sizes[updated_sizes.notna()]
    
    # Replace 'MIX' with '23UP' in the RAPSIZE column
    if 'RAPSIZE' in data.columns:
        data['RAPSIZE'] = data['RAPSIZE'].replace('MIX', '23UP')

    # Drop Length, Width, and Depth MM before saving
    columns_to_drop = ['Length', 'Width', 'Depth MM']
    data = data.drop(columns=[col for col in columns_to_drop if col in data.columns])

    output_directory = r"G:\Certificate\Python Output"
    os.makedirs(output_directory, exist_ok=True)
    datetime_str = dt.datetime.now().strftime('%d-%m-%Y')
    output_path = os.path.join(output_directory, f'Stock File {datetime_str}.xlsx')

    data.to_excel(output_path, index=False)

    # Apply the font colors to the required columns
    apply_font_colors(output_path)

    open_file(output_path)

if __name__ == '__main__':
    main()
