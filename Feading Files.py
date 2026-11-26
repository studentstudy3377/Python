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


def add_columns(df, lab_columns, report_no_columns):
    report_no_col = next((col for col in report_no_columns if col in df.columns), None)
    if report_no_col is None:
        raise KeyError(f"Missing Report No columns: {report_no_columns}")

    lab_col = next((col for col in lab_columns if col in df.columns), None)
    if lab_col is None:
        raise KeyError(f"Missing Lab columns: {lab_columns}")

    df['Report No'] = df[report_no_col].apply(
        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x)
    )
    df[lab_col] = df[lab_col].astype(str).str.strip()
    df['Inscription'] = df[lab_col] + ' ' + df['Report No']
    return df


def process_measurement(df, column_name):
    if column_name not in df.columns:
        return df

    def extract_values(measurement):
        numbers = re.findall(r'\d*\.\d+|\d+', str(measurement))
        return map(float, numbers) if len(numbers) == 3 else (None, None, None)

    df[['Length', 'Width', 'Hights Measurement']] = df.apply(
        lambda row: pd.Series(extract_values(row[column_name])), axis=1
    )
    return df


# ================= Normalizers ===================
def process_cut_polish_symmetry(v):
    v = str(v).strip().upper()
    return (
        'GD' if v in {'G', 'GOOD'} else
        'FAIR' if v == 'FR' else
        'EX' if v in {'EXCELLENT', '3EX', 'EXCL', 'EXC'} else
        'VG' if v in {'VERY GOOD', 'VGOOD', 'VERY'} else v
    )


def replace_comma_with_hyphen(v):
    return re.sub(r'\s*,\s*', '-', re.sub(r'\(s\)', '', str(v))).strip() if pd.notna(v) else v


def remove_sign(v):
    return re.sub(r'\s*[+-]\s*', '', str(v)) if pd.notna(v) else v


def fluorescence_intensity(v):
    lookup = {
        'NN': 'NONE', 'NON': 'NONE', 'N': 'NONE', 'NIL': 'NONE',
        'F': 'FAINT', 'FNT': 'FAINT', 'FA': 'FAINT',
        'MED': 'MED. BLUE', 'MEDIUM': 'MED. BLUE', 'MD': 'MED. BLUE','M':'MED',
        'SLIGHT': 'SLIGHT', 'SL': 'SLIGHT', 'STG': 'STRONG',
        'V.STRONG': 'V.STRONG', 'VERY SLIGHT': 'VERY SLIGHT'
    }
    return lookup.get(str(v).strip().upper(), str(v).strip().upper() if pd.notna(v) else 'NONE')


def culet(v):
    return {'N': 'NON', 'V': 'VSM', 'MIN': 'NON'}.get(str(v).strip().upper(), str(v).strip().upper() if pd.notna(v) else v)


def girdle(v):
    if pd.isna(v): return v
    value = str(v).strip().upper()
    mapping = {
        'MEDIUM TO SLIGHTLY THICK': 'MED TO STK',
        'MEDIUM TO SLIGHTLY THICK (FACETED)': 'MED TO STK',
        'STK - THK': 'MED TO STK',
        'M.S': 'MED TO STK',
        'THN - STK': 'THN TO STK',
        'SLIGHTLY THICK TO THICK': 'STK TO THK',
        'MEDIUM TO THICK': 'MED TO THK',
        'MEDIUM-SLIGHTLY THICK': 'MED TO STK',
        'MED - THK': 'MED TO THK',
        'MED-THK': 'MED TO THK',
        'MEDIUM TO VERY THICK': 'MED TO VTHK',
        'SLIGHTLY THICK': 'STK',
        'SLIGHT THICK': 'STK',
        'ST': 'STK',
        'MEDIUM': 'MED',
        'VERY THICK': 'VTHK',
        'VERY THIK': 'VTHK',
        'VERY THIN': 'VTN',
        'TK': 'THK',
        'EXTREMLY THICK': 'ETK',
        'EXTREMLY THIN': 'ETN'
    }
    return mapping.get(value, value)

def shape(v):
    if pd.isna(v): return v
    val = str(v).strip().upper()
    return (
        'ROUND' if val in {'BR', 'RBC', 'RD', 'ROUND BRILLIANT', 'BRILLIANT'} else
        'PEAR' if val in {'PR','PB'} else
        'OVAL' if val in {'OL', 'OV'} else
        'MARQUIES' if val == 'MQ' else
        'EMRALD' if val in {'EM', 'EMD'} else
        'CUSHION CUT' if val in {'CUS CUT', 'CUSH'} else val
        # ROUND EMRALD PEAR OVAL MARQUIES PRINCESS HEART RADIANT CUSHION CUT

    )

def top_open(v): return 'N/A' if str(v).strip() == '' else v

def eyeclean(v): return 'YES' if str(v).strip() == '' else v

def cost(v): return '1' if str(v).strip() == '' else v

def milky(v): return 'NORMAL' if str(v).strip() == '' else v

def girdel_condition(v): return 'Faceted' if str(v).strip() == '' else v

def Crn_Ag(v): return re.sub(r'[^0-9.]', '', str(v)) if pd.notna(v) else '0'

def reorder_length_width_if_round(row):
    if str(row.get('Shape', '')).upper() == 'ROUND':
        try:
            l, w = float(row['Length']), float(row['Width'])
            return min(l, w), max(l, w)
        except Exception:
            pass
    return row.get('Length'), row.get('Width')

def main():
    lab_columns = ['CERTI','Lab', 'LABNAME', 'Lab', 'Cert', 'LAB', 'Cert Lab', 'Crt', 'Cert.','Certi']
    measurement_columns = ['Diam.', 'Dia', 'MM', 'L*W/Diam', 'Measurements(L*W*D)', 'MEASUREMENTS', 'Measurements', 'Measurement', 'L/W Or Dia', 'Meas MM', 'Meas', 'Measurement','Measure.']
    report_no_columns = ['REPORT NO','Reportnumber','Cert.No.','Cert_No','Certificate No','CERT NO','Cert','CertiNo',' REF NO.', 'Cert No', 'Report.No', 'IGI Report Number', 'cert', 'Cert NO', 'Report', 'Report No', 'CERTNO', 'Report Number', ' REF NO.', 'Certificate Number', 'CERT NUMBER', 'Lab Report No', 'Report no ', 'REPORT NUM', 'Certi Number', 'GIA No', 'REPORT', 'REPORT #', 'Certificate ', 'CertNo', 'CertificateNo', 'Certificate', 'Certificate No', 'Id', 'Report #', 'ReportNo', 'Cert.No', 'REPORT NO.', 'Rep No', 'Cert No.', 'Certi No.', 'Cert.No.', 'Certificate #', 'Report No.']

    dfs = select_and_load_excel_files()
    if not dfs:
        raise ValueError("No files selected.")

    for file_name, df in dfs.items():
        try:
            dfs[file_name] = add_columns(df, lab_columns, report_no_columns)
        except KeyError as e:
            print(f"Skipping {file_name}: {e}")
            continue
        for col in measurement_columns:
            if col in df.columns:
                dfs[file_name] = process_measurement(df, col)
                break

        if 'GIRDLE THIN' in df.columns and 'GIRDLE THICK' in df.columns:
            df['Girdle_Concatenate'] = df['GIRDLE THIN'].fillna('') + ' - ' + df['GIRDLE THICK'].fillna('')

    data = pd.DataFrame()
    data['PARTY NAME'] = pd.concat(
        [pd.Series([name.split('.')[0]] * len(df)) for name, df in dfs.items()],
        ignore_index=True
    )

    columns_mapping = {
        'Job No': ['Job No'],
        'Control No': ['Control No'],
        'Diamond Dossier': ['Diamond Dossier'],
        'Report No': report_no_columns,
        'Report Dt': ['Valid From', 'Date', 'Report Dt', 'Certificate Date', 'Report Date'],
        'Client Ref': ['ID','Stock Id','StockID', 'SRD No', 'Client ref', 'Stk #', 'REFNO', 'Stock No.', 'StoneNo', 'RefNo', 'STOCK', 'Stock ID No', 'REF NO', 'STOCK NO ', 'StoneId', 'StockNo', 'Client Ref', 'Stone No', 'Stock #', 'Lot No.', 'Packet No', 'Stock ID', 'StockRef', 'Stock No', 'Refno', 'Stone No', 'STONE ID', 'Stone No', 'Stone ID', 'Stone Id', 'Stone ID', 'Stock Id', 'StoneID', 'Packet No', 'StockID'],
        'Memo No': ['Memo No'],
        'Shape': ['SHP','shape', 'ShapeName', 'Shp.', 'SHAPE', 'Shape', 'Shp', 'Shape', 'shp', 'Shape and Cutting Style'],
        'Length': ['Length','Len','Diam / LW Min', 'Max', 'Length', '', 'L', 'length', ''],      
        'Width': ['Width','Wid','Diam / LW Max', 'Min', 'Width', 'W', 'witdh'], 
        'Depth': ['Total Depth MM','Hei','Height', 'Height CAvg', 'Hgt','Depth mm', 'H', 'DEPTH', 'Hights Measurement', 'D'],      
        'Weight': ['CTS','WT','Ct.','Weight R', 'weight', 'SIZE', 'Carat Weight', 'CARAT', 'carat', 'WEIGHT', 'Cts.', 'Weight', 'cts', 'Cts', 'CRTWT', 'Size', 'CTS', 'Cts', 'Carats', 'Weight', 'Cts', 'Carat', 'Wgt'],      
        'Color': ['COL','Colour Grade', 'color', 'Color Grade', 'ColorName', 'col', 'Colour', 'COLOR', 'Color', 'DispColor', 'Color', 'Col', 'Color', 'COL'],
        'Color Descriptions': ['Color Descriptions'],
        'Clarity': ['CLA','Clrt','Quality', 'purity', 'CTY', 'CLAR', 'Clarity Grade', 'ClarityName', 'Clr', 'Clarity', 'CLARITY', 'Purity', 'CLR', 'DispClarity', 'Clr', 'Cl.', 'Clarity', 'Cla', 'Clar'],
        'Clarity Status': ['Clarity Status'],
        'Final Cut': ['C','Sh.Ap. / Cut','Proportions', 'Cut Grade', 'Cut', 'Cut Grade', 'Final Cut', 'CUT', ],
        'Polish': ['P','POL','polish', 'Polish', 'POLISH', 'PolishName', 'Pol', 'POL.', 'POL', 'Pl', 'Po', 'Pol.'],
        'Symmetry': ['S','SYM','Symm.','sym', 'Symmetry', 'SymmetryName', 'Sym', 'SYM.', 'SYM', 'Symm', 'Sym.', 'Sy'],
        'Fluorescence Intensity': ['FLUO','FLS','LW-fluo', 'fl', 'FLR', 'Fluorescence', 'Flu', 'FLIntName', 'Flor', 'FLO.', 'Fluorescence ', 'FlrIntens', 'Fluorescence Intensity', 'Fluor', 'Fluor.', 'Flur', 'Flour', 'Fluor', 'Flo', 'Fl', 'Flor', 'FLUOR', 'Flr', 'Fluo Int', 'Fluo', 'Fluorescent', 'Fluorescence', 'Fls', 'FL'],
        'Fluorescence Color': ['Fluorescence Color'],
        'Girdle': ['Girdle ID','Grdl','GirdleThick','Gridle','Girdle_Concatenate', 'Girdle Desc', 'Girdle description', 'Girdle Thick', 'GIRDLE', 'Girdle Name', 'Girdle', 'Girdle Thickness', 'Girdle Name', 'Girdle Desc', 'Gr', 'Girdle Thin', 'Gridle'],
        'Girdle Condition': ['Grdl.Cond.','Girdle nature','Girdle Condition', 'GIRDLE CONDITION', 'girdlecondition', 'GirdleCondition', 'Girdle Cond'],
        'Culet Size': ['Culet ID','culetsize','Culet nature','Culet', 'Culet Size'],
        'Depth %': ['DEPTH','Depth','TD','T.Dpt','Total Depth Per.','Depth(%)','Dep%','Depth%','Tot.Depth %','Tot Depth', 'T.Dep.', 'Dep', 'Total depth%', 'TD%', 'TD', 'Depth %', 'T.D.%', 'Depth%'],
        'Table %': ['TABLE','TA','Tbl','Table Size Per.','Table width %', 'TAB%', 'TA%', 'Tableprct', 'TBL', 'Tab', 'Table', 'Tab%', 'Tab%', 'Table Depth', 'Tbl', 'Table%', 'TA', 'Table(%)', 'Table %'],
        'Crn Ag': ['C.A','Crn Ang','Cr Ang','Cr.Ang','Crown Angel','Cr Ang', 'Cr Angle', 'CA', 'C Angle', 'CRAng', 'Crown angle°', 'Crn Ag', 'Crown Angle', 'Cr >', 'C.A.', 'C.Angle', 'C/A', 'Cr.Ang', 'Crown Angle', 'CA'],    
        'Crn Ht': ['C.H.','Crn Hgt','CrownHeight','Crown Height Per.','Cr Ht','Cr.Heig%','Cr Hgt', 'CH', 'CHeight', 'CRHgt', 'Cr. Height%', 'Crn Dp', 'Crn Ht', 'Crown Height', 'Cr Hgt%', 'Crown Height', 'C.H', 'C.Height', 'C/H', 'Cr.Heig %', 'Crown Height', 'CH'],
        'Pav Ag': ['P.A','Pavi. Angle','Pav Ang','Pav.Ang','Pav Angel','PA', 'P Angle', 'PVAng', 'Pavillion angle°', 'Pav Ag', 'Pav Angle', 'Pav >', ' Pavilion Angle ', 'P.A.', 'Pav. Angle', 'P/A', 'Pav.Ang', 'Pavilion Angle', 'PA'],
        'Pav Dp': ['P.H','Pav Dpt','PavillionHeight','Pavi. Depth Per.','Pav Ht','Pav.Dep','Pav Hgt', 'PH', 'PHeight', 'PVHgt', 'Pav. depth%', 'Pavilion Height', 'Pav Dp', 'Pav Hgt%', 'Pavilion Depth', 'P.D.', 'Pav Height', 'Pav. Height', 'P/H', 'Pav.Per', 'Pavilion Depth', 'PH'],
        'Str Ln': ['Star Length','Str Ln', 'StrLn'],
        'Lr Half': ['Lower Half','Lr Half', 'LrHalf'],
        'Painting': ['Painting'],
        'Proportion': ['Proportion'],
        'Paint Comm': ['Paint Comm'],
        'Key to Symbols': ['KET TO SYMBOLS','KeyToSym','Clarity Characteristics', 'Key to Symbols', 'KeyToSymbol', 'keytosymbols', 'KeyToSymbols', 'KeyToSymbole', 'Key to Sym', 'KTSVIEW', 'Key to symbols', 'Comment', 'Key to Symbols', 'Key To Symbol', 'KeytoSymbols', 'Key Symbol', 'KEY TO SYMBOL', 'K T S', 'Keytosymbols', 'Key To Symbols', 'Key To Sym', 'Key to symbols'],
        'Report Comments': ['Report Comment', 'Report Comments', 'Comments', 'Cert comment'],
        'Inscription': ['Inscription', 'Inscription(s)'], 
        'Synthetic Indicator': ['Synthetic Indicator'],
        'Girdle %': ['Grdl%','Gridle Thickness Per.','Girdle %','Girdle%','GIRDLE%', 'Girdle Per', 'Girdle Size', 'Girdle%', 'Girdle', 'Girdle%', 'Girdle size %', 'Girdle size %'],
        'Polish Features': ['Polish Features'],
        'Symmetry Features': ['Symmetry Features'],
        'Shape Description': [''],
        'Report Type': [''],
        'Sorting': [''],
        'Basket Status': [''],
        'natts': ['natts'],
        'milky': ['milky'],
        'top opens': [],
        'bot opens': [],
        'tinch': ['tinch'],
        'eyeclean': ['eyeclean'],
        'ha': ['ha'],
        'cost': ['cost'],
        'STOCK ID': ['STOCK ID'],
        'origin': ['origin'],
        'lab': lab_columns
    }

    for column, keys in columns_mapping.items():
        data[column] = concatenate_dataframes(dfs, keys)
        # Apply post-processing
        if column in ['Final Cut', 'Polish', 'Symmetry']:
            data[column] = data[column].apply(process_cut_polish_symmetry)
        elif column in ['Key to Symbols', 'Report Comments']:
            data[column] = data[column].apply(replace_comma_with_hyphen)
        elif column == 'Fluorescence Intensity':
            data[column] = data[column].apply(fluorescence_intensity)
        elif column == 'Girdle':
            data[column] = data[column].apply(girdle)
        elif column == 'Shape':
            data[column] = data[column].apply(shape)
        elif column == 'Color':
            data[column] = data[column].apply(lambda x: remove_sign(str(x)).upper() if pd.notna(x) else x)
        elif column == 'Clarity':
            data[column] = data[column].apply(lambda x: remove_sign(str(x)).upper() if pd.notna(x) else x)
        elif column in ['top opens','bot opens','ha','origin']:
            data[column] = data[column].apply(top_open)
        elif column == 'eyeclean':
            data[column] = data[column].apply(eyeclean)
        elif column == 'cost':
            data[column] = data[column].apply(cost)
        elif column == 'Girdle Condition':
            data[column] = data[column].apply(girdel_condition)
        elif column == 'milky':
            data[column] = data[column].apply(milky)
        elif column == 'Culet Size':
            data[column] = data[column].apply(culet)
        elif column in ['Girdle %','Crn Ag','Crn Ht','Pav Ag','Pav Dp','Str Ln','Lr Half']:
            data[column] = data[column].apply(Crn_Ag)
        elif column == 'Report Dt' and 'Report Dt' in data.columns:
            data['Report Dt'] = pd.to_datetime(data['Report Dt'], errors='coerce').dt.strftime('%b-%d-%y')

    if all(x in data.columns for x in ['Shape', 'Length', 'Width']):
        data[['Length', 'Width']] = data.apply(
            reorder_length_width_if_round, axis=1, result_type='expand')

    data = data[data['Report No'].astype(str).str.lower().isin(['', 'nan', 'none', 'null']) == False]

    output_dir = r"G:\Certificate\Python Output"
    # output_dir = r"C:\Users\cert07\Downloads"
    os.makedirs(output_dir, exist_ok=True)
    filename = f"Feading {dt.datetime.now().strftime('%d-%m-%Y')}.csv"
    output_path = os.path.join(output_dir, filename)
    data.to_csv(output_path, index=False)
    print(f"✅ File saved: {output_path}")
    open_file(output_path)

if __name__ == "__main__":
    main()
