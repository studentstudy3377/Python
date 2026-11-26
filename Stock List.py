import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime
import os
import platform
import subprocess
import numpy as np
import tkinter as tk
from tkinter import Tk, filedialog

def load_data_with_formulas(filepath):
    wb = load_workbook(filepath, data_only=False)
    ws = wb.active

    data_rows = []
    for row in ws.iter_rows(values_only=True):
        data_rows.append([cell for cell in row])

    headers = data_rows[0]
    data = pd.DataFrame(data_rows[1:], columns=headers)
    return data, wb, ws

def load_data_with_values(filepath):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    data_rows = []
    for row in ws.iter_rows(values_only=True):
        data_rows.append([cell for cell in row])

    headers = data_rows[0]
    data = pd.DataFrame(data_rows[1:], columns=headers)
    return data

def extract_formulas(filepath, column_letter):
    wb = load_workbook(filepath, data_only=False)
    ws = wb.active
    formulas = []
    for row in ws.iter_rows(min_row=2, max_col=ws.max_column):
        for cell in row:
            if cell.column_letter == column_letter:
                formulas.append(cell.value)
    return formulas

def reapply_formulas(ws, formulas, column_letter):
    for idx, formula in enumerate(formulas, start=2):
        ws[f"{column_letter}{idx}"].value = formula

def apply_openpyxl_styles(ws, row, col, fill_color=None, font_color=None, bold=False, font_name='Calibri', font_size=11):
    cell = ws.cell(row=row, column=col)
    cell.font = Font(name=font_name, size=font_size, color=font_color, bold=bold)
    if fill_color:
        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

def apply_styles_to_excel(ws, df):
    yellow_fill_pattern = PatternFill(start_color='FCE5CD', end_color='FCE5CD', fill_type='solid')  # Yellow

    for i, row in df.iterrows():
        excel_row = i + 2  # Adjust for header row
        for col_idx, col_name in enumerate(df.columns, start=1):
            val = row[col_name]
            cell = ws.cell(row=excel_row, column=col_idx)
            cell.value = val

            if col_name == 'STATUS' and val == 'New Arrival':
                for c in range(1, len(df.columns)+1):
                    current_fill = ws.cell(row=excel_row, column=c).fill
                    if current_fill.start_color.index != '00000000':  # Check if the cell is already highlighted
                        continue
                    ws.cell(row=excel_row, column=c).fill = yellow_fill_pattern

            # Additional conditions for styling
            clarity_val = row.get('CLARITY', None)
            depth_val = row.get('DEPTH', None)
            table_val = row.get('TABLE', None)
            rapsize_val = row.get('RAPSIZE', None)
            measurement_val = row.get('MEASUREMENT', None)

            if isinstance(measurement_val, str):
                measurement_val = pd.to_numeric(measurement_val, errors='coerce')
            
            if col_name == 'STATUS' and val == 'New Arrival':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
                if 'DISC' in df.columns:
                    disc_col_idx = df.columns.get_loc('DISC') + 1
                    ws.cell(row=excel_row, column=disc_col_idx, value=0)

            elif col_name == 'POL' and val == 'GD':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'CERT' and val in ['HRD', 'IGI']:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'SYM' and val == 'GD':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'CLARITY' and val == 'I1':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'CUT' and val == 'GD':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'TINCH' and val == 'B-1':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'FLUOR' and val != 'NONE':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'SHAPE' and val != 'ROUND':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'LUSTER' and val != 'EX':
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'BLACK' and isinstance(val, str) and 'H' in val:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'KEY TO SYMBOLS' and isinstance(val, str) and any(keyword in val for keyword in ['Feather- Chip- Surface Graining', 'Internal Graining- Feather- Cavity- Surface Graining']):
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'KEY TO SYMBOLS' and isinstance(val, str) and any(keyword in val for keyword in ['Cavity', 'Manufacturing Remnant', 'Graining', 'Internal Laser Drilling', 'Laser Drill-hole', 'Knot']):
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'KEY TO SYMBOLS' and isinstance(val, str) and 'Internal Graining' in val and clarity_val in ['IF', 'FL', 'VVS1', 'VVS2']:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'KEY TO SYMBOLS' and isinstance(val, str) and 'Surface Graining' in val and clarity_val in ['IF', 'FL', 'VVS1', 'VVS2']:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'DEPTH' and depth_val is not None and table_val is not None and depth_val < table_val:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'TABLE' and table_val is not None and depth_val is not None and depth_val < table_val:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FFCCCC', font_color='FF5050', bold=True)
            elif col_name == 'MEASUREMENT' and rapsize_val == '30UP' and measurement_val > 4.50:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FF9966')
            elif col_name == 'MEASUREMENT' and rapsize_val == '40UP' and measurement_val > 5.00:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FF9966')
            elif col_name == 'MEASUREMENT' and rapsize_val == '50UP' and measurement_val > 5.25:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FF9966')
            elif col_name == 'MEASUREMENT' and rapsize_val == '60UP' and measurement_val > 5.50:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FF9966')
            elif col_name == 'MEASUREMENT' and rapsize_val == '70UP' and measurement_val > 5.85:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FF9966')
            elif col_name == 'MEASUREMENT' and rapsize_val == '80UP' and measurement_val > 6.00:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FF9966')
            elif col_name == 'MEASUREMENT' and rapsize_val == '90UP' and measurement_val > 6.15:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FF9966')
            elif col_name == 'MEASUREMENT' and rapsize_val == '1.00UP' and measurement_val > 6.30:
                apply_openpyxl_styles(ws, excel_row, col_idx, fill_color='FF9966')
                
def save_data_with_formulas(data, workbook, output_path):
    if 'Matching' in workbook.sheetnames:
        del workbook['Matching']

    ws = workbook.create_sheet('Matching')

    for col_idx, header in enumerate(data.columns, start=1):
        ws.cell(row=1, column=col_idx, value=header)

    apply_styles_to_excel(ws, data)
    workbook.save(output_path)

def add_empty_row(results):
    empty_row = pd.DataFrame([pd.NA] * len(results.columns), index=results.columns).transpose()
    return pd.concat([results, empty_row], ignore_index=True)

def open_file(file_path):
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":
            subprocess.call(["open", file_path])
        else:
            subprocess.call(["xdg-open", file_path])
    except Exception as e:
        print(f"Error opening file: {e}")

def find_carat_range(carat):
    carat_ranges = [
        (0.30, 0.36), (0.37, 0.39), (0.40, 0.45), (0.46, 0.49), (0.50, 0.59),
        (0.60, 0.69), (0.70, 0.79), (0.80, 0.89), (0.90, 0.95), (0.96, 0.99),
        (1.00, 1.19), (1.20, 1.29), (1.30, 1.39), (1.40, 1.49), (1.50, 1.59), (1.60, 1.69), (1.70, 1.79),
        (1.80, 1.89), (1.90, 1.99), (2.00, 2.19), (2.20, 2.49), (2.50, 2.99),
        (3.00, 3.49), (3.50, 3.99), (4.00, 4.49), (4.50, 4.99), (5.00, 5.49),
        (5.50, 5.99)
    ]
    for start, end in carat_ranges:
        if start <= carat <= end:
            return (start, end)
    return None

def get_allowed_cut(cut):
    cut_hierarchy = {
        '3EX': ['3EX'],
        'EX': ['3EX', 'EX'],
        'VG': ['3EX', 'EX', 'VG'],
        'GD': ['3EX', 'EX', 'VG', 'GD'],
        'FR': ['3EX', 'EX', 'VG', 'GD', 'FR']
    }

    if pd.isna(cut):
        return [pd.NA]
    else:
        return cut_hierarchy.get(cut, ['3EX', 'EX', 'VG', 'GD', 'FR'])

def get_allowed_fluor(fluor):
    fluor_hierarchy = {
        'NONE': ['NONE'],
        'FNT': ['NONE', 'FNT', 'VSL'],
        'MED': ['NONE', 'FNT', 'VSL', 'MED', 'SL'],
        'STG': ['NONE', 'FNT', 'VSL', 'MED', 'SL', 'STG'],
        'VST': ['NONE', 'FNT', 'VSL', 'MED', 'SL', 'STG', 'VST']
    }
    return fluor_hierarchy.get(fluor, ['NONE', 'FNT', 'VSL', 'MED', 'SL', 'STG', 'VST'])

def replace_low(df):
    if 'Luster' in df.columns:
        df.loc[df['Luster'] == 'LOW', 'LUSTER'] = 'LOW'
    return df

def match_stones():
    input_file_path = filedialog.askopenfilename(title="Select  Stock File", filetypes=[("Excel files", "*.xlsx")],initialdir=r'G:')
    revision_file = r"G:\Certificate\PREMAL\REVISE DAYS TRACKING.xlsx"
    days_file = filedialog.askopenfilename(title="Select  Days File", filetypes=[("Excel files", "*.xlsx")], initialdir=r"G:")

    # Load the data from the input file
    try:
        data, wb, ws = load_data_with_formulas(input_file_path)  # Load data with formulas
        data_values = load_data_with_values(input_file_path)  # Load data with values
    except Exception as e:
        print(f"Error loading data from {input_file_path}: {e}")
        return  # Exit or proceed to next process if loading fails

    if data.empty:
        print("No data found in the selected file.")
        return  # Exit or continue with the next process

    # Clean the string data in the 'data' DataFrame
    try:
        for col in data.columns:
            data[col] = data[col].map(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)
    except Exception as e:
        print(f"Error processing the data: {e}")
        return  # Exit or continue with the next process

    # Proceed with the rest of the operations
    data = data[~data['CERT'].str.contains('IIDGR', case=False, na=False)]
    data['STOCKID_VALUES'] = data_values['STOCKID']
    data['STOCKID_VALUES'] = data['STOCKID_VALUES'].astype(str).str.strip()
    revision = pd.read_excel(revision_file)
    days = pd.read_excel(days_file)
    days.rename(columns={'Luster': 'Luster'}, inplace=True)
    
    # Assuming you have loaded the data using the correct function
    data, wb, ws = load_data_with_formulas(input_file_path)  # Load data with formulas
    data_values = load_data_with_values(input_file_path)  # Load data with values

    # Now you can access the columns of the 'data' DataFrame (not file path)
    for col in data.columns:
        data[col] = data[col].map(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)

    data = data[~data['CERT'].str.contains('IIDGR', case=False, na=False)]

    data['STOCKID_VALUES'] = data_values['STOCKID']
    data['STOCKID_VALUES'] = data['STOCKID_VALUES'].astype(str).str.strip()
    revision['Sr. No'] = revision['Sr. No'].astype(str).str.strip()
    days['Sr. No'] = days['Sr. No'].astype(str).str.strip()

    stockid_formulas = extract_formulas(input_file_path, 'A')

    merged_df2 = pd.merge(data, revision[['Sr. No', 'DAYS']], left_on='STOCKID_VALUES', right_on='Sr. No', how='left')
    merged_df2.rename(columns={'DAYS': 'REV DAYS'}, inplace=True)
    merged_df1 = pd.merge(merged_df2, days[['Sr. No', 'Stock Days']], left_on='STOCKID_VALUES', right_on='Sr. No', how='left')
    merged_df1.rename(columns={'Stock Days': 'OLD Days'}, inplace=True)
    merged_df = pd.merge(merged_df1, days[['Sr. No', 'Luster']], left_on='STOCKID_VALUES', right_on='Sr. No', how='left')

    merged_df = replace_low(merged_df)

    columns_to_drop = ['Sr. No_x', 'Sr. No_y', 'Sr. No', 'Luster', 'STOCKID_VALUES']
    merged_df.drop(columns=[col for col in columns_to_drop if col in merged_df.columns], inplace=True)

    headers = [cell.value for cell in ws[1]]
    report_comm_col_idx = headers.index('REPORT COMM') + 1

    if 'REV DAYS' not in headers:
        ws.insert_cols(report_comm_col_idx + 1)
        ws.cell(row=1, column=report_comm_col_idx + 1, value='REV DAYS')
        headers.insert(report_comm_col_idx, 'REV DAYS')

    if 'OLD Days' not in headers:
        ws.insert_cols(report_comm_col_idx + 2)
        ws.cell(row=1, column=report_comm_col_idx + 2, value='OLD Days')
        headers.insert(report_comm_col_idx + 1, 'OLD Days')

    for idx, row in merged_df.iterrows():
        excel_row = idx + 2
        ws.cell(row=excel_row, column=report_comm_col_idx + 1, value=row['REV DAYS'])
        ws.cell(row=excel_row, column=report_comm_col_idx + 2, value=row['OLD Days'])
        if 'LUSTER' in headers:
            luster_col_idx = headers.index('LUSTER') + 1
            ws.cell(row=excel_row, column=luster_col_idx, value=row['LUSTER'])

    results = pd.DataFrame(columns=merged_df.columns)
    matched_new_arrivals = set()

    for _, new_stone in merged_df[merged_df['STATUS'].isin(['New Arrival', 'New arrival'])].iterrows():
        stock_id = new_stone['STOCKID']
        if stock_id in matched_new_arrivals:
            continue

        matched_new_arrivals.add(stock_id)
        carat_range = find_carat_range(new_stone['CARAT'])
        if not carat_range:
            continue

        is_round = new_stone['SHAPE'] == "ROUND"

        query = [
            (merged_df['SHAPE'] == new_stone['SHAPE']),
            (merged_df['CARAT'].between(*carat_range)),
            (merged_df['COLOR'] == new_stone['COLOR']),
            (merged_df['CLARITY'] == new_stone['CLARITY']),
            (merged_df['FLUOR'].isin(get_allowed_fluor(new_stone['FLUOR']))),
            (~merged_df['STOCKID'].isin(matched_new_arrivals))
        ]
        if is_round:
            query.append(merged_df['CUT'].isin(get_allowed_cut(new_stone['CUT'])))
            query.append(merged_df['POL'].isin(get_allowed_cut(new_stone['POL'])))
            query.append(merged_df['SYM'].isin(get_allowed_cut(new_stone['SYM'])))

        matches = merged_df[np.all(query, axis=0)]
        matched_new_arrivals.update(matches['STOCKID'])
        combined_results = pd.concat([new_stone.to_frame().T, matches], ignore_index=True)
        results = pd.concat([results, combined_results], ignore_index=True)
        results = add_empty_row(results)

    current_time = datetime.now().strftime("%d-%m-%Y_%H-%M")
    output_file_path = os.path.join(os.path.dirname(input_file_path), f"Stock {current_time}.xlsx")
    save_data_with_formulas(results, wb, output_file_path)

    matching_ws = wb['Matching']
    reapply_formulas(matching_ws, stockid_formulas, 'A')

    open_file(output_file_path)
    print(f"The new file with matched stones is saved at: {output_file_path}")

if __name__ == "__main__":
    match_stones()
