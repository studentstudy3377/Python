import os
import pandas as pd
import tkinter as tk
import warnings
from tkinter import filedialog
from openpyxl import load_workbook

def merge_sheets(file_path):
    """
    Merge all sheets in an Excel file into a single DataFrame.
    """
    warnings.simplefilter("ignore", UserWarning)
    workbook = load_workbook(file_path)
    sheet_names = workbook.sheetnames

    merged_data = pd.DataFrame()

    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        data = sheet.values
        columns = next(data)[0:]
        df = pd.DataFrame(data, columns=columns)

        if not df.empty and df.dropna(how="all").shape[0] > 0:
            df = df.dropna(axis=1, how="all")
            merged_data = pd.concat([merged_data, df], ignore_index=True)

    return merged_data

def append_total_row(df):
    total = df['VALUE'].sum()

    # Create a dictionary with column names
    row_data = {}
    for col in df.columns:
        if col == 'VALUE':
            row_data[col] = total

    # Create DataFrame for the total row
    total_df = pd.DataFrame([row_data])

    # Concatenate without warning
    return pd.concat([df, total_df], ignore_index=True)

def filter_and_save_data(df, output_file_path):
    """
    Filter and process the DataFrame, then save it to an Excel file with multiple sheets.
    """
    required_columns = ['CERT', 'STATUS', 'RAPSIZE', 'LUSTER', 'SHAPE', 'COLOR', 'CLARITY', 'VALUE']
    for column in required_columns:
        if column not in df.columns:
            raise ValueError(f"Missing required column: {column}")

    df = df[df['STATUS'] != 'New Arrival']

    df = df[~df["STOCKID"].astype(str).str.contains("FSA", na=False , regex=True)]

    IIDGR = df[df['CERT'].str.contains('https://.*,"IIDGR"', na=False, regex=True) |
                df['CERT'].str.contains('IIDGR', na=False)]

    HRD_IGI = df[df['CERT'].str.contains('https://.*(?:HRD|IGI)', na=False, regex=True) |
                df['CERT'].str.contains('HRD|IGI', na=False)]

    GIR_RD_30_69 = df[
        (df['CERT'].str.contains('https://.*,"GIA"', na=False, regex=True) |
        df['CERT'].str.contains('GIA', na=False)) &
        (df['RAPSIZE'].isin(['30UP', '40UP', '50UP', '60UP'])) &
        (~df['LUSTER'].isin(['M-1', 'M-2', 'M-3', 'M-4'])) &
        (df['SHAPE'] == 'ROUND') &
        (df['DISC'] <= -49.5)
    ]

    GIR_RD_30_69.loc[
        GIR_RD_30_69['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'
    ] = GIR_RD_30_69.loc[
        GIR_RD_30_69['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'
    ].apply(lambda x: 'Blank' if x in ['IF', 'VVS1', 'VVS2'] else x)

    GIR_RD_30_69 = GIR_RD_30_69[GIR_RD_30_69['CLARITY'] != 'Blank']

    # Add total row here
    GIR_RD_30_69 = append_total_row(GIR_RD_30_69)

    GIR_RD_70_Above = df[
        (df['CERT'].str.contains('https://.*,"GIA"', na=False, regex=True) |
        df['CERT'].str.contains('GIA', na=False)) &
        (~df['RAPSIZE'].isin(['30UP', '40UP', '50UP', '60UP'])) &
        (~df['LUSTER'].isin(['M-1', 'M-2', 'M-3', 'M-4'])) &
        (df['SHAPE'] == 'ROUND') &
        (df['DISC'] <= -45) &
        (df['VALUE'] <= 2499)
    ]

    GIR_RD_70_Above.loc[
        GIR_RD_70_Above['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'
    ] = GIR_RD_70_Above.loc[
        GIR_RD_70_Above['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'
    ].apply(lambda x: 'Blank' if x in ['IF', 'VVS1', 'VVS2'] else x)

    GIR_RD_70_Above = GIR_RD_70_Above[GIR_RD_70_Above['CLARITY'] != 'Blank']

    GIR_RD_70_Above = append_total_row(GIR_RD_70_Above)

    GIR_Fancy = df[
        (df['CERT'].str.contains('https://.*,"GIA"', na=False, regex=True) |
        df['CERT'].str.contains('GIA', na=False)) &
        (~df['LUSTER'].isin(['M-1', 'M-2', 'M-3', 'M-4'])) &
        (df['SHAPE'] != 'ROUND') &
        (df['DISC'] <= -40) &
        (df['VALUE'] <= 2499)
    ]

    GIR_Fancy.loc[
        GIR_Fancy['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'
    ] = GIR_Fancy.loc[
        GIR_Fancy['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'
    ].apply(lambda x: 'Blank' if x in ['IF', 'VVS1', 'VVS2'] else x)

    GIR_Fancy = GIR_Fancy[GIR_Fancy['CLARITY'] != 'Blank']

    GIR_Fancy = append_total_row(GIR_Fancy)

    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        # Filter rows for the 'list' sheet
        filtered_df = df[~df['CERT'].str.contains('IIDGR|HRD|IGI', na=False, regex=True)]
        filtered_df.to_excel(writer, sheet_name='list', index=False)

        # Write filtered DataFrames to respective sheets
        IIDGR.to_excel(writer, sheet_name='IIDGR', index=False)
        HRD_IGI.to_excel(writer, sheet_name='HRD & IGI', index=False)
        GIR_RD_30_69.to_excel(writer, sheet_name='GIA RD 0.30 - 0.69 @-', index=False)
        GIR_RD_70_Above.to_excel(writer, sheet_name='GIA RD 0.70 & ABOVE @-', index=False)
        GIR_Fancy.to_excel(writer, sheet_name='GIA FANCY SHAPES @-', index=False)

        # Set tab colors
        writer.sheets['list'].set_tab_color('#ff9999')
        writer.sheets['IIDGR'].set_tab_color('#ffcc99')
        writer.sheets['HRD & IGI'].set_tab_color('#ffcccc')
        writer.sheets['GIA RD 0.30 - 0.69 @-'].set_tab_color('#66ffff')
        writer.sheets['GIA RD 0.70 & ABOVE @-'].set_tab_color('#99ff99')
        writer.sheets['GIA FANCY SHAPES @-'].set_tab_color('#ffccff')

def main():
    try:
        root = tk.Tk()
        root.withdraw()

        input_file_path = filedialog.askopenfilename(
            title="AHS File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not input_file_path:
            raise ValueError("No file selected.")

        output_folder = r"G:\Certificate\Python Output"
        output_file_path = f"{output_folder}/AHS DEAL.xlsx"

        merged_data = merge_sheets(input_file_path)
        filter_and_save_data(merged_data, output_file_path)

        os.startfile(output_file_path)
    except Exception as e:
        print("Error", str(e))

if __name__ == "__main__":
    main()