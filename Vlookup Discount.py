import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def select_file(title):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])
    return file_path if file_path else None

def apply_remark_color(file_path):
    # Load the workbook and select the active worksheet
    wb = load_workbook(file_path)
    ws = wb.active

    # Define the fill color
    fill_color = PatternFill(start_color='00ccff', end_color='00ccff', fill_type='solid')

    # Identify the column with the header "Remark"
    remark_col = None
    for col in ws.iter_cols(min_row=1, max_row=1):
        if col[0].value == 'Remark':
            remark_col = col[0].column_letter
            break

    if not remark_col:
        return  # "Remark" column not found

    # Apply color to cells in the "Remark" column where the cell is not empty
    for row in ws.iter_rows(min_col=ws[remark_col + '1'].column, max_col=ws[remark_col + '1'].column):
        cell = row[0]
        if cell.row > 1 and cell.value:  # Skip header and blank cells
            cell.fill = fill_color

    # Save the modified workbook
    wb.save(file_path)

def main():
    try:
        # Prompt user to select files
        file_path = select_file("Select the Hemil Cert Merge.xlsx file")
        promo_path = select_file("Select the PROMO 27-08.xlsx file")
        premal_path = select_file("Select the Premal.xlsx file")
        sachin_path = select_file("Select the Sachin.xlsx file")
        sagar_path = select_file("Select the Sagar.xlsx file")
        date_path = select_file("Select the PREMAL 2024 APR.xlsx file")

        if not all([file_path, promo_path, date_path]):
            print("Warning", "Some required files are missing. Exiting.")
            return

        # Load data from Excel files
        file = pd.read_excel(file_path)
        promo = pd.read_excel(promo_path)
        date = pd.read_excel(date_path, sheet_name='REVISE DAYS')


        
        # Merge data
        merged_df = pd.merge(file, promo[['STOCKID', 'P DISC']], left_on='Sr. No', right_on='STOCKID', how='left')
        merged_df.rename(columns={'P DISC': 'Promo Disc'}, inplace=True)
        merged_df.drop(columns=['STOCKID'], inplace=True)

        # Conditionally handle the Premal file
        if premal_path:
            premal = pd.read_excel(premal_path, sheet_name='Matching')
            merged_df1 = pd.merge(file, premal[['STOCKID', 'Premal']], left_on='Sr. No', right_on='STOCKID', how='left')
            # merged_df1.rename(columns={'Disc': 'Premal'}, inplace=True)
            merged_df1.drop(columns=['STOCKID'], inplace=True)
        else:
            merged_df['Premal'] = pd.NA
        
        # Conditionally handle the Sachin file
        if sachin_path:
            sachin = pd.read_excel(sachin_path, sheet_name='for upload')
            merged_df2 = pd.merge(file, sachin[['STOCKID', 'Sachin']], left_on='Sr. No', right_on='STOCKID', how='left')
            # merged_df2.rename(columns={'Disc': 'Sachin'}, inplace=True)
            merged_df2.drop(columns=['STOCKID'], inplace=True)
        else:
            merged_df['Sachin'] = pd.NA
        
        # Conditionally handle the Sagar file
        if sagar_path:
            sagar = pd.read_excel(sagar_path, sheet_name='SAGAR')
            merged_df3 = pd.merge(file, sagar[['STOCKID', 'Sagar']], left_on='Sr. No', right_on='STOCKID', how='left')
            merged_df3.drop(columns=['STOCKID'], inplace=True)
            merged_df['Sagar'] = merged_df3['Sagar']
        else:
            merged_df['Sagar'] = pd.NA

        # Column Days and Rename as LST Rev 
        merged_df4 = pd.merge(file, date[['Sr. No', 'DAYS']], left_on='Sr. No', right_on='Sr. No', how='left')
        merged_df4.rename(columns={'DAYS': 'LST Rev'}, inplace=True)
        merged_df4.drop(columns=['Sr. No'], inplace=True)

        # Add columns to merged_df
        merged_df['Premal'] = merged_df1.get('Premal', pd.NA)
        merged_df['Sachin'] = merged_df2.get('Sachin', pd.NA)
        merged_df['Sagar'] = merged_df.get('Sagar', pd.NA)

        # Create New Column name as SMALL and Value is Greater Number of All Three Column 
        merged_df['SMALL'] = merged_df[['Premal', 'Sachin', 'Sagar']].min(axis=1)
        
        # Explicitly convert columns to numeric types before filling NaN values
        merged_df[['Premal', 'Sachin', 'Sagar', 'SMALL']] = merged_df[['Premal', 'Sachin', 'Sagar', 'SMALL']].apply(pd.to_numeric, errors='coerce').fillna(0)

        # Identify position of the "Price" column and insert the new columns before it
        if 'Price' in merged_df.columns:
            position = merged_df.columns.get_loc('Price')
            # Insert columns before 'Price'
            columns_to_insert = ['Promo Disc', 'Premal', 'Sachin', 'Sagar', 'SMALL']
            for col in reversed(columns_to_insert):
                merged_df.insert(position, col, merged_df.pop(col))
        
        # Identify position of the "Stock Days" column and insert "LST Rev"
        if 'Stock Days' in merged_df.columns:
            position = merged_df.columns.get_loc('Stock Days') + 1
            merged_df.insert(position, 'LST Rev', merged_df4['LST Rev'])
        else:
            merged_df['LST Rev'] = merged_df4['LST Rev']

        # Rename Name DISC column to List
        merged_df.rename(columns={'DISC': 'List'}, inplace=True)

        # Remove unnecessary columns
        merged_df.drop(columns=['Pur Date', 'E Date'], inplace=True)

        # Add condition to replace "DIGITAL" in "Remark" column with blank
        if 'Remark' in merged_df.columns:
            merged_df['Remark'] = merged_df['Remark'].replace('DIGITAL', '')


        today_date = datetime.today().strftime('%d-%m-%Y-%I-%M %p')

        # Save the DataFrame to an Excel file
        output_file = f'D:\\Premal Sir Vlookup\\Output\\ Hemil Cert Merge {today_date}.xlsx'
        merged_df.to_excel(output_file, index=False)

        # Apply color to the "Remark" column
        apply_remark_color(output_file)

        print("Success", f"File saved successfully as '{output_file}'")

    except FileNotFoundError as e:
        print("File Error", f"File not found: {e}")
    except pd.errors.EmptyDataError as e:
        print("Data Error", f"Empty data error: {e}")
    except pd.errors.MergeError as e:
        print("Merge Error", f"Merge error: {e}")
    except ValueError as e:
        print("Value Error", f"Value error: {e}")
    except Exception as e:
        print("Unexpected Error", f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
