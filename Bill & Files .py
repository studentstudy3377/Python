import os

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from datetime import datetime
import pandas as pd
from tkinter.filedialog import askopenfilename

def select_excel_file():
    """Prompt user to select an Excel file."""
    return askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])

def read_excel(file_path):
    """Read and return the active sheet from the selected Excel file."""
    wb = load_workbook(filename=file_path, data_only=True)
    return wb.active

def get_columns_to_keep(sheet):
    """Identify the necessary columns and their indices from the sheet."""
    columns_to_keep = [
        'QualityCode', 'FOR', 'Size', 'Pcs', 'Carats', 'Price', 'Cost',
        'Disc 1%', 'Disc 2%', 'Add %', 'Remark', 'stoneid', 'PARTY NAME'
    ]
    header_row = list(sheet.iter_rows(values_only=True))[0]
    return {col_name: index for index, col_name in enumerate(header_row) if col_name in columns_to_keep}

def extract_data(sheet, columns_indexes):
    """Extract relevant data from the sheet."""
    data = []
    for row in sheet.iter_rows(values_only=True, min_row=2):
        filtered_row = {col: row[columns_indexes[col]] if columns_indexes[col] < len(row) else None for col in columns_indexes}
        data.append(filtered_row)
    return data

def process_data(data):
    """Process data by modifying 'Price' and clearing unwanted columns."""
    columns_to_clear = ['Cost', 'Disc 1%', 'Disc 2%', 'Add %', 'FOR', 'Size', 'Pcs', 'Remark']
    for row in data:
        if row['Cost'] is not None:
            row['Price'] = row['Cost']
        for col in columns_to_clear:
            row[col] = None
    return data

def get_unique_party_names(data):
    """Extract unique party names from the data."""
    return {row['PARTY NAME'] for row in data if row['PARTY NAME']}

def create_workbook(data, columns_to_keep, output_folder):
    """Create and save separate CSV files for each unique party."""
    
    for party_name in get_unique_party_names(data):
        party_data = [row for row in data if row['PARTY NAME'] == party_name]
        if len(party_data) > 1:
            df = pd.DataFrame(party_data)

            # Remove the 'PARTY NAME' column as it is used for filenames
            df = df.drop(columns=['PARTY NAME'], errors='ignore')

            # Ensure the output folder exists
            os.makedirs(output_folder, exist_ok=True)

            # Sanitize the party name for a safe filename
            safe_party_name = "".join(c if c.isalnum() or c in " " else "_" for c in str(party_name))
            output_filename = os.path.join(output_folder, f"{safe_party_name}.csv")

            # Save as CSV
            df.to_csv(output_filename, index=False, encoding="utf-8-sig")

            # Open the file (optional)
            os.startfile(output_filename)

def main():
    file_path = select_excel_file()
    if not file_path:
        print("No file selected.")
        return None

    sheet = read_excel(file_path)
    columns_indexes = get_columns_to_keep(sheet)
    data = extract_data(sheet, columns_indexes)
    processed_data = process_data(data)
    
    output_folder = "G:\\Certificate\\SAURABH SHAH\\Gaurav\\MPB Bill"
    os.makedirs(output_folder, exist_ok=True)
    
    create_workbook(processed_data, list(columns_indexes.keys()), output_folder)
    
    return file_path  # Return file path for further use in create_excel()

def create_excel(file_path):
    data = pd.read_excel(file_path)
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    thin_border = Border(left=Side(style='thin'), 
                        right=Side(style='thin'),
                        top=Side(style='thin'), 
                        bottom=Side(style='thin'))
    
    alignment_center = Alignment(horizontal='center', vertical='center')
    alignment_left = Alignment(horizontal='left', vertical='center')

    today_date = datetime.today().strftime("%d/%m/%Y")

    output_dir = "G:\\Certificate\\SAURABH SHAH\\Gaurav\\MPB Bill"
    os.makedirs(output_dir, exist_ok=True)
    current_datetime = datetime.now().strftime("%d.%m.%Y")
    output_file = os.path.join(output_dir, f"MPB Bill {current_datetime}.xlsx")

    start_column = 1  
    start_row = 1
    invoices_per_column = 4 
    row_offset = 12  
    column_offset = 4

    for index, party in enumerate(data['PARTY NAME'].unique()):
        party_data = data[data['PARTY NAME'] == party]
        carats_value = party_data['Carats'].sum()
        carats = "{:.2f}".format(carats_value)
        bank_values = party_data['BANK$'].dropna().astype(str).unique()
        total_bank = bank_values[0] if len(bank_values) == 1 else ', '.join(bank_values)
        stone_id = party_data['stoneid'].dropna().astype(str).unique()
        total_stone_id = stone_id[0] if len(stone_id) == 1 else ','.join('✔')
        broker_name = party_data['BROKER NAME'].iloc[0] if not party_data['BROKER NAME'].isna().all() else ""

        bill_type = "RS BILL" if total_bank.replace('.', '', 1).isdigit() else "$ BILL"
        bill_type = "RS BILL" if total_bank.replace('.', '', 1).isdigit() else "$ BILL"
        
        
        if index % invoices_per_column == 0 and index > 0:
            start_column += column_offset
            start_row = 1

        data_template = [
            ["DATE :-", today_date, "", "", "", ""],
            ["PARTY :-", party, "", "", "", ""],
            ["THRU:-", broker_name, "", "", "", ""],
            ["QUALITY", "CARAT", bill_type, "", "", ""],
            ["", carats, "", "", "", ""],
            ["", "", "", "", "", ""],
            ["LESS (-)", "NET", "", "", "", ""],
            ["X", total_bank, "", "", "", ""],
            ["TERM", " Days", "", "", "", ""],
            ["REMARK", total_stone_id, "", "", "", ""],
            ["BUYING PERSON", "SAURABH SHAH", "", "", "", ""],
        ]

        # Merge specific cells before setting values
        merge_cells_list = [
            start_row, start_row + 1, start_row + 2, start_row + 6,
            start_row + 7, start_row + 8, start_row + 9, start_row + 10, start_row + 11
        ]
        for row in merge_cells_list:
            ws.merge_cells(start_row=row, start_column=2 + start_column, end_row=row, end_column=3 + start_column)

        # Insert data with styling (after merging)
        for row_idx, row in enumerate(data_template, start=start_row):
            for col_idx, value in enumerate(row[:3], start=1):  # Only first 3 columns
                cell = ws.cell(row=row_idx, column=col_idx + start_column)
                
                # Only assign value to the top-left cell in merged ranges
                if not isinstance(cell, MergedCell):
                    cell.value = value
                # Apply alignment for specific rows
                if row[0] in ["LESS (-)", "LESS (-)"]:  # Left align the rows with "LESS (-)" and "NET"
                    cell.alignment = alignment_left
                else:
                    cell.alignment = alignment_center

                cell.border = thin_border

        # Set column widths explicitly
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 17

        # Move row offset for next invoice in the same column
        start_row += row_offset

    # Save and open the workbook
    wb.save(output_file)
    os.startfile(output_file)


if __name__ == "__main__":
    selected_file = main()
    if selected_file:
        create_excel(selected_file)
