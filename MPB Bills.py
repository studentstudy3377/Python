import os
from tkinter.filedialog import askopenfilename
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell  # Ensure we check for merged cells
from datetime import datetime
import pandas as pd

def select_excel_file():
    """Prompt user to select an Excel file."""
    return askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])

def create_excel():
    """Create an Excel invoice with a vertical arrangement of 3 invoices per column."""
    # Prompt to select the Excel file
    df_path = select_excel_file()
    if not df_path:
        print("No file selected.")
        return

    # Read data from the Excel file
    data = pd.read_excel(df_path)

    # Create a new workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    # Define border and alignment styles
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    alignment = Alignment(horizontal='center', vertical='center')

    # Get today's date
    today_date = datetime.today().strftime("%d/%m/%Y")

    # Define output directory
    output_dir = r"G:\Certificate\Python Output"
    os.makedirs(output_dir, exist_ok=True)  # Ensure directory exists

    # Generate filename with timestamp
    current_datetime = datetime.now().strftime("%d.%m.%Y.%S")
    file_path = os.path.join(output_dir, f"MPB Bill {current_datetime}.xlsx")

    # Initialize row and column positions
    start_column = 1  # Start in the first column
    start_row = 1
    invoices_per_column = 3  # Stack 3 invoices per column
    row_offset = 15  # Space needed for each invoice
    column_offset = 4  # Space between columns (increase for clarity)

    for index, party in enumerate(data['PARTY NAME'].unique()):
        party_data = data[data['PARTY NAME'] == party]
        carats = party_data['Carats'].sum()
        bank_values = party_data['BANK$'].dropna().astype(str).unique()
        total_bank = bank_values[0] if len(bank_values) == 1 else ', '.join(bank_values)
        stone_id = party_data['stoneid'].dropna().astype(str).unique()
        total_stone_id = stone_id[0] if len(stone_id) == 1 else ','.join('✔')

        # Move to next column after every 3 invoices
        if index % invoices_per_column == 0 and index > 0:
            start_column += column_offset  # Shift to next column
            start_row = 1  # Reset row to top of column

        # Invoice template data
        data_template = [
            ["DATE :-", today_date, "", "", "", ""],  
            ["PARTY :-", party, "", "", "", ""],
            ["THRU:-", "", "", "", "", ""],
            ["QUALITY", "CARAT", "RS BILL", "", "", ""],
            ["", carats, "", "", "", ""],
            ["", "", "", "", "", ""],
            ["LESS (-)", "", "", "", "", ""],
            ["LESS (-)", "", "", "", "", ""],
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

                cell.alignment = alignment
                cell.border = thin_border

        # Set column widths explicitly
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 20

        # Move row offset for next invoice in the same column
        start_row += row_offset

    # Save and open the workbook
    wb.save(file_path)
    print(f"File successfully saved at: {file_path}")
    os.startfile(file_path)

    return file_path

if __name__ == "__main__":
    file_path = create_excel()
    print(f"Invoice saved at: {file_path}")