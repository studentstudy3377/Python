import os
from tkinter import filedialog
import openpyxl
import datetime
from openpyxl.styles import PatternFill, Font

# Path to source file
data = file1 = filedialog.askopenfilename(title="Stock List files", filetypes=[("Excel files", "*.xlsx")])
wb = openpyxl.load_workbook(data)
ws = wb.active

# Keep only one sheet
for sheet in wb.sheetnames:
    if sheet != ws.title:
        del wb[sheet]

# Read headers and column indexes
headers = [cell.value for cell in ws[1]]
col_idx = {header: idx + 1 for idx, header in enumerate(headers)}

# Add missing columns
for col in ['VALUE', 'PRICE', 'DISC']:
    if col not in headers:
        ws.cell(row=1, column=ws.max_column + 1, value=col)
        col_idx[col] = ws.max_column

# Function to style cell and try numeric conversion
def clean_and_convert(cell):
    # Preserve hyperlink style
    if cell.hyperlink:
        cell.font = Font(color="0000FF", underline="single")
    else:
        cell.font = Font(color="000000")
    cell.fill = PatternFill(fill_type=None)
    
    # Try numeric conversion if cell has no hyperlink and looks like a number
    if not cell.hyperlink and isinstance(cell.value, str):
        try:
            val = cell.value.strip()
            if val.replace('.', '', 1).isdigit():
                cell.value = float(val) if '.' in val else int(val)
        except:
            pass

# Main logic
for row in ws.iter_rows(min_row=2, values_only=False):
    if row[col_idx['CERT'] - 1].value != 'GIA':
        continue

    carat = row[col_idx['CARAT'] - 1].value
    value = row[col_idx['VALUE'] - 1].value

    # Convert value to a number if it's a string
    if isinstance(value, str):
        try:
            value = float(value) if '.' in value else int(value)
        except ValueError:
            value = 0  # Fallback if the value is not a valid number

    # Carat addition logic
    if carat is not None:
        if 0.15 <= carat <= 0.22: addition = 24
        elif 0.23 <= carat <= 0.29: addition = 25.6
        elif 0.30 <= carat <= 0.39: addition = 31.2
        elif 0.40 <= carat <= 0.49: addition = 32.8
        elif 0.50 <= carat <= 0.69: addition = 33.6
        elif 0.70 <= carat <= 0.99: addition = 40.8
        elif 1.00 <= carat <= 1.49: addition = 68
        elif 1.50 <= carat <= 1.99: addition = 88
        else: addition = 0
    else:
        addition = 0

    new_value = (value or 0) + round(addition, 2)
    row[col_idx['VALUE'] - 1].value = new_value

    # PRICE
    price = round(new_value / carat, 2) if carat else 0
    row[col_idx['PRICE'] - 1].value = price

    # DISC
    rapprice = row[col_idx['RAPPRICE'] - 1].value
    disc = round((price - rapprice) / rapprice * 100, 2) if rapprice else 0
    row[col_idx['DISC'] - 1].value = disc

# Save to new file
today_str = datetime.datetime.now().strftime("%d-%m-%Y")
output_filename = fr"G:\Certificate\GIA Pricing {today_str}.xlsx"
wb.save(output_filename)
print(f"Updated file saved as: {output_filename}")
os.startfile(output_filename)