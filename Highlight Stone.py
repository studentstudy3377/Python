import os
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# File paths
file_main = filedialog.askopenfilename(title="Stock List files", filetypes=[("Excel files", "*.xlsx")])
file_sr   = r"G:\Certificate\JYOTI\Gaurav\Do not Delete\Highlight Stone\Sr.xlsx"
output    = r"G:\Certificate\Python Output\Highlight Stone.xlsx"

# Read input Excel files
df  = pd.read_excel(file_main)   # Main data
dfs = pd.read_excel(file_sr)     # Sr data

# Convert dates to dd/mm/yyyy format
df['Pur Date'] = pd.to_datetime(df['Pur Date'], errors='coerce').dt.strftime('%d/%m/%Y')
df['E Date']   = pd.to_datetime(df['E Date'], errors='coerce').dt.strftime('%d/%m/%Y')

# Create new column 'Revers' = Depth - Table
df['Revers'] = df['Depth'] - df['Table']

# Merge DataFrames on 'Cts' from df and 'Sr' from dfs
df = pd.merge( df, dfs[['Sr']],left_on='Cts',right_on='Sr',how='left')

# Save the output Excel
df.to_excel(output, index=False, engine='openpyxl')

# Load workbook for highlighting
wb = load_workbook(output)
ws = wb.active

# Define highlight style (yellow fill)
highlight = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
red_font = Font(color="FF0000")
# Find column indexes for 'Sr. No', 'CUT', 'Cts'
headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}

col_srno = headers.get("Sr. No")
col_cut  = headers.get("CUT")
col_pol =  headers.get("POLISH")
col_sym =  headers.get("Symetry")
col_cts  = headers.get("Cts")
col_sr   = headers.get("Sr") 
col_depth   = headers.get("Depth") 
col_table   = headers.get("Table") 
col_rev   = headers.get("Revers") 
col_key   = headers.get("Key Notes")

# Apply highlight based on condition
for row in range(2, ws.max_row + 1):
    cts_value = ws.cell(row=row, column=col_cts).value
    cut_value = ws.cell(row=row, column=col_cut).value
    pol_value = ws.cell(row=row, column=col_pol).value
    sym_value = ws.cell(row=row, column=col_sym).value
    sr_value  = ws.cell(row=row, column=col_sr).value
    rev_value = ws.cell(row=row, column=headers.get("Revers")).value
    
    if cts_value == 0.55 and str(cut_value).strip().upper() == "V. GOOD":
        # Highlight Sr. No and CUT cells
        ws.cell(row=row, column=col_srno).fill = highlight
        ws.cell(row=row, column=col_cut).fill = highlight
            
    elif sr_value is not None and cut_value == "V. GOOD":
        # Highlight Sr. No and CUT cells
        ws.cell(row=row, column=col_srno).fill = highlight
        ws.cell(row=row, column=col_cut).fill = highlight
    elif cut_value == "EXCELLENT" and pol_value in ["V. GOOD","GOOD"]:
        # Highlight Sr. No and CUT cells
        ws.cell(row=row, column=col_srno).fill = highlight
        ws.cell(row=row, column=col_cut).fill = highlight
        ws.cell(row=row, column=col_pol).fill = highlight
    elif cut_value == "EXCELLENT" and sym_value in ["V. GOOD","GOOD"]:
        # Highlight Sr. No and CUT cells
        ws.cell(row=row, column=col_srno).fill = highlight
        ws.cell(row=row, column=col_cut).fill = highlight
        ws.cell(row=row, column=col_sym).fill = highlight
       
    elif rev_value is not None and isinstance(rev_value, (int, float)):
        if rev_value < 0:
            # Highlight Sr. No and CUT cells
            ws.cell(row=row, column=col_srno).fill = highlight
            ws.cell(row=row, column=col_depth).fill = highlight
            ws.cell(row=row, column=col_table).fill = highlight
            ws.cell(row=row, column=col_depth).font = red_font
            ws.cell(row=row, column=col_table).font = red_font

    key_value = ws.cell(row=row, column=col_key).value if col_key else None

    # Highlight if "Key Notes" contains any of the keywords
    keywords = ["GRAINING", "PINPOINTS", "CHIP", "CAVITY", "MANUFACTURING REMNANT"]

    if key_value is not None:
        key_upper = str(key_value).upper()
        if any(keyword in key_upper for keyword in keywords):
            ws.cell(row=row, column=col_srno).fill = highlight
            ws.cell(row=row, column=col_key).fill = highlight

headers = {cell.value.strip().upper(): idx for idx, cell in enumerate(ws[1], start=1)}

col_shape = headers.get("SHAPE")
col_table = headers.get("TABLE")
col_depth = headers.get("DEPTH")
col_ratio = headers.get("RATIO")

for row in range(2, ws.max_row + 1):
    shape_value = ws.cell(row=row, column=col_shape).value
    table_value = ws.cell(row=row, column=col_table).value
    depth_value = ws.cell(row=row, column=col_depth).value
    ratio_value = ws.cell(row=row, column=col_ratio).value

    if shape_value is not None and str(shape_value).strip().upper() == "MARQUISE":
        table_in_range = (
            table_value is not None and 56 <= float(table_value) <= 62
        )
        depth_in_range = (
            depth_value is not None and 58.00 <= float(depth_value) <= 65.5
        )
        ratio_in_range = (
            ratio_value is not None and 1.75 <= float(ratio_value) <= 2.19
        )

         # Always highlight the serial number column (col_srno)
        # ws.cell(row=row, column=col_srno).fill = highlight

        if not table_in_range:
            ws.cell(row=row, column=col_table).fill = highlight
        if not depth_in_range:
            ws.cell(row=row, column=col_depth).fill = highlight
        if not ratio_in_range:
            ws.cell(row=row, column=col_ratio).fill = highlight


for row in range(2, ws.max_row + 1):
    shape_value = ws.cell(row=row, column=col_shape).value
    table_value = ws.cell(row=row, column=col_table).value
    depth_value = ws.cell(row=row, column=col_depth).value
    ratio_value = ws.cell(row=row, column=col_ratio).value

    if shape_value is not None and str(shape_value).strip().upper() == "PEAR":
        table_in_range = (
            table_value is not None and 56 <= float(table_value) <= 62
        )
        depth_in_range = (
            depth_value is not None and 58.00 <= float(depth_value) <= 65.5
        )
        ratio_in_range = ( 
            ratio_value is not None and 1.45 <= float(ratio_value) <= 1.65
        )

         # Always highlight the serial number column (col_srno)
        # ws.cell(row=row, column=col_srno).fill = highlight

        if not table_in_range:
            ws.cell(row=row, column=col_table).fill = highlight
        if not depth_in_range:
            ws.cell(row=row, column=col_depth).fill = highlight
        if not ratio_in_range:
            ws.cell(row=row, column=col_ratio).fill = highlight


for row in range(2, ws.max_row + 1):
    shape_value = ws.cell(row=row, column=col_shape).value
    table_value = ws.cell(row=row, column=col_table).value
    depth_value = ws.cell(row=row, column=col_depth).value
    ratio_value = ws.cell(row=row, column=col_ratio).value

    if shape_value is not None and str(shape_value).strip().upper() == "OVAL":
        table_in_range = (
            table_value is not None and 56 <= float(table_value) <= 62
        )
        depth_in_range = (
            depth_value is not None and 58.00 <= float(depth_value) <= 65.5
        )
        ratio_in_range = (
            ratio_value is not None and 1.30 <= float(ratio_value) <= 1.55
        )

         # Always highlight the serial number column (col_srno)
        # ws.cell(row=row, column=col_srno).fill = highlight

        if not table_in_range:
            ws.cell(row=row, column=col_table).fill = highlight
        if not depth_in_range:
            ws.cell(row=row, column=col_depth).fill = highlight
        if not ratio_in_range:
            ws.cell(row=row, column=col_ratio).fill = highlight


for row in range(2, ws.max_row + 1):
    shape_value = ws.cell(row=row, column=col_shape).value
    table_value = ws.cell(row=row, column=col_table).value
    depth_value = ws.cell(row=row, column=col_depth).value
    ratio_value = ws.cell(row=row, column=col_ratio).value

    if shape_value is not None and str(shape_value).strip().upper() == "EMRALD":
        table_in_range = (
            table_value is not None and 61 <= float(table_value) <= 69
        )
        depth_in_range = (
            depth_value is not None and 62.00 <= float(depth_value) <= 71.00
        )
        ratio_in_range = (
            ratio_value is not None and 1.30 <= float(ratio_value) <= 1.55
        )

        # Always highlight the serial number column (col_srno)
        # ws.cell(row=row, column=col_srno).fill = highlight

        if not table_in_range:
            ws.cell(row=row, column=col_table).fill = highlight
        if not depth_in_range:
            ws.cell(row=row, column=col_depth).fill = highlight
        if not ratio_in_range:
            ws.cell(row=row, column=col_ratio).fill = highlight
        
for row in range(2, ws.max_row + 1):
    shape_value = ws.cell(row=row, column=col_shape).value
    table_value = ws.cell(row=row, column=col_table).value
    depth_value = ws.cell(row=row, column=col_depth).value
    ratio_value = ws.cell(row=row, column=col_ratio).value

    if shape_value is not None and str(shape_value).strip().upper() == "CUSHION":
        table_in_range = (
            table_value is not None and 58 <= float(table_value) <= 65
        )
        depth_in_range = (
            depth_value is not None and 66.00 <= float(depth_value) <= 71.50
        )
        ratio_in_range = (
            ratio_value is not None and 1.00 <= float(ratio_value) <= 1.20
        )

         # Always highlight the serial number column (col_srno)
        # ws.cell(row=row, column=col_srno).fill = highlight

        if not table_in_range:
            ws.cell(row=row, column=col_table).fill = highlight
        if not depth_in_range:
            ws.cell(row=row, column=col_depth).fill = highlight
        if not ratio_in_range:
            ws.cell(row=row, column=col_ratio).fill = highlight


for row in range(2, ws.max_row + 1):
    shape_value = ws.cell(row=row, column=col_shape).value
    table_value = ws.cell(row=row, column=col_table).value
    depth_value = ws.cell(row=row, column=col_depth).value
    ratio_value = ws.cell(row=row, column=col_ratio).value

    if shape_value is not None and str(shape_value).strip().upper() == "PRINCESS":
        table_in_range = (
            table_value is not None and 68 <= float(table_value) <= 75
        )
        depth_in_range = (
            depth_value is not None and 69.00 <= float(depth_value) <= 78.00
        )
        ratio_in_range = (
            ratio_value is not None and 1.00 <= float(ratio_value) <= 1.05
        )

         # Always highlight the serial number column (col_srno)
        # ws.cell(row=row, column=col_srno).fill = highlight

        if not table_in_range:
            ws.cell(row=row, column=col_table).fill = highlight
        if not depth_in_range:
            ws.cell(row=row, column=col_depth).fill = highlight
        if not ratio_in_range:
            ws.cell(row=row, column=col_ratio).fill = highlight

# Save changes
wb.save(output)

# Open the file
os.startfile(output)