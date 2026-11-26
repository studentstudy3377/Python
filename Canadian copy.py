import os
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog

def remove_hyperlinks_except_v360_with_filters(
    file_path,
    column_to_remove_hyperlinks,   # e.g. 'STOCKID'
    column_to_preserve_hyperlinks, # e.g. 'V360'
    output_file_path
    
):
    """
    1) Removes hyperlinks from 'column_to_remove_hyperlinks' but preserves 
       hyperlinks in 'column_to_preserve_hyperlinks'.
    2) Applies filtering conditions (including background color on 'SHAPE')
       and outputs only rows that match.
    3) Moves the 'column_to_preserve_hyperlinks' column to the second column
       in the final spreadsheet.
    4) Removes the 'TRACR HYPERLINK' column from the final output.
    """
    # --- Step 1: Open original workbook and identify columns ---
    wb_original = openpyxl.load_workbook(file_path)
    ws_original = wb_original.active  # or specify a sheet name if needed
    
    # Get header from the first row
    header_row = [cell.value for cell in next(ws_original.iter_rows(min_row=1, max_row=1))]
    
    # Helper to get 0-based index of a column
    def get_col_index(col_name):
        if col_name not in header_row:
            raise ValueError(f"Column '{col_name}' not found in header.")
        return header_row.index(col_name)
    
    # Identify required columns (adjust names as needed)
    cert_col_index     = get_col_index("CERT")
    clarity_col_index  = get_col_index("CLARITY")
    cut_col_index      = get_col_index("CUT")
    pol_col_index      = get_col_index("POL")
    sym_col_index      = get_col_index("SYM")
    fluor_col_index    = get_col_index("FLUOR")
    status_col_index   = get_col_index("STATUS")
    rough_col_index    = get_col_index("ROUGH ORIGIN")
    location_col_index = get_col_index("STONE LOCATION")
    shape_col_index    = get_col_index("SHAPE")
    black_col_index    = get_col_index("BLACK")
    
    # Find the column index for the preserve-hyperlink column (1-based for openpyxl)
    col_preserve = None
    for col_idx, cell in enumerate(next(ws_original.iter_rows(min_row=1, max_row=1)), start=1):
        if cell.value == column_to_preserve_hyperlinks:
            col_preserve = col_idx
            break
    if col_preserve is None:
        raise ValueError(
            f"Could not find preserve-hyperlink column '{column_to_preserve_hyperlinks}' in header."
        )

    # --- Step 2: Build hyperlink map and determine rows to keep ---
    v360_hyperlink_map = {}   # Map: original Excel row -> (cell_value, hyperlink_url)
    keep_row_numbers = []     # Excel row numbers to keep
    
    # Restart iterator for rows (skip header)
    for row in ws_original.iter_rows(min_row=2, values_only=False):
        row_num = row[0].row  # Get Excel row number
        # Create a list of cell values from the row
        row_values = [cell.value for cell in row]
        if all(x is None for x in row_values):
            continue

        # Extract required cell values based on known indices
        cert_value     = row_values[cert_col_index]
        clarity_value  = row_values[clarity_col_index]
        cut_value      = row_values[cut_col_index]
        pol_value      = row_values[pol_col_index]
        sym_value      = row_values[sym_col_index]
        fluor_value    = row_values[fluor_col_index]
        status_value   = row_values[status_col_index]
        rough_value    = row_values[rough_col_index]
        location_value = row_values[location_col_index]
        black_value    = row_values[black_col_index]
        shape_value    = row_values[shape_col_index]

        # Capture V360 hyperlink info
        preserve_cell = row[col_preserve - 1]  # Adjust for 0-based index
        cell_value = preserve_cell.value
        hyperlink_url = preserve_cell.hyperlink.target if preserve_cell.hyperlink else None
        v360_hyperlink_map[row_num] = (cell_value, hyperlink_url)

        # Check the SHAPE cell background color (using the cell from the shape column)
        shape_cell = row[shape_col_index]
        fill_color = shape_cell.fill.start_color.rgb if shape_cell.fill.start_color.rgb else "00000000"
        if fill_color != "00000000":
            continue  # Skip rows with a fill color

        # Apply filtering conditions
        if (cert_value == "GIA" and
            cut_value != "GD" and
            pol_value != "GD" and
            sym_value != "GD" and
            rough_value == "CANADIAN" and
            location_value == "INDIA" and
            fluor_value in ["NONE", "FNT", "MED"] and
            clarity_value not in ["I1", "I2", "I3", "FL"] and
            (status_value in ["PROMOTION", None, ""]) and
            (black_value not in ["CENTER - H", "S & C - H", "SIDE - H", "RED SPOT", "HEAVY"]) and
            (shape_value != "ROUND" or cut_value == "3EX")
        ):
            keep_row_numbers.append(row_num)
    
    # --- Step 3: Load data with Pandas and filter by row numbers ---
    df = pd.read_excel(file_path)
    # Map Excel row number to DataFrame index: Excel row 2 -> index 0, etc.
    keep_df_indices = [r - 2 for r in keep_row_numbers if r - 2 < len(df)]
    df_filtered = df.iloc[keep_df_indices].copy()
    
    # --- Step 4: Remove hyperlink formatting from the specified column ---
    if column_to_remove_hyperlinks not in df_filtered.columns:
        raise ValueError(
            f"Column '{column_to_remove_hyperlinks}' not found in filtered DataFrame."
        )
    df_filtered[column_to_remove_hyperlinks] = df_filtered[column_to_remove_hyperlinks].astype(str)
    
    # Now, drop the "TRACR HYPERLINK" and "Blank_Column" columns (if they exist)
    cols_to_drop = [col for col in ["TRACR HYPERLINK", "Unnamed: 41"] if col in df_filtered.columns]
    if cols_to_drop:
        df_filtered.drop(columns=cols_to_drop, inplace=True)

    # Move the 'V360' (column_to_preserve_hyperlinks) to the second column
    cols = list(df_filtered.columns)
    if column_to_preserve_hyperlinks not in cols:
        raise ValueError(
            f"Column '{column_to_preserve_hyperlinks}' is not in the filtered DataFrame."
        )
    col_idx = cols.index(column_to_preserve_hyperlinks)
    cols.insert(1, cols.pop(col_idx))
    df_filtered = df_filtered[cols]
    
    # --- Step 6: Write the filtered DataFrame to the output file ---
    df_filtered.to_excel(output_file_path, index=False)
    
    # --- Step 7: Re-inject the V360 hyperlinks into the new file ---
    wb_new = openpyxl.load_workbook(output_file_path)
    ws_new = wb_new.active

    # Re-detect the 'column_to_preserve_hyperlinks' column index in the new file
    new_header = [cell.value for cell in next(ws_new.iter_rows(min_row=1, max_row=1))]
    if column_to_preserve_hyperlinks not in new_header:
        raise ValueError(
            f"Column '{column_to_preserve_hyperlinks}' not found in new file header."
        )
    col_preserve_new = new_header.index(column_to_preserve_hyperlinks) + 1  # 1-based indexing

    # Reinstate hyperlinks for each filtered row
    for i, original_df_index in enumerate(df_filtered.index, start=2):
        original_excel_row = original_df_index + 2  # Map back to original Excel row
        cell_value, hyperlink_url = v360_hyperlink_map.get(original_excel_row, (None, None))
        cell = ws_new.cell(row=i, column=col_preserve_new)
        cell.value = cell_value
        if hyperlink_url:
            cell.hyperlink = hyperlink_url
            cell.style = "Hyperlink"

    wb_new.save(output_file_path)

if __name__ == "__main__":
    # Let user select an Excel file
    root = tk.Tk()
    root.withdraw()
    selected_file = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    root.destroy()

    if not selected_file:
        print("No file selected. Exiting...")
        exit()

    # Build the output path in D:\Python\Canadian\Output
    output_folder = r"G:\Certificate\Python Output"
    os.makedirs(output_folder, exist_ok=True)
    base_name = os.path.basename(selected_file)          
    filename_no_ext, ext = os.path.splitext(base_name)   
    output_filename = f"Canadian.xlsx"
    output_file_path = os.path.join(output_folder, output_filename)
    
    # Run the filtering/hyperlink-preservation function
    remove_hyperlinks_except_v360_with_filters(
        file_path=selected_file,
        column_to_remove_hyperlinks="STOCKID",
        column_to_preserve_hyperlinks="V360",
        output_file_path=output_file_path
        
    )

    os.startfile(output_file_path)