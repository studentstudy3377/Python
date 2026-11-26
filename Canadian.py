import os
import time
import threading
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import openpyxl

def remove_hyperlinks_except_v360_with_filters(
    file_path,
    column_to_remove_hyperlinks,
    column_to_preserve_hyperlinks,
    output_file_path,
    progress_bar=None,
    root_window=None ):

    if progress_bar and root_window:
        progress_bar["value"] = 0
        root_window.update_idletasks()
        for i in range(10, 50, 5):
            time.sleep(0.1)
            progress_bar["value"] = i
            root_window.update_idletasks()

    wb_original = openpyxl.load_workbook(file_path)
    ws_original = wb_original.active  # or specify a sheet name if needed

    # Get header from the first row (assumes row 1 is the header)
    header_row = [cell.value for cell in next(ws_original.iter_rows(min_row=1, max_row=1))] # type: ignore

    def get_col_index(col_name):
        """Given a column name, return its index in header_row (0-based)."""
        if col_name not in header_row:
            raise ValueError(f"Column '{col_name}' not found in header.")
        return header_row.index(col_name)

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

    v360_hyperlink_map = {}
    keep_row_numbers = []

    for row in ws_original.iter_rows(min_row=2, values_only=False):
        row_num = row[0].row
        row_values = [cell.value for cell in row]
        if all(x is None for x in row_values):
            continue

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

        preserve_cell = row[col_preserve - 1]  # zero-based offset
        cell_value    = preserve_cell.value
        hyperlink_url = preserve_cell.hyperlink.target if preserve_cell.hyperlink else None
        v360_hyperlink_map[row_num] = (cell_value, hyperlink_url)

        # Exclude row if SHAPE cell is filled with a color
        shape_cell = row[shape_col_index]
        fill_color = shape_cell.fill.start_color.rgb or "00000000"
        if fill_color != "00000000":
            continue

        # Filtering conditions (example logic; adjust as needed)
        if (
            cert_value == "GIA" and
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

    df = pd.read_excel(file_path)
    keep_df_indices = [r - 2 for r in keep_row_numbers if (r - 2) < len(df)]
    df_filtered = df.iloc[keep_df_indices].copy()

    if column_to_remove_hyperlinks not in df_filtered.columns:
        raise ValueError(
            f"Column '{column_to_remove_hyperlinks}' not found in filtered DataFrame."
        )
    df_filtered[column_to_remove_hyperlinks] = df_filtered[column_to_remove_hyperlinks].astype(str)

    # Drop "TRACR HYPERLINK" or "Unnamed: 41" if present
    cols_to_drop = [c for c in ["TRACR HYPERLINK", "Unnamed: 41"] if c in df_filtered.columns]
    if cols_to_drop:
        df_filtered.drop(columns=cols_to_drop, inplace=True)

    df_filtered["CULET"] = (
        df_filtered["CULET"]
        .astype(str)       # ensure string type
        .apply(lambda x: x.strip())  # strip leading/trailing spaces
        .replace("nan", "None")           # or replace("", "None") to show literal text
    )
      # or replace("", "None") to show literal text
    cols = list(df_filtered.columns)
    if column_to_preserve_hyperlinks not in cols:
        raise ValueError(
            f"Column '{column_to_preserve_hyperlinks}' not found in the filtered DataFrame."
        )
    col_idx = cols.index(column_to_preserve_hyperlinks)
    cols.insert(1, cols.pop(col_idx))
    df_filtered = df_filtered[cols]

    df_filtered.to_excel(output_file_path, index=False)

    wb_new = openpyxl.load_workbook(output_file_path)
    ws_new = wb_new.active

    new_header = [cell.value for cell in next(ws_new.iter_rows(min_row=1, max_row=1))]
    if column_to_preserve_hyperlinks not in new_header:
        raise ValueError(f"Column '{column_to_preserve_hyperlinks}' not found in new file header.")
    col_preserve_new = new_header.index(column_to_preserve_hyperlinks) + 1

    for i, original_df_index in enumerate(df_filtered.index, start=2):
        original_excel_row = original_df_index + 2
        cell_value, hyperlink_url = v360_hyperlink_map.get(original_excel_row, (None, None))
        cell = ws_new.cell(row=i, column=col_preserve_new)
        cell.value = cell_value
        if hyperlink_url:
            cell.hyperlink = hyperlink_url
            cell.style = "Hyperlink"

    wb_new.save(output_file_path)

    if progress_bar and root_window:
        # Set progress bar to 100
        progress_bar["value"] = 100
        root_window.update_idletasks()

    # Attempt to open the file on Windows
    try:
        os.startfile(output_file_path)
    except AttributeError:
        # Not on Windows or startfile not available
        print(f"File saved to: {output_file_path}")
        print("os.startfile() is not available on this platform.")

    # Now remove the progress bar from the GUI
    if root and root_window:
        # Must schedule the destruction on the main thread
        root_window.after(0, root.destroy)


def start_processing():
    """Function to start processing in a separate thread."""
    selected_file = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not selected_file:
        return

    # Build the output path in G:\Certificate\Python Output (modify if needed)
    output_folder = r"G:\Certificate\Python Output"
    os.makedirs(output_folder, exist_ok=True)

    # We'll name the output file "Canadian.xlsx" in that folder
    output_filename = "Canadian.xlsx"
    output_file_path = os.path.join(output_folder, output_filename)

    # Run the processing in a separate thread to avoid freezing the GUI
    thread = threading.Thread(
        target=remove_hyperlinks_except_v360_with_filters,
        args=(selected_file, "STOCKID", "V360", output_file_path),
        kwargs={
            "progress_bar": progress,
            "root_window": root
        },
        daemon=True
    )
    thread.start()

root = tk.Tk()
root.title("")
root.geometry("400x200")

# Progress Bar
progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress.pack(pady=20)

# Start Button
start_button = ttk.Button(root, text="Select File & Process", command=start_processing)
start_button.pack(pady=10)

root.mainloop()