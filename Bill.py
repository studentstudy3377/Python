import datetime as dt
import os
import platform
from tkinter import Tk, filedialog

import pandas as pd
from openpyxl.utils import get_column_letter


def select_and_load_excel_files():
    """
    Opens a file dialog to select multiple Excel or CSV files.
    Returns {filename: DataFrame} with optional custom skiprows logic.
    """
    root = Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(title="Select Excel/CSV Files to Merge",filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All Files", "*.*")],)
        
    skiprows_map = {}  # e.g., {"SOMEFILENAME": 1}
    dfs = {}

    for path in file_paths:
        filename = os.path.basename(path)
        ext = os.path.splitext(filename)[1].lower()
        skiprows = next(
            (v for k, v in skiprows_map.items() if k in filename.upper()), 0
        )

        try:
            if ext in (".xlsx", ".xls"):
                engine = "openpyxl" if ext == ".xlsx" else "xlrd"
                dfs[filename] = pd.read_excel(path, engine=engine, skiprows=skiprows)
            elif ext == ".csv":
                dfs[filename] = pd.read_csv(path, skiprows=skiprows)
        except Exception as e:
            print(f"❌ Failed to read file: {filename} — {e}")

    return dfs


def open_file(file_path):
    """Opens the given file using the default application (Windows only)."""
    if platform.system() == "Windows":
        os.startfile(file_path)


def pick_first_present_column(df: pd.DataFrame, candidates):
    """Return the first column from candidates that exists in df, else None."""
    for c in candidates:
        if c in df.columns:
            return c
    return None


def standardize_dataframe(df: pd.DataFrame, columns_mapping: dict) -> pd.DataFrame:
    """
    Create a standardized DataFrame with target columns defined by columns_mapping.
    For each target column, pick the first matching source column from the candidates.
    If none found, fill with empty strings.
    """
    out = pd.DataFrame()
    for target_col, candidates in columns_mapping.items():
        src = pick_first_present_column(df, candidates) if candidates else None
        if src is not None:
            out[target_col] = df[src]
        else:
            out[target_col] = ""
    return out


# --- For reference only (mirrors the Excel formula we write below) ---
def adjusted_last_price(p_last, disc1_total, disc2_total, disc1_last):
    diff = disc1_total - disc2_total
    return p_last + (diff * p_last) / disc1_last
# ---------------------------------------------------------------------


def main():
    # Columns that can serve as the "ID" source
    report_no_columns = ["Stock No", "Stone ID", "Diamond ID","REPORT"]

    # Final schema and mapping from many possible inputs
    columns_mapping = {
        "QualityCode": [],
        "FOR": [],
        "Size": [],
        "Pcs": [],
        "Carats": ["Carats","CTS"],
        "Price": ["Price",'PPC'],
        "Cost": [],       # will be filled by formula
        "Disc 1%": [],
        "Disc 2%": [],
        "Add %": [],
        "Remark": [],
        "stoneid": ["stoneid"],
        "PARTY NAME": [],
        "ID": report_no_columns,
        "SHAPE": [
            "SHP","shape","ShapeName","Shp.","SHAPE","Shp","Shape","shp","Shape and Cutting Style","SHAPE"
        ],
        "Color": [
            "COL","Colour Grade","color","Color Grade","ColorName","col","Colour","COLOR",
            "Color","DispColor","Col","COLOR"
        ],
        "Clarity": [
            "CLA","Clrt","Quality","purity","CTY","CLAR","Clarity Grade","ClarityName","Clr",
            "Clarity","CLARITY","Purity","CLR","DispClarity","Cl.","Cla","Clar","CLARITY"
        ],
        "DIS%": ["Disc",'DIS'],
        "VALUE": ["Value",'TOTAL'],
        "BANK$": ["BANK$"],
        "PUR DATE": ["Report Date"],
        "SALESMAN": [],
        "SALES PRICE": [],
        "STONE ID": ["Client Ref",'REPORT'],
        "LOT ID": [],
        "Origin": [],
        "BROKER NAME": [],
    }

    # 1) Load input files
    dfs = select_and_load_excel_files()
    if not dfs:
        raise ValueError("No files selected.")

    # 2) Build standardized frames per file, then append
    standardized_frames = []
    for _, df in dfs.items():
        standardized_frames.append(standardize_dataframe(df, columns_mapping))

    data = (
        pd.concat(standardized_frames, ignore_index=True)
        if standardized_frames
        else pd.DataFrame(columns=columns_mapping.keys())
    )

    # Ensure all target columns exist
    for col in columns_mapping.keys():
        if col not in data.columns:
            data[col] = ""

    # 3) Row-level formulas (no rounding)
    if "BANK$" in data.columns and "Price" in data.columns:
        bank_col_letter = get_column_letter(data.columns.get_loc("BANK$") + 1)
        price_col_letter = get_column_letter(data.columns.get_loc("Price") + 1)
        data["Cost"] = [
            f"={bank_col_letter}{row}*{price_col_letter}{row}"
            for row in range(2, len(data) + 2)
        ]

    if "Carats" in data.columns and "Cost" in data.columns:
        carat_col_letter = get_column_letter(data.columns.get_loc("Carats") + 1)
        cost_col_letter = get_column_letter(data.columns.get_loc("Cost") + 1)
        data["Disc 1%"] = [
            f"={carat_col_letter}{row}*{cost_col_letter}{row}"
            for row in range(2, len(data) + 2)
        ]
        
    # 4) Positions BEFORE appending summary rows
    start_row = 1
    end_row = len(data) + 1  # last data row index in Excel
    disc1_col = get_column_letter(data.columns.get_loc("Disc 1%") + 1)
    disc2_col = get_column_letter(data.columns.get_loc("Disc 2%") + 1)
    price_col = get_column_letter(data.columns.get_loc("Price") + 1)

    subtotal_row = end_row + 1  # where SUBTOTALs will be placed

    if all(col in data.columns for col in ["Price", "Disc 1%", "Disc 2%"]):
        price_last_cell = f"{price_col}{end_row}"
        disc1_last_cell = f"{disc1_col}{end_row}"
        disc1_total_cell = f"{disc1_col}{subtotal_row}"
        disc2_total_cell = f"{disc2_col}{subtotal_row}"

        # Excel formula version of:
        # p_last + ((Disc1_total - Disc2_total) * p_last) / disc1_last
        adj_formula = (
            f"={price_last_cell}-(({disc1_total_cell}-{disc2_total_cell})*"
            f"{price_last_cell})/{disc1_last_cell}"
        )

    
    # 5) Summary rows (compute positions BEFORE appending)
    subtotal = {c: "" for c in data.columns}    
    subtotal["Disc 1%"] = f"=SUBTOTAL(9,{disc1_col}{start_row}:{disc1_col}{end_row})"
    subtotal["Disc 2%"] = f"=SUBTOTAL(9,{disc2_col}{start_row}:{disc2_col}{end_row})"

    difference = {c: "" for c in data.columns}
    difference["Disc 1%"] = f"={disc1_col}{subtotal_row}-{disc2_col}{subtotal_row}"

    # Create DataFrame without blank row
    summary_rows = pd.DataFrame([subtotal, difference])

    # Append only non-blank rows (if there's no need for a blank row)
    data = pd.concat([data, summary_rows], ignore_index=True)

    if all(col in data.columns for col in ["Price", "Disc 1%", "Disc 2%"]):
        price_last_cell = f"{price_col}{end_row}"
        disc1_last_cell = f"{disc1_col}{end_row}"
        disc1_total_cell = f"{disc1_col}{subtotal_row}"
        disc2_total_cell = f"{disc2_col}{subtotal_row}"

        # Excel formula version of:
        # p_last + ((Disc1_total - Disc2_total) * p_last) / disc1_last
        adj_formula = (
            f"={price_last_cell}-(({disc1_total_cell}-{disc2_total_cell})*"
            f"{price_last_cell})/{disc1_last_cell}"
        )

        data["Price"] = data["Price"].astype(object)
    data.at[end_row - 0, "Price"] = adj_formula

    # 6) Save output (formulas preserved; no rounding)
    output_dir = r"G:\Certificate\Python Output"
    os.makedirs(output_dir, exist_ok=True)
    filename = f"Bill {dt.datetime.now().strftime('%d-%m-%Y')}.csv"
    output_path = os.path.join(output_dir, filename)

    data.to_csv(output_path, index=False)
    open_file(output_path)

if __name__ == "__main__":
    main()
    