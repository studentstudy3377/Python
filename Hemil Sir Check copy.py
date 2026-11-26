import re
import openpyxl
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

# Expected columns and column mapping for standardization
expected_columns = [
    "STOCKID", "V360", "SHAPE", "RAPSIZE", "CERT", "CARAT", "COLOR", "CLARITY", "CUT",
    "POL", "SYM", "FLUOR", "DISC", "MEASUREMENT", "DEPTH", "TABLE", "BLACK", "LUSTER",
    "TINCH", "KEY TO SYMBOLS", "REPORT COMM", "RATIO", "CERTNO", "STATUS", "RAPPRICE"
]

column_mapping = {
    "STOCK ID": "STOCKID",
    "POLISH": "POL",
    "SYMMETRY": "SYM",
    "BLACK INCLUSION": "BLACK",
    "TINGE": "TINCH",
    "KEY TO SYMBOL": "KEY TO SYMBOLS",
    "CERTIFIED NO": "CERTNO",
    "FINAL CUSTOMER": "NAME",
    "DATE": "Sold Date",
    "RAPPRICE($)": "RAPPRICE"
}

def read_excel_standardized(file_path, skip_rows=0, add_disc=False):
    """Read and standardize Excel data using openpyxl."""
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    # Read header row (after skipping required rows)
    headers = [cell.value for cell in sheet[skip_rows + 1]]
    # Standardize column names
    standardized_headers = [column_mapping.get(h, h) for h in headers]

    # Read data into a list
    data = []
    for row in sheet.iter_rows(min_row=skip_rows + 2, values_only=True):
        row_dict = {standardized_headers[i]: row[i] for i in range(len(row))}
        data.append(row_dict)
    return data

def select_file(title: str) -> str:
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not file_path:
        return None
    return file_path

def load_workbook_data(filepath: str, data_only: bool = False):
    """Load workbook data and return a DataFrame, workbook, and worksheet."""
    wb = load_workbook(filepath, data_only=data_only)
    ws = wb.active
    data_rows = [[cell for cell in row] for row in ws.iter_rows(values_only=True)]
    headers, *data = data_rows
    df = pd.DataFrame(data, columns=headers)
    return df, wb, ws

def extract_formulas(filepath: str, column_letter: str) -> list:
    """Extract formulas from a specific column in an Excel file."""
    wb = load_workbook(filepath, data_only=False)
    ws = wb.active
    return [
        cell.value
        for row in ws.iter_rows(min_row=2, max_col=ws.max_column)
        for cell in row
        if cell.column_letter == column_letter
    ]

def reapply_formulas(ws, formulas: list, column_letter: str):
    """Reapply formulas to a worksheet in a specific column."""
    for idx, formula in enumerate(formulas, start=2):
        ws[f"{column_letter}{idx}"].value = formula

def apply_style(ws, row: int, col: int, fill_color: str = None, font_color: str = None, bold: bool = False):
    """Apply styling to a specific cell."""
    cell = ws.cell(row=row, column=col)
    cell.font = Font(name="Calibri", size=11, color=font_color, bold=bold)
    if fill_color:
        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

def is_valid_date(value):
    """Check if a value is a valid date."""
    try:
        pd.to_datetime(value)
        return True
    except Exception:
        return False

def apply_styles_to_excel(ws, df):
    """Apply conditional formatting and styling to the Excel worksheet based on DataFrame values."""
    green_fill_pattern = PatternFill(start_color="CFEAF7", end_color="CFEAF7", fill_type="solid")
    yellow_fill_pattern = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    combined_fill_pattern = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")

    for i, row in df.iterrows():
        excel_row = i + 2  # Adjust for header row (Excel rows start at 1)

        # Write all cell values for the row
        for col_idx, col_name in enumerate(df.columns, start=1):
            value = row[col_name]
            if pd.isna(value):
                value = None
            ws.cell(row=excel_row, column=col_idx, value=value)

        if col_name == "FLUOR" and val != "NONE":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)


        sold = "SOURCE" in df.columns and row.get("SOURCE") == "Sold"
        new_arrival = (
            "STATUS" in row and 
            isinstance(row["STATUS"], str) and 
            row["STATUS"].strip() == "New Arrival"
        )

        if sold and new_arrival:
            fill = combined_fill_pattern
        elif sold:
            fill = green_fill_pattern
        elif new_arrival:
            fill = yellow_fill_pattern
        else:
            fill = None

        if fill is not None:
            for col_idx in range(1, len(df.columns) + 1):
                ws.cell(row=excel_row, column=col_idx).fill = fill

        clarity_val = row.get("CLARITY", None)
        depth_val = row.get("DEPTH", None)
        table_val = row.get("TABLE", None)
        rapsize_val = row.get("RAPSIZE", None)
        measurement_val = row.get("MEASUREMENT", None)
        if isinstance(measurement_val, str):
            measurement_val = pd.to_numeric(measurement_val, errors="coerce")


        for col_idx, col_name in enumerate(df.columns, start=1):
            val = row[col_name]
            if col_name == "STATUS" and val == "New Arrival":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
                # if "DISC" in df.columns:
                #     disc_col_idx = df.columns.get_loc("DISC") + 1
                #     ws.cell(row=excel_row, column=disc_col_idx, value=0)

            elif col_name == "FLUOR" and val != "NONE":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
             # **Small** only when it has a real value
            elif col_name == "Small" and pd.notna(val):
                apply_style(ws, excel_row, col_idx, fill_color="99FF99")

            elif col_name == "POL" and val == "GD":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "CERT" and val in ["HRD", "IGI"]:
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "SYM" and val == "GD":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "CLARITY" and val == "I1":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "CUT" and val == "GD":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "TINCH" and val == "B-1":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "FLUOR" and val != "NONE":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "SHAPE" and val != "ROUND":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "LUSTER" and val != "EX":
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "BLACK" and isinstance(val, str) and "H" in val:
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif (
                col_name == "KEY TO SYMBOLS"
                and isinstance(val, str)
                and any(
                    keyword in val
                    for keyword in [
                        "Etch Channel- Cloud",
                        "Feather- Chip- Surface Graining",
                        "Internal Graining- Feather- Cavity- Surface Graining",
                    ]
                )
            ):
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif (
                col_name == "KEY TO SYMBOLS"
                and isinstance(val, str)
                and any(
                    keyword in val
                    for keyword in [
                        "Cavity",
                        "Clarity",
                        "Manufacturing Remnant",
                        "Graining",
                        "Internal Laser Drilling",
                        "Laser Drill-hole",
                        "Knot",
                    ]
                )
            ):
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif (
                col_name == "KEY TO SYMBOLS"
                and isinstance(val, str)
                and "Internal Graining" in val
                and clarity_val in ["IF", "FL", "VVS1", "VVS2"]
            ):
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif (
                col_name == "KEY TO SYMBOLS"
                and isinstance(val, str)
                and "Surface Graining" in val
                and clarity_val in ["IF", "FL", "VVS1", "VVS2"]
            ):
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif (
                col_name == "DEPTH"
                and depth_val is not None
                and table_val is not None
                and depth_val < table_val
            ):
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif (
                col_name == "TABLE"
                and table_val is not None
                and depth_val is not None
                and depth_val < table_val
            ):
            
                apply_style(ws, excel_row, col_idx, fill_color="FFCCCC", font_color="FF5050", bold=True)
            elif col_name == "MEASUREMENT" and rapsize_val == "30UP" and measurement_val > 4.50:
                apply_style(ws, excel_row, col_idx, fill_color="FF9966")
            elif col_name == "MEASUREMENT" and rapsize_val == "40UP" and measurement_val > 5.00:
                apply_style(ws, excel_row, col_idx, fill_color="FF9966")
            elif col_name == "MEASUREMENT" and rapsize_val == "50UP" and measurement_val > 5.25:
                apply_style(ws, excel_row, col_idx, fill_color="FF9966")
            elif col_name == "MEASUREMENT" and rapsize_val == "60UP" and measurement_val > 5.50:
                apply_style(ws, excel_row, col_idx, fill_color="FF9966")
            elif col_name == "MEASUREMENT" and rapsize_val == "70UP" and measurement_val > 5.85:
                apply_style(ws, excel_row, col_idx, fill_color="FF9966")
            elif col_name == "MEASUREMENT" and rapsize_val == "80UP" and measurement_val > 6.00:
                apply_style(ws, excel_row, col_idx, fill_color="FF9966")
            elif col_name == "MEASUREMENT" and rapsize_val == "90UP" and measurement_val > 6.15:
                apply_style(ws, excel_row, col_idx, fill_color="FF9966")
            elif col_name == "MEASUREMENT" and rapsize_val == "1.00UP" and measurement_val > 6.30:
                apply_style(ws, excel_row, col_idx, fill_color="FF9966")

def save_data_with_formulas(df: pd.DataFrame, wb, output_path: str):
    """Save the final DataFrame into a new sheet named 'Matching', and remove all other sheets."""
    if "Matching" in wb.sheetnames:
        del wb["Matching"]
    ws = wb.create_sheet("Matching")
    for col_idx, header in enumerate(df.columns, start=1):
        ws.cell(row=1, column=col_idx, value=header)
    apply_styles_to_excel(ws, df)
    unwanted_cols = [
        "FINAL PRICE", "DISCOUNT", "PRICE/Ct($)", "AMOUNT", "RAPPRICE($)",
        "ONLINE DISC.", "ONLINE AMOUNT", "SELLING DISC.", "SELLING PRICE",
        "SELLING AMOUNT", "CULET", "BLACK INCLU. REMARK", "EYE CLEAN",
        "TABLE CLEAN", "CROWN ANGLE", "CROWN HEIGHT", "PAVILION ANGLE",
        "PAVILION HEIGHT", "GIRDLE", "GIRDLE %", "STAR LENGTH", "LOWER HALF",
        "PROPORTION", "H & A", "TOP OPENS", "BOTTOM OPENS", "LOCATION",
        "CONSIGNEE PARTY", "LAB STATUS", "CERTIFICATE", "VIDEO LINK", "IMAGES", "SOURCE", None,
        "REPORT COMMENTS", "Display_Text"
    ]
    headers = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]
    cols_to_delete = [idx for idx, header in enumerate(headers, start=1) if header in unwanted_cols]
    for col_idx in sorted(cols_to_delete, reverse=True):
        ws.delete_cols(col_idx)
    for sheet in wb.sheetnames.copy():
        if sheet != "Matching":
            del wb[sheet]
    wb.save(output_path)

def add_empty_row(df: pd.DataFrame) -> pd.DataFrame:
    """Append a blank row (full of NA) at the bottom of a DataFrame."""
    empty = pd.DataFrame([[pd.NA] * len(df.columns)], columns=df.columns)
    return pd.concat([df, empty], ignore_index=True)

def open_file(file_path: str):
    """Open file using the default application (Windows-only for os.startfile)."""
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        else:
            subprocess.call(["xdg-open", file_path])
    except Exception as e:
        print(f"Error opening file: {e}")

def find_carat_range(carat: float, shape: str,
                     fluor=None, luster=None, tinch=None, black=None):
    """
    Find the appropriate carat range for a given carat value.
    - If shape != "ROUND", immediately use the broad non-round ranges.
    - Otherwise, if fluor, luster, tinch, and black meet the special criteria,
      use an alternate “tighter” set; else use the default round ranges.
    """
    # First, if not round, always use the broad non-round list:
    if shape is not None and str(shape).strip().upper() != "ROUND":
        ranges = [
            (0.23, 0.29), (0.30, 0.39), (0.40, 0.49), (0.50, 0.59), (0.60, 0.69),
            (0.70, 0.79), (0.80, 0.89), (0.90, 0.99), (1.00, 1.19), (1.20, 1.29),
            (1.30, 1.39), (1.40, 1.49), (1.50, 1.59), (1.60, 1.69), (1.70, 1.79),
            (1.80, 1.89), (1.90, 1.99), (2.00, 2.19), (2.20, 2.49), (2.50, 2.99),
            (3.00, 3.49), (3.50, 3.99), (4.00, 4.49), (4.50, 4.99), (5.00, 5.49),
            (5.50, 5.99)
        ]
    else:
        # For ROUND, decide between alternate and default based on the
        # fluor/luster/tinch/black criteria:
        special = (
            fluor is not None and str(fluor).strip().upper() == "MED" and
            luster is not None and str(luster).strip().upper() in ["M-1", "M-2","M1","M2"] and
            tinch is not None and str(tinch).strip().upper() in ["B1", "B2"] and
            black is not None and isinstance(black, str) and "H" in black
        )
        if special:
            ranges = [
                (0.23, 0.29), (0.30, 0.39), (0.40, 0.49), (0.50, 0.59), (0.60, 0.69),
                (0.70, 0.79), (0.80, 0.89), (0.90, 0.99), (1.00, 1.19), (1.20, 1.29),
                (1.30, 1.39), (1.40, 1.49), (1.50, 1.59), (1.60, 1.69), (1.70, 1.79),
                (1.80, 1.89), (1.90, 1.99), (2.00, 2.19), (2.20, 2.49), (2.50, 2.99),
                (3.00, 3.49), (3.50, 3.99), (4.00, 4.49), (4.50, 4.99), (5.00, 5.49),
                (5.50, 5.99)
            ]
        else:
            ranges = [
                (0.23, 0.29), (0.30, 0.36), (0.37, 0.39), (0.40, 0.45), (0.46, 0.49),
                (0.50, 0.59), (0.60, 0.69), (0.70, 0.79), (0.80, 0.89), (0.90, 0.99),
                (1.00, 1.19), (1.20, 1.29), (1.30, 1.39), (1.40, 1.49), (1.50, 1.59),
                (1.60, 1.69), (1.70, 1.79), (1.80, 1.89), (1.90, 1.99), (2.00, 2.19),
                (2.20, 2.49), (2.50, 2.99), (3.00, 3.49), (3.50, 3.99), (4.00, 4.49),
                (4.50, 4.99), (5.00, 5.49), (5.50, 5.99)
            ]
    # Now locate and return the matching bucket
    for start, end in ranges:
        if start <= carat <= end:
            return (start, end)
    return None

def get_allowed_cut(cut):
    """Return allowable cuts based on a hierarchy."""
    cut_hierarchy = {
        'EX': ['EX'],
        'VG': ['EX', 'VG'],
        'GD': ['EX', 'VG', 'GD'],
        'FR': ['EX', 'VG', 'GD', 'FR']
    }
    if pd.isna(cut):
        return [pd.NA]
    else:
        return cut_hierarchy.get(cut, ['EX', 'VG', 'GD', 'FR'])

def get_allowed_cert(cert):
    """Return allowable certificates based on the input cert."""
    cert_hierarchy = {
        'GIA':   ['GIA'],
        'HRD':   ['HRD'],
        'IGI':   ['IGI'],
        'IIDGR': ['IIDGR']
    }
    if pd.isna(cert):
        return [pd.NA]
    else:
        return cert_hierarchy.get(cert, ['GIA', 'HRD', 'IGI','IIDGR'])
    
def get_allowed_fluor(fluor):
    """Return allowable fluor values based on the input fluor."""
    fluor_hierarchy = {
        'NONE': ['NONE'],
        'FNT':  ['NONE', 'FNT', 'VSL'],
        'MED':  ['NONE', 'FNT', 'VSL', 'MED'],
        'SL':   ['NONE', 'FNT', 'VSL', 'MED', 'SL'],
        'STG':  ['NONE', 'FNT', 'VSL', 'MED', 'SL', 'STG'],
        'VST':  ['NONE', 'FNT', 'VSL', 'MED', 'SL', 'STG', 'VST']
    }
    return fluor_hierarchy.get(fluor, ['NONE', 'FNT', 'VSL', 'MED', 'SL', 'STG', 'VST'])

def replace_low(df: pd.DataFrame) -> pd.DataFrame:
    """If 'Luster' column says LOW, copy to 'LUSTER' column."""
    if "Luster" in df.columns:
        df.loc[df["Luster"] == "LOW", "LUSTER"] = "LOW"
    return df

def get_code(row):
    """
    Generate a code for a 'New Arrival' stone:
      - 'PZ' if shape != ROUND or carat >= 1.00 or clarity == I1
      - 'SK' if carat >= 1.00 OR (clarity in VS2, SI1, SI2, I1 and carat is 0.23-0.99)
      - 'SP' if clarity in (IF, VVS1, VVS2, VS1)
    """
    # Check for "New Arrival" exactly (case-sensitive, after stripping whitespace)
    if not (isinstance(row.get("STATUS"), str) and row["STATUS"].strip() == "New Arrival"):
        return ""
    if pd.isnull(row["CLARITY"]) or str(row["CLARITY"]).strip() == "":
        return ""
    clarity = str(row["CLARITY"]).upper()
    shape = str(row["SHAPE"]).upper() if pd.notnull(row["SHAPE"]) else ""
    try:
        carat = float(row["CARAT"])
    except (ValueError, TypeError):
        return ""
    codes = []
    if shape != "ROUND" or carat >= 1.00 or clarity == "I1":
        codes.append("PZ")
    if carat >= 1.00 or (clarity in ("VS2", "SI1", "SI2", "I1") and 0.23 <= carat <= 0.99):
        codes.append("SK")
    if clarity in ("IF", "VVS1", "VVS2", "VS1"):
        codes.append("SP")
    return "/".join(codes) if codes else ""

def calculate_sold_disc(row):
    """Calculate discount if SOURCE == 'Sold'."""
    if row.get("SOURCE") != "Sold":
        return None
    rap_price = row.get("RAPPRICE", 0)
    final_price = row.get("FINAL PRICE", 0)
    return round(((final_price - rap_price) / rap_price) * 100, 2) if rap_price != 0 else 0

def remove_sold_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Remove rows entirely that have SOURCE == 'Sold'."""
    return df[df['SOURCE'] != 'Sold'].reset_index(drop=True)

def append_sources_to_book2(premal, sachin, sagar, book2_path=r"G:\Certificate\PREMAL\NOT TO DELETE\Disc.xlsx"):
    """
    Appends Premal, Sachin, and Sagar data into a single sheet 'Combined' in Book2.xlsx.
    Columns: STOCKID, Value, Source, Date
    """
    today_str = datetime.now().strftime('%d-%m-%Y')

    # Standardize and tag source
    premal_df = premal[['stockid', 'pz']].copy()
    premal_df.columns = ['STOCKID', 'Disc']
    premal_df['Code'] = 'PZ'

    sachin_df = sachin[['stockid', 'sk']].copy()
    sachin_df.columns = ['STOCKID', 'Disc']
    sachin_df['Code'] = 'SK'

    sagar_df = sagar[['stockid', 'sp']].copy()
    sagar_df.columns = ['STOCKID', 'Disc']
    sagar_df['Code'] = 'SP'

    # Combine and add date
    combined_df = pd.concat([premal_df, sachin_df, sagar_df], ignore_index=True)
    combined_df.dropna(subset=['STOCKID', 'Disc'], inplace=True)
    combined_df['Date'] = today_str

    # Load or create workbook
    if os.path.exists(book2_path):
        wb = load_workbook(book2_path)
        ws = wb.active
        if ws.max_row == 1 and all(cell.value is None for cell in ws[1]):
            ws.append(['STOCKID', 'Disc', 'Code', 'Date'])
    else:
        ws = wb.active
        ws.title = 'Combined'
        ws.append(['STOCKID', 'Disc', 'Code', 'Date'])

    # Append rows
    for _, row in combined_df.iterrows():
        ws.append([row['STOCKID'], row['Disc'], row['Code'], row['Date']])

    wb.save(book2_path)

def match_stones():
    """
    Main function to match stones, group them, and sort groups + matched stones by ascending CARAT.
    """
    # ----------------- File Inputs ---------------------
    input_file = select_file("Select the Stock file")
    sold_file = select_file("Select the Sold file")
    days_file = select_file("Select the Days file")
    premal_file = select_file("Select the Premal file")
    sachin_file = select_file("Select the Sachin file")
    sagar_file = select_file("Select the Sagar file")
    revision_file = r"G:\Certificate\PREMAL\REVISE DAYS TRACKING.xlsx"

    # -------------- Read supporting files --------------
    if premal_file:
        premal = pd.read_excel(premal_file)
        premal.columns = premal.columns.str.lower()
    else:
        premal = pd.DataFrame(columns=["stockid", "pz"])

    if sachin_file:
        sachin = pd.read_excel(sachin_file)
        sachin.columns = sachin.columns.str.lower()
    else:
        sachin = pd.DataFrame(columns=["stockid", "sk"])

    if sagar_file:
        sagar = pd.read_excel(sagar_file)
        sagar.columns = sagar.columns.str.lower()
    else:
        sagar = pd.DataFrame(columns=["stockid", "sp"])

    # --------------- Read main input files ----------------
    data1, wb, ws = load_workbook_data(input_file, data_only=False)
    data1["SOURCE"] = "Stock"

    desired_order = [
        "STOCKID", "SHAPE", "CERT", "CARAT", "COLOR", "CLARITY",
        "CUT", "POL", "SYM", "FLUOR", "MEASUREMENT", "DEPTH",
        "TABLE", "KEY TO SYMBOLS", "DISC", "BLACK", "LUSTER",
        "TINCH", "STATUS", "RATIO",  "REPORT COMM", 
        "NAME", "Sold Date","CERTNO", "RAPPRICE","ONLINE PRICE",
        "V360", "RAPSIZE","SOURCE", 
    ]
    for col in desired_order:
        if col not in data1.columns:
            data1[col] = None
    data1 = data1.reindex(columns=desired_order)

    data_values, _, _ = load_workbook_data(input_file, data_only=True)
    data2 = read_excel_standardized(sold_file, skip_rows=2, add_disc=True)
    for row in data2:
        if "FINAL PRICE" in row:
            row["ONLINE PRICE"] = row.get("FINAL PRICE")
    data2_df = pd.DataFrame(data2)
    data2_df["SOURCE"] = "Sold"

    data = pd.concat([data1, data2_df], ignore_index=True)
    data = data[~data["CERT"].str.contains("IIDGR", case=False, na=False)]
    data["STOCKID_VALUES"] = data_values["STOCKID"].astype(str).str.strip()
    data['Display_Text'] = data['CERT'].str.extract(r'"([^\"]+)"\)$')[0].fillna(data['CERT'])
    stockid_formulas = extract_formulas(input_file, 'A')

    revision = pd.read_excel(revision_file, dtype=str)
    days = pd.read_excel(days_file, dtype=str)
    for df in (revision, days):
        df["Sr. No"] = df["Sr. No"].str.strip()

    merged_df = (
        data
        .merge(revision[['Sr. No', 'DAYS']], left_on='STOCKID_VALUES', right_on='Sr. No', how='left')
        .rename(columns={'DAYS': 'R D'})
        .merge(days[['Sr. No', 'Stock Days', 'Luster']], left_on='STOCKID_VALUES', right_on='Sr. No', how='left')
        .rename(columns={'Stock Days': 'O D'})
        .merge(days[['Sr. No', 'DISC']], left_on='STOCKID_VALUES', right_on='Sr. No', how='left')
        .rename(columns={'DISC_y': 'List Disc', 'DISC_x': 'DISC'})
        .merge(premal[['stockid', 'pz']], left_on='STOCKID_VALUES', right_on='stockid', how='left')
        .rename(columns={'pz': 'Premal'})
        .merge(sachin[['stockid', 'sk']], left_on='STOCKID_VALUES', right_on='stockid', how='left')
        .rename(columns={'sk': 'Sachin'})
        .merge(sagar[['stockid', 'sp']], left_on='STOCKID_VALUES', right_on='stockid', how='left')
        .rename(columns={'sp': 'Sagar'})
    )
    merged_df = replace_low(replace_low(merged_df))
    for col in ['Premal', 'Sachin', 'Sagar']:
        merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce')
    merged_df['Small'] = merged_df[['Premal', 'Sachin', 'Sagar']].min(axis=1)

    # blank out 'Small' for non–New Arrival rows where DISC equals Small (numeric comparison)
    disc_numeric = pd.to_numeric(merged_df['DISC'], errors='coerce')
    small_numeric = merged_df['Small']
    mask = (merged_df["STATUS"].str.lower() != "new arrival") & (disc_numeric == small_numeric)
    merged_df.loc[mask, 'Small'] = np.nan

    drop_cols = [c for c in ['Sr. No_x', 'StockID', 'stockid', 'stockid_x', 'stockid_y', 'Sr. No_y', 'Sr. No', 'Luster', 'STOCKID_VALUES'] if c in merged_df]
    merged_df.drop(columns=drop_cols, inplace=True)
    merged_df[['R D', 'O D', 'List Disc']] = merged_df[['R D', 'O D', 'List Disc']].apply(pd.to_numeric, errors='coerce')
    merged_df = merged_df[~((merged_df['STATUS'].str.strip() == 'New Arrival') & (merged_df['Small'].isna()))]

    groups = []
    matched_new_arrivals = set()
    merged_df['STATUS'] = merged_df['STATUS'].astype(str).str.strip()
    merged_df['SHAPE'] = merged_df['SHAPE'].astype(str).str.strip().str.upper()
    new_arrivals = merged_df[merged_df['STATUS'] == 'New Arrival']
    new_arrivals = pd.concat([new_arrivals[new_arrivals['SHAPE']=='ROUND'], new_arrivals[new_arrivals['SHAPE']!='ROUND']])
    
 

    for _, new_stone in new_arrivals.iterrows():
        if new_stone['STOCKID'] in matched_new_arrivals:
            continue
        matched_new_arrivals.add(new_stone['STOCKID'])
        carat_range = find_carat_range(
            new_stone['CARAT'], shape=new_stone['SHAPE'],
            fluor=new_stone.get('FLUOR'), luster=new_stone.get('LUSTER'),
            tinch=new_stone.get('TINCH'), black=new_stone.get('BLACK')
        )
        if not carat_range:
            grp_df = add_empty_row(new_stone.to_frame().T)
            groups.append((new_stone['CARAT'], grp_df))
            continue

        query = [
            (merged_df['SHAPE'] == new_stone['SHAPE']),
            (merged_df['CARAT'].between(*carat_range)),
            (merged_df['COLOR'] == new_stone['COLOR']),
            (merged_df['CLARITY'] == new_stone['CLARITY']),
            (merged_df['LUSTER'] == new_stone['LUSTER']),
            (merged_df['TINCH'] == new_stone['TINCH']),
            (merged_df['Display_Text'].str.contains('|'.join(get_allowed_cert(new_stone['Display_Text'])), na=False)),
            (~merged_df['STOCKID'].isin(matched_new_arrivals)),
        ]
        if new_stone['STATUS'].lower() == 'new arrival':
            for attr in ('CUT', 'POL', 'SYM', 'FLUOR', 'LUSTER', 'TINCH'):
                query.append((merged_df['SOURCE'] != 'Sold') | (merged_df[attr] == new_stone[attr]))
        if new_stone['SHAPE'] == 'ROUND':
            query.extend([
                merged_df['CUT'].isin(get_allowed_cut(new_stone['CUT'])),
                merged_df['POL'].isin(get_allowed_cut(new_stone['POL'])),
                merged_df['SYM'].isin(get_allowed_cut(new_stone['SYM'])),
                merged_df['FLUOR'].isin(get_allowed_fluor(new_stone['FLUOR']))
            ])
        else:
            query.append(merged_df['FLUOR'].isin(get_allowed_fluor(new_stone['FLUOR'])))

        matches = merged_df[np.all(query, axis=0)]
        matched_new_arrivals.update(matches['STOCKID'])
        main_df = new_stone.to_frame().T
        match_df = matches.sort_values('CARAT')

        # sort sold by date if present
        if 'Sold Date' in match_df.columns:
            match_df['Date'] = pd.to_datetime(match_df['Sold Date'], errors='coerce')
            match_df['SRC_ORDER'] = match_df['SOURCE'].map({'Stock': 0, 'Sold': 1})
            match_df.sort_values(['SRC_ORDER','Date'], ascending=[True, False], inplace=True)
            match_df.drop(columns=['Date','SRC_ORDER'], inplace=True)

        # recalc discounts
        def calc_disc(r):
            return round(((r.get('FINAL PRICE',0) - r.get('RAPPRICE',0)) / r.get('RAPPRICE',1)) * 100, 2) if r['SOURCE'] == 'Sold' else r.get('DISC')
        for df in (main_df, match_df):
            if 'RAPPRICE' in df.columns:
                df['RAPPRICE'] = main_df.iloc[0]['RAPPRICE']
            if 'FINAL PRICE' in df.columns:
                mask = df['SOURCE'] != 'Stock'
                df.loc[mask, 'DISC'] = df.loc[mask].apply(calc_disc, axis=1)

        grp = pd.concat([main_df, match_df], ignore_index=True)
        grp = add_empty_row(grp)

        # --- rebuild group rows: stock (sorted), sold, blank ---
        blanks = grp[grp.isnull().all(axis=1)].copy()
        sold_rows = grp[grp['SOURCE'] == 'Sold'].copy()
        stock_rows = grp[grp['SOURCE'] == 'Stock'].dropna(how='all').copy()

        # split new arrivals
        nas = stock_rows[stock_rows['STATUS'].str.strip() == 'New Arrival']
        if not nas.empty:
            main_na = nas.iloc[[0]].copy()
            other_na = nas.iloc[1:].sort_values('CARAT').copy()
        else:
            main_na = stock_rows.iloc[[0]].copy()
            other_na = stock_rows.iloc[1:0].copy()

        remaining = stock_rows.drop(main_na.index.union(other_na.index))

        cut_val, pol_val, sym_val, fluor_val = main_na.iloc[0][['CUT','POL','SYM','FLUOR']].values
        matching = remaining[
            (remaining['CUT']==cut_val) &
            (remaining['POL']==pol_val) &
            (remaining['SYM']==sym_val) &
            (remaining['FLUOR']==fluor_val)
        ].sort_values(by=['CUT','POL','SYM','FLUOR'], ascending=True, na_position='last').copy()

        other = remaining.drop(matching.index).sort_values(by=['CUT','POL','SYM','FLUOR'], ascending=True, na_position='last').copy()
        stock_sorted = pd.concat([main_na, other_na, matching, other], ignore_index=True)
        grp = pd.concat([stock_sorted, sold_rows, blanks], ignore_index=True)

        grp['Sold Disc'] = grp.apply(calculate_sold_disc, axis=1)
        sold_subset = grp[grp['SOURCE']=='Sold'].drop_duplicates(subset=['DISC','NAME'])
        if len(sold_subset)>5:
            sold_subset = sold_subset.iloc[:5]
        grp = pd.concat([grp[grp['SOURCE']!='Sold'], sold_subset], ignore_index=True)

        sold_discs = grp.loc[grp['SOURCE']=='Sold','Sold Disc'].dropna().astype(str)
        if not sold_discs.empty:
            grp.at[0,'Sold Disc'] = "/".join(sold_discs)
        grp.loc[grp['SOURCE']=='Sold','Sold Disc'] = None
        for col in ['NAME','Sold Date']:
            vals = grp.loc[grp['SOURCE']=='Sold',col].dropna().astype(str)
            if vals.any():
                grp.at[0,col] = "/".join(vals)
            grp.loc[grp['SOURCE']=='Sold',col] = None

        total_stock = (grp['SOURCE']=='Stock').sum()
        total_sold = (grp['SOURCE']=='Sold').sum()
        grp['Total Stock'] = np.where(grp['STATUS'].str.strip()=='New Arrival', total_stock, None)
        grp['Total Sold'] = np.where(grp['STATUS'].str.strip()=='New Arrival', total_sold, None)
        groups.append((new_stone['CARAT'], grp))

    # assemble final results
    groups.sort(key=lambda tpl: (tpl[1].iloc[0].get('SHAPE')!='ROUND', tpl[0]))
    final_results = pd.concat([g for _, g in groups], ignore_index=True)
    mask_3ex = (final_results['CUT'].str.upper()=='EX') & (final_results['POL'].str.upper()=='EX') & (final_results['SYM'].str.upper()=='EX')
    final_results.loc[mask_3ex,'CUT'] = '3EX'
    final_results['CODE'] = final_results.apply(get_code, axis=1)
    final_results = remove_sold_rows(final_results)

    append_sources_to_book2(premal, sachin, sagar)

    now = datetime.now().strftime("%S-%m-%Y")
    output_path = os.path.join(os.path.dirname(input_file), f"Hemil check {now}.xlsx")
    cols = [c for c in final_results.columns if c not in ["NAME", "Sold Date","CERTNO", "RAPPRICE","ONLINE PRICE", "V360", "RAPSIZE","SOURCE"]]
    final_results = final_results[cols+["NAME", "Sold Date","CERTNO", "RAPPRICE","ONLINE PRICE", "V360", "RAPSIZE","SOURCE"]]
    save_data_with_formulas(final_results, wb, output_path)
    reapply_formulas(wb['Matching'], stockid_formulas, 'A')
    open_file(output_path)
    print(f"The new file with matched stones is saved at: {output_path}")

if __name__ == "__main__":
    match_stones()