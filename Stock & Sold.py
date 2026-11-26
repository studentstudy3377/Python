from pathlib import Path
import re
import os
import platform
import subprocess
import numpy as np
import pandas as pd
import tkinter as tk
from openpyxl import Workbook
from datetime import datetime
from tkinter import Tk, filedialog
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

column_mapping = {
    "STOCK ID":       "STOCKID",
    "POLISH":         "POL",
    "SYMMETRY":       "SYM",
    "BLACK INCLUSION":"BLACK",
    "TINGE":          "TINCH",
    "KEY TO SYMBOL":  "KEY TO SYMBOLS",
    "CERTIFIED NO":   "CERTNO",
    "DATE":           "Sold Date",
    "RAPPRICE($)":    "RAPPRICE",
    "REPORT COMMENTS":"REPORT COMM"
}

desired_order = [
    "STOCKID", "V360", "SHAPE", "RAPSIZE", "CERT", "CARAT", "COLOR", "CLARITY",
    "CUT", "POL", "SYM", "FLUOR", "MEASUREMENT","RATIO", "DEPTH", "TABLE", 
    "DISC","GEN","Pricing","BLACK","LUSTER","TINCH","KEY TO SYMBOLS","REPORT COMM",
    "RAPPRICE","ONLINE PRICE","CERTNO","Sold Date","R D","O D","CODE",
    "STATUS","SOURCE"
]
def select_file(title: str) -> str:
    """Open a file dialog to select an Excel file."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        raise FileNotFoundError("No file selected.")
    return file_path

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

def apply_styles_to_excel(ws, df: pd.DataFrame) -> None:
   
    # Define fills
    sold_fill     = PatternFill(start_color="CFEAF7", end_color="CFEAF7", fill_type="solid")
    new_fill      = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    combined_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    od_fill       = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    red_fill      = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    green_fill    = PatternFill(start_color="66FFCC", end_color="66FFCC", fill_type="solid")
    pink_fill     = PatternFill(start_color="ff8ecb", end_color="66FFCC", fill_type="solid")
    red_font      = Font(color="FF5050", bold=True)
    black_font    = Font(color="000000", bold=True)

    # Precompute col indexes for styling
    col_index = {col: idx + 1 for idx, col in enumerate(df.columns)}
    cell_styles = {}

    for idx, row in df.iterrows():
        excel_row = idx + 2

        # Write data
        for col, col_idx in col_index.items():
            val = row[col]
            ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)

        # Extract fields
        status = str(row.get("STATUS", "")).strip().lower()
        is_new = status == "new arrival"
        is_sold = row.get("SOURCE") == "Sold"
        od_val = row.get("O D")

        # Row-based fill logic
        fill = None
        if is_new:
            fill = new_fill
            if pd.notna(od_val) and od_val >= 2:
                fill = od_fill
        if is_sold and is_new:
            fill = combined_fill
        elif is_sold:
            fill = sold_fill

        # Apply row fill
        if fill:
            for col_idx in range(1, len(df.columns) + 1):
                ws.cell(row=excel_row, column=col_idx).fill = fill

        # Precompute specific value-based conditions
        sta = row.get("STATUS")
        clarity = row.get("CLARITY")
        cert = row.get("CERT")
        pol = row.get("POL")
        sym = row.get("SYM")
        cut = row.get("CUT")
        fluor = str(row.get("FLUOR", "")).strip().upper()
        tinch = str(row.get("TINCH", "")).strip().upper()
        shape = str(row.get("SHAPE", "")).strip()
        luster = str(row.get("LUSTER", "")).strip().upper()
        report_comm = str(row.get("REPORT COMM", "") or "")
        key_symbol = str(row.get("KEY TO SYMBOLS", "") or "")
        black = str(row.get("BLACK", "") or "")
        stockid = str(row.get("STOCKID", "") or "")
        depth = row.get("DEPTH")
        table = row.get("TABLE")
        ratio = row.get("RATIO")
       
        # Conditions dictionary
        checks = {
            "STATUS": status == "new arrival",
            "STATUS": sta == "Hold-Not To See",
            "POL": pol == "GD",
            "SYM": sym == "GD",
            "CLARITY": clarity == "I1",
            "CUT": cut == "GD",
            "CERT": cert in {"HRD", "IGI"},
            "TINCH": tinch in {"B-1", "MIX"},
            "FLUOR": fluor in {"FNT", "VST", "MED", "SL", "STG", "VSL"},
            "SHAPE": shape in {'Oval','Pear','Emerald','Marquise','Princess','Cushion Brilliant','CUSHION CUT','Cushion Modified','Radiant','Heart','Square Radiant','SQ EMERALD','Sq. Emerald'},
            "LUSTER": luster in {'LOW',"M-1", "M1", "M-2", "M2"},
            "REPORT COMM": "Natural" in report_comm or "Clarity grade is based on clouds that are not shown." in report_comm,
            "KEY TO SYMBOLS": ("Internal Graining" in key_symbol and clarity in ["IF", "FL", "VVS1", "VVS2"] ),
            "KEY TO SYMBOLS": any(x in key_symbol for x in ["Bruise", "Cavity", "Cavity", "Grain center", "Etch Channel", "Chip", "Indented Natural","Needle- Natural","Manufacturing Remnant","Surface Graining"]),
            "BLACK": "H" in black,
            "STOCKID": any(x in stockid for x in ["V", "T"]),
            "TABLE": pd.notna(depth) and pd.notna(table) and depth < table,
            "DEPTH": pd.notna(depth) and pd.notna(table) and depth < table,
            }
        
        # Ratio
        if shape == "MARQUISE" and pd.notna(ratio) and ratio <= 1.85:
            checks["RATIO"] = True
        elif shape == "PEAR" and pd.notna(ratio) and ratio <= 1.44:
            checks["RATIO"] = True
        elif shape == "OVAL" and pd.notna(ratio) and ratio <= 1.34:
            checks["RATIO"] = True
        elif shape == "EMERALD" and pd.notna(ratio) and ratio <= 1.33:
            checks["RATIO"] = True
        elif shape == "SQ EMERALD" and pd.notna(ratio) and ratio <= 1.33:
            checks["RATIO"] = True
        elif shape == "SQ. EMERALD" and pd.notna(ratio) and ratio <= 1.33:
            checks["RATIO"] = True
        else:
            checks["RATIO"] = False
        
        # Depths 
        if shape == "MARQUISE" and pd.notna(depth) and depth >= 64:
            checks["DEPTH"] = True
        elif shape == "PRINCESS" and pd.notna(depth) and depth >= 75:
            checks["DEPTH"] = True
        elif shape == "OVAL" and pd.notna(depth) and depth >= 65:
            checks["DEPTH"] = True
        elif shape == "PEAR" and pd.notna(depth) and depth >= 65:
            checks["DEPTH"] = True
        elif shape == "EMERALD" and pd.notna(depth) and depth >= 70:
            checks["DEPTH"] = True
        elif shape == "SQ EMERALD" and pd.notna(depth) and depth >= 70:
            checks["DEPTH"] = True
        elif shape == "SQ. EMERALD" and pd.notna(depth) and depth >= 70:
            checks["DEPTH"] = True
        elif shape == "SQ. EMERALD" and pd.notna(depth) and depth >= 70:
            checks["DEPTH"] = True
        elif shape == "CUSHION BRILLIANT" and pd.notna(depth) and depth >= 70:
            checks["DEPTH"] = True
        elif shape == "CUSHION MOFIFIED" and pd.notna(depth) and depth >= 70:
            checks["DEPTH"] = True

        for col, condition in checks.items():
            if not condition or col not in col_index:
                continue

            # --- FIRST: violet condition for depth < table ---
            if col == "DEPTH" and pd.notna(depth) and pd.notna(table) and depth < table:
                fill = red_fill
                font = red_font

            # --- Other special cases ---
            elif col == "STOCKID":
                fill = green_fill
                font = black_font

            elif col == "RATIO":
                fill = pink_fill
                font = black_font

            # --- All other DEPTH conditions (>= thresholds) ---
            elif col == "DEPTH":
                fill = pink_fill
                font = black_font

            # --- Default red cases ---
            else:
                fill = red_fill
                font = red_font if col != "REPORT COMM" else black_font

            cell_styles[(excel_row, col_index[col])] = (fill, font)

    # Final style pass (batch)
    for (r, c), (f_fill, f_font) in cell_styles.items():
        cell = ws.cell(row=r, column=c)
        if f_fill:
            cell.fill = f_fill
        if f_font:
            cell.font = f_font

def apply_styles_excel(ws, results):

    fill_cut_3ex = PatternFill(start_color="33cccc", end_color="33cccc", fill_type="solid") # 3EX fill
    fill_vg = PatternFill(start_color="ff66ff", end_color="ff66ff", fill_type="solid")     # Light green
    fill_ex = PatternFill(start_color="ff9900", end_color="ff9900", fill_type="solid")     # Light green
    fill_fancy = PatternFill(start_color="ffc971",end_color="ffc971",fill_type="solid")    # Fancy Color
    fill_gd = PatternFill(start_color="E6D5F3",end_color="E6D5F3",fill_type="solid")      # GD Color 
    # Header index mapping
    header = [cell.value for cell in ws[1]]
    col_indexes = {col: header.index(col) + 1 for col in header}
    cell_styles = {} 
    for i in range(1, results.shape[0] + 1):
        excel_row_idx = i - 1
        row = results.iloc[excel_row_idx]

        try:
            if row["SOURCE"] != "Stock":
                continue
            if str(row.get("STATUS", "")).strip().lower() == "new arrival":
                continue
        except Exception:
            continue

        # Shape-based logic
        shape_val = str(row.get("SHAPE", "")).strip().upper()
        is_fancy = shape_val != "ROUND"
        if is_fancy:
            
            column_highlight = ["POL", "SYM"]
        else:
            column_highlight = ["CUT", "POL", "SYM"]

        # CUT == 3EX
        cut_val = str(row.get("CUT", "")).strip().upper()
        if cut_val == "3EX":
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_cut_3ex
        elif cut_val == "EX"  :
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_ex
        elif cut_val == "VG":
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_vg
  
        # POL == EX
        pol_val = str(row.get("POL", "")).strip().upper()
        if pol_val in ["EX","VG"]  and is_fancy:
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_fancy
        
        gd_value = str(row.get("POL","SYM")).strip().upper()
        if gd_value == "GD" and is_fancy:
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_gd
        
        cut_pol_sym = [str(row.get(col, "")).strip().upper() for col in ["CUT", "POL", "SYM"]]
        if "GD" in cut_pol_sym:
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_gd

    # Apply all styles at once
    for (row_idx, col_idx), fill in cell_styles.items():
        ws.cell(row=row_idx, column=col_idx).fill = fill

def add_empty_row(df: pd.DataFrame) -> pd.DataFrame:
    """Add an empty row to the DataFrame, dtype-safe."""
    empty = pd.DataFrame([{col: np.nan for col in df.columns}])
    return pd.concat([df, empty], ignore_index=True)

def open_file(file_path: str):
    """Open a file using the default application."""
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        else:
            subprocess.call(["xdg-open", file_path])
    except Exception as e:
        print(f"Error opening file: {e}")

def find_carat_range(carat: float, shape: str,fluor=None, luster=None, tinch=None, black=None):

    # normalize
    shape_u = str(shape).strip().upper() if shape is not None else ""
    luster_u = str(luster).strip().upper() if luster is not None else ""
    tinch_u = str(tinch).strip().upper().replace(" ", "")  # handles "B-1","B1", etc.
    black_u  = str(black).strip().upper() if black is not None else ""
    fluor_u  = str(fluor).strip().upper() if fluor is not None else ""

    # your uniform (“special”) ranges
    uniform = [
        (0.23, 0.29), (0.30, 0.39), (0.40, 0.49), (0.50, 0.59), (0.60, 0.69),
        (0.70, 0.79), (0.80, 0.89), (0.90, 0.99), (1.00, 1.19), (1.20, 1.29),
        (1.30, 1.39), (1.40, 1.49), (1.50, 1.59), (1.60, 1.69), (1.70, 1.79),
        (1.80, 1.89), (1.90, 1.99), (2.00, 2.19), (2.20, 2.49), (2.50, 2.99),
        (3.00, 3.49), (3.50, 3.99), (4.00, 4.49), (4.50, 4.99), (5.00, 5.49),
        (5.50, 5.99)
    ]

    # your finer “non-special” (ROUND) ranges
    fine = [
        (0.23, 0.29), (0.30, 0.34), (0.35, 0.39), (0.40, 0.45), (0.46, 0.49),
        (0.50, 0.59), (0.60, 0.69), (0.70, 0.79), (0.80, 0.89), (0.90, 0.99),
        (1.00, 1.19), (1.20, 1.29), (1.30, 1.39), (1.40, 1.49), (1.50, 1.59),
        (1.60, 1.69), (1.70, 1.79), (1.80, 1.89), (1.90, 1.99), (2.00, 2.19),
        (2.20, 2.49), (2.50, 2.99), (3.00, 3.49), (3.50, 3.99), (4.00, 4.49),
        (4.50, 4.99), (5.00, 5.49), (5.50, 5.99)
    ]

    # 1) non-round == always uniform
    if shape_u and shape_u != "ROUND":
        ranges = uniform

    else:
        # 2) any one of these four flags → uniform   
        special_checks = [
            fluor_u in {"MED","VST","STG"},
            luster_u in {"M-1", "M1", "M-2", "M2"},
            tinch_u in {"B-1", "B1", "B-2", "B2"},
            "H" in black_u,
        ]
        if any(special_checks):
            ranges = uniform
        else:
            ranges = fine

    # now pick your bucket
    for start, end in ranges:
        if start <= carat <= end:
            return (start, end)
    return None

def get_allowed_cut(cut):
    
    cut_hierarchy = {
        'EX': ['EX'],
        'VG': ['EX', 'VG'],
        'GD': ['EX', 'VG', 'GD'],
        'FR': ['EX', 'VG', 'GD', 'FR']
    }

    if pd.isna(cut):
        # If CUT is NaN (not a string), return a list with pd.NA
        return [pd.NA]
    else:
        # If CUT doesn't appear in cut_hierarchy, default to the broadest set
        return cut_hierarchy.get(cut, ['EX', 'VG', 'GD', 'FR'])

def get_allowed_cert(cert):
    """
    Return a list of allowed certification types based on the given certificate.
    """
    cert_hierarchy = {
        'GIA': ['GIA'],
        'HRD': ['HRD'],
        'IGI': ['IGI'],
        'IIDGR': ['IIDGR']
    }
    if pd.isna(cert):
        return []
    else:
        return cert_hierarchy.get(cert, ['GIA', 'HRD', 'IGI','IIDGR'])
    
def get_allowed_fluor(fluor):
    fluor_hierarchy = {
        'NONE': ['NONE'],
        'FNT': ['NONE', 'FNT'],
        'MED': ['NONE', 'FNT','MED'],
        'SL' : ['NONE', 'FNT', 'MED', 'SL'],
        'STG': ['NONE', 'FNT', 'MED', 'SL', 'STG'],
        'VST': ['NONE', 'FNT', 'VST', 'MED', 'SL', 'STG', 'VST'],
        'VSL': ['NONE', 'FNT', 'VST', 'MED', 'SL', 'STG', 'VST','VSL']
    }
    if pd.isna(fluor):
        return [pd.NA]
    else:
        return fluor_hierarchy.get(fluor, ['NONE', 'FNT', 'VSL', 'MED', 'SL', 'STG', 'VST','VSL'])  # Default to all if not found

def replace_low(df: pd.DataFrame) -> pd.DataFrame:
    """Replace values in the DataFrame for low luster stones."""
    if "Luster" in df.columns:
        df.loc[df["Luster"] == "LOW", "LUSTER"] = "LOW"
    return df

def get_code(row):
    # Only compute CODE for rows with STATUS "New Arrival"
    if not (isinstance(row.get("STATUS"), str) and row["STATUS"].strip().lower() == "new arrival"):
        return ""
    
    # If the CLARITY value is blank (or missing), return a blank CODE.
    if pd.isnull(row["CLARITY"]) or str(row["CLARITY"]).strip() == "":
        return ""
    
    # Standardize values and extract numeric value for CARAT.
    clarity = str(row["CLARITY"]).upper()
    shape = str(row["SHAPE"]).upper() if pd.notnull(row["SHAPE"]) else ""
    try:
        carat = float(row["CARAT"])
    except (ValueError, TypeError):
        return ""
    
    # Define the conditions.
    flag_PZ = (shape != "ROUND" or carat >= 1.00 or clarity == "I1")
    flag_SP = (clarity in ("IF", "VVS1", "VVS2", "VS1"))
    flag_SK = (carat >= 1.00 or (clarity in ("VS2", "SI1", "SI2", "I1") and 0.23 <= carat <= 0.99))
    
    # Combine the codes.
    codes = []
    if flag_PZ:
        codes.append("PZ")
    if flag_SK:
        codes.append("SK")
    if flag_SP:
        codes.append("SP")
    
    return "/".join(codes) if codes else ""

def calculate_disc(row):
    if row.get("SOURCE") == "Stock":
        return row.get("DISC", None)
    rap = row.get("RAPPRICE", 0) or 0
    final = row.get("ONLINE PRICE", 0) or 0
    return round(((final - rap) / rap) * 100, 2) if rap else 0

def assign_disc_from_exact_match(group, stone):
   
    if stone["SHAPE"].strip().upper() != "ROUND":
        return group  # Skip round stones

    match_col = ["SHAPE", "CERT", "COLOR", "CLARITY","CUT","POL","SYM","FLUOR", "LUSTER", "TINCH", "BLACK"]

    # Clean status safely (won't error on NaN)
    status = group["STATUS"]
    status_clean = status.where(status.notna(), "").str.strip().str.lower()

    # Filter the stock rows based on the conditions
    stock_rows = group[
        (group["SOURCE"] == "Stock")
        & (~group["RAPSIZE"].isin(["1.00UP", "1.50UP", "2.00UP", "3.00UP", "4.00UP", "5.00UP"]))
        & status.notna() 
        & (status_clean != "new arrival")  
        & (status_clean != "hold-not to see")  
    ].copy()

    disc_map = (stock_rows.dropna(subset=["DISC"]).groupby(match_col)["DISC"].max().reset_index())
    mask_new = group["STATUS"].str.strip().str.lower() == "new arrival"
    merged = group.loc[mask_new, match_col].merge(disc_map, on=match_col, how="left")
    group.loc[mask_new, "DISC"] = merged["DISC"].fillna(0).to_numpy()

    return group

def assign_disc_from_exact_match_non(group, stone):
   
    if stone["SHAPE"].strip().upper() == "ROUND":
        return group  # Skip non-round stones

    match_col = ["SHAPE", "CERT", "COLOR", "CLARITY", "FLUOR", "LUSTER", "TINCH", "BLACK"]

    # Clean status safely (won't error on NaN)
    status = group["STATUS"]
    status_clean = status.where(status.notna(), "").str.strip().str.lower()

    # Filter the stock rows based on the conditions
    stock_rows = group[
        (group["SOURCE"] == "Stock")
        & (~group["RAPSIZE"].isin(["90UP","1.00UP", "1.50UP", "2.00UP", "3.00UP", "4.00UP", "5.00UP"]))
        & status.notna()
        & (~group["POL"].isin(["GD","FR"]))
        & (~group["SYM"].isin(["GD","FR"])) 
        & (status_clean != "new arrival")  
        & (status_clean != "hold-not to see")  
    ].copy()

    disc_map = (stock_rows.dropna(subset=["DISC"]).groupby(match_col)["DISC"].max())
    mask_new = group["STATUS"].str.strip().str.lower() == "new arrival"
    merged = group.loc[mask_new, match_col].merge(disc_map, on=match_col, how="left")
    group.loc[mask_new, "DISC"] = merged["DISC"].fillna(0).to_numpy()

    return group

def append_with_date_to_excel(df: pd.DataFrame,file_path: str,sheet_name: str = "NewArrivals") -> None:
    # Prepare incoming data
    df = df.copy()
    today_str = datetime.now().strftime("%d-%m-%Y")
    df["date"] = today_str

    # Ensure dir exists
    p = Path(file_path)
    p.parent.mkdir(parents=True, exist_ok=True)

    # Load existing (if available)
    if p.exists():
        try:
            with pd.ExcelFile(p, engine="openpyxl") as xls:
                if sheet_name in xls.sheet_names:
                    existing_df = pd.read_excel(xls, sheet_name=sheet_name, dtype={"STOCKID": object})
                else:
                    existing_df = pd.DataFrame(columns=["STOCKID", "gen", "date"])
        except Exception:
            existing_df = pd.DataFrame(columns=["STOCKID", "gen", "date"])
    else:
        existing_df = pd.DataFrame(columns=["STOCKID", "gen", "date"])

    # Align columns
    all_cols = ["STOCKID", "gen", "date"]
    existing_df = existing_df.reindex(columns=all_cols, fill_value=pd.NA)
    df = df.reindex(columns=all_cols)

    # Combine & (optionally) dedupe STOCKID for the same date
    combined = pd.concat([existing_df, df], ignore_index=True)

    # Write back (keeps all rows by writing the combined frame)
    with pd.ExcelWriter(p, engine="openpyxl", mode="w") as writer:
        combined.to_excel(writer, sheet_name=sheet_name, index=False)

def match_stones():
    """Main function to match stones based on several criteria."""

    # 1) Select files
    input_file = select_file("Select Stock File")
    sold_file = select_file("Select the Sold file")
    days_file = select_file("Select Days File")
    revision_file = r"G:\Certificate\PREMAL\REVISE DAYS TRACKING.xlsx"
    hold_not_to_see = r"G:\Certificate\PROMO\PROMO 05-11.xls"

    # 2) Load stock data and mark as 'Stock'
    data1 = pd.read_excel(input_file)
    data1.loc[data1["STATUS"].str.lower() == "new arrival", "DISC"] = 0
    data1.loc[data1["STATUS"].str.lower() == "hold-not to see", "STATUS"] = "PROMOTION"
    data1["SOURCE"] = "Stock"

    # 3) Load sold data, map columns, and mark as 'Sold'
    sold_data = pd.read_excel(sold_file, skiprows=2)
    sold_data.rename(columns=column_mapping, inplace=True)
    sold_data["SOURCE"] = "Sold"

    # 4) Combine stock and sold data
    data1 = pd.concat([data1, sold_data], ignore_index=True)

    # 5) Merge in revision & days data
    revision = pd.read_excel(revision_file, dtype=str)
    hold_not = pd.read_excel(hold_not_to_see, dtype=str)
    days = pd.read_excel(days_file, dtype=str)
    for df in (revision, days):
        df["Sr. No"] = df["Sr. No"].str.strip()

    merged = (
        data1
        .merge( revision[["Sr. No", "DAYS"]],left_on="STOCKID",right_on="Sr. No", how="left",)
        .rename(columns={"DAYS": "R D"})   
        .merge(days[["Sr. No", "Stock Days", "Luster"]],left_on="STOCKID", right_on="Sr. No",how="left",)
        .rename(columns={"Stock Days": "O D"})    
        .merge(hold_not[["STOCKID","PDISC"]],left_on="STOCKID",right_on="STOCKID",how="left",)
        )
   
    merged = replace_low(merged)
    
    merged["PDISC"].fillna(merged['DISC'], inplace=True)
    merged['DISC'] = merged['PDISC']
    merged["DISC"] = pd.to_numeric(merged["DISC"], errors='coerce')

    merged.drop(columns=["Sr. No_x", "Sr. No_y", "Sr. No", "Luster"], errors="ignore",inplace=True, )
      
    merged[["R D", "O D", "DEPTH", "TABLE"]] = ( merged[["R D", "O D", "DEPTH", "TABLE"]].apply(pd.to_numeric, errors="coerce"))
       
    merged[["R D", "O D"]] = (merged[["R D", "O D"]] .fillna(0) .apply(pd.to_numeric, errors="coerce"))
        
    # 6) Prepare for matching loop
    seen_new = set()
    new_arrivals = merged[merged["STATUS"].str.lower() == "new arrival"]

    groups = []  # collect groups instead of appending to results directly

    for _, stone in new_arrivals.iterrows():
        sid = stone["STOCKID"]
        if sid in seen_new:
            continue
        seen_new.add(sid)

        # Determine carat range
        cr = find_carat_range(
            stone["CARAT"],
            shape=stone["SHAPE"],
            fluor=stone.get("FLUOR"),
            luster=stone.get("LUSTER"),
            tinch=stone.get("TINCH"),
            black=stone.get("BLACK"),
        )
        if not cr:
            grp = stone.to_frame().T
            groups.append(pd.concat([grp, add_empty_row(grp)], ignore_index=True))
            continue

        # Build filter conditions
        base_filters = [
            merged["SHAPE"] == stone["SHAPE"],
            merged["CARAT"].between(*cr),
            merged["CERT"] == stone["CERT"],
            merged["COLOR"] == stone["COLOR"],
            merged["CLARITY"] == stone["CLARITY"],
            # merged["CUT"] == stone["CUT"],
            # merged["POL"] == stone["POL"],
            merged["LUSTER"] == stone["LUSTER"],
            merged["TINCH"] == stone["TINCH"],
        ]

        seen_filter = ((~merged["STOCKID"].isin(seen_new)) | ((merged["SOURCE"] == "Sold") & (merged["STOCKID"] == sid)))
   
        base_filters.append(seen_filter)

        if stone["STATUS"].strip().lower() == "new arrival":
           
            base_filters.append(merged["SOURCE"] != "Sold")

        # Build ROUND vs non-ROUND filters
        base_mask = np.all(base_filters, axis=0)

        if stone["SHAPE"] == "ROUND":
            fluor_mask = merged["FLUOR"].isin(get_allowed_fluor(stone["FLUOR"]))

            # 1) For New Arrival rows: require exact CUT/POL/SYM
            na_mask = (
                (merged["STATUS"].str.lower() == "new arrival")
                & (merged["CUT"] == stone["CUT"])
                & (merged["POL"] == stone["POL"])
                & (merged["SYM"] == stone["SYM"])
            )

            # 2) For non–New Arrival rows: allow a range for CUT/POL/SYM
            non_na_mask = (
                (merged["STATUS"].str.lower() != "new arrival")
                & merged["CUT"].isin(get_allowed_cut(stone["CUT"])))
                # & merged["SYM"].isin(get_allowed_cut(stone["SYM"])))
            
            matches_df = merged[(base_mask) & (fluor_mask) & (na_mask | non_na_mask)].copy()
        else:
            matches_df = merged[base_mask & merged["FLUOR"].isin(get_allowed_fluor(stone["FLUOR"]))].copy()

        if all(stone[attr] == "EX" for attr in ("CUT", "POL", "SYM")):
            matches_df = matches_df[
                (matches_df["CUT"] == "EX")
                & (matches_df["POL"] == "EX")
                & (matches_df["SYM"] == "EX")
            ].copy()

        seen_new.update(matches_df["STOCKID"])
        new_arrival_df = stone.to_frame().T

        # Build initial group
        group = new_arrival_df.copy()
        extra_na = matches_df[matches_df["STATUS"].str.lower() == "new arrival"]
        if not extra_na.empty:
            extra_na = extra_na.dropna(axis=1, how="all")
            group = pd.concat([group, extra_na], ignore_index=True)

        # Exact stock matches (3EX priority)
        mask_stock = matches_df["SOURCE"] == "Stock"
        mask_3ex = (
            (matches_df["CUT"] == stone["CUT"]) &
            (matches_df["POL"] == stone["POL"]) &
            (matches_df["SYM"] == stone["SYM"])
        )
        exact = matches_df[mask_stock & mask_3ex]
        if not exact.empty:
            group = pd.concat([group, exact], ignore_index=True)

        # Other partial stock matches
        partial = matches_df[mask_stock & ~mask_3ex].copy()
        if not partial.empty:
            partial = partial.sort_values(
                by=["CUT", "POL", "SYM"],
                ascending=True,
                na_position="first",
            )
            common_cols = group.columns.intersection(partial.columns)
            group = pd.concat([group[common_cols], partial[common_cols]], ignore_index=True)

        # Sort remaining stock rows by hit-priority
        stock_rows = group[group["SOURCE"] == "Stock"].copy()
        other_rows = group[group["SOURCE"] != "Stock"].copy()
        new_stock = stock_rows[stock_rows["STATUS"].str.lower() == "new arrival"]
        sortable = stock_rows.drop(new_stock.index)
        if not sortable.empty:
            for hit in ("FLUOR", "CUT", "POL", "SYM", "BLACK"):
                sortable[f"{hit.lower()}_hit"] = (
                    (sortable[hit] == stone[hit]).astype(int)
                )
            sort_keys = ["fluor_hit","cut_hit","pol_hit","sym_hit", "black_hit",] + [c for c in ("CUT", "POL", "SYM", "BLACK") if c in sortable]
              
            ascending = [False] * 5 + [True] * (len(sort_keys) - 5)
            sortable = sortable.sort_values(
                by=sort_keys,
                ascending=ascending,
                na_position="last",
            )
            sortable = sortable.drop(
                columns=[k for k in sortable.columns if k.endswith("_hit")],
                errors="ignore",
            )
        stock_rows = pd.concat([new_stock, sortable], ignore_index=True)
        group = pd.concat([stock_rows, other_rows], ignore_index=True)

        # Append sold rows for this stone
        sold_rows = merged[
            (merged["SOURCE"] == "Sold") & (merged["STOCKID"] == sid)
        ].copy()
        if not sold_rows.empty:
            common_cols = group.columns.intersection(sold_rows.columns)
            group = pd.concat([group, sold_rows[common_cols]], ignore_index=True)

        # Include last few sold matches
        if stone["SHAPE"] == "ROUND":
            match_cols = ["SHAPE","CARAT","CERT","COLOR","CLARITY","CUT","POL","SYM","FLUOR","LUSTER","TINCH","BLACK",]    
        else:
            match_cols = ["SHAPE","CARAT","CERT","COLOR","CLARITY","POL","SYM","FLUOR","LUSTER","TINCH","BLACK",]
            
        sold_matches = sold_data.copy()
        for col in match_cols:
            if col != "CARAT":
                sold_matches = sold_matches[sold_matches[col] == stone[col]]
        sold_matches = sold_matches[sold_matches["CARAT"].between(*cr)]

        # FIX: sort the candidate sold_matches by Sold Date (not sold_data)
        if "Sold Date" in sold_matches.columns:
            sold_matches["Sold Date"] = pd.to_datetime(
                sold_matches["Sold Date"], errors="coerce"
            )
            sold_matches = sold_matches.sort_values(
                "Sold Date", ascending=False
            ).head(5)
            sold_matches["Sold Date"] = sold_matches["Sold Date"].dt.strftime(
                "%d-%b-%Y"
            )

        if not sold_matches.empty:
            common_cols = group.columns.intersection(sold_matches.columns)
            group = pd.concat([group, sold_matches[common_cols]], ignore_index=True)

        group = group.drop_duplicates(
            subset=["STOCKID", "SOURCE"], keep="first"
        ).reset_index(drop=True)
        group = add_empty_row(group)
        group = assign_disc_from_exact_match(group, stone)
        group = assign_disc_from_exact_match_non(group, stone)

        groups.append(group)

    # Separate groups into ROUND and non-ROUND
    round_groups = []
    non_round_groups = []
    for g in groups:
        if not g.empty and g.iloc[0]["SHAPE"] == "ROUND":
            round_groups.append(g)
        else:
            non_round_groups.append(g)

    # Sort ONLY ROUND groups by CARAT of first row (ascending)
    round_groups.sort(
        key=lambda g: ( float(pd.to_numeric(g.iloc[0]["CARAT"], errors="coerce"))if not g.empty else -float("inf") ),reverse=False, )
    
    # Combine back: sorted ROUND first, then non-round in original order
    results = pd.concat(round_groups + non_round_groups, ignore_index=True)

    results["CARAT"] = results["CARAT"].apply(lambda x: f"{float(x):.2f}")
    results["CARAT"] = results["CARAT"].replace("nan","")

    mask_final = results["SOURCE"] != "Stock"
    results.loc[mask_final, "DISC"] = results.loc[mask_final].apply(calculate_disc, axis=1)
    
    na_mask = results["STATUS"].str.strip().str.lower() == "new arrival"

    # Ensure DISC is numeric, treat NaNs as 0, then filter non-zero
    
    nz_mask = results["DISC"] != 0

    export_df = (
        results.loc[na_mask & nz_mask, ["STOCKID", "DISC"]]
        .dropna(subset=["STOCKID"])            # safety
        .drop_duplicates(subset=["STOCKID"])   # one row per stock id
        .sort_values(by="STOCKID")
        .reset_index(drop=True)
        .rename(columns={"DISC": "gen"})
    )

    # Finalize and save
    results["CODE"] = results.apply(get_code, axis=1)

    # For rows where STATUS is "new arrival", set DISC equal to GEN in the original DataFrame
    results.loc[results["STATUS"].str.lower() == "new arrival", 'GEN'] = results['DISC']

    mask_3ex_final = (results[["CUT", "POL", "SYM"]] == "EX").all(axis=1)
    results.loc[mask_3ex_final, "CUT"] = "3EX"

    for col in desired_order:
        if col not in results.columns:
            results[col] = None

    results = results[desired_order]

    wb = Workbook()
    ws = wb.active
    ws.title = "Matched Stones"

    for row in dataframe_to_rows(results, index=False, header=True):
        ws.append(row)

    apply_styles_to_excel(ws, results)
    apply_styles_excel(ws, results)
    auto_gent_file = r"C:\Users\cert07\Downloads\NEW26 11\Book1.xlsx"
    append_with_date_to_excel(export_df, auto_gent_file, sheet_name="New Arrivals")
    output_name = f"Stock {datetime.now():%d-%m-%Y}.xlsx"
    outp = os.path.join(os.path.dirname(input_file), output_name)
    wb.save(outp)
    open_file(outp)
    print(f"Saved styled matched stones file to: {outp}")

if __name__ == "__main__":
    match_stones()  