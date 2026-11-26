import os
import re
import platform
import subprocess
from datetime import datetime
import numpy as np
import pandas as pd
import tkinter as tk
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from tkinter import filedialog


# =========================
# Config / Column Mappings
# =========================

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
    "FINAL CUSTOMER": "NAME",
    "REPORT COMMENTS":"REPORT COMM"
}

desired_order = [
    "STOCKID","SHAPE", "CERT", "CARAT", "COLOR", "CLARITY","CUT", "POL", "SYM", "FLUOR", 
    "MEASUREMENT", "DEPTH", "TABLE", "DISC","BLACK","LUSTER","TINCH",
    "KEY TO SYMBOLS","REPORT COMM","R D","O D","RATIO","List Disc","Gen",
    "Premal", "Sachin", "Sagar","Small","NAME","Sold Disc","Sold Date","CODE",
    "Total Stock","Total Sold","STATUS", "RAPPRICE","RAPSIZE","V360","ONLINE PRICE","CERTNO"
]


# =========================
# Utilities / Helpers
# =========================

def select_file(title: str) -> str:
    """Open a file dialog to select an Excel file. Returns None if cancelled."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx")]
    )
    return file_path or None


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
    """Row and cell-level conditional styling for the 'Matched Stones' sheet."""
    # Define fills
    sold_fill     = PatternFill(start_color="CFEAF7", end_color="CFEAF7", fill_type="solid")
    new_fill      = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    combined_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    red_fill      = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    green_fill    = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")

    red_font      = Font(color="FF5050", bold=True)
    black_font    = Font(color="000000", bold=True)

    # Helper function for one-off cell fill
    def _fill_cell(_ws, r, c, color_hex):
        fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        _ws.cell(row=r, column=c).fill = fill
        _ws.cell(row=r, column=c).font = black_font

    # Map columns
    col_index = {col: idx + 1 for idx, col in enumerate(df.columns)}
    cell_styles = {}

    for idx, row in df.iterrows():
        excel_row = idx + 2  # header = row 1

        # Write data
        for col, col_idx in col_index.items():
            val = row[col]
            ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)

        # Extract values
        status  = str(row.get("STATUS", "")).strip().lower()
        is_new  = status == "new arrival"
        is_sold = row.get("SOURCE") == "Sold"

        # Row-level fills
        row_fill = None
        if is_new:
            row_fill = new_fill
        if is_new and is_sold:
            row_fill = combined_fill
        elif is_sold:
            row_fill = sold_fill

        if row_fill:
            for c in range(1, len(df.columns) + 1):
                ws.cell(row=excel_row, column=c).fill = row_fill

        # Value-based checks
        clarity     = row.get("CLARITY")
        cert        = row.get("CERT")
        pol         = row.get("POL")
        sym         = row.get("SYM")
        cut         = row.get("CUT")
        fluor       = str(row.get("FLUOR", "")).strip().upper()
        tinch       = str(row.get("TINCH", "")).strip().upper()
        shape       = str(row.get("SHAPE", "")).strip()
        luster      = str(row.get("LUSTER", "")).strip().upper()
        report_comm = str(row.get("REPORT COMM", "") or "")
        key_symbol  = str(row.get("KEY TO SYMBOLS", "") or "")
        black       = str(row.get("BLACK", "") or "")
        stockid     = str(row.get("STOCKID", "") or "")
        depth       = row.get("DEPTH")
        table       = row.get("TABLE")

        checks = {
            "STATUS": is_new,
            "POL": pol == "GD",
            "SYM": sym == "GD",
            "CLARITY": clarity == "I1",
            "CUT": cut == "GD",
            "CERT": cert in {"HRD", "IGI"},
            "TINCH": tinch in {"B-1", "MIX"},
            "FLUOR": fluor in {"FNT", "VST", "MED", "SL", "STG", "VSL"},
            "SHAPE": shape in {
                'Oval','Pear','Emerald','Marquise','Princess',
                'Cushion Brilliant','CUSHION CUT','Cushion Modified',
                'Radiant','Heart','Square Radiant','SQ EMERALD','Sq. Emerald'
            },
            "LUSTER": luster in {'LOW',"M-1", "M1", "M-2", "M2"},
            "REPORT COMM": (
                "Natural" in report_comm or
                "Clarity grade is based on clouds that are not shown." in report_comm
            ),
            "KEY TO SYMBOLS": (
                ("Internal Graining" in key_symbol and clarity in ["IF","FL","VVS1","VVS2"])
                or any(k in key_symbol for k in [
                    "Bruise","Cavity","Grain center","Etch Channel","Chip","Indented Natural"
                ])
            ),
            "BLACK": "H" in black,
            "STOCKID": any(x in stockid for x in ["V", "T"]),
            "DEPTH": pd.notna(depth) and pd.notna(table) and depth < table,
            "TABLE": pd.notna(depth) and pd.notna(table) and depth < table,
        }

        for col, condition in checks.items():
            if condition and col in col_index:
                style_fill = red_fill
                style_font = red_font if col not in {"STOCKID", "REPORT COMM"} else black_font
                cell_styles[(excel_row, col_index[col])] = (style_fill, style_font)

        # Extra: MEASUREMENT threshold when FLUOR is NONE/FNT
        if "MEASUREMENT" in col_index and "RAPSIZE" in df.columns:
            fluor_val = str(row.get("FLUOR", "")).strip().upper()
            rapsize_val = str(row.get("RAPSIZE", "")).strip().upper()
            measurement_val = row.get("MEASUREMENT")

            if isinstance(measurement_val, str):
                nums = re.findall(r"[\d\.]+", measurement_val)
                measurement_val = max(map(float, nums)) if nums else None
                measurement_val = pd.to_numeric(measurement_val, errors="coerce")

            if pd.notna(measurement_val) and fluor_val in ["NONE", "FNT"]:
                threshold_map = {
                    "30UP": 4.50,
                    "40UP": 5.00,
                    "50UP": 5.25,
                    "60UP": 5.50,
                    "70UP": 5.85,
                    "80UP": 6.00,
                    "90UP": 6.15,
                    "1.00UP": 6.30
                }
                if rapsize_val in threshold_map and measurement_val > threshold_map[rapsize_val]:
                    _fill_cell(ws, excel_row, col_index["MEASUREMENT"], "33ccff")

    # Handle "Small" column green highlight
    if "Small" in col_index:
        col_idx = col_index["Small"]
        small_col = df["Small"]
        mask = small_col.notna() & (small_col.astype(str).str.strip() != "")
        for idx in df[mask].index:
            excel_row = idx + 2
            ws.cell(row=excel_row, column=col_idx).fill = green_fill
            ws.cell(row=excel_row, column=col_idx).font = black_font

    # Apply stored styles
    for (r, c), (f_fill, f_font) in cell_styles.items():
        cell = ws.cell(row=r, column=c)
        cell.fill = f_fill
        cell.font = f_font


def add_empty_row(df: pd.DataFrame) -> pd.DataFrame:
    """Add an empty row to the DataFrame, dtype-safe."""
    empty = pd.DataFrame([{col: np.nan for col in df.columns}])
    return pd.concat([df, empty], ignore_index=True)


def open_file(file_path: str):
    """Open a file using the default application."""
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)  # type: ignore[attr-defined]
        else:
            subprocess.call(["xdg-open", file_path])
    except Exception as e:
        print(f"Error opening file: {e}")


def find_carat_range(carat: float, shape: str, fluor=None, luster=None, tinch=None, black=None):
    """Return (start, end) carat range based on shape/flags."""
    # normalize
    shape_u = str(shape).strip().upper() if shape is not None else ""
    luster_u = str(luster).strip().upper() if luster is not None else ""
    tinch_u = str(tinch).strip().upper().replace(" ", "")  # handles "B-1","B1", etc.
    black_u  = str(black).strip().upper() if black is not None else ""
    fluor_u  = str(fluor).strip().upper() if fluor is not None else ""

    # uniform (“special”) ranges
    uniform = [
        (0.23, 0.29), (0.30, 0.39), (0.40, 0.49), (0.50, 0.59), (0.60, 0.69),
        (0.70, 0.79), (0.80, 0.89), (0.90, 0.99), (1.00, 1.19), (1.20, 1.29),
        (1.30, 1.39), (1.40, 1.49), (1.50, 1.59), (1.60, 1.69), (1.70, 1.79),
        (1.80, 1.89), (1.90, 1.99), (2.00, 2.19), (2.20, 2.49), (2.50, 2.99),
        (3.00, 3.49), (3.50, 3.99), (4.00, 4.49), (4.50, 4.99), (5.00, 5.49),
        (5.50, 5.99)
    ]

    # finer (“non-special” ROUND) ranges
    fine = [
        (0.23, 0.29), (0.30, 0.36), (0.37, 0.39), (0.40, 0.45), (0.46, 0.49),
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
        # 2) any one of flags → uniform
        special_checks = [
            fluor_u in {"MED","VST","STG"},
            luster_u in {"M-1", "M1", "M-2", "M2"},
            tinch_u in {"B-1", "B1", "B-2", "B2"},
            "H" in black_u,
        ]
        ranges = uniform if any(special_checks) else fine

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
        return [pd.NA]
    else:
        return cut_hierarchy.get(cut, ['EX', 'VG', 'GD', 'FR'])


def get_allowed_cert(cert):
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
        return fluor_hierarchy.get(fluor, ['NONE', 'FNT', 'VSL', 'MED', 'SL', 'STG', 'VST','VSL'])


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


def append_sources_to_book2(premal, sachin, sagar, gen_df,
                            book2_path=r"G:\Certificate\PREMAL\NOT TO DELETE\Disc.xlsx"):
    """
    Appends Premal, Sachin, Sagar and GEN data into 'Combined' in Disc.xlsx.
    Columns: STOCKID, Disc, Code, Date  (Date = DD-MM-YYYY)
    """
    today_str = datetime.now().strftime('%d-%m-%Y')

    # Standardize and tag sources
    premal_df = premal[['stockid', 'pz']].copy() if 'pz' in premal.columns else pd.DataFrame(columns=['stockid','pz'])
    premal_df.columns = ['STOCKID', 'Disc']
    premal_df['Code'] = 'PZ'

    sachin_df = sachin[['stockid', 'sk']].copy() if 'sk' in sachin.columns else pd.DataFrame(columns=['stockid','sk'])
    sachin_df.columns = ['STOCKID', 'Disc']
    sachin_df['Code'] = 'SK'

    sagar_df = sagar[['stockid', 'sp']].copy() if 'sp' in sagar.columns else pd.DataFrame(columns=['stockid','sp'])
    sagar_df.columns = ['STOCKID', 'Disc']
    sagar_df['Code'] = 'SP'

    gen_part = gen_df[['STOCKID', 'gen']].copy() if 'gen' in gen_df.columns else pd.DataFrame(columns=['STOCKID','gen'])
    gen_part.columns = ['STOCKID', 'Disc']
    gen_part['Code'] = 'GEN'

    combined_df = pd.concat([premal_df, sachin_df, sagar_df, gen_part], ignore_index=True)
    combined_df.dropna(subset=['STOCKID', 'Disc'], inplace=True)
    combined_df['Date'] = today_str

    # Load or create workbook
    if os.path.exists(book2_path):
        wb = load_workbook(book2_path)
        ws = wb.active
        if ws.max_row == 1 and all(cell.value is None for cell in ws[1]):
            ws.append(['STOCKID', 'Disc', 'Code', 'Date'])
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = 'Combined'
        ws.append(['STOCKID', 'Disc', 'Code', 'Date'])

    # Append rows
    for _, row in combined_df.iterrows():
        ws.append([row['STOCKID'], row['Disc'], row['Code'], row['Date']])

    wb.save(book2_path)

def match_stones():
    """Main function to match stones based on several criteria."""

    # 1) Select files
    input_file = select_file("Select the Stock file")
    sold_file = select_file("Select the Sold file")
    days_file = select_file("Select the Days file")
    premal_file = select_file("Select the Premal file")
    sachin_file = select_file("Select the Sachin file")
    sagar_file = select_file("Select the Sagar file")
    gen_file_path = select_file("Select the Auto Generate file")
    revision_file = r"G:\Certificate\PREMAL\REVISE DAYS TRACKING.xlsx"

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

    # 2) Load stock and mark
    data1 = pd.read_excel(input_file)
    data1["SOURCE"] = "Stock"

    # 3) Load sold, remap columns, mark
    sold_data = pd.read_excel(sold_file, skiprows=2)
    sold_data.rename(columns=column_mapping, inplace=True)
    sold_data["SOURCE"] = "Sold"

    # 4) Combine
    data1 = pd.concat([data1, sold_data], ignore_index=True)

    # 5) Merge revision, days & GEN (GEN filtered to today & yesterday only)
    revision = pd.read_excel(revision_file, dtype=str)
    days = pd.read_excel(days_file, dtype=str)

    # --- GEN: keep only today & previous day; ensure DD-MM-YYYY format ---
    gen_raw = pd.read_excel(gen_file_path)  # don't force dtype; we need to parse dates
    gen = gen_raw.copy()
    gen.columns = gen.columns.str.strip()

    # find 'date' column (case-insensitive)
    date_col = next((c for c in gen.columns if c.strip().lower() == "date"), None)

    if date_col:
        # normalize to string and trim
        date_str = gen[date_col].astype(str).str.strip()

        # first pass: parse common strings (DD-MM-YYYY etc) with dayfirst
        parsed = pd.to_datetime(date_str, dayfirst=True, errors="coerce")

        # second pass: Excel serial numbers (if any)
        as_num = pd.to_numeric(date_str, errors="coerce")
        need_serial = parsed.isna() & as_num.notna()
        if need_serial.any():
            parsed.loc[need_serial] = pd.to_datetime(
                as_num[need_serial], unit="D", origin="1899-12-30", errors="coerce"
            )

        # compare by date only
        date_only = parsed.dt.date

        # target dates (system local date)
        today = datetime.now().date()
        yesterday = today - pd.Timedelta(days=1)

        keep_mask = date_only.isin([today, yesterday])

        # filter
        gen = gen.loc[keep_mask].copy()

        # write back in DD-MM-YYYY format
        parsed_f = pd.to_datetime(gen[date_col], dayfirst=True, errors="coerce")
        gen[date_col] = parsed_f.dt.strftime("%d-%m-%Y")
    else:
        # no date column -> avoid pulling stale data
        gen = pd.DataFrame(columns=["STOCKID", "gen", "date"])

    # ensure 'gen' column exists / proper name
    if "gen" not in gen.columns:
        maybe_gen = next((c for c in gen.columns if c.strip().lower() == "gen"), None)
        if maybe_gen:
            gen = gen.rename(columns={maybe_gen: "gen"})
        else:
            gen["gen"] = pd.NA

    merged = (
        data1
        .merge(revision[["Sr. No", "DAYS"]], left_on="STOCKID", right_on="Sr. No", how="left")
        .rename(columns={"DAYS": "R D"})
        .merge(days[["Sr. No", "Stock Days", "Luster", "DISC"]], left_on="STOCKID", right_on="Sr. No", how="left")
        .rename(columns={"Stock Days": "O D", "DISC_y": "List Disc", "DISC_x": "DISC"})
        .merge(premal[['stockid', 'pz']], left_on='STOCKID', right_on='stockid', how='left')
        .rename(columns={'pz': 'Premal'})
        .merge(sachin[['stockid', 'sk']], left_on='STOCKID', right_on='stockid', how='left')
        .rename(columns={'sk': 'Sachin'})
        .merge(sagar[['stockid', 'sp']], left_on='STOCKID', right_on='stockid', how='left')
        .rename(columns={'sp': 'Sagar'})
        .merge(gen[['STOCKID', 'gen']], on='STOCKID', how='left')
        .rename(columns={'gen': 'Gen'})
    )

    merged = replace_low(merged)

    for col in ['Premal', 'Sachin', 'Sagar', "R D", "O D", "CERTNO", "DEPTH", "TABLE", "List Disc", "DISC", "Gen"]:
        merged[col] = pd.to_numeric(merged[col], errors='coerce')

    merged['Small'] = merged[['Gen', 'Premal', 'Sachin', 'Sagar']].min(axis=1)

    disc_numeric = merged['DISC']
    small_numeric = merged['Small']
    mask = (merged["STATUS"].str.lower() != "new arrival") & (disc_numeric == small_numeric)
    merged.loc[mask, 'Small'] = np.nan

    merged = merged[~((merged['STATUS'].str.strip() == 'New Arrival') & (merged['Small'].isna()))]

    # 6) Prepare for matching loop
    group_list = []
    seen_new = set()
    new_arrivals = merged[merged["STATUS"].str.lower() == "new arrival"]

    for _, stone in new_arrivals.iterrows():
        sid = stone["STOCKID"]
        if sid in seen_new:
            continue
        seen_new.add(sid)

        cr = find_carat_range(
            stone["CARAT"], shape=stone["SHAPE"],
            fluor=stone.get("FLUOR"), luster=stone.get("LUSTER"),
            tinch=stone.get("TINCH"), black=stone.get("BLACK")
        )
        if not cr:
            grp = stone.to_frame().T
            grp = add_empty_row(grp)
            group_list.append((stone["SHAPE"], stone["CARAT"], grp))
            continue

        base_filters = [
            merged["SHAPE"] == stone["SHAPE"],
            merged["CARAT"].between(*cr),
            merged["CERT"] == stone["CERT"],
            merged["COLOR"] == stone["COLOR"],
            merged["CLARITY"] == stone["CLARITY"],
            merged["LUSTER"] == stone["LUSTER"],
            merged["TINCH"] == stone["TINCH"]
        ]
        seen_filter = (
            (~merged["STOCKID"].isin(seen_new)) | ((merged["SOURCE"] == "Sold") & (merged["STOCKID"] == sid)))
        base_filters.append(seen_filter)

        if stone["STATUS"].strip().lower() == "new arrival":
            for attr in ["CUT", "POL", "SYM", "FLUOR", "LUSTER", "TINCH"]:
                base_filters.append((merged["SOURCE"] != "Sold") | (merged[attr] == stone[attr]))

        if stone["SHAPE"] == "ROUND":
            base_filters += [
                merged["CUT"].isin(get_allowed_cut(stone["CUT"])),
                merged["FLUOR"].isin(get_allowed_fluor(stone["FLUOR"]))
            ]
        else:
            base_filters.append(merged["FLUOR"].isin(get_allowed_fluor(stone["FLUOR"])) )

        matches_df = merged[np.all(base_filters, axis=0)].copy()

        if all(stone[attr] == "EX" for attr in ("CUT", "POL", "SYM")):
            matches_df = matches_df[
                (matches_df["CUT"] == "EX") &
                (matches_df["POL"] == "EX") &
                (matches_df["SYM"] == "EX")
            ].copy()

        seen_new.update(matches_df["STOCKID"])
        new_arrival_df = stone.to_frame().T

        group = new_arrival_df.copy()
        extra_na = matches_df[matches_df["STATUS"].str.lower() == "new arrival"]
        if not extra_na.empty:
            extra_na = extra_na.dropna(axis=1, how="all")
            group = pd.concat([group, extra_na], ignore_index=True)

        mask_stock = matches_df["SOURCE"] == "Stock"
        mask_3ex = (
            (matches_df["CUT"] == stone["CUT"]) &
            (matches_df["POL"] == stone["POL"]) &
            (matches_df["SYM"] == stone["SYM"])
        )
        exact = matches_df[mask_stock & mask_3ex]
        if not exact.empty:
            group = pd.concat([group, exact], ignore_index=True)

        partial = matches_df[mask_stock & ~mask_3ex].copy()
        if not partial.empty:
            partial = partial.sort_values(by=["CUT", "POL", "SYM"], ascending=True, na_position='first')
            common_cols = group.columns.intersection(partial.columns)
            group = pd.concat([group[common_cols], partial[common_cols]], ignore_index=True)

        stock_rows = group[group["SOURCE"] == "Stock"].copy()
        other_rows = group[group["SOURCE"] != "Stock"].copy()
        new_stock = stock_rows[stock_rows["STATUS"].str.lower() == "new arrival"]
        sortable = stock_rows.drop(new_stock.index)
        if not sortable.empty:
            for hit in ("FLUOR", "CUT", "POL", "SYM"):
                sortable[f"{hit.lower()}_hit"] = (sortable[hit] == stone[hit]).astype(int)
            sort_keys = ["fluor_hit", "cut_hit", "pol_hit", "sym_hit"] + [c for c in ("CUT", "POL", "SYM") if c in sortable]
            ascending = [False]*4 + [True]*(len(sort_keys)-4)
            sortable = sortable.sort_values(by=sort_keys, ascending=ascending, na_position='last')
            sortable = sortable.drop(columns=[k for k in sortable.columns if k.endswith("_hit")], errors="ignore")
        stock_rows = pd.concat([new_stock, sortable], ignore_index=True)
        group = pd.concat([stock_rows, other_rows], ignore_index=True)

        sold_rows = merged[(merged["SOURCE"] == "Sold") & (merged["STOCKID"] == sid)].copy()
        if not sold_rows.empty:
            common_cols = group.columns.intersection(sold_rows.columns)
            group = pd.concat([group, sold_rows[common_cols]], ignore_index=True)

        match_cols = ["SHAPE","CARAT","CERT","COLOR","CLARITY","CUT","POL","SYM","FLUOR","LUSTER","TINCH"]
        sold_matches = sold_data.copy()
        for col in match_cols:
            if col != "CARAT":
                sold_matches = sold_matches[sold_matches[col] == stone[col]]
        sold_matches = sold_matches[sold_matches["CARAT"].between(*cr)]
        if "Sold Date" in sold_data.columns:
            sold_data["Sold Date"] = pd.to_datetime(sold_data["Sold Date"], errors="coerce")
            sold_data.sort_values("Sold Date", ascending=False, inplace=True)
            sold_matches = sold_matches.head(5)
            sold_data["Sold Date"] = sold_data["Sold Date"].dt.strftime("%d-%b-%Y")
        if not sold_matches.empty:
            common_cols = group.columns.intersection(sold_matches.columns)
            group = pd.concat([group, sold_matches[common_cols]], ignore_index=True)

        sold_subset = group[group["SOURCE"] == "Sold"]
        sold_subset = sold_subset.drop_duplicates(subset="NAME")
        if len(sold_subset) > 5:
            sold_subset = sold_subset.iloc[:5]
        group = pd.concat([group[group["SOURCE"] != "Sold"], sold_subset], ignore_index=True)

        if not sold_subset.empty:
            sold_subset["Sold Disc"] = sold_subset.apply(calculate_disc, axis=1)
            sold_discs = sold_subset["Sold Disc"].dropna().astype(str).tolist()
            names = sold_subset["NAME"].dropna().astype(str).unique().tolist()
            dates = sold_subset["Sold Date"].dropna().astype(str).unique().tolist()
            if sold_discs:
                group.at[0, "Sold Disc"] = "/".join(sold_discs)
            if names:
                group.at[0, "NAME"] = "/".join(names)
            if dates:
                group.at[0, "Sold Date"] = "/".join(dates)

        group = group.drop_duplicates(subset=["STOCKID", "SOURCE"], keep="first").reset_index(drop=True)

        total_stock = (group['SOURCE'] == 'Stock').sum()
        total_sold = (group['SOURCE'] == 'Sold').sum()
        group['Total Stock'] = np.where(group['STATUS'].str.strip().str.lower() == 'new arrival', total_stock, None)
        group['Total Sold']  = np.where(group['STATUS'].str.strip().str.lower() == 'new arrival', total_sold, None)

        group = add_empty_row(group)
        group_list.append((stone["SHAPE"], stone["CARAT"], group))

    # --- Group Sorting Logic ---
    sorted_groups = []
    for shape, carat, group in group_list:
        if shape == "ROUND":
            sorted_groups.append((carat, group))
        else:
            sorted_groups.append((None, group))
    sorted_groups.sort(key=lambda x: (x[0] is None, x[0]))
    results = pd.concat([g for _, g in sorted_groups], ignore_index=True)

    results["CODE"] = results.apply(get_code, axis=1)
    mask_3ex_final = ((results[["CUT","POL","SYM"]] == "EX").all(axis=1))
    results.loc[mask_3ex_final, "CUT"] = "3EX"
    results = results.loc[results["SOURCE"].str.lower() != "sold"].reset_index(drop=True)
    for col in desired_order:
        if col not in results.columns:
            results[col] = None
    results = results[desired_order]

    # Write workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Matched Stones"
    for row in dataframe_to_rows(results, index=False, header=True):
        ws.append(row)
    apply_styles_to_excel(ws, results)

    # Append sources to Disc.xlsx (uses filtered GEN)
    append_sources_to_book2(premal, sachin, sagar, gen)

    # Save output
    output_name = f"Hemil Check {datetime.now():%d-%m-%Y}.xlsx"
    outp = os.path.join(os.path.dirname(input_file), output_name)
    wb.save(outp)
    open_file(outp)
    print(f"Saved styled matched stones file to: {outp}")


if __name__ == "__main__":
    match_stones()
