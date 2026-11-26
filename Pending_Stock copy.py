import os
import re
from datetime import datetime
from tkinter import filedialog

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


# Mapping and expected schema
COLUMN_MAPPING = {
    "STOCK ID": "STOCKID",
    "POLISH": "POL",
    "SYMMETRY": "SYM",
    "BLACK INCLUSION": "BLACK",
    "TINGE": "TINCH",
    "KEY TO SYMBOL": "KEY TO SYMBOLS",
    "CERTIFIED NO": "CERTNO",
    "DATE": "Sold Date",
    "RAPPRICE($)": "RAPPRICE",
    "REPORT COMMENTS": "REPORT COMM",
    "FINAL CUSTOMER": "RAPSIZE",
}

EXPECTED_COLUMNS = [
    "STOCKID", "V360", "SHAPE", "RAPSIZE", "CERT", "CARAT", "COLOR",
    "CLARITY", "CUT", "POL", "SYM", "FLUOR", "DISC", "MEASUREMENT",
    "DEPTH", "TABLE", "BLACK", "LUSTER", "TINCH", "KEY TO SYMBOLS",
    "REPORT COMM", "RATIO", "CERTNO", "STATUS", "RAPPRICE", "ONLINE PRICE",
    "SOURCE",
]

# Custom sort orders
FLUOR_ORDER = {"NONE": 1, "FNT": 2, "VSL": 3, "MED": 4, "SL": 5, "STG": 6, "VST": 7}
CERT_ORDER = {"GIA": 1, "HRD": 2, "IGI": 3, "IIDGR": 4}
LUSTER_ORDER = {("EX","LOW"),("M-1","LOW"),("M-2","LOW") }
CLARITY_ORDER = {"FL": 1, "IF": 2, "VVS1": 3, "VVS2": 4, "VS1": 5, "VS2": 6, "SI1": 7, "SI2": 8}
BLACK_ORDER = {
    "NO BLK": 1, "SIDE - L": 2, "CENTER - L": 3, "S & C - L": 4,
    "SIDE - H": 5, "CENTER - H": 6, "S & C - H": 7, "RED SPOT": 8, "HEAVY": 9,
}

# Carat ranges
CARAT_RANGES = [
    (0.23, 0.29), (0.30, 0.36), (0.37, 0.39), (0.40, 0.45), (0.46, 0.49),
    (0.50, 0.54), (0.55, 0.59), (0.60, 0.69), (0.70, 0.79), (0.80, 0.89),
    (0.90, 0.99), (1.00, 1.19), (1.20, 1.29), (1.30, 1.39), (1.40, 1.49),
    (1.50, 1.59), (1.60, 1.69), (1.70, 1.79), (1.80, 1.89), (1.90, 1.99),
    (2.00, 2.19), (2.20, 2.49), (2.50, 2.99), (3.00, 3.49), (3.50, 3.99),
    (4.00, 4.49), (4.50, 4.99),
]
CARAT_RANGE_LABELS = [f"{s:.2f}-{e:.2f}" for s, e in CARAT_RANGES]

STATUS_PRIORITY = {"PROMOTION": 0, "NEW": 1, "CLEARANCE": 2}


def get_carat_label(row):
    try:
        carat = float(row.get("CARAT_NUM", 0))
    except (TypeError, ValueError):
        return "Other"

    fluor = str(row.get("FLUOR", "")).strip().upper()
    if fluor in ("NONE", "FNT"):
        if 0.50 <= carat <= 0.54:
            return "0.50-0.54"
        if 0.55 <= carat <= 0.59:
            return "0.55-0.59"

    for (start, end), label in zip(CARAT_RANGES, CARAT_RANGE_LABELS):
        if start <= carat <= end:
            return label

    return "Other"


def load_combined_stock_and_sold():
    stock_path = filedialog.askopenfilename(
        title="Select Stock File", filetypes=[("Excel Files", "*.xlsx")]
    )
    sold_path = filedialog.askopenfilename(
        title="Select Sold File", filetypes=[("Excel Files", "*.xlsx")]
    )
    days_path = filedialog.askopenfilename(
        title="Select Days File ",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    df_stock = pd.read_excel(stock_path)
    df_sold = pd.read_excel(sold_path, skiprows=2)
    df_days = pd.read_excel(days_path)
    df_sold = df_sold.rename(columns=COLUMN_MAPPING)

    # Ensure all expected columns present
    for df in (df_stock, df_sold):
        for col in EXPECTED_COLUMNS:
            if col not in df.columns:
                df[col] = pd.NA

    # Set statuses and sources
    df_stock["STATUS"] = df_stock["STATUS"].fillna("IN STOCK")
    df_sold["STATUS"] = df_sold["Sold Date"]
    if "FINAL PRICE" in df_sold.columns:
        df_sold["ONLINE PRICE"] = df_sold["FINAL PRICE"]
    df_stock["SOURCE"] = "Stock"
    df_sold["SOURCE"] = "Sold"
    df_days.rename(columns={"DISC": "List Disc"}, inplace=True)

    # Calculate discount
    def calc_disc(row):
        if row["SOURCE"] == "Stock":
            return row.get("DISC", pd.NA)
        rap = row.get("RAPPRICE", 0) or 0
        fin = row.get("FINAL PRICE", 0) or 0
        return round(((fin - rap) / rap) * 100, 2) if rap else 0
    for df in (df_stock, df_sold):
        df["DISC"] = df.apply(calc_disc, axis=1)

    combined = pd.concat([
        df_stock[EXPECTED_COLUMNS],
        df_sold[EXPECTED_COLUMNS]
    ], ignore_index=True)

    # Merge in days data by STOCKID -> Sr. No
    combined = combined.merge(df_days[["Sr. No", "Stock Days", "Luster","List Disc"]],left_on="STOCKID", right_on="Sr. No",how="left")
    combined["CARAT_NUM"] = pd.to_numeric(combined["CARAT"], errors="coerce")
    combined["CARAT_RANGE"] = combined.apply(get_carat_label, axis=1)

    # Sort by status or date
    def parse_status(val):
        s = str(val).strip().upper()
        if s in STATUS_PRIORITY:
            return (0, STATUS_PRIORITY[s])
        try:
            dt = datetime.strptime(s, "%d-%b-%Y")
            return (1, -dt.timestamp())
        except ValueError:
            return (2, s)
    combined["STATUS_SORT"] = combined["STATUS"].apply(parse_status)
    combined = combined.sort_values(by="STATUS_SORT", ascending=True)
    combined = combined.drop(columns=["STATUS_SORT"])

    # 3) Reorder so that "List Disc" sits just after "DISC"
    cols = combined.columns.tolist()
    if "List Disc" in cols and "DISC" in cols:
        # remove List Disc from where it is
        cols.pop(cols.index("List Disc"))
        # re-insert it right after DISC
        insert_at = cols.index("DISC") + 1
        cols.insert(insert_at, "List Disc")
        combined = combined[cols]

    return combined


def remove_all_sold_groups(df, group_cols):
    mask = df.groupby(group_cols)["SOURCE"].transform( lambda x: x.nunique() == 2)

    return df[mask.fillna(False)]


def filter_groups(df, group_cols):
    df2 = df.copy()
    df2["GROUP_SIZE"] = df2.groupby(group_cols)["DISC"].transform("size")
    candidates = df2[df2["GROUP_SIZE"] > 1]
    out = []

    for _, g in candidates.groupby(group_cols, sort=False):
        if g["DISC"].nunique(dropna=False) > 1:
            out.append(g)

    return pd.concat(out, ignore_index=True) if out else df2.head(0)

def remove_new_arrival(df):
    return df[~df["STATUS"].str.strip().str.lower().eq("new arrival")]

def apply_style(ws, row, col, fill_color="FFCCCC",
                font_color="FF5050", bold=True):
    cell = ws.cell(row=row, column=col)
    cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
    cell.font = Font(color=font_color, bold=bold)

def write_highlighted_excel(df, group_cols, output_path):
    dfc = df.copy()
    dfc["CERT_SORT"] = dfc["CERT"].str.upper().map(CERT_ORDER).fillna(len(CERT_ORDER) + 1).astype(int)
        
    dfc["CLARITY_SORT"] = dfc["CLARITY"].str.upper().map(CLARITY_ORDER).fillna(max(CLARITY_ORDER.values()) + 1).astype(int)
            
    dfc["FLUOR_SORT"] = dfc["FLUOR"].str.upper().map( FLUOR_ORDER).fillna(max(FLUOR_ORDER.values()) + 1).astype(int)
          
    dfc["BLACK_SORT"] = dfc["BLACK"].str.upper().map(BLACK_ORDER).fillna(max(BLACK_ORDER.values()) + 1).astype(int)
        
    def parse_status(val):
        s = str(val).strip().upper()
        if s in STATUS_PRIORITY:
            return (0, STATUS_PRIORITY[s])
        try:
            dt = datetime.strptime(s, "%d-%b-%Y")
            return (1, -dt.timestamp())
        except ValueError:
            return (2, s)

    dfc["STATUS_SORT"] = dfc["STATUS"].apply(parse_status)

    primary = [c for c in group_cols if c not in ("CERT", "CLARITY", "FLUOR", "BLACK")]
    secondary = ["CUT", "POL", "SYM", "STATUS", "TINCH","LUSTER"]
    sort_keys = ["STATUS_SORT", "CARAT", "CARAT_RANGE", *primary,
                 "CERT_SORT", "CLARITY_SORT", "FLUOR_SORT", "BLACK_SORT",
                 *secondary, "DISC"]
    ascending = [True] * (len(sort_keys) - 1) + [False]

    dfs = dfc.sort_values(by=sort_keys, ascending=ascending).drop(
        columns=["CARAT_NUM", "CERT_SORT", "CLARITY_SORT","FLUOR_SORT", "BLACK_SORT", "GROUP_SIZE", "STATUS_SORT"],  errors="ignore")
                 
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for i, col in enumerate(dfs.columns, start=1):
        ws.cell(row=1, column=i, value=col)

    row_idx = 2
    for _, grp in dfs.groupby(group_cols, sort=False):
        grp = grp.sort_values(by="SOURCE", key=lambda x: x.eq("Sold"))

        stock = grp[grp["SOURCE"] == "Stock"]
        top_disc = stock["DISC"].max() if not stock.empty else None
        high_mask = (grp["DISC"] == top_disc) & (grp["SOURCE"] == "Stock")

        if len(stock) > 1 and high_mask.sum() > len(stock) / 2:
            high_mask = (~high_mask) & (grp["SOURCE"] == "Stock")

        if not high_mask.any():
            high_mask = high_mask.copy()
            if not stock.empty:
                high_mask.loc[stock.index[0]] = True
            else:
                high_mask.iloc[0] = True

        high = grp[high_mask]
        rest = grp[~high_mask]

        # Sort stock rows by CARAT within each subset
        def sort_subset(sub):
            stock_rows = sub[sub["SOURCE"] == "Stock"].sort_values(by="CARAT")
            sold_rows = sub[sub["SOURCE"] == "Sold"]
            return pd.concat([stock_rows, sold_rows], ignore_index=False)

        for subset in (sort_subset(high), sort_subset(rest)):
            for _, rec in subset.iterrows():
                nums = re.findall(r"[\d\.]+", str(rec.get("MEASUREMENT", "")))
                measure_val = max(map(float, nums)) if nums else None
                depth, table = rec.get("DEPTH"), rec.get("TABLE")
                rapsize = rec.get("RAPSIZE")
                is_high = rec.name in high.index

                for col_idx, col_name in enumerate(dfs.columns, start=1):
                    val = rec[col_name]
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)

                    if rec["SOURCE"] == "Sold":
                        cell.fill = PatternFill(fill_type="solid", fgColor="C2F0F0")
                    elif is_high:
                        cell.fill = PatternFill(fill_type="solid", fgColor="F9FAAF")

                    # Existing conditional highlights
                    if col_name == "KEY TO SYMBOLS" and val == "Etch Channel- Cloud":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "POL" and val == "GD":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "CERT" and val in ("HRD", "IGI"):
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "SYM" and val == "GD":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "CLARITY" and val == "I1":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "CUT" and val == "GD":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "TINCH" and val in ("B-1", "MIX"):
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "FLUOR" and val != "NONE":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "SHAPE" and val != "ROUND":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "LUSTER" and val != "EX":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "BLACK" and isinstance(val, str) and "H" in val:
                        apply_style(ws, row_idx, col_idx)
                    elif (col_name == "REPORT COMM" and isinstance(val, str)
                          and "Clarity grade is based on clouds that are not shown." in val):
                        apply_style(ws, row_idx, col_idx)
                    elif (col_name in ("DEPTH", "TABLE") and
                          isinstance(depth, (int, float)) and
                          isinstance(table, (int, float)) and depth < table):
                        apply_style(ws, row_idx, col_idx)
                    # New measurement-based conditional
                    elif (col_name == "MEASUREMENT" and measure_val is not None):
                        thresholds = {
                            "30UP": 4.50,
                            "40UP": 5.00,
                            "50UP": 5.25,
                            "60UP": 5.50,
                            "70UP": 5.85,
                            "80UP": 6.00,
                            "90UP": 6.15,
                            "1.00UP": 6.30,
                        }
                        limit = thresholds.get(rapsize)
                        if limit is not None and measure_val > limit:
                            apply_style(ws, row_idx, col_idx)

                row_idx += 1
        row_idx += 1
    wb.save(output_path)

def main():
    group_cols = [
        "SHAPE", "CERT", "COLOR", "CLARITY", "CUT",
        "POL", "SYM", "FLUOR", "BLACK","LUSTER",
        "TINCH", "CARAT_RANGE",
    ]
    now = datetime.now().strftime("%d-%m-%Y")
    output_dir = r"G:\Certificate\Python Output"
    out_path = os.path.join(output_dir, f"Pending Stock List {now}.xlsx")

    df = load_combined_stock_and_sold()
    df["DISC"] = pd.to_numeric(df["DISC"], errors="coerce")
    df = remove_new_arrival(df)
    df = remove_all_sold_groups(df, group_cols)
    df = filter_groups(df, group_cols)
     # —– FINAL LUSTER SYNC —–
    if "Luster" in df.columns:
        df.loc[df["Luster"] == "LOW", "LUSTER"] = "LOW"
     # Drop merge helper columns
    df = df.drop(columns=["Sr. No", "Luster"], errors='ignore')
    write_highlighted_excel(df, group_cols, out_path)
    os.startfile(out_path)

if __name__ == "__main__":
    main()

