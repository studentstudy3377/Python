import os
import re
from datetime import datetime
from tkinter import filedialog
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# Custom sort orders
FLUOR_ORDER = {
    "NONE": 1,
    "FNT": 2,
    "VSL": 3,
    "MED": 4,
    "SL": 5,
    "STG": 6,
    "VST": 7,
}
CERT_ORDER = {"GIA": 1, "HRD": 2, "IGI": 3, "IIDGR": 4}
CLARITY_ORDER = {
    "FL": 1,
    "IF": 2,
    "VVS1": 3,
    "VVS2": 4,
    "VS1": 5,
    "VS2": 6,
    "SI1": 7,
    "SI2": 8,
}
BLACK_ORDER = {
    "NO BLK": 1,
    "SIDE - L": 2,
    "CENTER - L": 3,
    "S & C - L": 4,
    "SIDE - H": 5,
    "CENTER - H": 6,
    "S & C - H": 7,
    "RED SPOT": 8,
    "HEAVY": 9,
}


# Define carat ranges and labels
CARAT_RANGES = [
    (0.23, 0.29), (0.30, 0.36), (0.37, 0.39), (0.40, 0.45), (0.46, 0.49),
    (0.50, 0.54), (0.55, 0.59), (0.60, 0.69), (0.70, 0.79), (0.80, 0.89),
    (0.90, 0.99), (1.00, 1.19), (1.20, 1.29), (1.30, 1.39), (1.40, 1.49),
    (1.50, 1.59), (1.60, 1.69), (1.70, 1.79), (1.80, 1.89), (1.90, 1.99),
    (2.00, 2.19), (2.20, 2.49), (2.50, 2.99), (3.00, 3.49), (3.50, 3.99),
    (4.00, 4.49), (4.50, 4.99),
]
CARAT_RANGE_LABELS = [f"{s:.2f}-{e:.2f}" for s, e in CARAT_RANGES]


def get_carat_label(row):
    
    try:
        c = float(row.get("CARAT_NUM", 0))
    except (TypeError, ValueError):
        return "Other"

    fluor = str(row.get("FLUOR", "")).strip().upper()
    if fluor in ("NONE", "FNT"):
        if 0.50 <= c <= 0.54:
            return "0.50-0.54"
        if 0.55 <= c <= 0.59:
            return "0.55-0.59"

    for (start, end), label in zip(CARAT_RANGES, CARAT_RANGE_LABELS):
        if start <= c <= end:
            return label

    return "Other"


def load_excel(path=None):
   
    path = filedialog.askopenfilename(  title="Select Stock File", filetypes=[("Excel Files", "*.xlsx")])
    days = filedialog.askopenfilename(  title="Select Days File", filetypes=[("Excel Files", "*.xlsx")])
      
    df = pd.read_excel(path)
    df_days = pd.read_excel(days)
    df_days.rename(columns={"DISC": "List Disc"}, inplace=True)
    # Convert key columns to numeric
    for col in ("CARAT", "DEPTH", "TABLE"):
        df[col] = pd.to_numeric(df.get(col), errors="coerce")

    df["CARAT_NUM"] = df["CARAT"]
    df["CARAT_RANGE"] = df.apply(get_carat_label, axis=1)

    df = df.merge(df_days[["Sr. No", "Stock Days", "Luster","List Disc"]],left_on="STOCKID", right_on="Sr. No",how="left")


    return df

def filter_groups(df, group_cols):
    """
    Keep only groups with more than one distinct DISC value.
    """
    df_copy = df.copy()
    df_copy["GROUP_SIZE"] = df_copy.groupby(group_cols)["DISC"].transform("size")
    multi = df_copy[df_copy["GROUP_SIZE"] > 1]

    filtered = []
    for _, group in multi.groupby(group_cols, sort=False):
        if group["DISC"].nunique(dropna=False) > 1:
            filtered.append(group)

    if not filtered:
        return df_copy.head(0)
    return pd.concat(filtered, ignore_index=True)


def remove_new_arrival(df):
    """Remove rows where STATUS is 'New Arrival'."""
    return df[~df["STATUS"].str.strip().str.lower().eq("new arrival")]


def apply_style(
    ws,
    row,
    col,
    fill_color="FFCCCC",
    font_color="FF5050",
    bold=True,
):
    """Apply fill and font styling to a cell."""
    cell = ws.cell(row=row, column=col)
    cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
    cell.font = Font(color=font_color, bold=bold)


def write_highlighted_excel(df, group_cols, output_path):
    """
    Sort data, highlight worst values per group, highlight best rows in yellow,
    and write to Excel with invert‐highlight logic.
    """
    df_copy = df.copy()
    # Create sort keys
    df_copy["CERT_SORT"] = (
        df_copy["CERT"].str.upper().map(CERT_ORDER)
        .fillna(len(CERT_ORDER) + 1).astype(int)
    )
    df_copy["CLARITY_SORT"] = (
        df_copy["CLARITY"].str.upper().map(CLARITY_ORDER)
        .fillna(max(CLARITY_ORDER.values()) + 1).astype(int)
    )
    df_copy["FLUOR_SORT"] = (
        df_copy["FLUOR"].str.upper().map(FLUOR_ORDER)
        .fillna(max(FLUOR_ORDER.values()) + 1).astype(int)
    )
    df_copy["BLACK_SORT"] = (
        df_copy["BLACK"].str.upper().map(BLACK_ORDER)
        .fillna(max(BLACK_ORDER.values()) + 1).astype(int)
    )

    secondary_cols = ["CUT", "POL", "SYM", "STATUS", "LUSTER", "TINCH"]
    primary_keys = [c for c in group_cols if c not in (
        "CERT", "CLARITY", "FLUOR", "BLACK"
    )]
    sort_keys = [
        "CARAT", "CARAT_RANGE", *primary_keys,
        "CERT_SORT", "CLARITY_SORT", "FLUOR_SORT", "BLACK_SORT",
        *secondary_cols, "DISC"
    ]
    ascending = [True] * (len(sort_keys) - 1) + [False]
    df_sorted = df_copy.sort_values(by=sort_keys, ascending=ascending)

    # Drop helper columns
    drop_helpers = [
        'CARAT_NUM', 'CERT_SORT', 'CLARITY_SORT',
        'FLUOR_SORT', 'BLACK_SORT', 'Unnamed: 26', 'GROUP_SIZE','Sr. No'
    ]
    df_sorted = df_sorted.drop(columns=drop_helpers, errors='ignore')

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Write header
    for idx, col in enumerate(df_sorted.columns, start=1):
        ws.cell(row=1, column=idx, value=col)

    row_idx = 2
    for _, group in df_sorted.groupby(group_cols, sort=False):
        # 1) Find the maximum DISC in the group
        top_disc = group["DISC"].max()
        # 2) Mark rows equal to that top value
        mask = group["DISC"] == top_disc
        # 3) If more than half the group is “top,” invert the mask
        if mask.sum() > len(group) / 2:
            mask = ~mask

        # 4) Split into the to‐be‐highlighted and the rest
        high = group[mask]
        rest = group[~mask]

        for subset in (high, rest):
            for _, record in subset.iterrows():
                raw_meas = record.get("MEASUREMENT", "")
                nums = re.findall(r"[\d\.]+", str(raw_meas))
                measure_val = max(map(float, nums)) if nums else None

                depth = record.get("DEPTH")
                table = record.get("TABLE")
                rapsize = record.get("RAPSIZE")

                is_high = record.name in high.index

                for col_idx, col_name in enumerate(df_sorted.columns, start=1):
                    val = record[col_name]
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)

                    # Highlight inverted‐best rows in yellow
                    if is_high:
                        cell.fill = PatternFill(fill_type="solid", fgColor="F9FAAF")

                    # (existing conditional highlights follow)
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
                    elif (
                        col_name == "REPORT COMM" and isinstance(val, str) and
                        "Clarity grade is based on clouds that are not shown." in val
                    ):
                        apply_style(ws, row_idx, col_idx)
                    elif (
                        col_name in ("DEPTH", "TABLE") and
                        isinstance(depth, (int, float)) and
                        isinstance(table, (int, float)) and depth < table
                    ):
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "MEASUREMENT" and measure_val is not None:
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
        # blank line between groups
        row_idx += 1

    wb.save(output_path)

def main():
    """Main process to filter and write highlighted stock list."""
    group_cols = [
        "SHAPE", "RAPSIZE", "CERT", "COLOR", "CLARITY", "CUT",
        "POL", "SYM", "FLUOR", "BLACK", "STATUS", "LUSTER",
        "TINCH", "CARAT_RANGE"
    ]

    desired_order = [
        "STOCKID", "V360", "SHAPE", "RAPSIZE", "CERT", "CARAT", "COLOR", "CLARITY",
        "CUT", "POL", "SYM", "FLUOR", "MEASUREMENT", "DEPTH", "TABLE",
        "DISC","List Disc","Stock Days","BLACK", "Luster", "TINCH", "KEY TO SYMBOLS", "REPORT COMM",
        "RAPPRICE", "ONLINE PRICE", "CERTNO", "RATIO",
        "STATUS", "SOURCE"
    ]

    date_str = datetime.now().strftime("%d-%m-%Y")
    output_path = fr"D:\Python\Pending Stock List\Output\Pending Stock List {date_str}.xlsx"

    df = load_excel()
    df["DISC"] = pd.to_numeric(df["DISC"], errors="coerce")
    df["Luster"] = df["Luster"].str.strip().str.upper().replace({"NORMAL": "EX","MILKY":"M-2","LIGHT MILKY":"M-1"})
    df = remove_new_arrival(df)

    # Create filtered BEFORE using it
    filtered = filter_groups(df, group_cols)

    # Ensure all desired columns exist on the filtered dataframe
    for col in desired_order:
        if col not in filtered.columns:
            filtered[col] = None

    # Reorder columns (keep extras at the end, if any)
    cols_present = [c for c in desired_order if c in filtered.columns]
    extras = [c for c in filtered.columns if c not in cols_present]
    filtered = filtered[cols_present + extras]

    # Write out
    write_highlighted_excel(filtered, group_cols, output_path)
    os.startfile(output_path)

if __name__ == "__main__":
    main()