import os
import re
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
import streamlit as st

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

def load_excel(stock_file, days_file):
    df = pd.read_excel(stock_file)
    df_days = pd.read_excel(days_file)
    df_days.rename(columns={"DISC": "List Disc"}, inplace=True)
    
    # Convert key columns to numeric
    for col in ("CARAT", "DEPTH", "TABLE"):
        df[col] = pd.to_numeric(df.get(col), errors="coerce")

    df["CARAT_NUM"] = df["CARAT"]
    df["CARAT_RANGE"] = df.apply(get_carat_label, axis=1)
    df = df.merge(df_days[["Sr. No", "Stock Days", "Luster","List Disc"]], left_on="STOCKID", right_on="Sr. No", how="left")
    
    return df

def filter_groups(df, group_cols):
    """Keep only groups with more than one distinct DISC value."""
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

def apply_style(ws, row, col, fill_color="FFCCCC", font_color="FF5050", bold=True):
    """Apply fill and font styling to a cell."""
    cell = ws.cell(row=row, column=col)
    cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
    cell.font = Font(color=font_color, bold=bold)

from io import BytesIO

def write_highlighted_excel(df, group_cols):
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

                    # Apply existing conditional highlights
                    if col_name == "KEY TO SYMBOLS" and val == "Etch Channel- Cloud":
                        apply_style(ws, row_idx, col_idx)
                    elif col_name == "POL" and val == "GD":
                        apply_style(ws, row_idx, col_idx)
                    # Add your other conditional styles here

                row_idx += 1
        # blank line between groups
        row_idx += 1

    # Save to BytesIO instead of directly saving to file
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def run_pending_stock():
    st.title("Pending Stock List")
    with st.sidebar:
        st.markdown("### Upload File")
        uploaded_stock_file = st.file_uploader("Select Stock File", type=["xlsx", "xls"])
        uploaded_days_file = st.file_uploader("Select Days File", type=["xlsx", "xls"])
    
    with st.spinner("Processing files..."):   
    
        if not uploaded_stock_file or not uploaded_days_file:
            st.info("⬆️ Please upload both files to continue.")
            st.stop()

        if uploaded_stock_file and uploaded_days_file:
            df = load_excel(uploaded_stock_file, uploaded_days_file)

            df["DISC"] = pd.to_numeric(df["DISC"], errors="coerce")
            df["Luster"] = df["Luster"].str.strip().str.upper().replace({"NORMAL": "EX", "MILKY": "M-2", "LIGHT MILKY": "M-1"})
            df = remove_new_arrival(df)

            group_cols = [
                "SHAPE", "RAPSIZE", "CERT", "COLOR", "CLARITY", "CUT",
                "POL", "SYM", "FLUOR", "BLACK", "STATUS", "LUSTER",
                "TINCH", "CARAT_RANGE"
            ]
            desired_order = [
                "STOCKID", "V360", "SHAPE", "RAPSIZE", "CERT", "CARAT", "COLOR", "CLARITY",
                "CUT", "POL", "SYM", "FLUOR", "MEASUREMENT", "DEPTH", "TABLE",
                "DISC", "List Disc", "Stock Days", "BLACK", "Luster", "TINCH", "KEY TO SYMBOLS", "REPORT COMM",
                "RAPPRICE", "ONLINE PRICE", "CERTNO", "RATIO",
                "STATUS", "SOURCE"
            ]
            filtered = filter_groups(df, group_cols)

            # Ensure all desired columns exist
            for col in desired_order:
                if col not in filtered.columns:
                    filtered[col] = None

            # Reorder columns
            cols_present = [c for c in desired_order if c in filtered.columns]
            extras = [c for c in filtered.columns if c not in cols_present]
            filtered = filtered[cols_present + extras]
            
            st.dataframe(filtered)
            
            output_path = f"Pending Stock List {datetime.now().strftime('%d-%m-%Y')}.xlsx"
            
            # Prepare Excel file in memory
            excel_file = write_highlighted_excel(filtered, group_cols)

            # Add a download button
            st.sidebar.download_button(
                label="⬇️ Download Excel",
                data=excel_file,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    run_pending_stock()