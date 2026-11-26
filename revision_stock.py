import base64
from pathlib import Path

from openpyxl.styles.fills import PatternFill
from openpyxl.styles.fonts import Font
import streamlit as st
import pandas as pd
import numpy as np
import io


def read_file(uploaded_file, skiprows: int = 0):
    # Read an Excel file into a pandas DataFrame
    # skiprows      : number of rows to skip at the top of the file
    return pd.read_excel(uploaded_file, skiprows=skiprows)

# Convert a specific column to numeric, coercing invalid values to NaN
def numeric(df, col):
    # Check if the column exists in the DataFrame
    if col in df.columns:
        # Convert column values to numeric (invalid values become NaN)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    # Return the modified DataFrame
    return df

# Normalize and standardize the LUSTER column
def replace_low(df: pd.DataFrame) -> pd.DataFrame:
    # If column is named "Luster" but not "LUSTER", create "LUSTER"
    if "Luster" in df.columns and "LUSTER" not in df.columns:
        df["LUSTER"] = df["Luster"]

    # If "LUSTER" exists, normalize values equal to "LOW"
    if "LUSTER" in df.columns:
        # Convert values to string, uppercase them, and match "LOW"
        df.loc[
            df["LUSTER"].astype(str).str.upper().eq("LOW"),
            "LUSTER"
        ] = "LOW"

    # Return the cleaned DataFrame
    return df

# Desired final column order for output DataFrame
desired_order = [
    "STOCKID", "V360", "SHAPE", "RAPSIZE", "CERT", "CARAT", "COLOR", "CLARITY",
    "CUT", "POL", "SYM", "FLUOR", "MEASUREMENT", "Trend", "RATIO", "DEPTH",
    "TABLE", "DISC", "GEN", "Pricing", "BLACK", "LUSTER", "TINCH",
    "KEY TO SYMBOLS", "REPORT COMM", "Sold Date", "R D", "O D", "CODE",
    "CERTNO", "STATUS", "RAPPRICE", "ONLINE PRICE", "SOURCE"
]

# Mapping of inconsistent column names to standardized names
column_mapping = {
    "STOCK ID":        "STOCKID",        # Normalize stock ID naming
    "POLISH":          "POL",             # Polish abbreviation
    "SYMMETRY":        "SYM",             # Symmetry abbreviation
    "BLACK INCLUSION": "BLACK",           # Black inclusion field
    "TINGE":           "TINCH",           # Tinge naming standard
    "KEY TO SYMBOL":   "KEY TO SYMBOLS",  # Expand to plural
    "CERTIFIED NO":    "CERTNO",           # Certificate number
    "DATE":            "Sold Date",        # Sale date
    "RAPPRICE($)":     "RAPPRICE",         # Rap price column
    "REPORT COMMENTS": "REPORT COMM"       # Shortened report comments
}


def get_special_symbol_group(value, fluor=None) -> str:
    """
    Return a grouping key for special values in KEY TO SYMBOLS.

    DON'T MAP these values into the NORMAL group:
        CHIP
        CAVITY
        MANUFACTRING REMANANT

    If KEY TO SYMBOLS contains any of these values, the row stays in
    that special group regardless of FLUOR. Rows without these values
    return "NORMAL".

    Both the correct spelling "MANUFACTURING REMNANT" and the older
    spelling "MANUFACTRING REMANANT" are recognized as the same
    manufacturing-remnant special group.
    """

    # FLUOR intentionally does NOT change special-symbol grouping.
    # CHIP / CAVITY / MANUFACTRING REMANANT always stay separate.
    if pd.isna(value):
        return "NORMAL"

    text = str(value).strip().upper()
    found = []

    if "CHIP" in text:
        found.append("CHIP")

    if "CAVITY" in text:
        found.append("CAVITY")

    if (
        "MANUFACTURING REMNANT" in text
        or "MANUFACTRING REMANANT" in text
    ):
        found.append("MANUFACTURING REMNANT")

    if not found:
        return "NORMAL"

    # If a row contains more than one special value,
    # that exact combination forms its own separate group.
    return "|".join(sorted(set(found)))


# Filter sold stones based on matching New Arrival stones within a group
def filter_sold_by_new_arrival(group: pd.DataFrame) -> pd.DataFrame:
    """
    Keep non-sold rows and attach up to the latest 5 matching Sold stones.

    In addition to the normal SHAPE/COLOR/CLARITY/etc. matching rules,
    KEY TO SYMBOLS is split into special groups:

        CHIP
        CAVITY
        MANUFACTURING REMNANT
        NORMAL

    Therefore a New Arrival containing one of the special symbols only
    receives Sold stones from the same special-symbol group.
    """

    # Normalize SOURCE safely
    source_norm = (
        group["SOURCE"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Keep every non-Sold row
    non_sold = group[source_norm.ne("SOLD")]

    # Sold rows
    sold = group[source_norm.eq("SOLD")].copy()

    # Normalize STATUS safely
    status_norm = (
        group["STATUS"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # New Arrival rows
    new_arrival = group[status_norm.eq("NEW ARRIVAL")]

    # No New Arrival -> remove Sold rows
    if new_arrival.empty:
        return non_sold

    # No Sold rows -> nothing else to add
    if sold.empty:
        return non_sold

    def norm_series(s: pd.Series) -> pd.Series:
        return (
            s.fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    def norm_val(v) -> str:
        return str(v).strip().upper() if pd.notna(v) else ""

    # Pre-calculate special group for Sold stones
    if "KEY TO SYMBOLS" in sold.columns:
        sold["_SPECIAL_SYMBOL_GROUP"] = sold.apply(
            lambda row: get_special_symbol_group(
                row.get("KEY TO SYMBOLS"),
                row.get("FLUOR")
            ),
            axis=1
        )
    else:
        sold["_SPECIAL_SYMBOL_GROUP"] = "NORMAL"

    keep_sold_chunks = []

    # Match Sold rows against every New Arrival in this group
    for _, na_row in new_arrival.iterrows():

        shape = norm_val(na_row.get("SHAPE"))

        # Special-symbol group for this New Arrival
        na_special_group = get_special_symbol_group(
            na_row.get("KEY TO SYMBOLS"),
            na_row.get("FLUOR")
        )

        if shape == "ROUND":
            match_cols = [
                "SHAPE", "COLOR", "CLARITY", "CUT",
                "POL", "SYM", "FLUOR", "LUSTER",
                "TINCH", "BLACK"
            ]
        else:
            # Fancy shapes do not compare CUT
            match_cols = [
                "SHAPE", "COLOR", "CLARITY",
                "POL", "SYM", "FLUOR", "LUSTER",
                "TINCH", "BLACK"
            ]

        mask = pd.Series(True, index=sold.index)

        # IMPORTANT:
        # Special-symbol group must match exactly.
        #
        # CHIP -> CHIP
        # CAVITY -> CAVITY
        # MANUFACTURING REMNANT -> MANUFACTURING REMNANT
        # NORMAL -> NORMAL
        mask &= sold["_SPECIAL_SYMBOL_GROUP"].eq(na_special_group)

        # Existing normal matching rules
        for col in match_cols:
            if col not in sold.columns or col not in group.columns:
                continue

            mask &= norm_series(sold[col]).eq(
                norm_val(na_row.get(col))
            )

        if mask.any():
            keep_sold_chunks.append(
                sold.loc[mask].copy()
            )

    # Combine matching Sold stones
    if keep_sold_chunks:
        keep_sold = pd.concat(
            keep_sold_chunks,
            axis=0
        ).drop_duplicates()
    else:
        keep_sold = sold.iloc[0:0].copy()

    # Remove temporary helper field
    if "_SPECIAL_SYMBOL_GROUP" in keep_sold.columns:
        keep_sold = keep_sold.drop(
            columns=["_SPECIAL_SYMBOL_GROUP"]
        )

    # Sort newest Sold Date first
    if not keep_sold.empty and "Sold Date" in keep_sold.columns:
        keep_sold["_SORT_SOLD_DATE"] = pd.to_datetime(
            keep_sold["Sold Date"],
            errors="coerce"
        )

        keep_sold = keep_sold.sort_values(
            "_SORT_SOLD_DATE",
            ascending=False,
            na_position="last"
        )

        keep_sold = keep_sold.drop(
            columns=["_SORT_SOLD_DATE"]
        )

    # Keep only latest 5 Sold stones
    keep_sold = keep_sold.head(5)

    return pd.concat(
        [non_sold, keep_sold],
        axis=0,
        ignore_index=False
    )


def find_carat_range(carat: float,shape: str,fluor=None,luster=None,tinch=None,black=None,key_to_symbol=None):
   
    # Normalize inputs
    shape_u = str(shape).strip().upper() if shape is not None else ""
    luster_u = str(luster).strip().upper() if luster is not None else ""
    tinch_u = str(tinch).strip().upper().replace(" ", "") if tinch is not None else ""
    black_u = str(black).strip().upper() if black is not None else ""
    fluor_u = str(fluor).strip().upper() if fluor is not None else ""
    key_to_symbol_u = str(key_to_symbol).strip().upper() if key_to_symbol is not None else ""

    # True if KEY TO SYMBOLS contains one of the special symbols
    has_special_symbol = (
        get_special_symbol_group(key_to_symbol_u, fluor_u) != "NORMAL"
    )

    uniform = [
        (0.18, 0.22), (0.23, 0.29), (0.30, 0.39), (0.40, 0.49),
        (0.50, 0.59), (0.60, 0.69), (0.70, 0.79), (0.80, 0.89),
        (0.90, 0.99), (1.00, 1.19), (1.20, 1.29), (1.30, 1.39),
        (1.40, 1.49), (1.50, 1.59), (1.60, 1.69), (1.70, 1.79),
        (1.80, 1.89), (1.90, 1.99), (2.00, 2.19), (2.20, 2.49),
        (2.50, 2.99), (3.00, 3.49), (3.50, 3.99), (4.00, 4.49),
        (4.50, 4.99), (5.00, 5.49), (5.50, 5.99)
    ]

    fine = [
        (0.18, 0.19), (0.20, 0.22), (0.23, 0.25), (0.26, 0.29),
        (0.30, 0.32), (0.33, 0.35), (0.36, 0.39), (0.40, 0.42),
        (0.43, 0.45), (0.46, 0.49), (0.50, 0.52), (0.53, 0.55),
        (0.56, 0.59), (0.60, 0.69), (0.70, 0.79), (0.80, 0.89),
        (0.90, 0.94), (0.95, 0.99), (1.00, 1.19),
        (1.20, 1.29), (1.30, 1.39), (1.40, 1.49), (1.50, 1.59),
        (1.60, 1.69), (1.70, 1.79), (1.80, 1.89), (1.90, 1.99),
        (2.00, 2.19), (2.20, 2.49), (2.50, 2.99), (3.00, 3.49),
        (3.50, 3.99), (4.00, 4.49), (4.50, 4.99),
        (5.00, 5.49), (5.50, 5.99)
    ]

    # Non-round always uses uniform
    if shape_u and shape_u != "ROUND":
        ranges = uniform

    else:
        special_checks = [
            fluor_u in {"MED", "VST", "STG"},
            luster_u in {"M-1", "M1", "M-2", "M2"},
            tinch_u in {"B-1", "B1", "B-2", "B2"},
            "H" in black_u,
            has_special_symbol,
        ]

        ranges = uniform if any(special_checks) else fine

    rapsize_map = {}

    c = float(carat) if pd.notna(carat) else 0.0

    for start, end in ranges:
        if start <= c <= end:
            rap = rapsize_map.get((start, end))
            return f"{start:.2f}-{end:.2f}", rap

    return None, None
def calculate_disc(row):
    # If the stone source is "Stock",
    # return the already-existing DISC value as-is
    if row.get("SOURCE") == "Stock":
        return row.get("DISC", None)

    # Get RAP price; default to 0 if missing or None
    rap = row.get("RAPPRICE", 0) or 0

    # Get final online price; default to 0 if missing or None
    final = row.get("ONLINE PRICE", 0) or 0

    # Calculate discount percentage:
    # ((Final Price - RAP Price) / RAP Price) * 100
    # - Rounded to 2 decimal places
    # - If RAP price is 0, return 0 to avoid division error
    return round(((final - rap) / rap) * 100, 2) if rap else 0

def assign_gen_from_stock_disc(stock: pd.DataFrame) -> pd.DataFrame:
    # Work on a copy to avoid mutating the original DataFrame
    df = stock.copy()

    # Normalize SHAPE values for reliable comparisons
    shape_norm = (df["SHAPE"].fillna("").astype(str).str.strip().str.upper())
         
    # Normalize STATUS values for case-insensitive matching
    status_clean = (df["STATUS"].fillna("").astype(str).str.strip().str.lower())

    # Common grouping keys used for DISC matching
    # Includes Carat Range to preserve pricing tiers

    match_keys = ["Carat Range", "SHAPE", "CERT", "COLOR", "CLARITY","CUT", "FLUOR", "LUSTER", "TINCH", "BLACK"]
        
    # Boolean mask for ROUND stones
    is_round = shape_norm.eq("ROUND")
    # ROUND: identify eligible STOCK stones for mapping
    eligible_stock_round = (
        (df["SOURCE"] == "Stock")                              # Must be stock stones
        & (~df["RAPSIZE"].isin(["1.00UP", "1.50UP", "2.00UP","3.00UP", "4.00UP", "5.00UP"])) # Exclude large RAP slabs                       
        & df["STATUS"].notna()                                 # STATUS must exist
        & (~status_clean.isin(["new arrival", "hold-not to see"])) # Exclude NA & holds
        & is_round                                             # Only ROUND stones
    )

    # Build DISC lookup table for ROUND stones
    disc_map_round = (
        df.loc[eligible_stock_round]                # Filter eligible stock rows
        .dropna(subset=["DISC"])                    # Require DISC to exist
        .groupby(match_keys, dropna=False)["DISC"]  # Group by full match key
        .max()                                      # Use maximum DISC per group
        .reset_index()                              # Convert index to columns
        .rename(columns={"DISC": "_DISC_MATCH"})    # Rename for merge safety
    )

    # Identify ROUND New Arrival stones
    mask_new_round = is_round & status_clean.eq("new arrival")

    # Only proceed if there are matching rows
    if mask_new_round.any():
        # Merge New Arrival stones with ROUND DISC map
        merged_round = (df.loc[mask_new_round, match_keys].merge(disc_map_round, on=match_keys, how="left"))
         
        # Assign matched DISC to GEN
        # - Convert to numeric
        # - Replace NaN with 0
        # - Align via numpy array to preserve row order
        df.loc[mask_new_round, "GEN"] = (pd.to_numeric(merged_round["_DISC_MATCH"], errors="coerce").fillna(0).to_numpy())
          
    # NON-ROUND stones (Fancy shapes)
    
    # NON-ROUND

    is_non_round = ~is_round

    bad_grade = df["POL"].isin(["GD", "FR"]) | df["SYM"].isin(["GD", "FR"])
    # NULLs will be <NA> here; treat them as NOT bad
    bad_grade = bad_grade.fillna(False)

    eligible_stock_non = (
        (df["SOURCE"] == "Stock")
        & (~df["RAPSIZE"].isin(["90UP", "1.00UP", "1.50UP", "2.00UP", "3.00UP", "4.00UP", "5.00UP"]))
        & (df["STATUS"].notna())
        & (~status_clean.isin(["new arrival", "hold-not to see"]))
        & (is_non_round)
        & (~bad_grade)   # ✅ this is the “if GD/FR then NON” logic
    )

    # Build DISC lookup table for NON-ROUND stones
    disc_map_non = (
        df.loc[eligible_stock_non]
        .dropna(subset=["DISC"])
        .groupby(match_keys, dropna=False)["DISC"]
        .max()
        .reset_index()
        .rename(columns={"DISC": "_DISC_MATCH"})
    )

    # Identify NON-ROUND New Arrival stones
    mask_new_non = is_non_round & status_clean.eq("new arrival")

    # Proceed only if such rows exist
    if mask_new_non.any():
        # Merge New Arrival fancy stones with DISC map
        merged_non = ( df.loc[mask_new_non, match_keys].merge(disc_map_non, on=match_keys, how="left"))
          
        # Assign matched DISC to GEN
        df.loc[mask_new_non, "GEN"] = (pd.to_numeric(merged_non["_DISC_MATCH"], errors="coerce").fillna(0).to_numpy())
      
    # Return the updated DataFrame
    return df

def get_code(row):

    # Only compute CODE for rows with STATUS = "New Arrival"
    if not (isinstance(row.get("STATUS"), str)and row["STATUS"].strip().lower() == "new arrival"):
        return ""

    # If CLARITY is missing or blank, CODE must be empty
    if pd.isnull(row["CLARITY"]) or str(row["CLARITY"]).strip() == "":
        return ""

    # Normalize CLARITY to uppercase
    clarity = str(row["CLARITY"]).upper()

    # Normalize SHAPE to uppercase (handle missing values safely)
    shape = str(row["SHAPE"]).upper() if pd.notnull(row["SHAPE"]) else ""

    # Safely parse CARAT as float
    try:
        carat = float(row["CARAT"])
    except (ValueError, TypeError):
        # If CARAT is invalid, do not assign any CODE
        return ""

    # Define business rule flags
    # PZ applies when:
    # - Stone is not ROUND, OR
    # - Carat is 1.00 or higher, OR
    # - Clarity is I1
    flag_PZ = (shape != "ROUND" or carat >= 1.00 or clarity == "I1")

    # SP applies to high-clarity stones
    flag_SP = (clarity in ("IF", "VVS1", "VVS2", "VS1"))

    # SK applies when:
    # - Carat is 1.00 or higher, OR
    # - Mid/low clarity stones in the 0.23–0.99 ct range
    flag_SK = (carat >= 1.00 or (clarity in ("VS2", "SI1", "SI2", "I1") and 0.23 <= carat <= 0.99))
      
    # Combine applicable codes
    codes = []

    # Add PZ code if condition is met
    if flag_PZ:
        codes.append("PZ")

    # Add SK code if condition is met
    if flag_SK:
        codes.append("SK")

    # Add SP code if condition is met
    if flag_SP:
        codes.append("SP")

    # Return slash-separated CODE string, or blank if none apply
    return "/".join(codes) if codes else ""

def apply_styles_excel(ws, results):
    # Fills
    fill_cut_3ex = PatternFill(start_color="33CCCC", end_color="33CCCC", fill_type="solid")  # 3EX
    fill_vg      = PatternFill(start_color="FF66FF", end_color="FF66FF", fill_type="solid")  # VG
    fill_ex      = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")  # EX
    fill_fancy   = PatternFill(start_color="FFC971", end_color="FFC971", fill_type="solid")  # Fancy
    fill_gd      = PatternFill(start_color="E6D5F3", end_color="E6D5F3", fill_type="solid")  # GD

    # Header -> column index mapping
    header = [cell.value for cell in ws[1]]
    col_indexes = {col: idx + 1 for idx, col in enumerate(header) if col is not None}

    cell_styles = {}

    for i in range(1, results.shape[0] + 1):
        row = results.iloc[i - 1]

        # Only style Stock rows, excluding New Arrival
        try:
            if row.get("SOURCE") != "Stock":
                continue
            if str(row.get("STATUS", "")).strip().lower() == "new arrival":
                continue
        except Exception:
            continue

        # Shape logic
        shape_val = str(row.get("SHAPE", "")).strip().upper()
        is_fancy = shape_val != "ROUND"

        # Columns to highlight
        column_highlight = ["POL", "SYM"] if is_fancy else ["CUT", "POL", "SYM"]

        # CUT-based coloring
        cut_val = str(row.get("CUT", "")).strip().upper()
        if cut_val == "3EX":
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_cut_3ex
        elif cut_val == "EX":
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_ex
        elif cut_val == "VG":
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_vg

        # Fancy rules: POL EX/VG => fancy fill
        pol_val = str(row.get("POL", "")).strip().upper()
        if is_fancy and pol_val in ["EX", "VG"]:
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_fancy

        # GD rules (any of CUT/POL/SYM contains GD)
        cut_pol_sym = [str(row.get(col, "")).strip().upper() for col in ["CUT", "POL", "SYM"]]
        if "GD" in cut_pol_sym:
            for col in column_highlight:
                if col in col_indexes:
                    cell_styles[(i + 1, col_indexes[col])] = fill_gd

    # Apply styles
    for (row_idx, col_idx), fill in cell_styles.items():
        ws.cell(row=row_idx, column=col_idx).fill = fill

def apply_styles_to_excel(ws, df: pd.DataFrame) -> None:
   
    # Define fills
    sold_fill     = PatternFill(start_color="CFEAF7", end_color="CFEAF7", fill_type="solid")
    new_fill      = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    combined_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    od_fill       = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    red_fill      = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    green_fill    = PatternFill(start_color="66FFCC", end_color="66FFCC", fill_type="solid")
    pink_fill     = PatternFill(start_color="ff8ecb", end_color="66FFCC", fill_type="solid")
    trend_fill     = PatternFill(start_color="99e6ff", end_color="99e6ff", fill_type="solid")
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
        shape = str(row.get("SHAPE", "")).strip().upper()
        luster = str(row.get("LUSTER", "")).strip().upper()
        report_comm = str(row.get("REPORT COMM", "") or "")
        key_symbol = str(row.get("KEY TO SYMBOLS", "") or "")
        black = str(row.get("BLACK", "") or "")
        stockid = str(row.get("STOCKID", "") or "")
        depth = row.get("DEPTH")
        table = row.get("TABLE")
        ratio = row.get("RATIO")
        trend = row.get("Trend")
       
        # Conditions dictionary
        checks = {
            "STATUS": status in {"new arrival", "hold-not to see"},
            "POL": pol == "GD",
            "SYM": sym == "GD",
            "CLARITY": clarity == "I1",
            "CUT": cut == "GD",
            "CERT": cert in {"HRD", "IGI"},
            "TINCH": tinch in {"B-1", "MIX"},
            "FLUOR": fluor in {"FNT", "VST", "MED", "SL", "STG", "VSL"},
            "SHAPE": shape in {
                "OVAL", "PEAR", "EMERALD", "MARQUISE", "PRINCESS",
                "CUSHION BRILLIANT", "CUSHION CUT", "CUSHION MODIFIED",
                "RADIANT", "HEART", "SQUARE RADIANT", "SQ EMERALD", "SQ. EMERALD"
            },
            "LUSTER": luster in {'LOW',"M-1", "M1", "M-2", "M2"},
            "REPORT COMM": "Natural" in report_comm or "Clarity grade is based on clouds that are not shown." in report_comm,
            "KEY TO SYMBOLS": (
                ("INTERNAL GRAINING" in key_symbol.upper() and str(clarity).upper() in ["IF", "FL", "VVS1", "VVS2"])
                or any(
                    x in key_symbol.upper()
                    for x in [
                        "BRUISE", "CAVITY", "GRAIN CENTER", "ETCH CHANNEL", "CHIP",
                        "INDENTED NATURAL", "NEEDLE- NATURAL",
                        "MANUFACTURING REMNANT", "SURFACE GRAINING"
                    ]
                )
            ),
            "BLACK": "H" in black,
            "Trend": pd.notna(trend),
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
        elif shape == "CUSHION MODIFIED" and pd.notna(depth) and depth >= 70:
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
            
            elif col == "Trend":
                fill = trend_fill
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

def get_allowed_fluor(fluor) -> list:
    fluor_hierarchy = {
        "NONE": ["NONE"],
        "FNT": ["NONE", "FNT"],
        "MED": ["NONE", "FNT", "MED"],
        "SL": ["NONE", "FNT", "MED", "SL"],
        "STG": ["NONE", "FNT", "MED", "SL", "STG"],
        "VST": ["NONE", "FNT", "MED", "SL", "STG", "VST"],
        "VSL": ["NONE", "FNT", "MED", "SL", "STG", "VST", "VSL"],
    }

    if pd.isna(fluor):
        return [pd.NA]

    fluor_u = str(fluor).strip().upper()
    # If unknown fluor grade, allow all (so we don't kill matching)
    return fluor_hierarchy.get(fluor_u, list(fluor_hierarchy.keys()))

def get_allowed_cut(cut) -> list:
    cut_hierarchy = {
        "3EX": ["3EX"],
        "EX": ["3EX", "EX"],
        "VG": ["3EX", "EX", "VG"],
        "GD": ["3EX", "EX", "VG", "GD"],
        "FR": ["3EX", "EX", "VG", "GD", "FR"],
    }

    if pd.isna(cut):
        return [pd.NA]

    cut_u = str(cut).strip().upper()
    # If unknown cut, allow all
    return cut_hierarchy.get(cut_u,list(cut_hierarchy.keys()))

def get_allowed_luster(luster) -> list:
    luster_hierarchy = {
        "EX": ["EX"],
        "M-1": ["M-1", "M-2"],
        "M-2": ["M-2", "M-1"],
    }

    if pd.isna(luster):
        return [pd.NA]

    lustre_u = str(luster).strip().upper()
    # If unknown cut, allow all
    return luster_hierarchy.get(lustre_u,list(luster_hierarchy.keys()))

def equal(col, val):
    # If the comparison value is NaN,
    # treat it as a wildcard and match all rows
    if pd.isna(val):
        return pd.Series(True, index=col.index)

    # Otherwise, perform element-wise equality comparison
    return col.eq(val)

def value_in(col: pd.Series, values: list[str]) -> pd.Series | bool:
    # If no values are provided,
    # treat it as no filtering condition (match everything)
    if not values:
        return True

    # Otherwise, return a boolean Series indicating membership
    return col.isin(values)

def order_group_values(stock_n: pd.DataFrame,group_idx: pd.Index,na_row: pd.Series,status_order: dict[str, int]) -> pd.Index:

    # Extract New Arrival reference attributes
    na_cut = na_row.get("CUT")      # Reference CUT value
    na_pol = na_row.get("POL")      # Reference POLISH value
    na_sym = na_row.get("SYM")      # Reference SYMMETRY value
    na_fluor = na_row.get("FLUOR")  # Reference FLUORESCENCE value

    # Build a temporary working DataFrame
    # Only includes rows from the current group
    tmp = stock_n.loc[group_idx,["CUT", "POL", "SYM", "FLUOR", "STATUS"]].copy()
    
    # Preserve original index order for final output
    tmp["_i"] = group_idx

    # Compute STATUS ranking (lower = higher priority)
    # Default rank = 5 if status not found
    status_rank = tmp["STATUS"].map(lambda x: status_order.get(x, 5))

    # Identify rows matching CUT / POL / SYM
    # (NaN values act as wildcards via equal())
    cps_mask = (
        equal(tmp["CUT"], na_cut)
        & equal(tmp["POL"], na_pol)
        & equal(tmp["SYM"], na_sym)
    )

    # Rows matching CUT/POL/SYM
    cps = tmp[cps_mask]

    # Remaining rows
    rest = tmp[~cps_mask]

    # Within CPS matches, further split by FLUOR
    fluor_match = equal(cps["FLUOR"], na_fluor)
    # equal() may return a single boolean if na_fluor is NaN
    # Convert it to a Series for consistent indexing
    if isinstance(fluor_match, bool):
        fluor_match = pd.Series([fluor_match] * len(cps),index=cps.index)
     
    # Exact FLUOR matches (highest priority)
    cps_fluor_exact = cps[fluor_match]

    # CPS rows with non-matching FLUOR
    # Sorted stably by FLUOR value
    cps_fluor_rest = cps[~fluor_match].sort_values( by=["FLUOR"],ascending=True,kind="stable")
     
    # Sort remaining non-CPS rows
    remaining = rest.sort_values(by=["CUT", "POL", "SYM", "FLUOR"],ascending=True,kind="stable")
 
    # Combine all groups in priority order
    ordered_idx = pd.Index(
        list(cps_fluor_exact["_i"])
        + list(cps_fluor_rest["_i"])
        + list(remaining["_i"])
    )

    # Final stable sort by STATUS rank
    # Ensures priority statuses (e.g. New Arrival)
    # float to the top when otherwise tied
    sr = status_rank.reindex(ordered_idx)

    ordered_idx = ordered_idx[np.argsort(sr.to_numpy(), kind="stable")]
    # Return the reordered index
    return ordered_idx

def same_value(x, val) -> bool:
    # If the first value is NaN,
    # return True only if the comparison value is also NaN
    if pd.isna(x):
        return pd.isna(val)
    # Otherwise, compare both values as normalized strings
    # (strip whitespace and convert to uppercase)
    return str(x).strip().upper() == str(val).strip().upper()


def sort_carat(df: pd.DataFrame) -> float:
  
    # Sort key function:
    # Returns the minimum CARAT value in the group
    # Used for ordering carat-range groups

    # Convert CARAT column to numeric, coercing invalid values to NaN
    s = pd.to_numeric(df["CARAT"], errors="coerce")

    # Get the minimum carat value in the group
    v = s.min()

    # Return the value if valid; otherwise return infinity
    # (so empty/invalid groups sort last)
    return float(v) if pd.notna(v) else float("inf")


def add_blank_row(s: pd.Series) -> pd.Series:
    # Normalize a Series for filtering:
    # - Convert to pandas string dtype
    # - Strip whitespace
    # - Convert to uppercase
    s = s.astype("string").str.strip().str.upper()

    # Replace empty strings with NA
    # This allows blank values to behave like "no filter"
    return s.replace("", pd.NA)

def other_sort(df: pd.DataFrame) -> int:
    
    SHAPE_ORDER = ["EMERALD", "PRINCESS", "OVAL", "PEAR", "MARQUIES", "HEART","MARQUISE","RADIANT","CUSHION CUT"]
    # --- Shape (Primary Sort) ---
    shape_cat = pd.Categorical(df["SHAPE"],categories=SHAPE_ORDER,ordered=True)

    shape_min = shape_cat.min()
    shape_rank = (SHAPE_ORDER.index(shape_min)
        if pd.notna(shape_min)
        else len(SHAPE_ORDER)
    )
    # --- Carat (Secondary Sort) ---
    carat_series = pd.to_numeric(df["CARAT"], errors="coerce")
    carat_min = carat_series.min()
    carat_value = float(carat_min) if pd.notna(carat_min) else float("inf")

    color_series = df["COLOR"].astype(str)
    color_min = color_series.min()
    
    clarity_series = df["CLARITY"].astype(str)
    clarity_min = clarity_series.min()
    
    # Shape first, then Carat
    return (shape_rank, carat_value, color_min, clarity_min)

def group(stock: pd.DataFrame) -> pd.DataFrame:

    # Fix column naming issue: "FLUOR " → "FLUOR"
    if "FLUOR " in stock.columns and "FLUOR" not in stock.columns:
        stock = stock.rename(columns={"FLUOR ": "FLUOR"})

    # Reset index to ensure clean, sequential indexing
    stock = stock.reset_index(drop=True)

    # Work on a normalized copy for matching logic
    stock_n = stock.copy()

    # ---------------------------------------------------------
    # SPECIAL KEY TO SYMBOLS GROUP
    # ---------------------------------------------------------
    # Stones containing CHIP, CAVITY, or MANUFACTURING REMNANT
    # must form separate groups from NORMAL stones.
    #
    # Example:
    #   Manufacturing Remnant -> MANUFACTURING REMNANT group
    #   Feather               -> NORMAL group
    #
    # The helper also accepts the older misspelling:
    #   MANUFACTRING REMANANT
    if "KEY TO SYMBOLS" in stock_n.columns:
        stock_n["_SPECIAL_SYMBOL_GROUP"] = stock_n.apply(
            lambda row: get_special_symbol_group(
                row.get("KEY TO SYMBOLS"),
                row.get("FLUOR")
            ),
            axis=1
        )
    else:
        stock_n["_SPECIAL_SYMBOL_GROUP"] = "NORMAL"

    # Columns to normalize for consistent comparisons
    norm_cols = ["SHAPE", "Carat Range", "CERT", "COLOR", "CLARITY","LUSTER", "TINCH", "CUT", "POL", "SYM", "FLUOR", "STATUS"]
      
    # Normalize each column:
    # - uppercase
    # - strip whitespace
    # - convert blanks to NA
    for c in norm_cols:
        if c in stock_n.columns:
            stock_n[c] = add_blank_row(stock_n[c])

    # Containers for grouped results
    round_first: list[pd.DataFrame] = []  # ROUND groups first
    others: list[pd.DataFrame] = []       # Non-round groups
    used_idx: set[int] = set()             # Track indices already grouped

    # Extract all NEW ARRIVAL rows
    new_arrivals_all = stock_n[stock_n["STATUS"] == "NEW ARRIVAL"]

    # Ensure Series-safe comparison for FLUOR == "NONE"
    new_arrivals_none = new_arrivals_all[equal(new_arrivals_all["FLUOR"], "NONE")]
    new_arrivals_rest = new_arrivals_all[~equal(new_arrivals_all["FLUOR"], "NONE")]

    # Process NONE-fluor new arrivals first
    new_arrivals = pd.concat([new_arrivals_none, new_arrivals_rest], axis=0)
      
    # STATUS priority mapping (lower = higher priority)
    status_order = {"NEW ARRIVAL": 0}

    # Iterate through each NEW ARRIVAL stone
    for idx, na_row in new_arrivals.iterrows():
        # Skip if this row was already grouped
        if idx in used_idx:
            continue

        # Extract reference values
        na_fluor = na_row.get("FLUOR")
        na_luster = na_row.get("LUSTER")
        na_cut = na_row.get("CUT")

        # Determine allowed CUT values
        if same_value(na_cut, "3EX"):
            # 3EX only matches 3EX
            allowed_cuts = ["3EX"]
        else:
            # Expand allowed cuts based on business rules
            allowed_cuts = [
                str(x).strip().upper()
                for x in get_allowed_cut(na_cut)
                if pd.notna(x) and str(x).strip().upper() != "3EX"
            ]

        # Determine allowed FLUOR values
        if same_value(na_fluor, "NONE"):
            # NONE only matches NONE
            allowed_fluors = ["NONE"]
        else:
            # Expand allowed fluorescence options
            allowed_fluors = [
                str(x).strip().upper()
                for x in get_allowed_fluor(na_fluor)
                if pd.notna(x)
            ]
        
        # Determine allowed FLUOR values
        if same_value(na_luster, "EX"):
            # NONE only matches NONE
            allowed_luster = ["EX"]
        else:
            # Expand allowed fluorescence options
            allowed_luster = [
                str(x).strip().upper()
                for x in get_allowed_luster(na_luster)
                if pd.notna(x)
            ]

        # Build base matching mask for grouping
        base_mask = (
            equal(stock_n["SHAPE"], na_row.get("SHAPE"))
            & equal(stock_n["Carat Range"], na_row.get("Carat Range"))
            & equal(stock_n["CERT"], na_row.get("CERT"))
            & equal(stock_n["COLOR"], na_row.get("COLOR"))
            & equal(stock_n["CLARITY"], na_row.get("CLARITY"))
            & value_in(stock_n["LUSTER"], allowed_luster)
            & equal(stock_n["TINCH"], na_row.get("TINCH"))
            & value_in(stock_n["CUT"], allowed_cuts)
            & value_in(stock_n["FLUOR"], allowed_fluors)

            # IMPORTANT:
            # Keep special KEY TO SYMBOLS in separate groups.
            & equal(
                stock_n["_SPECIAL_SYMBOL_GROUP"],
                na_row.get("_SPECIAL_SYMBOL_GROUP")
            )
        )

        # Collect indices matching the base mask
        group_idx = stock_n.index[base_mask]

        # Exclude rows already assigned to another group
        if used_idx:group_idx = group_idx[~group_idx.isin(list(used_idx))]
            
        # Skip if no rows remain
        if group_idx.empty:
            continue

        # Ensure at least one NEW ARRIVAL exists in the group
        if not stock_n.loc[group_idx, "STATUS"].eq("NEW ARRIVAL").any():
            continue

        # Order rows inside the group
        ordered_idx = order_group_values(stock_n,group_idx,na_row,status_order)
            
        # Retrieve original (non-normalized) rows
        result = stock.loc[ordered_idx].copy()

        # Apply sold/new-arrival filtering rules
        result = filter_sold_by_new_arrival(result)

        # Place group into ROUND or NON-ROUND bucket
        is_round = stock_n.loc[group_idx, "SHAPE"].eq("ROUND").any()
        (round_first if is_round else others).append(result)

        # Mark indices as used
        used_idx.update(group_idx.tolist())

    # Sort groups by minimum CARAT
    round_first = sorted(round_first, key=sort_carat)
    others = sorted(others, key=other_sort)

    # Combine all groups with blank row separators
    final_result: list[pd.DataFrame] = []
    for g in round_first + others:
        final_result.append(g)

        # Add a blank separator row after each group
        final_result.append(pd.DataFrame(np.nan, columns=g.columns, index=[0]))
    # Return final grouped DataFrame
    # If no groups were created, return original stock
    return pd.concat(final_result, ignore_index=True) if final_result else stock


def arrange_entire_groups(df: pd.DataFrame) -> pd.DataFrame:
    """
    Final group arrangement.

    Each group is separated by a blank row.  The complete group is moved as
    one block; rows inside the group keep exactly the same order.

    Group order is decided only from the FIRST ROW of each group, using these
    columns in priority order:
        SHAPE, RAPSIZE, CERT, COLOR, CLARITY, CUT, POL, SYM, FLUOR,
        BLACK, LUSTER, TINCH

    Values are compared case-insensitively. Blank first-row values sort after
    non-blank values. A blank separator row is kept after every group.

    Only groups with MORE THAN 5 non-SOLD rows are kept in the output.
    SOLD stones attached for reference do not count toward the group size.
    """

    arrange_columns = [
        "SHAPE", "RAPSIZE", "CERT", "COLOR", "CLARITY", "CUT",
        "POL", "SYM", "FLUOR", "BLACK", "LUSTER", "TINCH"
    ]

    if df.empty:
        return df

    # A separator row can contain NaN/None or empty strings created later
    # (for example CARAT/CODE). Treat both as blank.
    def is_blank_row(row: pd.Series) -> bool:
        return all(
            pd.isna(value) or str(value).strip() == ""
            for value in row
        )

    # Count rows in a group that are NOT SOLD.
    def non_sold_count(g: pd.DataFrame) -> int:
        if "SOURCE" not in g.columns:
            return len(g)
        src = g["SOURCE"].fillna("").astype(str).str.strip().str.upper()
        return int(src.ne("SOLD").sum())

    # Split the DataFrame back into whole groups.
    groups: list[pd.DataFrame] = []
    current_indexes: list[int] = []

    for idx, row in df.iterrows():
        if is_blank_row(row):
            if current_indexes:
                groups.append(df.loc[current_indexes].copy())
                current_indexes = []
        else:
            current_indexes.append(idx)

    # Keep a final group even if the DataFrame does not end with a blank row.
    if current_indexes:
        groups.append(df.loc[current_indexes].copy())

    # Keep only groups with MORE THAN 5 non-SOLD rows.
    groups = [g for g in groups if non_sold_count(g) > 5]

    if not groups:
        return df

    # Build the sort key ONLY from the first row of each whole group.
    def group_sort_key(group_df: pd.DataFrame):
        first_row = group_df.iloc[0]
        key = []

        for col in arrange_columns:
            value = first_row.get(col)

            if pd.isna(value) or str(value).strip() == "":
                # Blank values come after populated values.
                key.append((1, ""))
            else:
                key.append((0, str(value).strip().upper()))

        return tuple(key)

    # Python sort is stable, so completely equal first-row keys retain their
    # original relative order.
    groups = sorted(groups, key=group_sort_key)

    # Rebuild the result with one blank separator row after each group.
    arranged_parts: list[pd.DataFrame] = []
    for group_df in groups:
        arranged_parts.append(group_df)
        arranged_parts.append(
            pd.DataFrame([[np.nan] * len(df.columns)], columns=df.columns)
        )

    return pd.concat(arranged_parts, ignore_index=True)

def _find_sample(filename: str):
    """Locate a file inside a 'Sample Files' folder, searching upward from this
    script. Works on Linux (Streamlit Cloud) and Windows regardless of the
    current working directory or where this page sits in a multipage app."""
    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        candidate = parent / "Sample Files" / filename
        if candidate.exists():
            return candidate
    return None

def revision_stock_list():

    # Page title
    st.title("New Stock List")

    # Sidebar file uploads
    sample_path = _find_sample("New Stone File.xlsx")
    
    with st.sidebar:
    
        if sample_path is not None:
            with open(sample_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            href = (
                f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" '
                f'download="{sample_path.name}">Sample File</a>'
            )
            st.markdown(href, unsafe_allow_html=True)
        else:
            st.warning("Sample file not found on server.")
            
        st.markdown("### Upload Files")

        # Upload required Excel files
        stock = st.file_uploader("Stock File", type=["xlsx", "xls"])
        sold = st.file_uploader("Sold File", type=["xlsx", "xls"])
        days = st.file_uploader("Days File", type=["xlsx", "xls"])
        revsion = st.file_uploader("Revsion File", type=["xlsx", "xls"])
        promo = st.file_uploader("Promo File", type=["xlsx", "xls"])

    # Main processing block with spinner
    with st.spinner("Processing..."):

        # Ensure all required files are uploaded
        if stock is None or sold is None or days is None or revsion is None or promo is None:
            st.info("Upload Files.")
            st.stop()

        try:
            # Initialize progress bar
            progress_bar = st.progress(0)

            # Read uploaded files-
            stock = read_file(stock)
            progress_bar.progress(10)

            # Sold file has header offset
            sold = read_file(sold, skiprows=2)
            progress_bar.progress(20)

            days = read_file(days)
            progress_bar.progress(30)

            revision = read_file(revsion)
            promo = read_file(promo)
            progress_bar.progress(40)

        except Exception as e:
            # Handle file read errors
            st.error(f"Error reading files: {e}")
            st.stop()

        # Merge auxiliary data into stock
        stock = (
            stock
            # Merge revision days
            .merge(revision[["Sr. No", "DAYS"]],left_on="STOCKID",right_on="Sr. No",how="left")
            .rename(columns={"DAYS": "R D"})
            .merge(days[["Sr. No", "Stock Days", "Luster", "Trend"]],left_on="STOCKID",right_on="Sr. No",how="left")
            .rename(columns={"Stock Days": "O D"})
            .merge(promo[["STOCKID", "PDISC"]],left_on="STOCKID",right_on="STOCKID",how="left",))

        progress_bar.progress(45)

        # Normalize LUSTER values
        stock = replace_low(stock)

        # Use promo discount if available, else DISC
        stock["PDISC"] = stock["PDISC"].fillna(stock["DISC"])
        stock["DISC"] = stock["PDISC"]

        # New arrivals always start with DISC = 0
        stock.loc[stock["STATUS"].str.lower() == "new arrival", "DISC"] = 0

        # Mark source as Stock
        stock["SOURCE"] = "Stock"
        sold.rename(columns=column_mapping, inplace=True)
        if "FINAL PRICE" in sold.columns:
            sold["ONLINE PRICE"] = sold["FINAL PRICE"]

        sold["SOURCE"] = "Sold"

        progress_bar.progress(50)

        # Combine stock and sold into one DataFrame
        stock = pd.concat([stock, sold], axis=0)

        # Normalize 3EX (CUT + POL + SYM all EX)
        mask_3ex_final = (stock[["CUT", "POL", "SYM"]] == "EX").all(axis=1)
        stock.loc[mask_3ex_final, "CUT"] = "3EX"

        progress_bar.progress(60)
        # Compute Carat Range
        stock["Carat Range"] = stock.apply(
            lambda row: find_carat_range(
                row["CARAT"],
                row["SHAPE"],
                row.get("FLUOR"),
                row.get("LUSTER"),
                row.get("TINCH"),
                row.get("BLACK"),
                row.get("KEY TO SYMBOLS"),
            )[0],
            axis=1,
        )

        progress_bar.progress(70)
        # Group stock by New Arrival logic
        stock = group(stock)

        progress_bar.progress(75)
        # Recalculate DISC for sold stones only--
        sold_disc = stock["SOURCE"] != "Stock"
        stock.loc[sold_disc, "DISC"] = stock.loc[sold_disc].apply(calculate_disc,axis=1)
            
        # Sort sold dates (used indirectly in grouping)
        stock.loc[sold_disc, "Sold Date"].sort_values(ascending=False)

        # Assign GEN based on stock DISC
        stock = assign_gen_from_stock_disc(stock)

        # Generate CODE values
        stock["CODE"] = stock.apply(get_code, axis=1)

        # Format CARAT to two decimals
        stock["CARAT"] = stock["CARAT"].apply(lambda x: f"{float(x):.2f}")
        stock["CARAT"] = stock["CARAT"].replace("nan", "")

        progress_bar.progress(80)

        # Ensure all desired columns exist
        for col in desired_order:
            if col not in stock.columns:
                stock[col] = None

        # Reorder columns
        stock = stock[desired_order]

        # FINAL PROCESS: arrange complete groups by the FIRST ROW of each group.
        # The rows inside every group stay unchanged.
        # Only groups with more than 5 non-SOLD rows are kept.
        stock = arrange_entire_groups(stock)

        progress_bar.progress(90)
        # Display final DataFrame
        st.dataframe(stock)

        # Prepare Excel file for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            stock.to_excel(writer, index=False, sheet_name="Sheet1")
            ws = writer.book["Sheet1"]

            # Apply Excel styling
            apply_styles_to_excel(ws, stock)
            apply_styles_excel(ws, stock)

        output.seek(0)

        # Generate filename with current date
        file_name = f"Stock List {pd.Timestamp.now():%d.%m.%Y}.xlsx"

        progress_bar.progress(100)
        # Download button
        st.download_button(
            label="Download",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    revision_stock_list()