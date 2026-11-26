import base64
import datetime
from io import BytesIO
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles.fills import PatternFill
from openpyxl.styles.fonts import Font
import streamlit as st
import pandas as pd
import numpy as np
import io

def read_file(uploaded_file, skiprows: int = 0):
    return pd.read_excel(uploaded_file, skiprows=skiprows)

def replace_low(df: pd.DataFrame) -> pd.DataFrame:
    if "Luster" in df.columns and "LUSTER" not in df.columns:
        df["LUSTER"] = df["Luster"]

    if "LUSTER" in df.columns:
        df.loc[df["LUSTER"].astype(str).str.upper().eq("LOW"), "LUSTER"] = "LOW"

    return df

desired_order = [
    "STOCKID","SHAPE", "CERT", "CARAT", "COLOR", "CLARITY","CUT", "POL", "SYM",  
    "FLUOR","MEASUREMENT", "DEPTH", "TABLE", "DISC","BLACK","LUSTER","TINCH",
    "KEY TO SYMBOLS","REPORT COMM","R D","O D","RATIO","List Disc","Gen","Trend",
    "Premal", "Sagar","Small","NAME","Sold Disc","Sold Date","CODE","Total Stock",
    "Total Sold","STATUS", "RAPPRICE","RAPSIZE","V360","ONLINE PRICE","CERTNO"]

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

def get_special_symbol_group(value, fluor=None) -> str:
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

    return "|".join(sorted(set(found)))

def filter_sold_by_new_arrival(group: pd.DataFrame) -> pd.DataFrame:
    non_sold = group[group["SOURCE"].ne("Sold")]
    sold = group[group["SOURCE"].eq("Sold")]
    new_arrival = group[group["STATUS"].eq("New Arrival")]

    if new_arrival.empty:
        return non_sold

    def norm_series(s: pd.Series) -> pd.Series:
        return s.fillna("").astype(str).str.strip().str.upper()

    def norm_val(v) -> str:
        return str(v).strip().upper() if pd.notna(v) else ""

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

    for _, na_row in new_arrival.iterrows():
        shape = norm_val(na_row.get("SHAPE"))

        match_cols = (
            ["SHAPE", "COLOR", "CLARITY", "CUT", "POL", "SYM", "FLUOR", "LUSTER", "TINCH", "BLACK"]
            if shape == "ROUND"
            else ["SHAPE", "COLOR", "CLARITY", "POL", "SYM", "FLUOR", "LUSTER", "TINCH", "BLACK"]
        )

        mask = pd.Series(True, index=sold.index)

        for col in match_cols:
            if col not in sold.columns:
                continue
            mask &= norm_series(sold[col]).eq(norm_val(na_row.get(col)))

        if mask.any():
            keep_sold_chunks.append(sold.loc[mask])

    keep_sold = (
        pd.concat(keep_sold_chunks, axis=0).drop_duplicates()
        if keep_sold_chunks
        else sold.iloc[0:0]
    )

    if not keep_sold.empty and "Sold Date" in keep_sold.columns:
        keep_sold = keep_sold.sort_values("Sold Date", ascending=False)

    keep_sold = keep_sold.head(5)

    sold_discs = keep_sold.get("Sold Disc", pd.Series(dtype="object")).dropna().astype(str).unique().tolist()
    names     = keep_sold.get("NAME",      pd.Series(dtype="object")).dropna().astype(str).unique().tolist()
    dates     = keep_sold.get("Sold Date", pd.Series(dtype="object")).dropna().astype(str).unique().tolist()

    # Update ALL New Arrival rows that exist in non_sold
    na_mask = group["STATUS"].astype(str).str.strip().str.lower().eq("new arrival")
    na_idxs = group.index[na_mask]

    non_sold = non_sold.copy()

    for col in ["Sold Disc", "NAME", "Sold Date"]:
        if col in non_sold.columns:
            non_sold[col] = non_sold[col].astype("string")

    for idx in na_idxs:
        if idx in non_sold.index:
            if sold_discs:
                non_sold.at[idx, "Sold Disc"] = "/".join(sold_discs)
            if names:
                non_sold.at[idx, "NAME"] = "/".join(names)
            if dates:
                non_sold.at[idx, "Sold Date"] = "/".join(dates)

   # Totals (must be SCALARS)
    status_is_na = non_sold["STATUS"].astype(str).str.strip().str.lower().eq("new arrival")

    total_sold  = (keep_sold["SOURCE"] == "Sold").sum()

    total_stock = ((group["SOURCE"] == "Stock")& (~group["STATUS"].astype(str).str.strip().str.lower().eq("new arrival"))).sum()
    
    non_sold["Total Stock"] = np.where(status_is_na, total_stock, None)
    non_sold["Total Sold"]  = np.where(status_is_na, total_sold, None)

    return pd.concat([non_sold, keep_sold], axis=0, ignore_index=False)

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
    if row.get("SOURCE") == "Stock":
        return row.get("DISC", None)
    rap = row.get("RAPPRICE", 0) or 0
    final = row.get("ONLINE PRICE", 0) or 0
    return round(((final - rap) / rap) * 100, 2) if rap else 0


def assign_gen_from_stock_disc(stock: pd.DataFrame) -> pd.DataFrame:

    df = stock.copy()

    # Normalize helpers
    shape_norm = df["SHAPE"].fillna("").astype(str).str.strip().str.upper()
    status_clean = df["STATUS"].fillna("").astype(str).str.strip().str.lower()

    # Common match keys (include Carat Range to preserve your group restriction)
    match_keys = ["Carat Range","SHAPE","CERT","COLOR","CLARITY","CUT","FLUOR","LUSTER","TINCH","BLACK"]
  
    is_round = shape_norm.eq("ROUND")
    eligible_stock_round = (
        (df["SOURCE"] == "Stock")
        & (~df["RAPSIZE"].isin(["1.00UP", "1.50UP", "2.00UP", "3.00UP", "4.00UP", "5.00UP"]))
        & df["STATUS"].notna()
        & (~status_clean.isin(["new arrival", "hold-not to see"]))
        & is_round
    )

    disc_map_round = (
        df.loc[eligible_stock_round]
        .dropna(subset=["DISC"])
        .groupby(match_keys, dropna=False)["DISC"]
        .max()
        .reset_index()
        .rename(columns={"DISC": "_DISC_MATCH"})
    )

    mask_new_round = is_round & status_clean.eq("new arrival")
    if mask_new_round.any():
        merged_round = df.loc[mask_new_round, match_keys].merge(
            disc_map_round, on=match_keys, how="left"
        )
        df.loc[mask_new_round, "Gen"] = (
            pd.to_numeric(merged_round["_DISC_MATCH"], errors="coerce").to_numpy()
        )

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

    disc_map_non = (
        df.loc[eligible_stock_non]
        .dropna(subset=["DISC"])
        .groupby(match_keys, dropna=False)["DISC"]
        .max()
        .reset_index()
        .rename(columns={"DISC": "_DISC_MATCH"})
    )

    mask_new_non = is_non_round & status_clean.eq("new arrival")
    if mask_new_non.any():
        merged_non = df.loc[mask_new_non, match_keys].merge(
            disc_map_non, on=match_keys, how="left" )
       
        df.loc[mask_new_non, "Gen"] = (
            pd.to_numeric(merged_non["_DISC_MATCH"], errors="coerce").to_numpy())
        
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
    small_fill    = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")
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
        shape = str(row.get("SHAPE", "")).strip()
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
            
            elif col == "Trend":
                fill = trend_fill
                font = black_font

            # --- Default red cases ---
            else:
                fill = red_fill
                font = red_font if col != "REPORT COMM" else black_font

            cell_styles[(excel_row, col_index[col])] = (fill, font)

    if "Small" in col_index:
        col_idx = col_index["Small"]
        small_col = df["Small"]
        mask = small_col.notna() & (small_col.astype(str).str.strip() != "")
        for idx in df[mask].index:
            excel_row = idx + 2
            ws.cell(row=excel_row, column=col_idx).fill = small_fill
            ws.cell(row=excel_row, column=col_idx).font = black_font

    # Final style pass (batch)
    for (r, c), (f_fill, f_font) in cell_styles.items():
        cell = ws.cell(row=r, column=c)
        if f_fill:
            cell.fill = f_fill
        if f_font:
            cell.font = f_font

def append_sources_to_book2(premal, sagar, gen_df):
    """
    Appends Premal, Sagar, and GEN data into a 'Combined' sheet in an Excel file.
    Columns: STOCKID, Disc, Code, Date  (Date = DD-MM-YYYY)
    """
    today_str = datetime.datetime.now().strftime("%d-%m-%Y")

    premal_df = (
        premal[["stockid", "pz"]].copy()
        if "pz" in premal.columns
        else pd.DataFrame(columns=["stockid", "pz"])
    )
    premal_df.columns = ["STOCKID", "Disc"]
    premal_df["Code"] = "PZ"

    sagar_df = (
        sagar[["stockid", "sp"]].copy()
        if "sp" in sagar.columns
        else pd.DataFrame(columns=["stockid", "sp"])
    )
    sagar_df.columns = ["STOCKID", "Disc"]
    sagar_df["Code"] = "SP"

    gen_part = (
        gen_df[["STOCKID", "gen"]].copy()
        if "gen" in gen_df.columns
        else pd.DataFrame(columns=["STOCKID", "gen"])
    )
    gen_part.columns = ["STOCKID", "Disc"]
    gen_part["Code"] = "GEN"

    combined_df = pd.concat([premal_df, sagar_df, gen_part], ignore_index=True)
    combined_df.dropna(subset=["STOCKID", "Disc"], inplace=True)
    combined_df["Date"] = today_str

    output_disc = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Combined"
    ws.append(["STOCKID", "Disc", "Code", "Date"])

    for _, row in combined_df.iterrows():
        ws.append([row["STOCKID"], row["Disc"], row["Code"], row["Date"]])

    wb.save(output_disc)
    output_disc.seek(0)
    return output_disc

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

def _norm_col(s: pd.Series) -> pd.Series:
    # normalize, and treat blanks as NA so we can "skip filtering"
    s = s.astype("string").str.strip().str.upper()
    return s.replace("", pd.NA)

def smart_eq(col, val):
    if pd.isna(val):
        return pd.Series(True, index=col.index)
    return col.eq(val)


def smart_in(col: pd.Series, values: list[str]) -> pd.Series | bool:
   
    if not values:
        return True
    return col.isin(values)

def _order_group_indices(
    stock_n: pd.DataFrame,
    group_idx: pd.Index,
    na_row: pd.Series,
    status_order: dict[str, int],
 ) -> pd.Index:
    na_cut = na_row.get("CUT")
    na_pol = na_row.get("POL")
    na_sym = na_row.get("SYM")
    na_fluor = na_row.get("FLUOR")

    tmp = stock_n.loc[group_idx, ["CUT", "POL", "SYM", "FLUOR", "STATUS"]].copy()
    tmp["_i"] = group_idx

    status_rank = tmp["STATUS"].map(lambda x: status_order.get(x, 5))

    cps_mask = (
        smart_eq(tmp["CUT"], na_cut)
        & smart_eq(tmp["POL"], na_pol)
        & smart_eq(tmp["SYM"], na_sym)
    )

    cps = tmp[cps_mask]
    rest = tmp[~cps_mask]

    fluor_match = smart_eq(cps["FLUOR"], na_fluor)
    if isinstance(fluor_match, bool):
        fluor_match = pd.Series([fluor_match] * len(cps), index=cps.index)

    cps_fluor_exact = cps[fluor_match]
    cps_fluor_rest = cps[~fluor_match].sort_values(
        by=["FLUOR"],
        ascending=True,
        kind="stable",
    )

    remaining = rest.sort_values(
        by=["CUT", "POL", "SYM", "FLUOR"],
        ascending=True,
        kind="stable",
    )

    ordered_idx = pd.Index(
        list(cps_fluor_exact["_i"])
        + list(cps_fluor_rest["_i"])
        + list(remaining["_i"])
    )

    # Stable status-rank sort (NEW ARRIVAL floats to top if tied)
    sr = status_rank.reindex(ordered_idx)
    ordered_idx = ordered_idx[np.argsort(sr.to_numpy(), kind="stable")]

    return ordered_idx

def _smart_eq_scalar(x, val) -> bool:
    if pd.isna(x):
        return pd.isna(val)
    return str(x).strip().upper() == str(val).strip().upper()

def _group_carat_key(df: pd.DataFrame) -> float:
    # Sort by the minimum CARAT in the group (common for “group range” ordering)
    s = pd.to_numeric(df["CARAT"], errors="coerce")
    v = s.min()
    return float(v) if pd.notna(v) else float("inf")

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
    return (shape_rank, carat_value,color_min,clarity_min,get_allowed_fluor)

def group(stock: pd.DataFrame) -> pd.DataFrame:
    if "FLUOR " in stock.columns and "FLUOR" not in stock.columns:
        stock = stock.rename(columns={"FLUOR ": "FLUOR"})

    stock = stock.reset_index(drop=True)
    stock_n = stock.copy()

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

    norm_cols = [
        "SHAPE", "Carat Range", "CERT", "COLOR", "CLARITY", "LUSTER", "TINCH",
        "CUT", "POL", "SYM", "FLUOR", "STATUS",
    ]
    for c in norm_cols:
        if c in stock_n.columns:
            stock_n[c] = _norm_col(stock_n[c])

    round_first: list[pd.DataFrame] = []
    others: list[pd.DataFrame] = []
    used_idx: set[int] = set()

    new_arrivals_all = stock_n[stock_n["STATUS"] == "NEW ARRIVAL"]

    # Series-safe
    new_arrivals_none = new_arrivals_all[smart_eq(new_arrivals_all["FLUOR"], "NONE")]
    new_arrivals_rest = new_arrivals_all[~smart_eq(new_arrivals_all["FLUOR"], "NONE")]
    new_arrivals = pd.concat([new_arrivals_none, new_arrivals_rest], axis=0)

    status_order = {"NEW ARRIVAL": 0}

    for idx, na_row in new_arrivals.iterrows():
        if idx in used_idx:
            continue

        na_fluor = na_row.get("FLUOR")
        na_cut = na_row.get("CUT")

        if _smart_eq_scalar(na_cut, "3EX"):
            allowed_cuts = ["3EX"]
        else:
            allowed_cuts = [
                str(x).strip().upper()
                for x in get_allowed_cut(na_cut)
                if pd.notna(x) and str(x).strip().upper() != "3EX"
            ]

        if _smart_eq_scalar(na_fluor, "NONE"):
            allowed_fluors = ["NONE"]
        else:
            allowed_fluors = [
                str(x).strip().upper()
                for x in get_allowed_fluor(na_fluor)
                if pd.notna(x)
            ]

        base_mask = (
            smart_eq(stock_n["SHAPE"], na_row.get("SHAPE"))
            & smart_eq(stock_n["Carat Range"], na_row.get("Carat Range"))
            & smart_eq(stock_n["CERT"], na_row.get("CERT"))
            & smart_eq(stock_n["COLOR"], na_row.get("COLOR"))
            & smart_eq(stock_n["CLARITY"], na_row.get("CLARITY"))
            & smart_eq(stock_n["LUSTER"], na_row.get("LUSTER"))
            & smart_eq(stock_n["TINCH"], na_row.get("TINCH"))
            & smart_in(stock_n["CUT"], allowed_cuts)
            & smart_in(stock_n["FLUOR"], allowed_fluors)
            & smart_eq(stock_n["_SPECIAL_SYMBOL_GROUP"],
                      na_row.get("_SPECIAL_SYMBOL_GROUP"))         
        )

        group_idx = stock_n.index[base_mask]

        if used_idx:
            group_idx = group_idx[~group_idx.isin(list(used_idx))]

        if group_idx.empty:
            continue

        if not stock_n.loc[group_idx, "STATUS"].eq("NEW ARRIVAL").any():
            continue

        ordered_idx = _order_group_indices(stock_n, group_idx, na_row, status_order)
        result = stock.loc[ordered_idx].copy()

        result = filter_sold_by_new_arrival(result)

        is_round = stock_n.loc[group_idx, "SHAPE"].eq("ROUND").any()
        (round_first if is_round else others).append(result)

        used_idx.update(group_idx.tolist())


    round_first = sorted(round_first, key=_group_carat_key)
    others = sorted(others, key=other_sort)

    final_result: list[pd.DataFrame] = []
    for g in round_first + others:
        final_result.append(g)
        final_result.append(pd.DataFrame(np.nan, columns=g.columns, index=[0]))

    return pd.concat(final_result, ignore_index=True) if final_result else stock

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

def run_hemil_check():
    st.title("Hemil Check File")

    sample_path = _find_sample("Hemil Check.xlsx")
    
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
        stock = st.file_uploader("Stock File", type=["xlsx", "xls"])
        sold = st.file_uploader("Sold File", type=["xlsx", "xls"])
        days = st.file_uploader("Days File", type=["xlsx", "xls"])
        revsion = st.file_uploader("Revsion File", type=["xlsx", "xls"])
        promo = st.file_uploader("Promo File", type=["xlsx", "xls"])
        premal_file = st.file_uploader("Premal  File", type=["xlsx", "xls"])
        sagar_file = st.file_uploader("Sagar  File", type=["xlsx", "xls"])
    
    with st.spinner("Processing..."):
        if stock is None or sold is None or days is None or revsion is None or promo is None:
            st.info("Upload Files.")
            st.stop()

        try:
            # Initialize progress bar
            progress_bar = st.progress(0)
            
            # Read the files with progress updates
            stock = read_file(stock)
            progress_bar.progress(10)
            sold = read_file(sold,skiprows=2)
            progress_bar.progress(20)
            days = read_file(days)
            progress_bar.progress(30)
            revision = read_file(revsion)
            promo = read_file(promo)
            progress_bar.progress(40)

            # Loading optional files if they are uploaded, else empty DataFrames
            if premal_file:
                premal = pd.read_excel(premal_file)
                premal.columns = premal.columns.str.lower()
            else:
                premal = pd.DataFrame(columns=["stockid", "pz"])

            if sagar_file:
                sagar = pd.read_excel(sagar_file)
                sagar.columns = sagar.columns.str.lower()
            else:
                sagar = pd.DataFrame(columns=["stockid", "sp"])

        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()

        stock = (
            stock
            .merge(revision[["Sr. No", "DAYS"]],left_on="STOCKID",right_on="Sr. No", how="left",)
            .rename(columns={"DAYS": "R D"})   
            .merge(days[["Sr. No", "Stock Days", "Luster", "DISC","Trend"]], left_on="STOCKID", right_on="Sr. No", how="left")
            .rename(columns={"Stock Days": "O D", "DISC_y": "List Disc", "DISC_x": "DISC"})
            .merge(promo[["STOCKID","PDISC"]],left_on="STOCKID",right_on="STOCKID",how="left",)
            .merge(premal[['stockid', 'pz']], left_on='STOCKID', right_on='stockid', how='left')
            .rename(columns={'pz': 'Premal'})
            .merge(sagar[['stockid', 'sp']], left_on='STOCKID', right_on='stockid', how='left')
            .rename(columns={'sp': 'Sagar'})
            )
        
        progress_bar.progress(45)
        stock = replace_low(stock)
        stock["PDISC"] = stock["PDISC"].fillna(stock["DISC"])
        stock["DISC"] = stock["PDISC"]

        sold.rename(columns=column_mapping, inplace=True)
        sold["SOURCE"] = "Sold"
        stock["SOURCE"] = "Stock"

        if "FINAL PRICE" in sold.columns:
            sold["FINAL PRICE"] = sold["ONLINE PRICE"]

        progress_bar.progress(50)
        stock = pd.concat([stock, sold],axis=0) 

        sold_disc = stock["SOURCE"] != "Stock"
        stock.loc[sold_disc, "Sold Disc"] = stock.loc[sold_disc].apply(calculate_disc, axis=1)

        mask_3ex_final = (stock[["CUT", "POL", "SYM"]] == "EX").all(axis=1)
        stock.loc[mask_3ex_final, "CUT"] = "3EX"
        progress_bar.progress(60)

        stock['Carat Range'] = stock.apply(
            lambda row: find_carat_range(row['CARAT'], row['SHAPE'],row['KEY TO SYMBOLS'], row.get('FLUOR'), row.get('LUSTER'), row.get('TINCH'), row.get('BLACK'))[0],
            axis=1
        )
        progress_bar.progress(70) 
        stock = group(stock)
        
        progress_bar.progress(75)
        
        stock = assign_gen_from_stock_disc(stock)

        stock["CODE"] = stock.apply(get_code, axis=1)

        stock['Small'] = stock[['Gen', 'Premal',  'Sagar']].min(axis=1)
        disc_numeric = stock['DISC']
        small_numeric = stock['Small']
        mask = (stock["STATUS"].str.lower() != "new arrival") & (disc_numeric == small_numeric)
        stock.loc[mask, 'Small'] = np.nan

        stock = stock[~((stock['STATUS'].str.strip() == 'New Arrival') & (stock['Small'].isna()))]

        premal_input = stock[["STOCKID", "Premal"]].copy()
        premal_input.rename(columns={"STOCKID": "stockid", "Premal": "pz"}, inplace=True)

        sagar_input = stock[["STOCKID", "Sagar"]].copy()
        sagar_input.rename(columns={"STOCKID": "stockid", "Sagar": "sp"}, inplace=True)

        gen_input = stock[["STOCKID", "Gen"]].copy()
        gen_input.rename(columns={"Gen": "gen"}, inplace=True)

        combined_output = append_sources_to_book2(premal_input, sagar_input, gen_input)

        progress_bar.progress(70) 
        stock = group(stock)

        progress_bar.progress(80)
        stock = stock.loc[stock["SOURCE"] != "Sold"].reset_index(drop=True)
        for col in desired_order:
            if col not in stock.columns:
                stock[col] = None
        stock = stock[desired_order]

        progress_bar.progress(90)
        st.dataframe(stock)

        # Prepare the file for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            stock.to_excel(writer, index=False, sheet_name="Sheet1")
            ws = writer.book["Sheet1"]
            apply_styles_to_excel(ws, stock)
            apply_styles_excel(ws, stock)
        output.seek(0)

        file_name = f"Hemil Check {pd.Timestamp.now():%d.%m.%Y}.xlsx"
        combined_name = f"Disc {pd.Timestamp.now():%d.%m.%Y}.xlsx"
        progress_bar.progress(95)
        # Add a download button
        st.download_button(
            label="Download",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        progress_bar.progress(100)
        # Add a download button
        st.download_button(
            label="Disc File",
            data=combined_output,
            file_name=combined_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    run_hemil_check()
