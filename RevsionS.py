import os
from pathlib import Path
from tkinter import filedialog

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


# ------------------- Carat Range + RAPSIZE (fine-only) -------------------
def find_carat_range(
    carat: float,
    shape: str,
    fluor=None,
    luster=None,
    tinch=None,
    cut=None,
):
    """
    Returns (carat_range, rap_size).
    - carat_range: formatted range string (e.g., "0.30-0.34") or "Other"
    - rap_size: RAPSIZE label only when ROUND uses fine ranges (non-special); otherwise None
    """
    # Normalize
    shape_u = str(shape).strip().upper() if shape is not None else ""
    luster_u = str(luster).strip().upper() if luster is not None else ""
    tinch_u = str(tinch).strip().upper().replace(" ", "") if tinch is not None else ""
    fluor_u = str(fluor).strip().upper() if fluor is not None else ""
    cut = str(cut).strip().upper() if cut is not None else ""

    # Uniform ranges
    uniform = [
        (0.23, 0.29), (0.30, 0.39), (0.40, 0.49), (0.50, 0.59), (0.60, 0.69),
        (0.70, 0.79), (0.80, 0.89), (0.90, 0.99), (1.00, 1.19), (1.20, 1.29),
        (1.30, 1.39), (1.40, 1.49), (1.50, 1.59), (1.60, 1.69), (1.70, 1.79),
        (1.80, 1.89), (1.90, 1.99), (2.00, 2.19), (2.20, 2.49), (2.50, 2.99),
        (3.00, 3.49), (3.50, 3.99), (4.00, 4.49), (4.50, 4.99), (5.00, 5.49),
        (5.50, 5.99),
    ]

    # Fine ranges (ROUND only, if not special)
    fine = [
        (0.23, 0.29), (0.30, 0.34), (0.35, 0.39), (0.40, 0.45), (0.46, 0.49),
        (0.50, 0.59), (0.60, 0.69), (0.70, 0.79), (0.80, 0.89), (0.90, 0.95),
        (0.96, 0.99), (1.00, 1.19), (1.20, 1.29), (1.30, 1.39), (1.40, 1.49),
        (1.50, 1.59), (1.60, 1.69), (1.70, 1.79), (1.80, 1.89), (1.90, 1.99), 
        (2.00, 2.19), (2.20, 2.49), (2.50, 2.99), (3.00, 3.49), (3.50, 3.99), 
        (4.00, 4.49),(4.50, 4.99), (5.00, 5.49), (5.50, 5.99),
    ]

    # RAPSIZE mapping (applies only to fine ranges)
    rapsize_map = {
        (0.30, 0.34): "34UP",
        (0.35, 0.39): "35UP",
        (0.40, 0.45): "45UP",
        (0.46, 0.49): "46UP",
        (0.50, 0.54): "54UP",
        (0.55, 0.59): "55UP",
        
    }

    # Decide ranges
    if shape_u and shape_u != "ROUND":
        ranges = uniform
        use_fine = False
    else:
        special_checks = [
            fluor_u in {"MED", "VST", "STG"},
            luster_u in {"M-1", "M1", "M-2", "M2"},
            tinch_u in {"B-1", "B1", "B-2", "B2","MIX"},
            cut == "GD",
        ]
        if any(special_checks):
            ranges = uniform
            use_fine = False
        else:
            ranges = fine
            use_fine = True

    # Find range and optional RAPSIZE
    c = float(carat) if pd.notna(carat) else 0.0
    for start, end in ranges:
        if start <= c <= end:
            rap = rapsize_map.get((start, end)) if use_fine else None
            return f"{start:.2f}-{end:.2f}", rap

    return "", None


def add_psc_count(df_in: pd.DataFrame) -> pd.DataFrame:
    """
    Add a PSC column that shows how many times each SR# value occurs.
    For example, if SR#=1 appears 4 times, PSC=4 for those rows.
    """
    df = df_in.copy()
    if "SR#" in df.columns:
        counts = df["SR#"].value_counts(dropna=False)
        df["PSC"] = df["SR#"].map(counts)
    else:
        df["PSC"] = None
    return df

def main():
    # ------------------- File Paths -------------------
    file_main = filedialog.askopenfilename(
        title="Stock List files",
        filetypes=[("Excel files", "*.xlsx")],
    )
    file_lookup = filedialog.askopenfilename(
        title="Select Trend File",
        filetypes=[("Excel files", "*.xlsx")],
    )

    # file_main = r"D:\Python\Revsion\Files\Stock List.xlsx"
    # file_lookup = r"D:\Python\Revsion\Files\Trend.xlsx"

    rev_file = r"G:\Certificate\PREMAL\REVISE DAYS TRACKING.xlsx"
    out_path = r"G:\Certificate\Python Output\Stock.xlsx"

    # Ensure output directory exists
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)

    # ------------------- Read Excel Files -------------------
    df = pd.read_excel(file_main)
    dfs = pd.read_excel(file_lookup)
    dfss = pd.read_excel(rev_file, sheet_name="REVISE DAYS")

    # ------------------- Clean Column Names -------------------
    df.columns = df.columns.str.strip()
    dfs.columns = dfs.columns.str.strip()
    dfss.columns = dfss.columns.str.strip()

    # ------------------- Ensure Key Columns Match -------------------
    df["STOCKID"] = df["STOCKID"].astype(str).str.strip()
    dfs["Sr. No"] = dfs["Sr. No"].astype(str).str.strip()
    dfss["Sr. No"] = dfss["Sr. No"].astype(str).str.strip()

    # ------------------- Ensure Numeric Columns -------------------
    df["VALUE"] = pd.to_numeric(df.get("VALUE"), errors="coerce")
    df["CARAT"] = pd.to_numeric(df.get("CARAT"), errors="coerce")

    # ------------------- Merge Data -------------------
    data = (
        df.merge(
            dfs[["Sr. No", "Quality Desc", "Trend", "DISC", "Stock Days"]],
            left_on="STOCKID",
            right_on="Sr. No",
            how="left",
        )
        .rename(
            columns={
                "DISC_x": "DISC",
                "DISC_y": "List Disc",
                "Stock Days": "OLD DAYS",
            }
        )
    )

    data = data.merge(
        dfss[["Sr. No", "DAYS"]],
        left_on="STOCKID",
        right_on="Sr. No",
        how="left",
    ).rename(columns={"DAYS": "LAST REV"})

    # ------------------- Convert DISC Columns to Numeric -------------------
    data["DISC"] = pd.to_numeric(data.get("DISC"), errors="coerce")
    data["List Disc"] = pd.to_numeric(data.get("List Disc"), errors="coerce")

    # ------------------- Build final_cut only (>=1 match rule) -------------------
    cut_norm = data["CUT"].astype(str).str.strip().str.upper()
    pol_norm = data["POL"].astype(str).str.strip().str.upper()
    sym_norm = data["SYM"].astype(str).str.strip().str.upper()

    ex_3 = (cut_norm.eq("3EX")).astype(int)
    gd_count = (cut_norm.eq("GD")).astype(int)
    vg_count = (pol_norm.eq("VG")).astype(int)
    vg2_count = (sym_norm.eq("VG")).astype(int)
    ex_count = (cut_norm.eq("EX")).astype(int)
    cut_count = (cut_norm.eq("NAN")).astype(int) 
    pol_count = (pol_norm.eq("GD")).astype(int) 
    sym_count = (sym_norm.eq("GD")).astype(int)
    pol_ex_count = (pol_norm.eq("EX")).astype(int)
    sym_ex_count = (sym_norm.eq("EX")).astype(int)
    sym_fr_count = (sym_norm.eq("FR")).astype(int)

    final_cut = cut_norm.copy()
    final_cut.loc[ex_3 > 1] = "3EX"
    final_cut.loc[ex_count >= 1] = "EX"
    final_cut.loc[(gd_count >= 1)] = "GD"
    final_cut.loc[(pol_count >= 1)] = "GD"
    final_cut.loc[(sym_count >= 1)] = "GD"
    final_cut.loc[(cut_count >= 1) & (pol_ex_count >= 1)] = "EX"
    final_cut.loc[(cut_count >= 1) & (sym_ex_count >= 1)] = "EX"
    final_cut.loc[(cut_count >= 1) & (vg_count >= 1)] = "VG"
    final_cut.loc[(cut_count >= 1) & (vg2_count >= 1)] = "VG"
    final_cut.loc[(cut_count >= 1) & (pol_count >= 1)] = "GD"
    final_cut.loc[(cut_count >= 1) & (sym_count >= 1)] = "GD"
    final_cut.loc[(cut_count >= 1) & (sym_fr_count >= 1)] = "GD"

    data["1CUT"] = final_cut

    # -------------------------------------------------------------------
    
    if "RAPSIZE" not in data.columns:
        data["RAPSIZE"] = pd.NA

    def range_and_rap(row):
        carat_range, rap = find_carat_range(
            row.get("CARAT", 0),
            row.get("SHAPE"),
            fluor=row.get("FLUOR"),
            luster=row.get("LUSTER"),
            tinch=row.get("TINCH"),
            cut=row.get("1CUT"),
        )
        return pd.Series([carat_range, rap])

    temp = data.apply(range_and_rap, axis=1)
    temp.columns = ["Carat Range", "New RAPSIZE"]

    # Always set Carat Range
    data["Carat Range"] = temp["Carat Range"]

    # Overwrite RAPSIZE only where a new mapping exists; otherwise keep old
    data["RAPSIZE"] = data["RAPSIZE"].where(
        temp["New RAPSIZE"].isna(), temp["New RAPSIZE"]
    )

    # Normalize RAPSIZE to uppercase (keeps existing text if not overwritten)
    data["RAPSIZE"] = data["RAPSIZE"].astype("string").str.upper()

    # ------------------- Split ROUND vs NON-ROUND -------------------
    shape_norm = data["SHAPE"].astype(str).str.strip().str.upper()
    round_mask = shape_norm.eq("ROUND")

    data_round_all = data.loc[round_mask].copy()
    data_nonround = data.loc[~round_mask].copy()

    # >>> NEW: create explicit carat splits for ROUND <<<
    data_round_sub = data_round_all.loc[data_round_all["CARAT"].astype(float) < 1.00].copy()   # 0.01–0.99
    data_round_1up = data_round_all.loc[data_round_all["CARAT"].astype(float) >= 1.00].copy()  # 1.00+

    # ------------------- Helper to build ranking & summary per sheet ------------------- "Quality Desc",
    def build_rank_and_summary(df_in: pd.DataFrame):
        group_cols = ["SHAPE","COLOR","CLARITY","FLUOR","CERT", "TINCH", "LUSTER", "RAPSIZE","1CUT"]

        df_in["Values"] = (
            df_in.groupby(group_cols, dropna=False)["VALUE"].transform("sum").round(2)
        )
        df_in["Avg DISC"] = (
            df_in.groupby(group_cols, dropna=False)["DISC"].transform("mean").round(2)
        )
        df_in["Avg List Disc"] = (
            df_in.groupby(group_cols, dropna=False)["List Disc"]
            .transform("mean")
            .round(2)
        )
        df_in["Diff"] = (df_in["Avg List Disc"] - df_in["Avg DISC"]).round(2)

        df_in["Values"] = df_in["Values"].fillna(0)
        df_in.sort_values(by="Values", ascending=False, inplace=True)
        df_in["SR#"] = df_in["Values"].rank(method="dense", ascending=False).astype(int)
        df_in.reset_index(drop=True, inplace=True)

        summary_out = (
            df_in.groupby(group_cols, dropna=False)
            .agg(
                Values=("VALUE", "sum"),
                Avg_DISC=("DISC", "mean"),
                Avg_List_Disc=("List Disc", "mean"),
            )
            .reset_index()
        )
        summary_out["Diff"] = summary_out["Avg_List_Disc"] - summary_out["Avg_DISC"]
        summary_out[["Values", "Avg_DISC", "Avg_List_Disc", "Diff"]] = summary_out[
            ["Values", "Avg_DISC", "Avg_List_Disc", "Diff"]
        ].round(2)
        return df_in, summary_out

    # >>> Build for each ROUND slice and for NON-ROUND <<<
    data_round_sub, summary_round_sub = build_rank_and_summary(data_round_sub)   # 0.01–0.99 RD
    data_round_1up, summary_round_1up = build_rank_and_summary(data_round_1up)   # 1.00+ RD
    data_nonround, summary_nonround = build_rank_and_summary(data_nonround)      # Fancy

    # ------------------- Reorder Columns -------------------
    desired_order = ["STOCKID","SR#","PSC","Values","Trend","SHAPE","Carat Range","CARAT","COLOR","CLARITY","CERT","1CUT","CUT",
                        "POL","SYM","FLUOR","MEASUREMENT","RAPPRICE","RAP V","LAST REV","OLD DAYS","List Disc","DISC",
                        "PRICE","VALUE", "CERTNO","DEPTH","TABLE","RATIO","BLACK","TINCH","LUSTER","CA","CH","PA","PH",
                        "GIRDLE","GIRDLE%","CULET","KEY TO SYMBOLS","REPORT COMM","STATUS","SECTION NO","ROUGH ORIGIN",
                        "DIGITAL CERT","IMAGES/VIDEO","V360","STONE LOCATION","TRACR HYPERLINK","DETAILS LINK","Quality Desc",
                        "Avg DISC","Avg List Disc","Diff"]

    # --- Fill PSC from SR# count on each sheet ---
    data_round_sub  = add_psc_count(data_round_sub)
    data_round_1up  = add_psc_count(data_round_1up)
    data_nonround   = add_psc_count(data_nonround)

    # Ensure columns present in each DF before reindexing
    for col in desired_order:
        if col not in data_round_sub.columns:
            data_round_sub[col] = None
        if col not in data_round_1up.columns:
            data_round_1up[col] = None
        if col not in data_nonround.columns:
            data_nonround[col] = None

    data_round_sub = data_round_sub[desired_order]
    data_round_1up = data_round_1up[desired_order]
    data_nonround = data_nonround[desired_order]

    # ------------------- Save Results -------------------
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # 0.01–0.99 ROUND (data + summary limited to this slice)
        data_round_sub.to_excel(writer, index=False, sheet_name="0.01-0.99 RD")
        summary_round_sub.to_excel(writer, index=False, sheet_name="Round Summary")

        # FANCY (unchanged)
        data_nonround.to_excel(writer, index=False, sheet_name="FANCY SHAPES")
        summary_nonround.to_excel(writer, index=False, sheet_name="Fancy Summary")

        # 1.00+ ROUND (data + its own summary)
        data_round_1up.to_excel(writer, index=False, sheet_name="1.00 & ABOVE RD")
        summary_round_1up.to_excel(
            writer, index=False, sheet_name="1.00 & ABOVE RD Summary"
        )

    # ------------------- Apply Header Colors on data sheets -------------------
    header_colors = {
        "SR#": "00FFFF",        # Cyan
        "Values": "FF00FF",     # Magenta
        "LAST REV": "FF69B4",   # Pink
        "OLD DAYS": "ff6666",   # Blue
        "List Disc": "ADFF2F",  # Green-yellow
        "DISC": "FF8C00",       # Orange
        "1CUT": "FFFF00",       # Yellow highlight for the new indicator
    }

    wb = load_workbook(out_path)

    # Color headers on the actual saved sheet names (data sheets)
    for sheet_name in ["0.01-0.99 RD", "FANCY SHAPES", "1.00 & ABOVE RD"]:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for col_idx in range(1, ws.max_column + 1):
                header = ws.cell(row=1, column=col_idx).value
                if header in header_colors:
                    ws.cell(row=1, column=col_idx).fill = PatternFill(
                        start_color=header_colors[header],
                        end_color=header_colors[header],
                        fill_type="solid",
                    )

    # ------------------- Conditional Highlighting -------------------
    # TINCH: highlight if value is not empty
    # LUSTER: highlight if value is not "EX"
    highlight_fill = PatternFill(start_color="BFFF00", end_color="BFFF00", fill_type="solid")
    green_fill = PatternFill(start_color="00CC66", end_color="00CC66", fill_type="solid")
    def _is_nonempty(val) -> bool:
        if val is None:
            return False
        s = str(val).strip()
        if s == "" or s.upper() in {"NONE", "NAN"}:
            return False
        return True

    for sheet_name in ["0.01-0.99 RD", "FANCY SHAPES", "1.00 & ABOVE RD"]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        # Build header map
        header_map = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
        tinch_col = header_map.get("TINCH")
        luster_col = header_map.get("LUSTER")
        old_day = header_map.get("OLD DAYS")

        for r in range(2, ws.max_row + 1):  # Skip header
            if tinch_col:
                tinch_val = ws.cell(row=r, column=tinch_col).value
                if _is_nonempty(tinch_val):
                    ws.cell(row=r, column=tinch_col).fill = highlight_fill

            if luster_col:
                luster_val = ws.cell(row=r, column=luster_col).value
                lv = str(luster_val).strip().upper() if luster_val is not None else ""
                # Highlight if not blank and not "EX"
                if lv and lv != "EX":
                    ws.cell(row=r, column=luster_col).fill = highlight_fill

            if old_day:
                old_val = ws.cell(row=r, column=old_day).value
                try:
                    if old_val is not None and float(old_val) > 90:
                        ws.cell(row=r, column=old_day).fill = green_fill
                except (ValueError, TypeError):
                    pass  # Ignore non-numeric entries

    # ------------------- Save with formatting -------------------
    wb.save(out_path)

    # ------------------- Auto Open File -------------------
    try:
        os.startfile(out_path)  # type: ignore[attr-defined]
    except AttributeError:
        # Non-Windows platforms
        pass
if __name__ == "__main__":
    main()
