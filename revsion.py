import base64
import io
from pathlib import Path
import pandas as pd
import streamlit as st
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
            (0.50, 0.59), (0.96, 0.99), (0.60, 0.69), (0.70, 0.79), (0.80, 0.89),
            (0.90, 0.95), (1.00, 1.19), (1.20, 1.29), (1.30, 1.39), (1.40, 1.49),
            (1.50, 1.59),(1.60, 1.69), (1.70, 1.79), (1.80, 1.89), (1.90, 1.99), 
            (2.00, 2.19),(2.20, 2.49), (2.50, 2.99), (3.00, 3.49), (3.50, 3.99), 
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
            tinch_u in {"B-1", "B1", "B-2", "B2", "MIX"},
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

    return "Other", None


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

def run_revsion():
    st.title("Revision File")
    # File upload
    sample_path = _find_sample("Revision.xlsx")
    
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
            
        file_main = st.file_uploader("Select Stock File", type=["xlsx", "xls"], key="file_main")
        file_lookup = st.file_uploader("Select Trend File", type=["xlsx", "xls"], key="file_lookup")
        rev_file = st.file_uploader("Select Revision Days File", type=["xlsx", "xls"], key="rev_file")
        disc = st.file_uploader("Select  All Disc File", type=["xlsx", "xls"], key="disc_file")

    if not file_main or not file_lookup or not rev_file:
        st.info("Upload Files")
        st.stop()

    with st.spinner("Processing the file..."):
        try:
            progress_bar = st.progress(10)
            df = pd.read_excel(file_main)
            progress_bar.progress(15)
            dfs = pd.read_excel(file_lookup)
            progress_bar.progress(20)
            dfsss = pd.read_excel(disc)
            progress_bar.progress(25)
            dfss = pd.read_excel(rev_file, sheet_name="REVISE DAYS")
        except Exception as e:
            st.error(f"Error reading file: {e}")
            st.stop()

        # Clean column names
        df.columns = df.columns.str.strip()
        dfs.columns = dfs.columns.str.strip()
        dfss.columns = dfss.columns.str.strip()
        dfsss.columns = dfsss.columns.str.strip()

        # Ensure key columns match
        df["STOCKID"] = df["STOCKID"].astype(str).str.strip()
        dfs["Sr. No"] = dfs["Sr. No"].astype(str).str.strip()
        dfss["Sr. No"] = dfss["Sr. No"].astype(str).str.strip()
        dfsss["STOCKID"] = dfsss["STOCKID"].astype(str).str.strip()

        # Ensure numeric columns
        df["VALUE"] = pd.to_numeric(df.get("VALUE"), errors="coerce")
        df["CARAT"] = pd.to_numeric(df.get("CARAT"), errors="coerce")

        dfsss['STOCKID'].drop_duplicates(keep="last")

        # Merge data
        progress_bar.progress(30)
        data = df.merge(
            dfs[["Sr. No", "Quality Desc", "Trend", "DISC", "Stock Days"]],
            left_on="STOCKID",
            right_on="Sr. No",
            how="left"
        ).rename(columns={"DISC_x": "DISC", "DISC_y": "List Disc", "Stock Days": "OLD DAYS"})

        progress_bar.progress(40)
        data = data.merge(
            dfss[["Sr. No", "DAYS"]],
            left_on="STOCKID",
            right_on="Sr. No",
            how="left"
        ).rename(columns={"DAYS": "LAST REV"})
        
        progress_bar.progress(50)
        data = data.merge(
            dfsss[["STOCKID", "PDISC"]],
            left_on="STOCKID",
            right_on="STOCKID",
            how="left"
        )

        # Convert DISC columns to numeric
        data["DISC"] = pd.to_numeric(data["DISC"], errors="coerce")
        data["List Disc"] = pd.to_numeric(data["List Disc"], errors="coerce")

        # Build final_cut based on criteria
        cut_norm = data["CUT"].astype(str).str.strip().str.upper()
        pol_norm = data["POL"].astype(str).str.strip().str.upper()
        sym_norm = data["SYM"].astype(str).str.strip().str.upper()
        
        progress_bar.progress(60)
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

        progress_bar.progress(70)
        final_cut = cut_norm.copy()
        final_cut.loc[ex_3 > 1] = "3EX"
        final_cut.loc[ex_count >= 1] = "EX"
        final_cut.loc[(gd_count >= 1)] = "GD"
        final_cut.loc[(pol_count >= 1)] = "GD"
        final_cut.loc[(sym_count >= 1)] = "GD"
        final_cut.loc[(cut_count >= 1) & (pol_ex_count >= 1)] = "VG"
        final_cut.loc[(cut_count >= 1) & (sym_ex_count >= 1)] = "VG"
        final_cut.loc[(cut_count >= 1) & (vg_count >= 1)] = "VG"
        final_cut.loc[(cut_count >= 1) & (vg2_count >= 1)] = "VG"
        final_cut.loc[(cut_count >= 1) & (pol_count >= 1)] = "GD"
        final_cut.loc[(cut_count >= 1) & (sym_count >= 1)] = "GD"
        final_cut.loc[(cut_count >= 1) & (sym_fr_count >= 1)] = "GD"

        data["1CUT"] = final_cut

        if "RAPSIZE" not in data.columns:
            data["RAPSIZE"] = pd.NA

        def range_and_rap(row):
            # Find carat range and RAP size
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

        data["Carat Range"] = temp["Carat Range"]
        data["RAPSIZE"] = data["RAPSIZE"].where(
            temp["New RAPSIZE"].isna(), temp["New RAPSIZE"]
        )

        # Normalize RAPSIZE to uppercase
        data["RAPSIZE"] = data["RAPSIZE"].astype("string").str.upper()

        # Split ROUND vs NON-ROUND
        shape_norm = data["SHAPE"].astype(str).str.strip().str.upper()
        round_mask = shape_norm.eq("ROUND")

        data_round_all = data.loc[round_mask].copy()
        data_nonround = data.loc[~round_mask].copy()

        # Create explicit carat splits for ROUND
        data_round_sub = data_round_all.loc[data_round_all["CARAT"].astype(float) < 1.00].copy()  # 0.01–0.99
        data_round_1up = data_round_all.loc[data_round_all["CARAT"].astype(float) >= 1.00].copy()  # 1.00+

        # Helper to build ranking & summary per sheet
        progress_bar.progress(80)
        def build_rank_and_summary(df_in: pd.DataFrame):
            group_cols = ["SHAPE","COLOR","CLARITY","FLUOR","CERT", "TINCH", "LUSTER", "RAPSIZE","1CUT"]

            df_in["Values"] = df_in.groupby(group_cols, dropna=False)["VALUE"].transform("sum").round(2)
            df_in["Avg DISC"] = df_in.groupby(group_cols, dropna=False)["DISC"].transform("mean").round(2)
            df_in["Avg List Disc"] = df_in.groupby(group_cols, dropna=False)["List Disc"].transform("mean").round(2)
            df_in["Diff"] = (df_in["Avg List Disc"] - df_in["Avg DISC"]).round(2)
            
            rank_cols = ["SHAPE","COLOR","CLARITY","FLUOR","CERT", "TINCH", "LUSTER", "RAPSIZE","1CUT","Values"]
            df_in["Values"] = df_in["Values"].fillna(0)
            df_in.sort_values(by="Values", ascending=False, inplace=True)
            df_in["SR#"] = (df_in.groupby(rank_cols, dropna=False, sort=False).ngroup().add(1))
    
            df_in.reset_index(drop=True, inplace=True)

            summary_out = df_in.groupby(group_cols, dropna=False).agg(
                Values=("VALUE", "sum"),
                Avg_DISC=("DISC", "mean"),
                Avg_List_Disc=("List Disc", "mean"),
            ).reset_index()
            summary_out["Diff"] = summary_out["Avg_List_Disc"] - summary_out["Avg_DISC"]
            summary_out[["Values", "Avg_DISC", "Avg_List_Disc", "Diff"]] = summary_out[["Values", "Avg_DISC", "Avg_List_Disc", "Diff"]].round(2)
            return df_in, summary_out

        data_round_sub, summary_round_sub = build_rank_and_summary(data_round_sub)
        data_round_1up, summary_round_1up = build_rank_and_summary(data_round_1up)
        data_nonround, summary_nonround = build_rank_and_summary(data_nonround)

        # Reorder Columns
        progress_bar.progress(90)
        desired_order = ["STOCKID", "SR#", "PSC", "Values", "Trend", "SHAPE", "Carat Range", "CARAT", "COLOR", "CLARITY", "CERT", "1CUT", "CUT",
                        "POL", "SYM", "FLUOR", "MEASUREMENT", "RAPPRICE", "RAP V", "LAST REV", "OLD DAYS", "List Disc", "DISC","PDISC",
                        "PRICE", "VALUE", "CERTNO", "DEPTH", "TABLE", "RATIO", "BLACK", "TINCH", "LUSTER", "CA", "CH", "PA", "PH",
                        "GIRDLE", "GIRDLE%", "CULET", "KEY TO SYMBOLS", "REPORT COMM", "STATUS", "SECTION NO", "ROUGH ORIGIN",
                        "DIGITAL CERT", "IMAGES/VIDEO", "V360", "STONE LOCATION", "TRACR HYPERLINK", "DETAILS LINK", "Quality Desc",
                        "Avg DISC", "Avg List Disc", "Diff"]

        data_round_sub = add_psc_count(data_round_sub)
        data_round_1up = add_psc_count(data_round_1up)
        data_nonround = add_psc_count(data_nonround)

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

        # Define header colors
        header_colors = {
            "SR#": "00FFFF",        # Cyan
            "Values": "FF00FF",     # Magenta
            "LAST REV": "FF69B4",   # Pink
            "OLD DAYS": "ff6666",   # Blue
            "List Disc": "ADFF2F",  # Green-yellow
            "DISC": "FF8C00",       # Orange
            "1CUT": "FFFF00",       # Yellow highlight for the new indicator
        }

        # Conditional Formatting for specific columns
        highlight_fill = PatternFill(start_color="BFFF00", end_color="BFFF00", fill_type="solid")
        green_fill = PatternFill(start_color="00CC66", end_color="00CC66", fill_type="solid")

        # Function to check if the value is non-empty
        def _is_nonempty(val) -> bool:
            if val is None:
                return False
            s = str(val).strip()
            if s == "" or s.upper() in {"NONE", "NAN"}:
                return False
            return True

        # Save results to a buffer (in-memory file)
        output_buffer = io.BytesIO()
        
        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            # Write data to Excel sheets
            data_round_sub.to_excel(writer, index=False, sheet_name="0.01-0.99 RD")
            summary_round_sub.to_excel(writer, index=False, sheet_name="Round Summary")
            data_nonround.to_excel(writer, index=False, sheet_name="FANCY SHAPES")
            summary_nonround.to_excel(writer, index=False, sheet_name="Fancy Summary")
            data_round_1up.to_excel(writer, index=False, sheet_name="1.00 & ABOVE RD")
            summary_round_1up.to_excel(writer, index=False, sheet_name="1.00 & ABOVE RD Summary")

            # Access the workbook and sheets to modify headers and apply conditional formatting
            wb = writer.book

            for sheet_name in ["0.01-0.99 RD", "FANCY SHAPES", "1.00 & ABOVE RD"]:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]

                    # Highlight headers based on header_colors
                    for col_idx in range(1, ws.max_column + 1):
                        header = ws.cell(row=1, column=col_idx).value
                        if header in header_colors:
                            ws.cell(row=1, column=col_idx).fill = PatternFill(
                                start_color=header_colors[header],
                                end_color=header_colors[header],
                                fill_type="solid",
                            )

                    # Build header map for conditional formatting
                    header_map = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
                    tinch_col = header_map.get("TINCH")
                    luster_col = header_map.get("LUSTER")
                    old_day = header_map.get("OLD DAYS")

                    # Apply conditional formatting
                    for r in range(2, ws.max_row + 1):  # Skip header row
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
                            except ValueError:
                                pass  # Handle any conversion errors if the value is non-numeric

    # Show dataframes in Streamlit
    st.success("✅ File Generate...")
    st.subheader("Output:")
    st.dataframe(data_round_sub)
    
    output_buffer.seek(0)
    progress_bar.progress(100)
    # Provide download button for the user to download the file
    st.download_button(
        label="Download File",
        data=output_buffer,
        file_name="Revision File.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
if __name__ == "__main__":
    run_revsion()
