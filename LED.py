import os
import io
import time
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

# =========================
# Fast I/O helpers
# =========================
def read_file(uploaded_file):
    """Reads an Excel file with optional row skipping."""
    return pd.read_excel(uploaded_file)

def to_excel(df: pd.DataFrame) -> bytes:
    """Converts a DataFrame to Excel bytes for download."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    buffer.seek(0)
    return buffer.getvalue()

def run_led_mod():
    st.title("LED Files")

    with st.sidebar:
        st.markdown("### Upload File")
        file1 = st.file_uploader("Select File", type=["xlsx", "xls"], key="f1")
        st.markdown("### Price File Path")
        price_file_path = st.file_uploader("Select File Disc", type=["xlsx", "xls"])

    if not file1 or not price_file_path:
        st.info("⬆️ Please upload a file to continue.")
        st.stop()

    progress = st.progress(0)
    pct = 0

    def step(msg, bump=8):
        nonlocal pct
        pct = min(100, pct + bump)
        status.update(label=msg, state="running")
        progress.progress(pct)

    with st.status("⏳ Processing…", expanded=True) as status:
        try:
            step("Reading uploaded stock file…", 10)
            df_stock = read_file(file1)

            step("Reading Discount file…", 10)
            df_disc = read_file(price_file_path)

            # Filters
            step("Filtering HRD/GD and applying fluo filter…", 10)
            mask_gd   = (df_stock["POL"] == "GD") | (df_stock["SYM"] == "GD")
            mask_hrd  = df_stock["CERT"] == "HRD"
            mask_fluo = df_stock["FLUOR"] == "NONE"
            df_stock  = df_stock.loc[~(mask_gd | mask_hrd) & mask_fluo].reset_index(drop=True)

            mask_ex = (df_stock["POL"] == "EX") & (df_stock["SYM"] == "EX")
            df_stock["CUT"] = np.where(mask_ex, "2EX", "VG")

            # Discount mapping
            step("Mapping discount values…", 10)
            cols_key = ["SHAPE","RAPSIZE","COLOR","CLARITY","CUT"]
            df_stock[cols_key] = df_stock[cols_key].fillna("").astype("string")
            df_stock["Concnate"] = df_stock[cols_key].agg(" - ".join, axis=1)

            vg_map  = dict(zip(df_disc["CON1"], df_disc["VG DISC"]))
            ex2_map = dict(zip(df_disc["CON2"], df_disc["2EX DISC"]))

            df_stock["New Disc"] = np.where(
                df_stock["CUT"].eq("2EX"),
                df_stock["Concnate"].map(ex2_map),
                df_stock["Concnate"].map(vg_map)
            )

            df_stock = df_stock[df_stock["New Disc"].notna()].copy()
            df_stock["CERTNO"] = df_stock["CERTNO"].fillna("").astype(str)
            df_stock["New Disc"] = pd.to_numeric(df_stock["New Disc"], errors="coerce")

            # Carat calculation
            step("Computing New Carat…", 6)
            clean_size = df_stock["RAPSIZE"].str.replace("UP", "", regex=False)
            size_vals  = pd.to_numeric(clean_size, errors="coerce")
            no_dot     = ~clean_size.str.contains(r"\.", na=False)
            new_carat  = np.where(no_dot, size_vals / 100.0, size_vals)
            df_stock.insert(int(df_stock.columns.get_loc("CARAT")) + 1, "New Carat", new_carat)

            # Reorder columns
            step("Reordering columns & adding formulas…", 10)
            desired_order = [
                "STOCKID", "SHAPE", "RAPSIZE", "CARAT", "New Carat", "COLOR", "CLARITY", "CERT",
                "CUT", "POL", "SYM", "FLUOR", "MEASUREMENT", "RAPPRICE", "RAP V", "DISC", "PRICE",
                "VALUE", "New Disc", "New Price", "New Value", "Diff Value", "Diff %", "CERTNO",
                "DEPTH", "TABLE", "RATIO", "BLACK", "TINCH", "LUSTER", "CA", "CH", "PA", "PH",
                "GIRDLE", "GIRDLE%", "CULET", "KEY TO SYMBOLS", "REPORT COMM", "STATUS", "SECTION NO",
                "ROUGH ORIGIN", "DIGITAL CERT", "IMAGES/VIDEO", "V360", "STONE LOCATION", "TRACR HYPERLINK",
                "DETAILS LINK"
            ]
            for col in desired_order:
                if col not in df_stock.columns:
                    df_stock[col] = None
            df_stock = df_stock[desired_order]

            # Formulas
            col_disc_new_idx = int(df_stock.columns.get_loc("New Disc")) + 1
            try:
                col_rappr_idx = int(df_stock.columns.get_loc("RAPPRICE")) + 1
                df_stock["New Price"] = [
                    f"={get_column_letter(col_rappr_idx)}{r}*(1+{get_column_letter(col_disc_new_idx)}{r}/100)"
                    for r in range(2, len(df_stock) + 2)
                ]
            except Exception as e:
                st.error(f"Error in formula generation for 'New Price': {e}")
                raise


            rappr_idx  = int(df_stock.columns.get_loc("RAPPRICE")) + 1
            col_disc_i = int(df_stock.columns.get_loc("DISC")) + 1
            df_stock["PRICE"] = [
                f"={get_column_letter(rappr_idx)}{r}*(1+{get_column_letter(col_disc_i)}{r}/100)"
                for r in range(2, len(df_stock) + 2)
            ]

            if "VALUE" in df_stock.columns:
                df_stock.drop("VALUE", axis=1, inplace=True)
            col_carat = get_column_letter(int(df_stock.columns.get_loc("CARAT")) + 1)
            col_price_o = get_column_letter(int(df_stock.columns.get_loc("PRICE")) + 1)
            value_fm = [f"={col_carat}{r}*{col_price_o}{r}" for r in range(2, len(df_stock) + 2)]
            idx_val = int(df_stock.columns.get_loc("PRICE")) + 1
            df_stock.insert(idx_val, "VALUE", value_fm)

            col_np2   = get_column_letter(int(df_stock.columns.get_loc("New Price")) + 1)
            col_ncar2 = get_column_letter(int(df_stock.columns.get_loc("New Carat")) + 1)
            df_stock["New Value"] = [f"={col_np2}{r}*{col_ncar2}{r}" for r in range(2, len(df_stock) + 2)]

            col_nv3  = get_column_letter(int(df_stock.columns.get_loc("New Value")) + 1)
            col_val3 = get_column_letter(int(df_stock.columns.get_loc("VALUE")) + 1)
            df_stock["Diff Value"] = [f"={col_nv3}{r}-{col_val3}{r}" for r in range(2, len(df_stock) + 2)]

            col_dv4  = get_column_letter(int(df_stock.columns.get_loc("Diff Value")) + 1)
            col_val4 = get_column_letter(int(df_stock.columns.get_loc("VALUE")) + 1)
            df_stock["Diff %"] = [
                f"=IFERROR(ROUND(({col_dv4}{r}/{col_val4}{r})*100, 0), 0) & \"%\""
                for r in range(2, len(df_stock) + 2)
            ]

            # Subtotal row
            step("Adding subtotal row…", 5)
            start_row = 2
            end_row   = len(df_stock) + 1
            letters = {
                "VALUE":     get_column_letter(int(df_stock.columns.get_loc("VALUE")) + 1),
                "New Value": get_column_letter(int(df_stock.columns.get_loc("New Value")) + 1),
                "Diff Value":get_column_letter(int(df_stock.columns.get_loc("Diff Value")) + 1),
            }
            subtotal = {c: "" for c in df_stock.columns}
            subtotal["VALUE"]      = f"=SUBTOTAL(9,{letters['VALUE']}{start_row}:{letters['VALUE']}{end_row})"
            subtotal["New Value"]  = f"=SUBTOTAL(9,{letters['New Value']}{start_row}:{letters['New Value']}{end_row})"
            subtotal["Diff Value"] = f"=SUBTOTAL(9,{letters['Diff Value']}{start_row}:{letters['Diff Value']}{end_row})"
            subtotal["Diff %"]     = f"=IFERROR(ROUND({letters['Diff Value']}{end_row + 1}/{letters['VALUE']}{end_row + 1} * 100, 0), 0) & \"%\""
            st.dataframe(df_stock,hide_index=True)
            df_stock = pd.concat([df_stock, pd.DataFrame([subtotal])], ignore_index=True)

            # Download
            step("Generating Excel file…", 8)
            excel_bytes = to_excel(df_stock)

            status.update(label="✅ Completed successfully!", state="complete")
            progress.progress(100)
        except Exception as e:
            status.update(label="❌ Error occurred", state="error")
            st.error(f"Error: {e}")
            st.stop()
    
    st.download_button(
                label="Download File",
                data=excel_bytes,
                file_name="LED Stock.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
               
            )

if __name__ == "__main__":
    run_led_mod()
