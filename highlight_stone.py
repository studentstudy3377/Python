import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- Static Sr values from your provided list ---
sr_values = [
    0.46, 1.26, 0.37, 1.21, 0.57, 0.54, 0.47, 0.55, 0.38, 1.43,
    1.25, 0.49, 1.66, 0.66, 0.39, 0.59, 0.65, 0.56, 1.2, 0.68,
    1.06, 0.58, 1.27, 1.35, 1.7, 1.05, 1.22, 0.64, 1.62, 1.36,
    1.08, 1.72, 1.23, 1.15, 0.67, 1.74, 0.48, 1.1, 1.73, 1.12,
    1.6, 0.69, 1.9, 1.34, 1.91, 1.07, 1.8, 1.58, 1.55, 1.28, 1.61
]
df_sr = pd.DataFrame({"Sr": sr_values})


def highlight_excel(df_main: pd.DataFrame, df_sr: pd.DataFrame) -> bytes:
    """Apply highlighting rules and return the workbook as Excel bytes."""
    # Convert date columns
    df_main["Pur Date"] = pd.to_datetime(df_main["Pur Date"], errors="coerce").dt.strftime("%d/%m/%Y")
    df_main["E Date"] = pd.to_datetime(df_main["E Date"], errors="coerce").dt.strftime("%d/%m/%Y")

    # Create Revers column
    df_main["Revers"] = df_main["Depth"] - df_main["Table"]

    # Merge on Cts vs Sr
    df = pd.merge(df_main, df_sr, left_on="Cts", right_on="Sr", how="left")

    # Write to buffer
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    highlight = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    red_font = Font(color="FF0000")

    headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}

    col_srno = headers.get("Sr. No")
    col_cut = headers.get("CUT")
    col_pol = headers.get("POLISH")
    col_sym = headers.get("Symetry")
    col_cts = headers.get("Cts")
    col_sr = headers.get("Sr")
    col_depth = headers.get("Depth")
    col_table = headers.get("Table")
    col_rev = headers.get("Revers")
    col_key = headers.get("Key Notes")

    # Highlighting logic
    for row in range(2, ws.max_row + 1):
        cts_value = ws.cell(row=row, column=col_cts).value if col_cts else None
        cut_value = ws.cell(row=row, column=col_cut).value if col_cut else None
        pol_value = ws.cell(row=row, column=col_pol).value if col_pol else None
        sym_value = ws.cell(row=row, column=col_sym).value if col_sym else None
        sr_value = ws.cell(row=row, column=col_sr).value if col_sr else None
        rev_value = ws.cell(row=row, column=col_rev).value if col_rev else None

        # Highlight rules
        if cts_value == 0.55 and str(cut_value).strip().upper() == "V. GOOD":
            ws.cell(row=row, column=col_srno).fill = highlight
            ws.cell(row=row, column=col_cut).fill = highlight

        elif sr_value is not None and cut_value == "V. GOOD":
            ws.cell(row=row, column=col_srno).fill = highlight
            ws.cell(row=row, column=col_cut).fill = highlight

        elif cut_value == "EXCELLENT" and pol_value in ["V. GOOD", "GOOD"]:
            ws.cell(row=row, column=col_srno).fill = highlight
            ws.cell(row=row, column=col_cut).fill = highlight
            ws.cell(row=row, column=col_pol).fill = highlight

        elif cut_value == "EXCELLENT" and sym_value in ["V. GOOD", "GOOD"]:
            ws.cell(row=row, column=col_srno).fill = highlight
            ws.cell(row=row, column=col_cut).fill = highlight
            ws.cell(row=row, column=col_sym).fill = highlight

        if isinstance(rev_value, (int, float)) and rev_value < 0:
            ws.cell(row=row, column=col_srno).fill = highlight
            ws.cell(row=row, column=col_depth).fill = highlight
            ws.cell(row=row, column=col_table).fill = highlight
            ws.cell(row=row, column=col_depth).font = red_font
            ws.cell(row=row, column=col_table).font = red_font

        # Key Notes highlight
        key_value = ws.cell(row=row, column=col_key).value if col_key else None
        if key_value:
            key_upper = str(key_value).upper()
            keywords = ["GRAINING", "PINPOINTS", "CHIP", "CAVITY", "MANUFACTURING REMNANT"]
            if any(k in key_upper for k in keywords):
                ws.cell(row=row, column=col_srno).fill = highlight
                ws.cell(row=row, column=col_key).fill = highlight

    # Shape rule highlights
    headers = {cell.value.strip().upper(): idx for idx, cell in enumerate(ws[1], start=1)}
    col_shape = headers.get("SHAPE")
    col_table = headers.get("TABLE")
    col_depth = headers.get("DEPTH")
    col_ratio = headers.get("RATIO")

    def shape_rule(row, name, table_rng, depth_rng, ratio_rng):
        shape_value = ws.cell(row=row, column=col_shape).value
        if shape_value and str(shape_value).strip().upper() == name:
            table_value = ws.cell(row=row, column=col_table).value
            depth_value = ws.cell(row=row, column=col_depth).value
            ratio_value = ws.cell(row=row, column=col_ratio).value
            # ws.cell(row=row, column=col_srno).fill = highlight
            if not (table_value and table_rng[0] <= float(table_value) <= table_rng[1]):
                ws.cell(row=row, column=col_table).fill = highlight
                ws.cell(row=row, column=col_srno).fill = highlight
            if not (depth_value and depth_rng[0] <= float(depth_value) <= depth_rng[1]):
                ws.cell(row=row, column=col_depth).fill = highlight
                ws.cell(row=row, column=col_srno).fill = highlight
            if not (ratio_value and ratio_rng[0] <= float(ratio_value) <= ratio_rng[1]):
                ws.cell(row=row, column=col_ratio).fill = highlight
                ws.cell(row=row, column=col_srno).fill = highlight

    for row in range(2, ws.max_row + 1):
        shape_rule(row, "MARQUISE", (56, 62), (58.00, 65.5), (1.75, 2.19))
        shape_rule(row, "PEAR", (56, 62), (58.00, 65.5), (1.45, 1.65))
        shape_rule(row, "OVAL", (56, 62), (58.00, 65.5), (1.30, 1.55))
        shape_rule(row, "EMRALD", (61, 69), (62.00, 71.00), (1.30, 1.55))
        shape_rule(row, "CUSHION", (58, 65), (66.00, 71.50), (1.00, 1.20))
        shape_rule(row, "PRINCESS", (68, 75), (69.00, 78.00), (1.00, 1.05))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def run_highlight():
    st.title(" Highlight Stone")

    with st.sidebar:
        st.markdown("### Upload Excel File")
        file_main = st.file_uploader("Select Main Excel File", type=["xlsx"], key="main_file")
        skip1 = st.number_input("Skip Rows", 0, 1000, 0, 1)

    if not file_main:
        st.info("⬆️ Please upload your main Excel file to continue.")
        st.stop()

    try:
        df_main = pd.read_excel(file_main, skiprows=skip1)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    st.success("✅ File uploaded successfully!")
    st.dataframe(df_main)

    excel_bytes = highlight_excel(df_main, df_sr)

    with st.sidebar:
        st.download_button(
            label="⬇️ Download File",
            data=excel_bytes,
            file_name="Highlight Stone.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
if __name__ == "__main__":
    run_highlight()
