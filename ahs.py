import base64
from pathlib import Path
import pandas as pd
import warnings
from openpyxl import load_workbook
import tempfile
import streamlit as st

def merge_sheets(file_path):
    """Merge all sheets in an Excel file into a single DataFrame."""
    warnings.simplefilter("ignore", UserWarning)
    workbook = load_workbook(file_path)
    sheet_names = workbook.sheetnames
    merged_data = pd.DataFrame()

    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        data = sheet.values
        columns = next(data)[0:]
        df = pd.DataFrame(data, columns=columns)
        if not df.empty and df.dropna(how="all").shape[0] > 0:
            df = df.dropna(axis=1, how="all")
            merged_data = pd.concat([merged_data, df], ignore_index=True)
    return merged_data


def append_total_row(df):
    total = df['VALUE'].sum()
    row_data = {}
    for col in df.columns:
        if col == 'VALUE':
            row_data[col] = total
    total_df = pd.DataFrame([row_data])
    return pd.concat([df, total_df], ignore_index=True)


def ahs(df, output_file_path):
    """Filter and process the DataFrame, then save it to an Excel file with multiple sheets."""
    required_columns = ['CERT', 'STATUS', 'RAPSIZE', 'LUSTER', 'SHAPE', 'COLOR', 'CLARITY', 'VALUE']
    for column in required_columns:
        if column not in df.columns:
            raise ValueError(f"Missing required column: {column}")

    df = df[df['STATUS'] != 'New Arrival']
    df = df[~df["STOCKID"].astype(str).str.contains("FSA", na=False, regex=True)]

    IIDGR = df[df['CERT'].str.contains('https://.*,"IIDGR"', na=False, regex=True) |
               df['CERT'].str.contains('IIDGR', na=False)]
    HRD_IGI = df[df['CERT'].str.contains('https://.*(?:HRD|IGI)', na=False, regex=True) |
                 df['CERT'].str.contains('HRD|IGI', na=False)]

    # --- GIA ROUND 0.30–0.69 ---
    GIR_RD_30_69 = df[
        (df['CERT'].str.contains('https://.*,"GIA"', na=False, regex=True) |
         df['CERT'].str.contains('GIA', na=False)) &
        (df['CARAT'] >= 0.30) & (df['CARAT'] <= 0.69) &
        (~df['LUSTER'].isin(['M-1', 'M-2', 'M-3', 'M-4'])) &
        (df['SHAPE'] == 'ROUND') &
        (df['DISC'] <= -49.5)
    ].copy()
    GIR_RD_30_69.loc[GIR_RD_30_69['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'] = \
        GIR_RD_30_69.loc[GIR_RD_30_69['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'] \
        .apply(lambda x: 'Blank' if x in ['IF', 'VVS1', 'VVS2'] else x)
    GIR_RD_30_69 = GIR_RD_30_69[GIR_RD_30_69['CLARITY'] != 'Blank']
    GIR_RD_30_69 = append_total_row(GIR_RD_30_69)

    # --- GIA ROUND 0.70 & ABOVE ---
    GIR_RD_70_Above = df[
        (df['CERT'].str.contains('https://.*,"GIA"', na=False, regex=True) |
         df['CERT'].str.contains('GIA', na=False)) &
        (df['CARAT'] >= 0.70) &
        (~df['LUSTER'].isin(['M-1', 'M-2', 'M-3', 'M-4'])) &
        (df['SHAPE'] == 'ROUND') &
        (df['DISC'] <= -45) &
        (df['VALUE'] <= 2499)
    ].copy()
    GIR_RD_70_Above.loc[GIR_RD_70_Above['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'] = \
        GIR_RD_70_Above.loc[GIR_RD_70_Above['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'] \
        .apply(lambda x: 'Blank' if x in ['IF', 'VVS1', 'VVS2'] else x)
    GIR_RD_70_Above = GIR_RD_70_Above[GIR_RD_70_Above['CLARITY'] != 'Blank']
    GIR_RD_70_Above = append_total_row(GIR_RD_70_Above)

    # --- GIA FANCY SHAPES ---
    GIR_Fancy = df[
        (df['CERT'].str.contains('https://.*,"GIA"', na=False, regex=True) |
         df['CERT'].str.contains('GIA', na=False)) &
        (~df['LUSTER'].isin(['M-1', 'M-2', 'M-3', 'M-4'])) &
        (df['SHAPE'] != 'ROUND') &
        (df['DISC'] <= -40) &
        (df['VALUE'] <= 2499)
    ].copy()
    GIR_Fancy.loc[GIR_Fancy['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'] = \
        GIR_Fancy.loc[GIR_Fancy['COLOR'].isin(['D', 'E', 'F']), 'CLARITY'] \
        .apply(lambda x: 'Blank' if x in ['IF', 'VVS1', 'VVS2'] else x)
    GIR_Fancy = GIR_Fancy[GIR_Fancy['CLARITY'] != 'Blank']
    GIR_Fancy = append_total_row(GIR_Fancy)

    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        filtered_df = df[~df['CERT'].str.contains('IIDGR|HRD|IGI', na=False, regex=True)]
        filtered_df.to_excel(writer, sheet_name='list', index=False)
        IIDGR.to_excel(writer, sheet_name='IIDGR', index=False)
        HRD_IGI.to_excel(writer, sheet_name='HRD & IGI', index=False)
        GIR_RD_30_69.to_excel(writer, sheet_name='GIA RD 0.30 - 0.69 @-', index=False)
        GIR_RD_70_Above.to_excel(writer, sheet_name='GIA RD 0.70 & ABOVE @-', index=False)
        GIR_Fancy.to_excel(writer, sheet_name='GIA FANCY SHAPES @-', index=False)
        writer.sheets['list'].set_tab_color('#ff9999')
        writer.sheets['IIDGR'].set_tab_color('#ffcc99')
        writer.sheets['HRD & IGI'].set_tab_color('#ffcccc')
        writer.sheets['GIA RD 0.30 - 0.69 @-'].set_tab_color('#66ffff')
        writer.sheets['GIA RD 0.70 & ABOVE @-'].set_tab_color('#99ff99')
        writer.sheets['GIA FANCY SHAPES @-'].set_tab_color('#ffccff')
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

def run_ahs_mod():
    """Streamlit UI entrypoint."""
    st.subheader("AHS DEAL")

    sample_path = _find_sample("AHS.xlsx")

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

        uploaded = st.file_uploader("Select AHS Excel (.xlsx)", type=["xlsx"])

    if uploaded is None:
        st.info(" Upload the Excel file")
        st.stop()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
        tmp_in.write(uploaded.read())
        temp_input_path = tmp_in.name

    try:
        with st.spinner("sheets... ⏳"):
            merged_df = merge_sheets(temp_input_path)
    except Exception as e:
        st.error(f"Error reading/merging sheets: {e}")
        st.stop()

    st.dataframe(merged_df.head(100))

    # Processing with loader
    try:
        with st.spinner("output file..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
                temp_output_path = tmp_out.name

            ahs(merged_df, temp_output_path)
        st.success("Processing complete ✅")

        with open(temp_output_path, "rb") as f:
            xlsx_bytes = f.read()

        with st.sidebar:
            st.download_button(
                label="📥 Download File",
                data=xlsx_bytes,
                file_name="AHS DEAL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"Processing error: {e}")
# Standard Python entrypoint
if __name__ == "__main__":
    run_ahs_mod()

