import os
from io import BytesIO

import pandas as pd
import streamlit as st

def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    """Read an uploaded Excel file."""
    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        skiprows=skiprows,
        dtype={"Job No": "string"},
        keep_default_na=False,
    )

    # Remove leading and trailing spaces from column names.
    df.columns = [str(col).strip() for col in df.columns]

    return df

def make_streamlit_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Convert mixed object columns to strings for Streamlit display."""
    safe_df = df.copy()

    for col in safe_df.columns:
        if safe_df[col].dtype == "object":
            safe_df[col] = safe_df[col].astype("string")

    return safe_df

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert a DataFrame into downloadable Excel bytes."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Result",
        )

    output.seek(0)
    return output.getvalue()

def run_web_fromat():
    st.subheader("Web Format ")

    with st.sidebar:
        st.markdown("### web Format")

        file1 = st.file_uploader(
            "Select File 1",
            type=["xlsx", "xls"],
            key="file1",
        )

        skip1 = st.number_input(
            "Skip Rows for File 1",
            min_value=0,
            max_value=1000,
            value=0,
            step=1,
            key="skip_rows_file1",
        )

        file2 = st.file_uploader(
            "Select File 2",
            type=["xlsx", "xls"],
            key="file2",
        )

        skip2 = st.number_input(
            "Skip Rows for File 2",
            min_value=0,
            max_value=1000,
            value=0,
            step=1,
            key="skip_rows_file2",
        )

    # Both files are required for this merge.
    if file1 is None or file2 is None:
        st.info("Please select both files.")
        st.stop()

    with st.spinner("Processing..."):
        progress_bar = st.progress(10)

        try:
            df1 = read_file(file1, skiprows=skip1)
            progress_bar.progress(35)

            df2 = read_file(file2, skiprows=skip2)
            progress_bar.progress(60)

        except Exception as e:
            st.error(f"Failed to read file: {e}")
            st.stop()

        df1 = df1[~df1['FANCY_CUT'].isin(["IDEAL"])]
        df1 = df1[~df1['KEEP_CUT'].isin(["N/A"])]
        df1["Location"] = df1["Location"].replace("SACHIN", "DUBAI")
        mask = df1['Shape'].str.upper().ne('ROUND')
        df1.loc[mask, 'Measurement'] = (
            df1.loc[mask, 'Measurement']
            .str.replace(r'\s*-\s*', ' X ', regex=True)
        )

        df1["1"] = ""

        df1_display = make_streamlit_safe(df1)
        df2_display = make_streamlit_safe(df2)
        
        progress_bar.progress(80)

        st.markdown(f"### File 1: {file1.name}")
        st.dataframe(
            df1_display)
            
        st.markdown(f"### File 2: {file2.name}")
        st.dataframe(df2_display)

        # Keep every row from File 2 and matching rows from File 1.
        result_df = df1[df1["Stone No"].isin(df2["STOCKID"])].copy()

        st.markdown("### Result")
        result_display = make_streamlit_safe(result_df)
        st.dataframe(result_display)


            
        excel_bytes = to_excel_bytes(result_df)
        excel_file_name = f"Web Format {pd.Timestamp.now():%d.%m.%Y}.xlsx"

        progress_bar.progress(100)

    with st.sidebar:
        st.download_button(
            label="Download File",
            data=excel_bytes,
            file_name=excel_file_name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

if __name__ == "__main__":
    run_web_fromat()
