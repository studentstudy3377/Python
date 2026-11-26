import os
from io import BytesIO
import pandas as pd
import streamlit as st


def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    uploaded_file.seek(0)
    return pd.read_excel(
        uploaded_file,
        skiprows=skiprows,
        dtype={"Job No": "string"}
    )


def make_streamlit_safe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype("string")

    return df


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DEAL File")

    output.seek(0)
    return output.getvalue()


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "FANCY_CUT" in df.columns:
        df = df[~df['FANCY_CUT'].isin(["IDEAL"])]

    if "KEEP_CUT" in df.columns:
        df = df[~df['KEEP_CUT'].isin(["N/A"])]

    mask = df["SHAPE"].str.upper().ne("ROUND")
    df.loc[mask, 'MEASUREMENT'] = (
        df.loc[mask, 'MEASUREMENT']
        .str.replace(r'\s*-\s*', ' X ', regex=True)
    )

    df["1"] = ""

    return df

def run_export():
    st.subheader("Export File")

    with st.sidebar:
        st.markdown("### Upload Files")

        file = st.file_uploader(
            "Select File",
            type=["xlsx", "xls"],
            key="f2"
        )

        skip2 = st.number_input(
            "Skip Rows",
            min_value=0,
            max_value=1000,
            value=0,
            step=1
        )

    if file is None:
        st.success("Select File.")
        st.stop()

    with st.spinner("Processing..."):
        try:
            progress_bar = st.progress(10)

            progress_bar.progress(20)
            df = read_file(file, skiprows=skip2)

        except Exception as e:
            st.error(f"Failed to read file: {e}")
            st.stop()

        st.markdown(f"**Uploaded File:** {file.name}")
        st.dataframe(make_streamlit_safe(df))

        progress_bar.progress(30)

        df.columns = df.columns.str.upper()

        df = clean_data(df)
        df = make_streamlit_safe(df)

        st.markdown("### Result")
        st.dataframe(df)

        progress_bar.progress(90)

        excel_bytes = to_excel_bytes(df)
        name = os.path.splitext(file.name)[0]
        excel_file_name = f"{name}.xlsx"

        progress_bar.progress(100)

        with st.sidebar:
            st.download_button(
                label="Download File",
                data=excel_bytes,
                file_name=excel_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

if __name__ == "__main__":
    run_export()
