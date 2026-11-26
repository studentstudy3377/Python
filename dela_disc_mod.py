from datetime import datetime
import os
from io import BytesIO

import pandas as pd
import streamlit as st


def clean_key(col):
    return (
        col.astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.upper()
        .str.replace(r"[\W_]+", "", regex=True)
    )


def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, skiprows=skiprows)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DEAL File")
    output.seek(0)
    return output.getvalue()


def run_dela_disc_mod():
    st.subheader("Black Vlookup File")

    with st.sidebar:
        st.markdown("### Upload Files")
        file1 = st.file_uploader("Select DEAL File", type=["xlsx", "xls"], key="f1")
        file2 = st.file_uploader("Select ORG File", type=["xlsx", "xls"], key="f2")

        skip2 = st.number_input(
            "Skip Rows (ORG File)",
            min_value=0,
            max_value=1000,
            value=0,
            step=1,
        )

    if file1 is None or file2 is None:
        st.info("Please upload both files.")
        st.stop()

    try:
        df1_raw = read_file(file1)
        df2 = read_file(file2, skiprows=skip2)
    except Exception as e:
        st.error(f"Error reading files: {e}")
        st.stop()

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**{file1.name}**")
        st.dataframe(df1_raw)

    with c2:
        st.markdown(f"**{file2.name}**")
        st.dataframe(df2)

    cc1, cc2 = st.columns(2)
    with cc1:
        key1 = st.selectbox("File-1 Key Column", df1_raw.columns)

    with cc2:
        key2 = st.selectbox("File-2 Key Column", df2.columns)

    st.markdown("### Select Column")
    col_to_add = st.selectbox("Select Column", df2.columns)

    df1 = df1_raw.copy()
    df2 = df2.copy()

    df1[key1] = clean_key(df1[key1])
    df2[key2] = clean_key(df2[key2])

    today_str = datetime.today().strftime("%d-%m-%Y")
    temp_col = "__selected_col__"

    # Prepare merge source safely
    if col_to_add == key2:
        df2_merge = df2[[key2]].copy()
        temp_col = key2
    else:
        df2_merge = df2[[key2, col_to_add]].copy()
        df2_merge.rename(columns={col_to_add: temp_col}, inplace=True)

    # Merge
    df1 = df1.merge(
        df2_merge,
        left_on=key1,
        right_on=key2,
        how="left",
    )

    # Drop duplicate key column only
    if key2 in df1.columns and key2 != key1:
        df1.drop(columns=[key2], inplace=True)

    # Rename only imported selected column to today's date
    if temp_col in df1.columns and temp_col != key1:
        df1.rename(columns={temp_col: today_str}, inplace=True)

    # Calculations
    required_cols = ["dRap", "dWt", "dDisc1", today_str]
    missing_cols = [col for col in required_cols if col not in df1.columns]

    if missing_cols:
        st.warning(f"Missing columns for calculation: {missing_cols}")
    else:
        for col in required_cols:
            df1[col] = pd.to_numeric(df1[col], errors="coerce")

        if "PRICE" in df1.columns:
            df1.drop(columns=["PRICE"], inplace=True)
        df1["PRICE"] = df1["dRap"] + (df1["dRap"] * df1[today_str] / 100)

        if "VALUE" in df1.columns:
            df1.drop(columns=["VALUE"], inplace=True)
        df1["VALUE"] = df1["PRICE"] * df1["dWt"]

        if "DIFF" in df1.columns:
            df1.drop(columns=["DIFF"], inplace=True)
        df1["DIFF"] = df1["dDisc1"] - df1[today_str]

    st.markdown("### Result")
    st.dataframe(df1)

    excel_bytes = to_excel_bytes(df1)
    name = os.path.splitext(file1.name)[0]
    excel_file_name = f"{name}.xlsx"

    with st.sidebar:
        st.download_button(
            label="Download File",
            data=excel_bytes,
            file_name=excel_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    run_dela_disc_mod()
