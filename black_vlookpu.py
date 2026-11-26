import os
from io import BytesIO
import pandas as pd
import streamlit as st

COLUMNS = [
    "dLot", "Stock Id", "dSize", "dWt", "Lab", "dShape", "dColor",
    "dClarity", "dFlrc", "cutName", "dCut", "dPol", "dSym", "dTable",
    "dDepth", "dDisc", "dRap", "Rap value", "dValue", "dDisc1", "dValue1",
    "M_DAYS", "M_TREND", "tot_Pcs", "dSoldPrice", "dSoldDisc",
    "dNatts_1", "dNatts_2"
]

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


def run_black_deal_mod():

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
            step=1
        )

    if file1 is None or file2 is None:
        st.success("Select Files.")
        st.stop()
    
    with st.spinner("Processing..."):
        try:
            progress_bar = st.progress(10)
            df1_raw = read_file(file1)

            progress_bar.progress(20)
            df2 = read_file(file2, skiprows=skip2)

        except Exception as e:
            st.error(f"Failed to read file: {e}")
            st.stop()

        progress_bar.progress(30)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"**{file1.name}**")
            st.dataframe(df1_raw)
        with c2:
            st.markdown(f"**{file2.name}**")
            st.dataframe(df2)
        progress_bar.progress(40)
        cc1, cc2 = st.columns(2)
        with cc1:
            key1 = st.selectbox("File-1 Key Column", df1_raw.columns)
        with cc2:
            key2 = st.selectbox("File-2 Key Column", df2.columns)
        progress_bar.progress(50)
        st.markdown("### Select Columns")
        cols_to_add = st.multiselect(
            "Choose up to 3 values for Stock Id / dNatts_1 / dNatts_2",
            ["Blank"] + list(df2.columns)
        )
        progress_bar.progress(60)
        target_names = ["Stock Id", "dNatts_1", "dNatts_2"]

        if len(cols_to_add) > 3:
            st.warning("Choose only up to 3 values.")
            cols_to_add = cols_to_add[:3]

        rename_map = {}
        blank_targets = []

        for i, selected in enumerate(cols_to_add):
            target_col = target_names[i]

            if selected == "Blank":
                blank_targets.append(target_col)
            elif selected in df2.columns:
                rename_map[selected] = target_col

        real_selected_cols = [col for col in cols_to_add if col in df2.columns]
        progress_bar.progress(60)
        df1 = df1_raw.copy()
        df2 = df2.copy()

        for col in COLUMNS:
            if col not in df1.columns:
                df1[col] = None
        progress_bar.progress(70)
        df1[key1] = clean_key(df1[key1])
        df2[key2] = clean_key(df2[key2])

        merge_cols = [key2] + [col for col in real_selected_cols if col != key2]
        merge_cols = list(dict.fromkeys(merge_cols))

        right_df = df2[merge_cols].copy().rename(columns=rename_map)

        cols_from_right_after_rename = [c for c in right_df.columns if c != key2]
        drop_from_left = [
            c for c in cols_from_right_after_rename
            if c in df1.columns and c != key1
        ]

        if drop_from_left:
            df1 = df1.drop(columns=drop_from_left)
        progress_bar.progress(80)
        df1 = df1.merge(
            right_df,
            left_on=key1,
            right_on=key2,
            how="left"
        )

        if key2 in df1.columns and key2 != key1:
            df1 = df1.drop(columns=[key2])

        for col in blank_targets:
            df1[col] = None

        for col in target_names:
            if col not in df1.columns:
                df1[col] = None

        for col in COLUMNS:
            if col not in df1.columns:
                df1[col] = None

        df1 = df1[COLUMNS]

        st.markdown("### Result")
        progress_bar.progress(90)
        st.dataframe(df1)

        excel_bytes = to_excel_bytes(df1)
        name = os.path.splitext(file1.name)[0]
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
    run_black_deal_mod()
