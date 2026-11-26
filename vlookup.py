import io
import streamlit as st
import pandas as pd

def read_file(uploaded_file, skiprows: int = 0):
    return pd.read_excel(uploaded_file, skiprows=skiprows)

def make_unique_with_base(right_df: pd.DataFrame, base_cols, key2, key1):
   
    existing = set(base_cols)
    new_cols = []
    for c in right_df.columns:
        if key1 == key2 and c == key2:
            # same-named join key: keep as-is; merge(on=key) will not duplicate it
            new_cols.append(c)
            continue

        new_name = c
        while new_name in existing:
            new_name = new_name + " "
        new_cols.append(new_name)
        existing.add(new_name)
    right_df = right_df.copy()
    right_df.columns = new_cols
    return right_df

def run_vlookup_mod():
    st.subheader("Vlookup File")

    with st.sidebar:
        st.markdown("### Upload Files")
        file1 = st.file_uploader("Select File 1", type=["xlsx", "xls"], key="f1")
        skip1 = st.number_input("Skip Rows (File1)", 0, 1000, 0, 1)

        file2 = st.file_uploader("Select File 2", type=["xlsx", "xls"], key="f2")
        skip2 = st.number_input("Skip Rows (File2)", 0, 1000, 0, 1)

    if file1 is None or file2 is None:
        st.info("Please upload both files.")
        st.stop()

    try:
        df1 = read_file(file1, skiprows=skip1)
        df2 = read_file(file2, skiprows=skip2)
    except Exception as e:
        st.error(f"Error reading files: {e}")
        st.stop()

    c1, c2 = st.columns(2)
    if file1 is not None and file1:
        with c1:
            st.markdown(f"### 📄 File 1: `{file1.name}` (skip {skip1})")
            st.dataframe(df1, width='stretch', height=320)
    if file2 is not None and file2:
        with c2:
            st.markdown(f"### 📄 File 2: `{file2.name}` (skip {skip2})")
            st.dataframe(df2, width='stretch', height=320)

    with st.form("vlookup_main_controls", clear_on_submit=False):
        cc1, cc2 = st.columns(2)
        with cc1:
            key1 = st.selectbox("Lookup Column in File 1", df1.columns)
        with cc2:
            key2 = st.selectbox("Match Column in File 2", df2.columns)
        cols_to_add = st.multiselect("Select Columns from File 2 to Add", df2.columns)
        vlookup_run = st.form_submit_button("Apply VLOOKUP")

    if vlookup_run:
       
        # always include key2 for merge; remove duplicates while keeping order
        right_cols = [key2] + list(cols_to_add)
        right_cols = list(dict.fromkeys(right_cols))
        right_df = df2[right_cols].copy()

        # Always rename overlaps before merging (no toggle)
        right_df = make_unique_with_base(right_df, df1.columns, key2, key1)

        if key1 == key2:
            merged = pd.merge(df1, right_df, how="left", on=key1)
        else:
            merged = pd.merge(df1, right_df, how="left", left_on=key1, right_on=key2)

        st.success("Vlookup Done ✅")
        st.dataframe(merged)

        # Downloads
        csv_bytes = merged.to_csv(index=False)
        csv_name = f"Vlookup {pd.Timestamp.now():%d-%m-%Y}.csv"
        excel_name = f"Vlookup {pd.Timestamp.now():%d-%m-%Y}.xlsx"

        # Build Excel in memory
        buffer = io.BytesIO()
        # Try xlsxwriter first; fall back to openpyxl if needed
        try:
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                merged.to_excel(writer, index=False, sheet_name="Merged")
        except Exception:
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                merged.to_excel(writer, index=False, sheet_name="Merged")
        
        row1,row2 = st.columns(2)
        with row1:
            st.download_button(
                "📥 Download Merged CSV",
                data=csv_bytes,
                file_name=csv_name,
                mime="csv"
            )
        with row2:
            st.download_button(
                "📥 Download Merged Excel",
                data=buffer.getvalue(),  # <-- bytes!
                file_name=excel_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    run_vlookup_mod()
