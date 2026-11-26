import streamlit as st
import camelot
import pandas as pd
import os

def to_numeric_if_possible(val):
    """
    Try to clean and convert a string to int or float.
    Leaves non-numeric strings untouched.
    """
    if not isinstance(val, str):
        return val
    s = val.strip().replace(',', '')  # remove thousands separators
    # empty strings stay empty
    if s == "":
        return ""
    # try integer
    try:
        return int(s)
    except ValueError:
        pass
    # try float
    try:
        return float(s)
    except ValueError:
        return val
def dosser_run():
    # Streamlit app
    st.title("Dossires Report & 2.00UP Report PDF To Excel ")
    
    with st.sidebar:
        # File uploader
        pdf_file = st.file_uploader("Upload PDF", type="pdf")

    if pdf_file is not None:
        # Display progress bar for loading
        with st.spinner("Processing your PDF..."):
            try:
                # 2) Extract tables in stream mode
                tables = camelot.read_pdf(pdf_file, pages="all", flavor="stream")
            except Exception as e:
                st.error(f"Failed to parse PDF:\n{e}")
                st.stop()

            if not tables:
                st.error("No tables found in the PDF.")
                st.stop()

            # 3) Concatenate all tables and convert cells
            df = pd.concat([t.df for t in tables], ignore_index=True)
            
            # Use apply for each column, convert data types to numeric if possible
            for col in df.columns:
                df[col] = df[col].apply(to_numeric_if_possible)
            
            # Ensure consistent column types by converting to string or float explicitly
            # We'll convert all columns to float where applicable and string where not
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col], errors='raise')  # Try to convert numeric
                except Exception:
                    df[col] = df[col].astype(str)  # Convert the rest to string

            # Preview the tables before saving
            st.subheader("Extracted Tables Preview")
            st.write(df)  # Display first few rows of the combined table

            # Option to download the processed Excel file
            st.subheader("Download Processed Excel File")
            base_name = os.path.splitext(pdf_file.name)[0]
            xlsx_file = base_name + ".xlsx"

            # Save to Excel in memory and provide download link
            try:
                with pd.ExcelWriter(xlsx_file, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="All_Data", index=False, header=False)

                st.download_button(
                    label="Download Excel File",
                    data=open(xlsx_file, 'rb').read(),
                    file_name=xlsx_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Failed to write Excel:\n{e}")
                st.stop()

if __name__ == "__main__":
    dosser_run()