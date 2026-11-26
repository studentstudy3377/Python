from io import BytesIO
import os
import pandas as pd
import streamlit as st

# Function to extract the last two values from a string
def extract_last_two_values(value):
    parts = str(value).split('/')  # split the string by '/'
    return '/'.join(parts[-2:])  # join the last two parts

# Streamlit app
def sale_file():
    st.title("Data File Processor")

    with st.sidebar:
        # File upload section
        uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["xlsx", "csv"])

    if uploaded_file is not None:
        # Show a loading spinner while processing
        with st.spinner('Processing your file...'):
            # Check the file type based on the extension
            if uploaded_file.name.endswith('.xlsx'):
                df = pd.read_excel(uploaded_file)  # Read Excel file
            elif uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='latin1')  # or try 'ISO-8859-1', 'cp1252', etc.

            # Apply the function to every cell in the dataframe
            df = df.applymap(extract_last_two_values)

            # Show a preview of the first few rows
            st.subheader("Preview of the Modified Data")
            st.dataframe(df.head())

            # Create an in-memory buffer
            output_file = BytesIO()
            output_csv = BytesIO()

            # Save as Excel or CSV based on the input file type
            if uploaded_file.name.endswith('.xlsx'):
                df.to_excel(output_file, index=False)
                output_file.seek(0)  # Go to the beginning of the buffer
                st.success(f"Excel file ready for download.")

            elif uploaded_file.name.endswith('.csv'):
                df.to_csv(output_csv, index=False)
                output_csv.seek(0)  # Go to the beginning of the buffer
                st.success(f"CSV file ready for download.")

            # Provide a download link for Excel
            st.download_button(
                label="Download Excel file",
                data=output_file,
                file_name="Sale Files.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Provide a download link for CSV
            st.download_button(
                label="Download CSV file",
                data=output_csv,
                file_name="Sale Files.csv",
                mime="csv"
            )

if __name__ == "__main__":
    sale_file()
