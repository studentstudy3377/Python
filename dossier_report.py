import pandas as pd
import os
import streamlit as st
from io import BytesIO
import time

# Streamlit App
def dossier_report():
    # Streamlit File Uploader for CSV files
    st.title("DOSSIERS Report Processing")
    with st.sidebar:
        # Upload the CSV files
        csv_files = st.file_uploader("Upload CSV Files", type=["csv"], accept_multiple_files=True)

    # If files are uploaded
    if csv_files:
        st.write("Processing files...")

        # Show progress bar
        with st.spinner("Processing... Please wait."):
            time.sleep(2)  # Simulating a delay for processing

            # List to store all the individual DataFrames
            dfs = []

            # Read all uploaded CSV files
            for csv_file in csv_files:
                df = pd.read_csv(csv_file)
                dfs.append(df)

            # Concatenate all CSV files into one DataFrame
            if dfs:
                data = pd.concat(dfs, ignore_index=True)

            # Initialize the 'AverageDiscountPercent' column in the match file to 0 if it doesn't exist
            if 'AverageDiscountPercent' not in data.columns:
                data['AverageDiscountPercent'] = 0

            # Ensure that the 'AverageDiscountPercent' column is treated as string and remove the '%' symbol
            # First, convert to string (if not already) and remove the '%' sign
            data['AverageDiscountPercent'] = data['AverageDiscountPercent'].astype(str).str.rstrip('%')

            # Now, convert the 'AverageDiscountPercent' to a float
            data['AverageDiscountPercent'] = pd.to_numeric(data['AverageDiscountPercent'], errors='coerce')

            # Loop through each row and update the count
            for i, match_row in data.iterrows():
                # Filter the rows in 'data' that match all specified conditions
                matching_rows = data[(data['Clarity'] == match_row['Clarity']) & 
                                      (data['Color'] == match_row['Color']) & 
                                      (data['FromSize'] == match_row['FromSize']) & 
                                      (data['ToSize'] == match_row['ToSize']) & 
                                      (data['Shape'] == match_row['Shape'])]

                # If any rows match, update the 'AverageDiscountPercent' column in the 'data' dataframe
                if not matching_rows.empty:
                    data.at[i, 'AverageDiscountPercent'] = matching_rows['AverageDiscountPercent'].sum()

            # Show a preview of the processed data
            st.subheader("Preview of Processed Data:")
            st.write(data)

            # Save the updated data to a BytesIO buffer
            output_file = BytesIO()
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                data.to_excel(writer, index=False, sheet_name='ProcessedData')

            # Rewind the buffer to the beginning so it can be downloaded
            output_file.seek(0)

            # Show success message
            st.success("File saved successfully!")

            # Provide download button for the processed file
            st.download_button(
                label="Download Processed File",
                data=output_file,
                file_name="DOSSIERS REPORT OUTPUT.xlsx",
                mime="application/vnd.ms-excel"
            )

# Run the function when the app starts
if __name__ == "__main__":
    dossier_report()
