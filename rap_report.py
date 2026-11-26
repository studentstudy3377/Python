import pandas as pd
import os
import streamlit as st
from io import BytesIO

# Streamlit App
def rap_report():
    # Title for the Streamlit App
    st.title("Rap Report Processing")

    # File uploader for CSV files
    with st.sidebar:
        st.subheader("Upload CSV Files")
        csv_files = st.file_uploader("Upload CSV Files", type=["csv"], accept_multiple_files=True)

    # Check if the files are uploaded
    if csv_files:
        # Show processing indicator (loader)
        with st.spinner("Processing... Please wait."):

            # List to store all individual DataFrames
            dfs = []

            # Read and concatenate all the CSV files
            for csv_file in csv_files:
                df = pd.read_csv(csv_file)
                dfs.append(df)

            # Concatenate all DataFrames into one
            data = pd.concat(dfs, ignore_index=True)

            # Assuming you want to create an 'output' DataFrame similar to the structure in the error code
            output = data.copy()  # Create a copy of the data for processing

            # Initialize necessary columns in the output DataFrame
            output['AverageDiscountPercent'] = 0
            output['DiamondsInCategory'] = 0

            # Clean and convert 'AverageDiscountPercent' column in the 'data' DataFrame
            data['AverageDiscountPercent'] = data['AverageDiscountPercent'].str.rstrip('%').astype(float)
            output['AverageDiscountPercent'] = output['AverageDiscountPercent'].astype(float)
            output['DiamondsInCategory'] = output['DiamondsInCategory'].astype(float)

            # Loop through each row in the output DataFrame and update the count
            for i, match_row in output.iterrows():
                # Filter matching rows from 'data'
                matching_rows = data[(data['Clarity'] == match_row['Clarity']) &
                                      (data['Color'] == match_row['Color']) &
                                      (data['FromSize'] == match_row['FromSize']) &
                                      (data['ToSize'] == match_row['ToSize']) &
                                      (data['Shape'] == match_row['Shape'])]

                # If any rows match, update the 'AverageDiscountPercent' column in 'output'
                if not matching_rows.empty:
                    output.at[i, 'AverageDiscountPercent'] = matching_rows['AverageDiscountPercent'].sum()
                    output.at[i, 'DiamondsInCategory'] = matching_rows['DiamondsInCategory'].sum()

            # Show a preview of the processed data
            st.subheader("Preview of Processed Data:")
            st.write(output.head())  # Display the first 5 rows of the processed DataFrame

            # Save the processed DataFrame to a BytesIO buffer
            output_file = BytesIO()
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                output.to_excel(writer, index=False, sheet_name='ProcessedData')

            # Rewind the buffer to the beginning so it can be downloaded
            output_file.seek(0)

            # Show success message
            st.success("File processed successfully!")

            # Provide a download button for the processed file
            st.download_button(
                label="Download Processed File",
                data=output_file,
                file_name="Rap Report OUTPUT.xlsx",
                mime="application/vnd.ms-excel"
            )

# Run the function when the app starts
if __name__ == "__main__":
    rap_report()
