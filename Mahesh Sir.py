import platform
import subprocess
import pandas as pd
import os
from tkinter.filedialog import askopenfilename

# Method to open a file
def open_file(file_path):
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":
            subprocess.call(["open", file_path])
        else:
            subprocess.call(["xdg-open", file_path])
    except Exception as e:
        print(f"Error opening file: {e}")

# Ask user to select the Excel file
file_path = askopenfilename(title="Mahesh Sir", filetypes=[("Excel files", "*.xlsx *.xls")])

# Read the selected Excel file into a pandas DataFrame
data = pd.read_excel(file_path)

# Columns to keep
columns_to_keep = ['QualityCode', 'FOR', 'Size', 'Pcs', 'Carats', 'Price', 'Cost', 'Disc 1%', 'Disc 2%', 'Add %', 'Remark', 'stoneid', 'PARTY NAME']
data = data[columns_to_keep]

# Update 'Price' column based on 'Cost'
data['Price'] = data.apply(lambda row: row['Cost'] if pd.notnull(row['Cost']) else row['Price'], axis=1)

# Ensure the specified columns are blank
columns_to_clear = ['Cost', 'Disc 1%', 'Disc 2%', 'Add %','FOR', 'Size', 'Pcs','Remark']
data[columns_to_clear] = ''

# Get unique PARTY NAMES
unique_party_names = data['PARTY NAME'].unique()

# Loop through each unique PARTY NAME
for party_name in unique_party_names:
    party_data = data[data['PARTY NAME'] == party_name]

    if len(party_data) > 1:
        # Drop the 'PARTY NAME' column
        party_data = party_data.drop(columns=['PARTY NAME'])

        # Sanitize the PARTY NAME for valid file name
        safe_party_name = party_name.replace(' ', ' ').replace('/', ' ').replace('\\', ' ')

        # Define the output file path
        output_filename = f"G:\\Certificate\\Python Output\\{safe_party_name}.csv"

        # Ensure the directory exists
        os.makedirs(os.path.dirname(output_filename), exist_ok=True)

        # Save as .csv file
        party_data.to_csv(output_filename, index=False)

        print(f"Saved file for PARTY NAME: {party_name} as {output_filename}")

        # Open the file after saving
        open_file(output_filename)
