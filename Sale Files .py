import os
from tkinter import filedialog
import pandas as pd

# Load the file (Excel or CSV)
file_path = filedialog.askopenfilename(title="Select Files", filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])

# Check if the file is Excel or CSV based on extension
if file_path.endswith('.xlsx'):
    df = pd.read_excel(file_path)  # Read Excel file
elif file_path.endswith('.csv'):
    # Try reading the CSV with a different encoding (e.g., 'latin1' or 'ISO-8859-1')
    df = pd.read_csv(file_path, encoding='latin1')  # or try 'ISO-8859-1', 'cp1252', etc.
else:
    print("Invalid file format. Please select a valid Excel or CSV file.")
    exit()

# Define a function to extract the last two values from a string
def extract_last_two_values(value):
    parts = str(value).split('/')  # split the string by '/'
    return '/'.join(parts[-2:])  # join the last two parts

# Apply the function to every cell in the dataframe
df = df.applymap(extract_last_two_values)

# Save the modified data to a new file
output_dir = r"G:\Certificate\Python Output"
output = os.path.join(output_dir, f"Sale Files .xlsx")
output_csv = os.path.join(output_dir, f"Sale Files .csv")

if file_path.endswith('.xlsx'):
    df.to_excel(output, index=False)

elif file_path.endswith('.csv'):
    df.to_csv(output, index=False)

os.startfile(output)