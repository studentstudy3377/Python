import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox
import warnings

# Suppress specific warnings related to openpyxl
warnings.simplefilter("ignore", category=UserWarning)

def merge_files():
    try:
        # Create a Tkinter root widget without displaying the GUI window
        root = Tk()
        root.withdraw()
        
        # Open file dialog to select multiple CSV or Excel files
        files = filedialog.askopenfilenames(
            title="Merge Files", 
            filetypes=[("All Files", "*.*"), ("CSV Files", "*.csv"), ("Excel Files", "*.xlsx;*.xls")]
        )
        
        # Convert the returned tuple of filepaths to a list
        files_list = list(files)
        
        # Initialize an empty list to store DataFrames
        dfs = []
        
        # Iterate through selected files and read them into Pandas DataFrames
        for file in files_list:
            if file.endswith('.csv'):
                # For CSV, skip first 10 rows and ignore the last row by calculating the number of rows
                df = pd.read_csv(file, skiprows=10)
                df = df.iloc[:-1]  # Remove the last row
            elif file.endswith('.xlsx') or file.endswith('.xls'):
                # For Excel, skip first 10 rows and ignore the last row by calculating the number of rows
                df = pd.read_excel(file, skiprows=10)
                df = df.iloc[:-1]  # Remove the last row
            else:
                continue  # Skip unsupported file types
            
            # Append the DataFrame to the list
            dfs.append(df)
        
        # Concatenate all DataFrames into a single DataFrame
        merged_df = pd.concat(dfs, ignore_index=True)
        
        # Fill any NaN values in 'Fluorescence' column with 'None' (if it exists)
        if 'Fluorescence' in merged_df.columns:
            merged_df['Fluorescence'] = merged_df['Fluorescence'].fillna('None')
        
        # Specify the output file name
        output_filename = r'G:\Certificate\Python Output\Merged_File.xlsx'  # Output file in Excel format
        
        # Save the concatenated DataFrame to an Excel file
        merged_df.to_excel(output_filename, index=False)
        
        # Show a message box indicating successful merge
        messagebox.showinfo("Merge Complete", "Files merged successfully!")
        
    except Exception as e:
        # Display an error message if any exception occurs
        messagebox.showerror("Error", f"Error during merge: {str(e)}")

if __name__ == "__main__":
    merge_files()
