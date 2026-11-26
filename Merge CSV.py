import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox

def merge_csv_files():
    try:
        # Create a Tkinter root widget without displaying the GUI window
        root = Tk()
        root.withdraw()
        
        # Open file dialog to select multiple CSV files
        files = filedialog.askopenfilenames(
            title="Merge CSV Files", filetypes=[("All Files","*.*")]
        )
        
        # Convert the returned tuple of filepaths to a list
        files_list = list(files)
        
        # Initialize an empty list to store DataFrames
        dfs = []
        
        # Iterate through selected files and read them into Pandas DataFrames
        for file in files_list:
            if file.endswith('.xlsx') or file.endswith('xls'):
                df = pd.read_excel(file,keep_default_na=False)
            elif file.endswith('.csv'):
                df = pd.read_csv(file,keep_date_col=False)
            else:
                continue
            dfs.append(df)
        
        # Concatenate all DataFrames into a single DataFrame
        df = pd.concat(dfs, ignore_index=True)
       
        # Specify the output directory and file name
        output_directory = r"G:\Certificate\Python Output"
        output_filename = 'Merge.csv'
        output_path = os.path.join(output_directory, output_filename)
        
        # Save the concatenated DataFrame to a CSV file
        df.to_csv(output_path, index=False)
        
    except Exception as e:
        # Display an error message if any exception occurs
        messagebox.showerror("Error", f"Error during merge: {str(e)}")

    os.startfile(output_path)

if __name__ == "__main__":
    merge_csv_files()