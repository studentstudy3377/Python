import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox

def merge_and_clean_csv():
    try:
        # Create a Tkinter root widget without displaying the GUI window
        root = Tk()
        root.withdraw()
        
        # Open file dialog to select multiple CSV files
        files = filedialog.askopenfilenames(
            title="Filter Select CSV files", filetypes=[("CSV files", "*.csv")]
        )
        
        # Convert the returned tuple of filepaths to a list
        files_list = list(files)
        
        # Initialize an empty list to store DataFrames
        dfs = []
        
        # Iterate through selected files and read them into Pandas DataFrames
        for file in files_list:
            combined_df = pd.read_csv(file)
            dfs.append(combined_df)
        
        # Concatenate all DataFrames into a single DataFrame
        df = pd.concat(dfs, ignore_index=True)
        
        # Fill NaN values in 'Fluorescence' column with 'None'
        df['Fluorescence'] = df['Fluorescence'].fillna('None')
        
        # Filter based on 'Fluorescence'
        fluorescence_values = ['Faint', 'Medium', 'MediumBlue', 'None', 'Slight', 'SlightBlue']
        df = df[df['Fluorescence'].isin(fluorescence_values)]
        
        # Replace values in 'Cut' column
        df['Cut'] = df['Cut'].replace(['Excellent', 'Ideal'], 'EX').replace('Very Good', 'VG').replace(['Good','G'], 'GD')
        
        # Filter based on 'Culet'
        culet_values = ['Large ', 'Very Small ', 'Medium ', 'Small ']
        df = df[~df['Culet'].isin(culet_values)]

        # Fill NaN values in 'Depth', 'Table', 'City' and 'Measurements' columns with 0 and blank
        df['Depth'] = df['Depth'].fillna(0)
        df['City'] = df['City'].fillna("Blank")
        df['Table'] = df['Table'].fillna(0)
        df['Measurements'] = df['Measurements'].fillna(0) 
        
        # Filter based on 'City'
        city_values = ['Mumbai', 'Bombay', 'SURAT', 'Blank']
        df = df[df['City'].isin(city_values)]
                
        # Filter based on 'Country' and 'Lab Comment' and 'Member Comment'
        df = df[df['Country'] == 'India']
        df = df[~df['Lab Comment'].str.contains('Clarity grade ', na=False)]
        df = df[~df['Member Comment'].str.contains('Only for Export Sales,No Local Delivery', na=False)]
        df = df[~df['Member Comment'].str.contains('No BGM,Only for Export Sales,No Local Delivery', na=False)]
        
        # Replace values in 'Cut' column
        df['City'] = df['City'].replace('Blank', '')

        # Drop duplicates based on 'Vendor Stock number'
        df = df.drop_duplicates(subset='Vendor Stock number')
        
        # Specify the output directory
        output_directory = r"G:\Certificate\Python Output"
        
        # Ensure the directory exists; create it if necessary
        os.makedirs(output_directory, exist_ok=True)
        
        # Specify the output file path with .csv extension
        output_filename = os.path.join(output_directory, 'Merge_With_Filter.csv')
        
        # Save the cleaned DataFrame to a CSV file
        df.to_csv(output_filename, index=False)
        
        # Show a message box indicating successful merge
        messagebox.showinfo("Merge Complete", f"CSV files merged and cleaned successfully! Saved as '{output_filename}'")
    
    except Exception as e:
        # Display an error message if any exception occurs
        messagebox.showerror("Error", f"Error during merge and cleaning: {str(e)}")

if __name__ == "__main__":
    merge_and_clean_csv()
