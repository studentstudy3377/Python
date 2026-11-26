import pandas as pd
from tkinter import filedialog, Tk, messagebox
import os

# Function to handle file selection
def select_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    try:
        # Ask user to input file path from specific directory
        file_path = filedialog.askopenfilename(
            initialdir=r'D:\Stock\Data',
            title="Final",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            process_data(file_path)
        else:
            print("Warning", "No file selected.")
    except Exception as e:
        print("Error", f"Error selecting file: {str(e)}")

# Function to process data
def process_data(file_path):
    try:
        # Read data from CSV file
        data = pd.read_csv(file_path)

        # Convert "ReportDate" to formatted string
        data['ReportDate'] = pd.to_datetime(data['ReportDate']).dt.strftime('%d-%b-%y')

        # Fill NaN values in "FluorescenceIntensity" column with 'None'
        data['FluorescenceIntensity'] = data['FluorescenceIntensity'].fillna('None')
       
        # Define save file path
        save_directory = r"D:\Data Backup\Gaurav\Set-Up  Files\Upload Data 2026\Final"
        save = r"D:\Data Backup\Gaurav\Set-Up  Files\Upload Data 2026\Final"
        
        # Generate file name using the formatted date
        date = pd.to_datetime(data['ReportDate']).dt.strftime('%m-%d-%y')
        save_file = date.iloc[0]  
        
        # Using the first date as part of the file name
        file_name = f"{save_file} Final.csv"
        save_path = os.path.join(save_directory, file_name)
        save_paths = os.path.join(save, file_name)

        # Save processed data to CSV file
        data.to_csv(save_path, index=False)
        data.to_csv(save_paths, index=False)

        print(f"Data saved to: {save_path}")
        messagebox.showinfo("Success", f"Data saved to:\n{save_path}")

    except Exception as e:
       print("Error", f"Error processing data: {str(e)}")

# Main GUI part
if __name__ == "__main__":
    select_file()
