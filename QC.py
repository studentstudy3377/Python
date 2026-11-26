import os
from tkinter import Tk, filedialog, messagebox
import pandas as pd

def process_files():
    try:
        # Initialize Tkinter root
        root = Tk()
        root.withdraw()  # Hide the main window
        
        # Step 1: Read the first CSV file into a DataFrame
        dfs1 = filedialog.askopenfilename(
            initialdir=r'G:\Certificate\SAURABH SHAH\Mearge',
            title="QC",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not dfs1:
            raise FileNotFoundError("No file selected for the main file.")
        
        df1 = pd.read_excel(dfs1)
        
        # Step 2: Read the second CSV file into another DataFrame
        dfs2 = filedialog.askopenfilename(
            initialdir=r'G:\Certificate\SAURABH SHAH\Mearge',
            title="Select the QC File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not dfs2:
            raise FileNotFoundError("No file selected for the QC file.")
        
        df2 = pd.read_excel(dfs2)
        
        # Step 3: Merge df1 with a subset of df2 on the 'Diamond ID' column using a left join
        merged_df = pd.merge(df1, df2[['Diamond ID', 'QC']], on='Diamond ID', how='left')
        
        # Step 4: Replace NaN or blank values in 'QC' column with a placeholder (e.g., 'Blank')
        merged_df['QC'].fillna('Blank', inplace=True)
        
        # Step 5: Reorder columns to place 'QC' as the second column
        columns = merged_df.columns.tolist()
        
        if 'QC' in columns:
            columns.remove('QC')
            columns.insert(1, 'QC')
        
        merged_df = merged_df[columns]
        
        # Step 8: Group data by 'QC' column
        grouped = merged_df.groupby('QC')
        
        # Specify the output directory
        output_directory = r"G:\Certificate\Python Output"
        
        # Specify the output file path
        excel_name = os.path.join(output_directory, 'QC_sheets.xlsx')
        
        # Write to Excel with multiple sheets
        with pd.ExcelWriter(excel_name) as writer:
            for qc, group in grouped:
                sheet_name = f'{qc}' if qc != 'Blank' else 'Blank'
                group.to_excel(writer, sheet_name=sheet_name, index=False)  
        
        messagebox.showinfo("Success", f"File saved to:\n{excel_name}")
    
    except FileNotFoundError as e:
        messagebox.showerror("File Error", str(e))
    except pd.errors.EmptyDataError:
        messagebox.showerror("Data Error", "One of the CSV files is empty.")
    except pd.errors.MergeError:
        messagebox.showerror("Merge Error", "Error during merging the DataFrames.")
    except Exception as e:
        messagebox.showerror("Unexpected Error", str(e))
    finally:
        root.quit()

if __name__ == "__main__":
    process_files()
