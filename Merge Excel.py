import os
import pandas as pd
from tkinter import Tk, filedialog


def merge_csv_files():
    root = Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="Merge CSV / Excel Files",
        filetypes=[
            ("Supported Files", "*.csv *.xlsx *.xls"),
            ("All Files", "*.*"),
        ],
    )

    root.destroy()

    if not files:
        print("No files selected.")
        return

    dfs = []

    for file in files:
        extension = os.path.splitext(file)[1].lower()

        if extension in (".xlsx", ".xls"):
            # Prevent text such as "None" from being interpreted as a missing value
            df = pd.read_excel(file, keep_default_na=False)

        elif extension == ".csv":
            # Prevent text such as "None" from being interpreted as a missing value
            df = pd.read_csv(file, keep_default_na=False)

        else:
            continue

        dfs.append(df)

    if not dfs:
        print("No valid CSV or Excel files were selected.")
        return

    # Merge all files
    merged_df = pd.concat(dfs, ignore_index=True)

    # Convert any remaining actual None/NaN values to the text "None"
    merged_df = merged_df.fillna("None")

    output_directory = r"G:\Certificate\Python Output"
    output_filename = "Merge.xlsx"
    output_path = os.path.join(output_directory, output_filename)

    os.makedirs(output_directory, exist_ok=True)

    merged_df.to_excel(output_path, index=False)

    os.startfile(output_path)


if __name__ == "__main__":
    merge_csv_files()