import os
import warnings
from tkinter import Tk, filedialog, messagebox

import pandas as pd

warnings.simplefilter("ignore", UserWarning)


def select_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    try:
        file_path = filedialog.askopenfilename(
            initialdir=r"D:\Stock\Stock",
            title="Stock List",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )

        if file_path:
            process_data(file_path)
        else:
            messagebox.showwarning("Warning", "No file selected.")
    except Exception as e:
        messagebox.showerror("Error", f"Error selecting file:\n{e}")
    finally:
        try:
            root.destroy()
        except Exception:
            pass


def process_data(file_path: str):
    try:
        # Read data from CSV file
        data = pd.read_csv(file_path, encoding="windows-1252")

        # Remove extra line breaks in all columns
        for col in data.columns:
            data[col] = data[col].map(
                lambda x: x.replace("\n", " ") if isinstance(x, str) else x
            )

        # Remove unnecessary columns (ignore if missing)
        data = data.drop(
            columns=["IMAGES/VIDEO", "TRACR HYPERLINK", "DETAILS LINK"],
            errors="ignore",
        )

        # ---- Date rule:
        # Convert ONLY if date is like 31/12/2024 (contains '/')
        # If already like 31-Dec-2024, keep as-is
        if "Date" in data.columns:
            mask = data["Date"].astype(str).str.contains("/", na=False)

            data.loc[mask, "Date"] = pd.to_datetime(
                data.loc[mask, "Date"],
                dayfirst=True,
                errors="coerce",
            ).dt.strftime("%d-%b-%Y")

        # Replace Blank row with defaults
        fill_map = {
            "CULET": "None",
            "KEY TO SYMBOLS": "Other",
            "STATUS": "Other",
            "REPORT COMM": "Other",
            "ROUGH ORIGIN": "Other",
            "V360": "Other",
        }
        for col, default in fill_map.items():
            if col in data.columns:
                data[col] = data[col].fillna(default)

        # Generate file name using the (first row) Date
        if "Date" not in data.columns or data["Date"].isna().all():
            raise ValueError("Date column missing or all dates are invalid/blank.")

        # If first date is already 31-Dec-2024, we need it for filename as 12-31-24
        first_date = pd.to_datetime(data["Date"].iloc[0], errors="coerce")
        if pd.isna(first_date):
            raise ValueError(f"Invalid first date value: {data['Date'].iloc[0]}")

        save_file = first_date.strftime("%m-%d-%y")
        file_name = f"{save_file} Stock List.csv"

        # Save file paths
        save_directory_1 = r"D:\Data Backup\Gaurav\Set-Up  Files\Upload Data 2026\Stock List"
        save_directory_2 = r"D:\Data Backup\Gaurav\Set-Up  Files\Upload Data 2026\Stock List"

        save_path_1 = os.path.join(save_directory_1, file_name)
        save_path_2 = os.path.join(save_directory_2, file_name)

        # Save processed data to CSV
        data.to_csv(save_path_1, index=False)
        data.to_csv(save_path_2, index=False)

        print(f"Data saved to: {save_path_1}")
        messagebox.showinfo("Success", f"Data saved to:\n{save_path_1}")

    except Exception as e:
        messagebox.showerror("Error", f"Error processing data:\n{e}")

if __name__ == "__main__":
    select_file()