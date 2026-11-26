import tkinter as tk
from tkinter import filedialog, messagebox
import camelot
import pandas as pd
import os

def to_numeric_if_possible(val):
    """
    Try to clean and convert a string to int or float.
    Leaves non-numeric strings untouched.
    """
    if not isinstance(val, str):
        return val
    s = val.strip().replace(',', '')  # remove thousands separators
    # empty strings stay empty
    if s == "":
        return ""
    # try integer
    try:
        return int(s)
    except ValueError:
        pass
    # try float
    try:
        return float(s)
    except ValueError:
        return val

if __name__ == "__main__":
    # 1) Let user pick a PDF
    root = tk.Tk()
    root.withdraw()
    pdf_file = filedialog.askopenfilename(
        title="Select PDF to convert",
        filetypes=[("PDF files","*.pdf")]
    )
    if not pdf_file:
        messagebox.showinfo("Cancelled", "No file selected.")
        exit()

    # 2) Extract tables in stream mode
    try:
        tables = camelot.read_pdf(pdf_file, pages="all", flavor="stream")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to parse PDF:\n{e}")
        exit()

    if not tables:
        messagebox.showerror("Error", "No tables found in that PDF.")
        exit()

    # 3) Concatenate all tables and convert cells
    df = pd.concat([t.df for t in tables], ignore_index=True)
    df = df.applymap(to_numeric_if_possible)

    # 4) Write to Excel alongside the PDF
    base, _ = os.path.splitext(pdf_file)
    xlsx_file = base + ".xlsx"
    try:
        with pd.ExcelWriter(xlsx_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="All_Data", index=False, header=False)
        os.startfile(xlsx_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write Excel:\n{e}")
        exit()
