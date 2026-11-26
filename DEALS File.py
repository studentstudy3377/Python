
import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox

# Initialize Tkinter
root = Tk()
root.withdraw()

# Open file dialog
file = filedialog.askopenfilename(
    title="Stock List files",
    filetypes=[("Excel files", "*.xlsx")]
)

# If user cancels, stop execution
if not file:
    messagebox.showwarning("Cancelled", "No file was selected.")
    exit()

# Read Excel file
df = pd.read_excel(file)

# Convert columns to numeric
df["RAPPRICE"] = pd.to_numeric(df["RAPPRICE"], errors="coerce")
df["DISC"] = pd.to_numeric(df["DISC"], errors="coerce")
df["CARAT"] = pd.to_numeric(df["CARAT"], errors="coerce")

# Add 3 to DISC
df["DISC"] = df["DISC"] + 3

# Price calculation
df["PRICE"] = df["RAPPRICE"] + (df["RAPPRICE"] * df["DISC"] / 100)

# Value calculation
df["VALUE"] = df["PRICE"] * df["CARAT"]

# Create new output file name
base, ext = os.path.splitext(file)
new_file = f"{base} - Done{ext}"

# Save to new Excel file
df.to_excel(new_file, index=False)

# Success message box
messagebox.showinfo(
    "Success",
    f"{new_file}"
)