import datetime
import os
from tkinter import filedialog
import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter

# Load the data from the Excel files
stock_file_path = filedialog.askopenfilename( title="Select Stock File", filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")], )
price_file_path = r"G:\Certificate\HEMIL\PRICE DISC SEARCH\PRICE UPDATES.xlsx"

df_stock = pd.read_excel(stock_file_path)
df_disc = pd.read_excel(price_file_path)

# 1. Mask conditions to filter out rows based on the given criteria
mask_gd = (df_stock["POL"] == "GD") | (df_stock["SYM"] == "GD")
mask_hrd = df_stock["CERT"] == "HRD"
mask_fluor = df_stock["FLUOR"] == "NONE"

# 2. Apply the combined mask to filter the dataframe
df_stock = df_stock.loc[~(mask_gd | mask_hrd) & mask_fluor].reset_index(drop=True)
# df_stock = df_stock.loc[~(mask_gd | mask_hrd)].reset_index(drop=True)

# 3. Identify EX/EX rows and assign CUT1 and CUT2 values in one combined column
mask_ex = (df_stock["POL"] == "EX") & (df_stock["SYM"] == "EX")
df_stock["CUT"] = np.where(mask_ex, "2EX", "VG")  # If EX/EX, assign "2EX", else assign "VG"

# 4. Create the "Concnate" column by concatenating the specified columns with ' - ' separator
cols = ["SHAPE", "RAPSIZE", "COLOR", "CLARITY", "CUT"]
df_stock["Concnate"] = df_stock[cols].fillna(" ").astype(str).agg(" - ".join, axis=1)

# 5. Perform the join with df_disc to get the "Disc1" and "Disc2" columns based on matching "Concnate" and "CON1"/"CON2"
df_stock = df_stock.merge(df_disc[['CON1', 'VG DISC']], left_on='Concnate', right_on='CON1', how='left')
df_stock = df_stock.merge(df_disc[['CON2', '2EX DISC']], left_on='Concnate', right_on='CON2', how='left')

# 6. Create a new column "New Disc" by combining "Disc1" and "Disc2"
df_stock["New Disc"] = df_stock["VG DISC"].fillna('').astype(str) + ' ' + df_stock["2EX DISC"].fillna('').astype(str)

# Convert CERTNO and New Disc to numeric, coercing errors
df_stock['CERTNO'] = pd.to_numeric(df_stock['CERTNO'], errors='coerce')
df_stock['New Disc'] = pd.to_numeric(df_stock['New Disc'], errors='coerce')

# 7. Drop rows without a discount
df_stock.dropna(subset=["New Disc"], inplace=True)

# Add the New Carat column using RAPSIZE
clean_size = df_stock["RAPSIZE"].str.replace("UP", "", regex=False)
size_vals = clean_size.astype(float)
new_carat = np.where(~clean_size.str.contains(r"\."), size_vals / 100, size_vals)
idx_carat = df_stock.columns.get_loc("CARAT") + 1
df_stock.insert(idx_carat, "New Carat", new_carat)

# Desired column order
desired_order = [
    "STOCKID", "SHAPE", "RAPSIZE", "CARAT", "New Carat", "COLOR", "CLARITY", "CERT",
    "CUT", "POL", "SYM", "FLUOR", "MEASUREMENT", "RAPPRICE", "RAP V", "DISC", "PRICE",
    "VALUE", "New Disc", "New Price", "New Value", "Diff Value", "Diff %", "CERTNO",
    "DEPTH", "TABLE", "RATIO", "BLACK", "TINCH", "LUSTER", "CA", "CH", "PA", "PH",
    "GIRDLE", "GIRDLE%", "CULET", "KEY TO SYMBOLS", "REPORT COMM", "STATUS", "SECTION NO",
    "ROUGH ORIGIN", "DIGITAL CERT", "IMAGES/VIDEO", "V360", "STONE LOCATION", "TRACR HYPERLINK",
    "DETAILS LINK"
]
# Add missing columns with None as value  ,"Concnate","2EX DISC","VG DISC"

for col in desired_order:
    if col not in df_stock.columns:
        df_stock[col] = None

# Reorder the dataframe to the desired column order
df_stock = df_stock[desired_order]

# Apply the Excel formula to the "New Price" column
col_rappr_idx = df_stock.columns.get_loc("RAPPRICE") + 1  # Excel column index for RAPPRICE
col_disc_new_idx = df_stock.columns.get_loc("New Disc") + 1  # Excel column index for New Disc

# Generate the Excel formula for New Price
new_price_fm = [
    f"={get_column_letter(col_rappr_idx)}{r}*(1+{get_column_letter(col_disc_new_idx)}{r}/100)"
    for r in range(2, len(df_stock) + 2)
]

# Add the formulas to the "New Price" column in the DataFrame
df_stock["New Price"] = new_price_fm

# Apply the Excel formula to the "New Price" column
rappr_idx = df_stock.columns.get_loc("RAPPRICE") + 1  # Excel column index for RAPPRICE
col_disc_idx = df_stock.columns.get_loc("DISC") + 1  # Excel column index for New Disc

# Generate the Excel formula for New Price
price_fm = [
    f"={get_column_letter(rappr_idx)}{r}*(1+{get_column_letter(col_disc_idx)}{r}/100)"
    for r in range(2, len(df_stock) + 2)
]

# Add the formulas to the "New Price" column in the DataFrame
df_stock["PRICE"] = price_fm

# FIX: remove existing "VALUE" column before reinserting it
df_stock.drop("VALUE", axis=1, inplace=True)

# Get the column letters for CARAT and PRICE
col_carat = get_column_letter(df_stock.columns.get_loc("CARAT") + 1)
col_price_o = get_column_letter(df_stock.columns.get_loc("PRICE") + 1)

# Generate the Excel formula for VALUE
value_fm = [
    f"={col_carat}{r}*{col_price_o}{r}"
    for r in range(2, len(df_stock) + 2)
]

# Insert the new "VALUE" column with the calculated formulas
idx_val = df_stock.columns.get_loc("PRICE") + 1
df_stock.insert(idx_val, "VALUE", value_fm)

col_np2 = get_column_letter(df_stock.columns.get_loc("New Price") + 1)  # Excel column letter for New Price
col_ncar2 = get_column_letter(df_stock.columns.get_loc("New Carat") + 1)  # Excel column letter for New Carat

# Generate the Excel formula for New Value
new_value_fm = [
    f"={col_np2}{r}*{col_ncar2}{r}"
    for r in range(2, len(df_stock) + 2)
]

# Insert the new "New Value" column with the calculated formulas
df_stock["New Value"] = new_value_fm

# Insert the new "Diff Value" column with the calculated formulas

col_nv3     = get_column_letter(df_stock.columns.get_loc("New Value") + 1)
col_val3    = get_column_letter(df_stock.columns.get_loc("VALUE")     + 1)
diff_val_fm = [
        f"={col_nv3}{r}-{col_val3}{r}"
        for r in range(2, len(df_stock)+2)
    ]

df_stock["Diff Value"] = diff_val_fm

# Get the Excel column letters for "Diff Value" and "VALUE"
col_dv4 = get_column_letter(df_stock.columns.get_loc("Diff Value") + 1)  # Excel column letter for Diff Value
col_val4 = get_column_letter(df_stock.columns.get_loc("VALUE") + 1)  # Excel column letter for VALUE

# Generate the Excel formula for Diff % with rounding and percentage formatting
diff_pct_fm = [
    f"=IFERROR(ROUND(({col_dv4}{r}/{col_val4}{r})*100, 0), 0) & \"%\""
    for r in range(2, len(df_stock) + 2)
]

# Insert the "Diff %" column with the calculated formulas
df_stock["Diff %"] = diff_pct_fm

# Get the start and end row indices
start_row = 2
end_row = len(df_stock) + 1

# Get the Excel column letters for "VALUE", "New Value", and "Diff Value"
letters = {
    "VALUE": get_column_letter(df_stock.columns.get_loc("VALUE") + 1),
    "New Value": get_column_letter(df_stock.columns.get_loc("New Value") + 1),
    "Diff Value": get_column_letter(df_stock.columns.get_loc("Diff Value") + 1),
}

# Initialize a dictionary to hold the subtotal formulas for each column
subtotal = {c: "" for c in df_stock.columns}

# Define the subtotal formulas for the relevant columns
subtotal["VALUE"] = f"=SUBTOTAL(9,{letters['VALUE']}{start_row}:{letters['VALUE']}{end_row})"
subtotal["New Value"] = f"=SUBTOTAL(9,{letters['New Value']}{start_row}:{letters['New Value']}{end_row})"
subtotal["Diff Value"] = f"=SUBTOTAL(9,{letters['Diff Value']}{start_row}:{letters['Diff Value']}{end_row})"

# Calculate the "Diff %" subtotal formula based on the new "Diff Value" and "VALUE" subtotals
subtotal["Diff %"] = f"=IFERROR(ROUND({letters['Diff Value']}{end_row + 1}/{letters['VALUE']}{end_row + 1} * 100, 0), 0) & \"%\""

# Append the subtotal row to the DataFrame
df_stock = pd.concat([df_stock, pd.DataFrame([subtotal])], ignore_index=True)

# Create the output filename based on the current date
today_str = datetime.datetime.now().strftime("%d-%m-%Y")
output_filename = f"G:\\Certificate\\Python Output\\Stock List LEDS {today_str}.xlsx"
# Save the dataframe to an Excel file
df_stock.to_excel(output_filename, index=False)
# Open the output file
os.startfile(output_filename)
