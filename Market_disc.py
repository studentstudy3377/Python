import datetime
import os
from tkinter import filedialog
import numpy as np
import pandas as pd

file = filedialog.askopenfilename(title="Market Files", filetypes=[("CSV Files", "*.csv")])
files = filedialog.askopenfilename(title="STOCK Files", filetypes=[("CSV Files", "*.csv")])

df = pd.read_csv(file)
dfs = pd.read_csv(files)

df = df[df['Country'].isin(["Hong Kong","India"])]

df = df[df['Lab'] == 'GIA']

fin = (df['Finish'].astype(str).str.strip().str.upper())

# --- Define mapping: Finish -> (POL, SYM) ---
finish_to_pol_sym = {
    "3X": ("EX", "2EX"),
    "EX":  ("VG", "2EX"),
    "VG+": ("VG", "2VG"),
    "GD+": ("GD", "2GD"),
}

mapped = fin.map(finish_to_pol_sym)
df["POL"] = mapped.apply(lambda t: t[0] if isinstance(t, tuple) else None)
df["SYM"] = mapped.apply(lambda t: t[1] if isinstance(t, tuple) else None)

shape = {
    "Round": "RD",
    "Emerald": "EME",
    "Heart": "H",
    "Marquise": "MQ",
    "Oval": "OV",
    "Princess": "PRI",
    "Radiant": "RDT",
    "Pear": "PRS",
}

shape_col = df['Shape'].astype(str).str.strip().str.title()
mapped_shape = shape_col.map(shape)
df["Shapes"] = mapped_shape

fluor = {
    "None": "N",
    "Faint": "F",
    "Medium": "M",
    "Strong": "S",
    "Very Strong": "S",
    "Slight": "SL",
    "Very Slight": "SL",
}

fluor_col = df['FluorescenceIntensity'].astype(str).str.strip().str.title()
mapped_fluor = fluor_col.map(fluor)
df["Fluor"] = mapped_fluor
df["Fluor"] = fluor_col.map(fluor).fillna("N")

carat_categories = {
    (0.23, 0.29): "0.30DN",
    (0.30, 0.36): "0.30UP",
    (0.37, 0.39): "0.70UP",
    (0.40, 0.45): "0.40UP",
    (0.46, 0.49): "0.46UP",
    (0.50, 0.59): "0.50UP",
    (0.60, 0.69): "0.60UP",
    (0.70, 0.79): "0.70UP",
    (0.80, 0.89): "0.80UP",
    (0.90, 0.95): "0.90UP",
    (0.96, 0.99): "0.96UP",
    (1.00, 1.19): "1.00UP",
    (1.20, 1.24): "1.20UP",
    (1.25, 1.34): "1.25UP",
    (1.35, 1.49): "1.35UP",
    (1.50, 1.59): "1.50UP",
    (1.60, 1.69): "1.60UP",
    (1.70, 1.89): "1.70UP",
    (1.90, 1.99): "1.90UP",
    (2.00, 2.19): "2.00UP",
    (2.00, 2.19): "2.00UP",
    (2.20, 2.49): "2.20UP",
    (2.50, 2.69): "2.50UP",
    (2.70, 2.89): "2.70UP",
    (2.90, 2.99): "2.90UP",
    (3.00, 3.49): "3.00UP",
    (3.50, 3.99): "3.50UP",
    (4.00, 4.49): "4.00UP",
    (4.50, 4.99): "4.50UP",
    (5.00, 5.49): "5.00UP",
    (5.50, 5.99): "5.50UP",
}

intervals = pd.IntervalIndex.from_tuples(list(carat_categories.keys()), closed="both")
labels = list(carat_categories.values())

lower = df["Size"].str.extract(r"^\s*([\d.]+)\s*-\s*")[0].astype(float)

idx = intervals.get_indexer(lower)
df["RAPSIZE"] = np.where(idx >= 0, np.array(labels, dtype=object)[idx], None)

# Create a new column 'Concatenate' with all values joined in your desired order
df["Concatenate"] = (
    df["RAPSIZE"].astype(str) + "-" +
    df["Shapes"].astype(str) + "-" +
    df["Color"].astype(str) + "-" +
    df["Clarity"].astype(str) + "-" +
    df["Fluor"].astype(str) + "-" +
    df["POL"].astype(str) + "-" +
    df["SYM"].astype(str)
)

# --- build the right table for merge ---
right_cols = ["Concatenate", "AvgRapListPercent25", "AvgRapListPercent50", "AvgRapListPercent100", "AvgRapListPercent200"]

# safety: ensure those columns exist in df
missing = [c for c in right_cols if c not in df.columns]
if missing:
    raise KeyError(f"Missing in df: {missing}")

right = df[right_cols].copy()

# normalize join keys
right["Concatenate"] = right["Concatenate"].astype(str).str.strip()
dfs["Qlty"] = dfs["Qlty"].astype(str).str.strip()

# --- LEFT MERGE and ASSIGN back ---
dfs = dfs.merge(right, left_on="Qlty", right_on="Concatenate", how="left")

# Ensure the relevant columns are numeric, handling any invalid data by coercing errors to NaN
columns_to_sum = ['AvgRapListPercent25', 'AvgRapListPercent50', 'AvgRapListPercent100', 'AvgRapListPercent200']
dfs[columns_to_sum] = dfs[columns_to_sum].apply(pd.to_numeric, errors='coerce')

# Optionally, if you only want to keep the calculated average for each SrNo, you can group and get the mean
average_df = dfs.groupby('SrNo', as_index=False)[['AvgRapListPercent25', 'AvgRapListPercent50', 'AvgRapListPercent100', 'AvgRapListPercent200']].mean()

# Merge the calculated averages back into the original dataframe (if you need them alongside other data)
dfs = dfs.merge(average_df, on='SrNo', suffixes=('', ' '))

# Remove the specified columns from the final DataFrame
dfs = dfs.drop(columns=right_cols, errors='ignore')

# Drop duplicates based on 'SrNo' and keep the first occurrence
dfs = dfs.drop_duplicates(subset='SrNo', keep='first')

# Save the updated DataFrame to Excel
today_str = datetime.datetime.now().strftime("%d-%m-%Y")
output_avg = fr"G:\Certificate\Python Output\Market {today_str}.xlsx"
dfs.to_excel(output_avg, index=False)

# Open the output file
os.startfile(output_avg)