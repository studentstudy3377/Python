import os
from pathlib import Path
import pandas as pd
from tkinter import filedialog

# Output file path
OUTPUT_PATH = Path(r"G:\Certificate\Python Output\Market Price.xlsx")

# Size bins for assigning Index & RAP SIZE
SIZE_RANGES = [
    (0.23, 0.29), (0.30, 0.39), (0.40, 0.49), (0.50, 0.59),
    (0.60, 0.69), (0.70, 0.79), (0.80, 0.89), (0.90, 0.99),
    (1.00, 1.19), (1.20, 1.29), (1.30, 1.39), (1.40, 1.49),
    (1.50, 1.59), (1.60, 1.69), (1.70, 1.79), (1.80, 1.89),
    (1.90, 1.99), (2.00, 2.19), (2.20, 2.49), (2.50, 2.99),
    (3.00, 3.49), (3.50, 3.99), (4.00, 4.49), (4.50, 4.99),
    (5.00, 5.49), (5.50, 5.99),
]

# Segments for TOP‐N
SEGMENT_RANGES = [
    (1, 25),
    (26, 75),
    (76, 175),
    (176, 325),
]

# Columns to drop from the final DataFrame
TO_DROP = [
    "Treatment", "Star Length", "Pavilion", "Crown", "Shade", "Inclusion", "Brand",
    "Rap Spec", "Key To Symbols", "Member Comment", "Vendor Stock number",
    "Unnamed: 29", "Seller ID", "Seller", "City", "State", "Notes", "Lab Comment",
    "Girdle", "Last Updated", "Country", "Item Page", "Rating", "Availability",
    "Diamond Type", "$/ct", "%Rap", "Total", "Cash: $/ct", "Cash: %Rap", "Cash: Total"
]

# Filter criteria for summary calculations
VALID_COLORS = []
VALID_CERTS  = []

def main():
    # 1) Prompt user to select input file
    input_path = filedialog.askopenfilename( title="Select File", filetypes=[("Excel files", "*.xlsx")])
       
    # 2) Read raw data
    df = pd.read_excel(input_path)

    # 3) Prepare new columns
    df["Index"] = pd.NA
    df["RAP SIZE"] = pd.NA

    # 4) Calculate base price + 2%
    price_base = df["$/ct"].fillna(df["Cash: $/ct"])
    df["New Price"] = (price_base * 1.02).round(2)

    # 5) Calculate new value
    df["New Value"] = df["New Price"] * df["Size"]

    # 6) Assign sequential Index and RAP SIZE per size‐bin
    for low, high in SIZE_RANGES:
        label = f"{low:.2f}-{high:.2f}"
        mask = df["Size"].between(low, high)
        count = mask.sum()
        if count:
            df.loc[mask, "Index"] = range(1, count + 1)
            df.loc[mask, "RAP SIZE"] = label

    # 7) Build the “Market Summary” as segmented TOP‐N,
    #    but only include stones with specified Color & cert in the avg_price calc
    summary_sizes = sorted(
        df["RAP SIZE"].dropna().unique(),
        key=lambda s: float(s.split("-")[0]))
    
    # Column names for summary
    cols = ["SIZE"] + [
        f"TOP {h - l + 1}"
        for l, h in SEGMENT_RANGES
    ] + ["Total pcs in mkt"]
    summary_df = pd.DataFrame(columns=cols)

    for idx, size_label in enumerate(summary_sizes):
        mask_size = df["RAP SIZE"] == size_label
        summary_df.loc[idx, "SIZE"] = size_label
        summary_df.loc[idx, "Total pcs in mkt"] = mask_size.sum()

        for l, h in SEGMENT_RANGES:
            top_n = h - l + 1
            col = f"TOP {top_n}"
            # include Color & cert filter here
            mask_segment = (
                mask_size &
                df["Index"].between(l, h) 
                # df["Color"].isin(VALID_COLORS) &
                # df["Lab"].isin(VALID_CERTS)
            )
            total_size  = df.loc[mask_segment, "Size"].sum()
            total_value = df.loc[mask_segment, "New Value"].sum()

            # avg_price only over the filtered subset
            avg_price = round(total_value / total_size, 2) if total_size else pd.NA
            summary_df.loc[idx, col] = avg_price

    # 8) Drop unused columns
    df.drop(columns=TO_DROP, inplace=True)

    # 9) Write results to Excel with formatting
    with pd.ExcelWriter(OUTPUT_PATH, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Processed Data", index=False, float_format="%.2f")
        summary_df.to_excel(writer, sheet_name="Market Summary", index=False)

        workbook  = writer.book
        worksheet = writer.sheets["Market Summary"]

        # Write header row
        for col_idx, header in enumerate(summary_df.columns):
            worksheet.write(0, col_idx, header)

        # Set column widths
        worksheet.set_column(0, 0, 12)
        worksheet.set_column(1, len(summary_df.columns) - 1, 15)

    # 10) Open the result file
    os.startfile(OUTPUT_PATH)

if __name__ == "__main__":
    main()
