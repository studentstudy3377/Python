import os
from datetime import datetime
from tkinter import filedialog
import pandas as pd
from openpyxl.utils import get_column_letter

def main():
    """Main function to process stock data and apply discounts."""
    # 1. Load data
    stock_path = filedialog.askopenfilename(
        title="Select Stock File",
        filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")],
    )
    if not stock_path:
        print("No stock file selected. Exiting.")
        return

    disc_path = r"G:\Certificate\HEMIL\PRICE DISC SEARCH\PRICE UPDATE.xlsx"

    df_disc = pd.read_excel(disc_path, skiprows=2)
    df_stock = pd.read_excel(stock_path, skiprows=10)

    # 2. Filter out unwanted rows
    mask_gd = (df_stock["Pol"] == "GD") & (df_stock["Sym"] == "GD")
    mask_hrd = df_stock["Lab"] == "HRD"
    mask_fluor = df_stock["Flr"] == "Non"
    df = df_stock.loc[~(mask_gd | mask_hrd) & mask_fluor].reset_index(drop=True)

    # 3. Clean & map shape codes to full names
    shape_map = {
        "PE": "PEAR",
        "MQ": "MARQUISE",
        "EM": "EMERALD",
        "OL": "OVAL",
        "PR": "PRINCESS",
    }
    df["Shp"] = df["Shp"].str.strip().replace(shape_map)

    # 4. Add RAPSIZE column based on Cts
    def rap_size(cts):
        if 0.3 <= cts < 0.39:
            return '30UP'
        if 0.5 <= cts < 0.59:
            return '50UP'
        if 1.00 <= cts < 1.99:
            return '1.00UP'
        return ''
    df['RAPSIZE'] = df['Cts'].apply(rap_size)
    
    # 4. Add RAPSIZE column based on Cts
    def ctr_size(cts):
        if 0.3 <= cts < 0.39:
            return '0.30'
        if 0.5 <= cts < 0.59:
            return '0.50'
        if 1.00 <= cts < 1.99:
            return '1.00'
        return ''
    df['New Cts'] = df['Cts'].apply(ctr_size)
    df['New Cts'] = pd.to_numeric(df['New Cts'], errors='coerce')
    # 5. Build merge keys
    key_cols = ["Shp", "RAPSIZE", "Col", "Clr", "Cut"]
    df["CON1"] = df[key_cols].fillna(" ").astype(str).agg(" - ".join, axis=1)
    df["CON2"] = df[key_cols].fillna(" ").astype(str).agg(" - ".join, axis=1)

    # 6. Merge in discounts
    df = df.merge(df_disc[["CON1", "Disc1"]], on="CON1", how="left")
    df = df.merge(df_disc[["CON2", "Disc2"]], on="CON2", how="left")

  
    # 7. Combine Disc1 and Disc2 into a new column 'New Disc' (text join)
    df['New Disc'] = (  df['Disc1'].fillna('').astype(str) + df['Disc2'].fillna('').astype(str))
    df['New Disc'] = pd.to_numeric(df['New Disc'], errors='coerce')
    # 8. Clean up helper columns
    df.drop(["CON1", "CON2", "Disc1", "Disc2","Rep No","Video Link", "Image Link", "Certificate Link", "CA","CH", "PA", "PD", "Str Ln", "Lr Half", "H&A", "Lus", "Col Tinge","Remarks","BYDs Status"],
            axis=1, inplace=True,)

    # 9. Insert new 'New Disc' column after 'VALUE'
    disc_series = df.pop("New Disc")
    insert_disc_idx = df.columns.get_loc("Lab") + 1
    df.insert(insert_disc_idx, "New Disc", disc_series)
    
    # 10. Insert new 'New Disc' column after 'VALUE'
    disc_series = df.pop("New Cts")
    insert_disc_idx = df.columns.get_loc("Cts") + 1
    df.insert(insert_disc_idx, "New Cts", disc_series)


    # 11. Insert new 'New Disc' column after 'VALUE'
    disc_series = df.pop("RAPSIZE")
    insert_disc_idx = df.columns.get_loc("New Cts") + 1
    df.insert(insert_disc_idx, "RAPSIZE", disc_series)
    
    # 12. Build Excel formulas for 'New Price'
    #    =<RAPPRICE cell> * (1 + <New Disc cell> / 100)
    col_rappr = get_column_letter(df.columns.get_loc("Rap") + 3)
    col_disc = get_column_letter(df.columns.get_loc("New Disc") + 1)
    df["New Price"] = [
        f"={col_rappr}{row}*(1+{col_disc}{row}/100)"
        for row in range(2, len(df) + 2)
    ]
    insert_price_idx = df.columns.get_loc("New Disc") + 1
    df.insert(insert_price_idx, "New Price", df.pop("New Price"))

    # 13. Build Excel formulas for 'New Value'
    #    =<New Price cell> * <New Carat cell>
    col_price = get_column_letter(df.columns.get_loc("New Price") + 1)
    col_carat = get_column_letter(df.columns.get_loc("New Cts") + 1)
    df["New Value"] = [
        f"={col_price}{row}*{col_carat}{row}" for row in range(2, len(df) + 2)
    ]
    insert_value_idx = df.columns.get_loc("New Price") + 1
    df.insert(insert_value_idx, "New Value", df.pop("New Value"))

    

    # . Write out with formulas preserved and format dates
    today_str = datetime.today().strftime("%d-%m-%Y")
    output_dir = r"G:\Certificate\Python Output"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"Stock_List_LEDS_{today_str}.xlsx")

    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="DD-MM-YYYY") as writer:
        df.to_excel(writer, index=False, sheet_name="Stock")

    # Open the file automatically
    os.startfile(output_path)


if __name__ == "__main__":
    main()
