from datetime import datetime
import os
import numpy as np
import pandas as pd
from tkinter import filedialog
from openpyxl.utils import get_column_letter

def main():
    # 1. Load data
    stock_path = filedialog.askopenfilename( title="Select Stock File", filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")], )

    disc_path = r"G:\Certificate\HEMIL\PRICE DISC SEARCH\PRICE UPDATE.xlsx"

    df_disc = pd.read_excel(disc_path, skiprows=2)
    df_stock = pd.read_excel(stock_path)

    # 2. Filter out unwanted rows
    mask_gd    = (df_stock["POL"] == "GD") | (df_stock["SYM"] == "GD")
    mask_hrd   = df_stock["CERT"] == "HRD"
    mask_fluor = df_stock["FLUOR"] == "NONE"
    df = df_stock.loc[~(mask_gd | mask_hrd) & mask_fluor].reset_index(drop=True)

    # 3. Identify EX/EX rows and set CUT1/CUT2
    mask_ex      = (df["POL"] == "EX") & (df["SYM"] == "EX")
    df["CUT1"]   = np.where(mask_ex, "2EX", "")
    df["CUT2"]   = np.where(~mask_ex, "VG", "")

    # 4. Build merge keys
    cols1 = ["SHAPE", "RAPSIZE", "COLOR", "CLARITY", "CUT1"]
    cols2 = ["SHAPE", "RAPSIZE", "COLOR", "CLARITY", "CUT2"]
    df["CON1"] = df[cols1].fillna(" ").astype(str).agg(" - ".join, axis=1)
    df["CON2"] = df[cols2].fillna(" ").astype(str).agg(" - ".join, axis=1)

    # 5. Merge in discounts
    df = df.merge(df_disc[["CON1", "Disc1"]], on="CON1", how="left")
    df = df.merge(df_disc[["CON2", "Disc2"]], on="CON2", how="left")

    # 6. Apply discounts and cuts
    df.loc[mask_ex,  "Disc"] = df.loc[mask_ex,  "Disc1"]
    df.loc[~mask_ex, "Disc"] = df.loc[~mask_ex, "Disc2"]
    df.loc[mask_ex,  "CUT"]  = df.loc[mask_ex,  "CUT1"]
    df.loc[~mask_ex, "CUT"]  = df.loc[~mask_ex, "CUT2"]

    # 7. Drop rows without a discount
    df.dropna(subset=["Disc"], inplace=True)

    # 8. Clean up helper columns
    df.drop(
        ["CUT1", "CUT2", "CON1", "CON2", "Disc1", "Disc2"],
        axis=1,
        inplace=True,
    )

    # 9. Convert CERTNO to numeric
    df['CERTNO'] = pd.to_numeric(df['CERTNO'], errors='coerce')

    # 10. Insert 'New Disc' after the existing 'VALUE' column
    disc_series = df.pop("Disc")
    idx_value   = df.columns.get_loc("VALUE") + 1
    df.insert(idx_value, "New Disc", disc_series)

    # 11. Compute and insert 'New Carat' after 'CARAT'
    clean_size = df["RAPSIZE"].str.replace("UP", "", regex=False)
    size_vals  = clean_size.astype(float)
    new_carat  = np.where(~clean_size.str.contains(r"\."), size_vals/100, size_vals)
    idx_carat  = df.columns.get_loc("CARAT") + 1
    df.insert(idx_carat, "New Carat", new_carat)

    # 12. Build Excel formulas for 'New Price'
    col_rappr    = get_column_letter(df.columns.get_loc("RAPPRICE") + 1)
    col_disc_new = get_column_letter(df.columns.get_loc("New Disc") + 1)
    new_price_fm = [
        f"={col_rappr}{r}*(1+{col_disc_new}{r}/100)"
        for r in range(2, len(df)+2)
    ]
    idx_np = df.columns.get_loc("New Disc") + 1
    df.insert(idx_np, "New Price", new_price_fm)

    # 13. DROP the old 'VALUE' and rebuild it
    df.drop("VALUE", axis=1, inplace=True)  # FIX: remove existing before re-insert
    col_carat   = get_column_letter(df.columns.get_loc("CARAT") + 1)
    col_price_o = get_column_letter(df.columns.get_loc("PRICE") + 1)
    value_fm    = [
        f"={col_carat}{r}*{col_price_o}{r}"
        for r in range(2, len(df)+2)
    ]
    idx_val = df.columns.get_loc("PRICE") + 1
    df.insert(idx_val, "VALUE", value_fm)

    # 14. Build Excel formulas for 'New Value'
    col_np2      = get_column_letter(df.columns.get_loc("New Price") + 1)
    col_ncar2    = get_column_letter(df.columns.get_loc("New Carat") + 1)
    new_value_fm = [
        f"={col_np2}{r}*{col_ncar2}{r}"
        for r in range(2, len(df)+2)
    ]
    idx_nv = df.columns.get_loc("New Price") + 1
    df.insert(idx_nv, "New Value", new_value_fm)

    # 15. DROP the old 'PRICE' and rebuild it
    df.drop("PRICE", axis=1, inplace=True)  # FIX: remove existing before re-insert
    col_disc_o2 = get_column_letter(df.columns.get_loc("DISC") + 1)
    price_fm    = [
        f"={col_rappr}{r}*(1+{col_disc_o2}{r}/100)"
        for r in range(2, len(df)+2)
    ]
    idx_pr2 = df.columns.get_loc("DISC") + 1
    df.insert(idx_pr2, "PRICE", price_fm)

    # 16. Build Excel formulas for 'Diff Value'
    col_nv3     = get_column_letter(df.columns.get_loc("New Value") + 1)
    col_val3    = get_column_letter(df.columns.get_loc("VALUE")     + 1)
    diff_val_fm = [
        f"={col_nv3}{r}-{col_val3}{r}"
        for r in range(2, len(df)+2)
    ]
    idx_dv3 = df.columns.get_loc("New Value") + 1
    df.insert(idx_dv3, "Diff Value", diff_val_fm)

    # 17. Build Excel formulas for 'Diff %'
    col_dv4      = get_column_letter(df.columns.get_loc("Diff Value") + 1)
    col_val4     = get_column_letter(df.columns.get_loc("VALUE")      + 1)
    diff_pct_fm  = [
        f"={col_dv4}{r}/{col_val4}{r}"
        for r in range(2, len(df)+2)
    ]
    idx_dp4 = df.columns.get_loc("Diff Value") + 1
    df.insert(idx_dp4, "Diff %", diff_pct_fm)

    # 18. Append ONE subtotal row (no blanks)
    start_row = 2
    end_row   = len(df) + 1
    letters   = {
        "VALUE"     : get_column_letter(df.columns.get_loc("VALUE")      + 1),
        "New Value" : get_column_letter(df.columns.get_loc("New Value")  + 1),
        "Diff Value": get_column_letter(df.columns.get_loc("Diff Value") + 1),
    }
    subtotal = {c: "" for c in df.columns}
    subtotal["VALUE"]      = f"=SUBTOTAL(9,{letters['VALUE']}{start_row}:{letters['VALUE']}{end_row})"
    subtotal["New Value"]  = f"=SUBTOTAL(9,{letters['New Value']}{start_row}:{letters['New Value']}{end_row})"
    subtotal["Diff Value"] = f"=SUBTOTAL(9,{letters['Diff Value']}{start_row}:{letters['Diff Value']}{end_row})"
    subtotal["Diff %"]     = (
        f"={letters['Diff Value']}{end_row+1}/"
        f"{letters['VALUE']}{end_row+1}"
    )
    df = pd.concat([df, pd.DataFrame([subtotal])], ignore_index=True)

    # 19. Write out, format Diff % column
    today_str   = datetime.today().strftime("%d-%m-%Y")
    output_path = f"G:\\Certificate\\Python Output\\Stock List LEDS {today_str}.xlsx"
    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="DD-MM-YYYY") as writer:
        df.to_excel(writer, index=False, sheet_name="Stock")
        workbook     = writer.book
        worksheet    = writer.sheets["Stock"]
        diff_pct_col = get_column_letter(df.columns.get_loc("Diff %") + 1)
        pct_fmt      = workbook.add_format({"num_format": "0 %"})
        worksheet.set_column(f"{diff_pct_col}:{diff_pct_col}", None, pct_fmt)

    # 20. Open the file
    os.startfile(output_path)

if __name__ == "__main__":
    main()
