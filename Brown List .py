from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog
import os

import pandas as pd
from xlsxwriter.utility import xl_col_to_name


def select_and_load_excel_files() -> dict[str, pd.DataFrame]:
    """
    Opens a file dialog to select multiple Excel or CSV files.
    Returns {filename: DataFrame}.
    """
    root = Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        title="Select Excel/CSV Files to Merge",
        filetypes=[("All Files", "*.*")],
    )

    skiprows_map: dict[str, int] = {}
    dfs: dict[str, pd.DataFrame] = {}

    for path in file_paths:
        filename = os.path.basename(path)
        ext = os.path.splitext(filename)[1].lower()
        skiprows = next(
            (v for k, v in skiprows_map.items() if k in filename.upper()), 0
        )

        try:
            if ext in [".xlsx", ".xls"]:
                engine = "openpyxl" if ext == ".xlsx" else "xlrd"
                dfs[filename] = pd.read_excel(path, engine=engine, skiprows=skiprows)
            elif ext == ".csv":
                dfs[filename] = pd.read_csv(path, skiprows=skiprows)
            else:
                print(f"ℹ️ Skipping unsupported file type: {filename}")
        except Exception as e:
            print(f"❌ Failed to read file: {filename} ({e})")

    return dfs


def categorize_carat(carat: float | int | None) -> str | None:
    """Map carat value to a RAPSIZE bucket."""
    if pd.isna(carat):
        return None

    carat_categories = {
        (0.23, 0.29): "23UP",
        (0.30, 0.39): "30UP",
        (0.40, 0.49): "40UP",
        (0.50, 0.59): "50UP",
        (0.60, 0.69): "60UP",
        (0.70, 0.79): "70UP",
        (0.80, 0.89): "80UP",
        (0.90, 0.99): "90UP",
        (1.00, 1.19): "1.00UP",
        (1.20, 1.49): "1.50UP",
        (1.50, 1.99): "2.00UP",
        (2.00, 2.99): "2.00UP",
        (3.00, 3.99): "3.00UP",
        (4.00, 4.99): "4.00UP",
        (5.00, 5.99): "5.00UP",
    }

    for (lower, upper), category in carat_categories.items():
        if lower <= float(carat) <= upper:
            return category

    return " "


def main() -> None:
    # ----- Load & merge data -----
    files = select_and_load_excel_files()
    if not files:
        print("❌ No files selected.")
        return

    data = pd.concat(files.values(), ignore_index=True, sort=False).fillna(" ")

    # ----- Numeric casting (safe) -----
    for col in data.columns:
        if data[col].dtype == "object":
            # Only cast if purely digits (keeps IDs with letters/leading zeros intact)
            if data[col].astype(str).str.isnumeric().all():
                data[col] = pd.to_numeric(data[col], errors="coerce")

    if "Depth" in data.columns:
        data["Depth"] = pd.to_numeric(data["Depth"], errors="coerce")
    if "Table" in data.columns:
        data["Table"] = pd.to_numeric(data["Table"], errors="coerce")

    if "Vendor Stock number" in data.columns:
        data = data[data["Vendor Stock number"].astype(str).str.strip() != ""]

    # ----- Filters -----
    CULET_ALLOWED = ["None", "None Pointed"]
    CITIES_ALLOWED = ["SURAT", "Mumbai", "Bombay"]
    EXCLUDE_INCLUSIONS = ["BT2", "BT3", "BC2", "BC3", "WC2", "WC3"]
    EXCLUDE_SELLERS = {
        "Narola Brothers",
        "NAROLA DIAMONDS PRIVATE LIMITED",
        "JANAM CORPORATION PVT. LTD.",
        "SINGH DIAMONDS - rajesh kankaria",
    }
    EXCLUDE_INCL_REGEX = "|".join(EXCLUDE_INCLUSIONS)

    if "Country" in data.columns:
        data = data[data["Country"] == "India"]

    # Cut/Polish/Symmetry normalization
    repl_cps = {"Excellent": "EX", "Very Good": "VG", "Good": "GD", "Ideal": "EX"}
    for c in ("Cut", "Polish", "Symmetry"):
        if c in data.columns:
            data[c] = data[c].replace(repl_cps)

    # Fluorescence normalization
    if "Fluorescence" in data.columns:
        data["Fluorescence"] = data["Fluorescence"].replace(
            {
                " ": "NON",
                "Faint": "FNT",
                "Slight": "FNT",
                "Medium": "MED",
                "MediumBlue": "MED",
                "MediumYellow": "MED",
                "MediumGreen": "MED",
                "NoneWhite": "NON",
                "MediumWhite": "MED",
            }
        )

    # Lab normalization
    if "Lab" in data.columns:
        data = data[data["Lab"] != "GIA DOR"]
        data["Lab"] = data["Lab"].replace({"IGI": "NON CERT", "HRD": "NON CERT"})

    # City & geometry constraints
    if "City" in data.columns:
        data["City"] = data["City"].astype(str).str.strip()
        data = data[data["City"].str.upper().isin([c.upper() for c in CITIES_ALLOWED])]
    if "Depth" in data.columns:
        data = data[data["Depth"] > 50]
    if "Table" in data.columns:
        data = data[data["Table"] > 50]

    # Other exclusions
    if "Culet" in data.columns:
        data["Culet"] = data["Culet"].astype(str).str.strip()
        data = data[data["Culet"].isin(CULET_ALLOWED)]
    if "Seller" in data.columns:
        data = data[~data["Seller"].isin(EXCLUDE_SELLERS)]
    if "Shade" in data.columns:
        data = data[data["Shade"] != "Milkey"]
    # if "Inclusion" in data.columns:
    #     data = data[~data["Inclusion"].str.contains(EXCLUDE_INCL_REGEX, na=False)]
    if "Key To Symbols" in data.columns:
        data = data[~data["Key To Symbols"].str.startswith("Cavity", na=False)]
        data = data[~data["Key To Symbols"].str.startswith("Chip", na=False)]
    if "Lab Comment" in data.columns:
        data = data[
            ~data["Lab Comment"].str.startswith(
                "Clarity grade is based on internal graining that is not shown.",
                na=False,
            )
        ]

    # ❌ Remove rows where Seller == "M.SURESH..." AND Member Comment == "Only Export"
    if "Seller" in data.columns and "Member Comment" in data.columns:
        mask = (data["Seller"].str.strip().eq("M.SURESH COMPANY PRIVATE LTD")) & (
            data["Member Comment"].str.strip().eq("Only Export")
        )
        data = data[~mask]

    # ----- RAPSIZE from Size -----
    if "Size" in data.columns:
        data["RAPSIZE"] = data["Size"].apply(categorize_carat)

    # ----- Column mapping (collapse variants inline, no helper) -----
    columns_mapping = {
        "Diamond ID": ["Diamond ID", "Stone ID"],
        "Shape": ["Shape"],
        "Size": ["Size"],
        "RAPSIZE": ["RAPSIZE"],
        "Color": ["Color"],
        "CLARITY": ["Clarity"],
        "Cut": ["Cut"],
        "Polish": ["Polish"],
        "Symmetry": ["Symmetry"],
        "Fluorescence": ["Fluorescence"],
        "Lab": ["Lab"],
        "Ratio": ["Ratio"],
        "Depth": ["Depth"],
        "Table": ["Table"],
        "Measurements": ["Measurements", "Measurement"],
        "Member Comment": ["Member Comment"],
        "Vendor Stock number": ["Vendor Stock number"],
        "$/ct": ["$/ct", "Cash: $/ct"],
        "%Rap": ["%Rap", "Cash: %Rap"],
        "Total": ["Total", "Cash: Total"],
        "Seller": ["Seller"],
        "City": ["City"],
    }

    output_data = pd.DataFrame(index=data.index)

    for out_col, srcs in columns_mapping.items():
        cols = [c for c in srcs if c in data.columns]
        if not cols:
            output_data[out_col] = pd.NA
        elif len(cols) == 1:
            output_data[out_col] = data[cols[0]]
        else:
            # Treat blanks/whitespace as NA, then take first non-NA from left to right
            temp = data[cols].replace(r"^\s*$", pd.NA, regex=True)
            s = temp[cols[0]].copy()
            for c in cols[1:]:
                s = s.where(s.notna(), temp[c])  # keep existing; fill from next
            output_data[out_col] = s

    # Make numeric if present (optional but handy for calculations/sorting)
    for col in ["$/ct", "%Rap", "Total", "Depth", "Table", "Ratio", "Size"]:
        if col in output_data.columns:
            output_data[col] = pd.to_numeric(output_data[col], errors="coerce")

    # ----- Remove duplicates in Diamond ID (if present) -----
    if "Diamond ID" in output_data.columns:
        output_data = output_data.drop_duplicates(subset=["Diamond ID"], keep="first")

    if "%Rap" in output_data.columns:
        output_data["%Rap"] = output_data["%Rap"].apply(lambda x: round(x * 100, 2) if -1 < x < 1 else x)
        
    # ----- Save to XLSX (text format for Vendor Stock number) -----
    output_dir = Path(r"G:\Certificate\Python Output")
    output_dir.mkdir(parents=True, exist_ok=True)

    current_datetime = datetime.now().strftime("%d.%m.%Y")
    file_path = output_dir / f"Brown List {current_datetime}.xlsx"

    # Ensure text dtype pre-write for VSN
    if "Vendor Stock number" in output_data.columns:
        output_data["Vendor Stock number"] = output_data["Vendor Stock number"].astype(
            str
        )

    with pd.ExcelWriter(file_path, engine="xlsxwriter", mode="w") as writer:
        sheet_name = "Brown_List"
        output_data.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Text format for Vendor Stock number column
        if "Vendor Stock number" in output_data.columns:
            vsn_idx = output_data.columns.get_loc("Vendor Stock number")
            col_letter = xl_col_to_name(vsn_idx)
            col_range = f"{col_letter}:{col_letter}"
            text_fmt = workbook.add_format({"num_format": "@"})
            worksheet.set_column(col_range, None, text_fmt)
            worksheet.ignore_errors({"number_stored_as_text": col_range})

    print(f"✅ File successfully saved at: {file_path}")
    os.startfile(str(file_path))

if __name__ == "__main__":
    main()
