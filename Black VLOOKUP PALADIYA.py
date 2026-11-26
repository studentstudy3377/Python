import pandas as pd
import os
from datetime import datetime
from tkinter import Tk, filedialog

def select_folder(prompt: str) -> str | None:
    """Ask the user to pick a folder and return its path (or None if canceled)."""
    root = Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title=prompt, mustexist=True)
    root.destroy()
    return folder_path or None

def find_latest_file(folder_path: str, file_prefix: str) -> str | None:
    file_prefix_upper = file_prefix.upper()
    files = [
        f for f in os.listdir(folder_path)
        if f.upper().startswith(file_prefix_upper) and f.lower().endswith(('.xls', '.xlsx'))
    ]
    if not files:
        print(f"Warning: No files with prefix '{file_prefix}' found in {folder_path}. Skipping.")
        return None
    files.sort(key=lambda f: os.path.getmtime(os.path.join(folder_path, f)), reverse=True)
    return os.path.join(folder_path, files[0])

def process_excel_files(
    org_file_path: str,
    deal_file_path: str,
    output_file_path: str,
    org_skiprows: int | None = None,
    org_merge_col: str | None = None,
    deal_merge_col: str | None = None,
    dNatts_1: str | None = None,
    dNatts_2: str | None = None,
    stock_id: str | None = None,
):
    try:
        # --- Read deal file ---
        deal_df = pd.read_excel(deal_file_path)

        # --- Read org file (with optional skiprows) ---
        if org_skiprows is not None:
            org_df = pd.read_excel(org_file_path, skiprows=org_skiprows)
        else:
            org_df = pd.read_excel(org_file_path)

        # --- Normalize merge columns (strip punctuation, keep word chars/space) ---
        if deal_merge_col and deal_merge_col in deal_df.columns:
            deal_df[deal_merge_col] = (
                deal_df[deal_merge_col].astype(str).str.replace(r'[^\w\s]', '', regex=True)
            )
        if org_merge_col and org_merge_col in org_df.columns:
            org_df[org_merge_col] = (
                org_df[org_merge_col].astype(str).str.replace(r'[^\w\s]', '', regex=True)
            )

        # --- Select columns from org_df to bring over ---
        merge_columns = []
        if org_merge_col:
            merge_columns.append(org_merge_col)
        if dNatts_1:
            merge_columns.append(dNatts_1)
        if dNatts_2:
            merge_columns.append(dNatts_2)
        if stock_id:
            merge_columns.append(stock_id)

        # Guard: if required columns missing, raise a clear error
        missing_cols = [c for c in merge_columns if c not in org_df.columns]
        if missing_cols:
            raise KeyError(f"Columns {missing_cols} not found in ORG file: {org_file_path}")

        if not deal_merge_col or deal_merge_col not in deal_df.columns:
            raise KeyError(f"Merge column '{deal_merge_col}' not found in DEAL file: {deal_file_path}")

        # --- Merge ---
        merged_df = pd.merge(
            deal_df,
            org_df[merge_columns],
            left_on=deal_merge_col,
            right_on=org_merge_col,
            how='left'
        )

        # --- Drop the extra merge column if it was added from org_df ---
        if org_merge_col and org_merge_col in merged_df.columns:
            final_df = merged_df.drop(columns=[org_merge_col])
        else:
            final_df = merged_df

        # --- Limit rows to non-null in deal merge column (optional behavior you had) ---
        dlot_length = deal_df[deal_merge_col].notna().sum()
        final_df = final_df.head(dlot_length)

        # --- Friendly output column names, then remove originals ---
        cols_to_remove = []
        if stock_id and stock_id in final_df.columns:
            final_df['Stock Id'] = final_df[stock_id]
            cols_to_remove.append(stock_id)
        if dNatts_1 and dNatts_1 in final_df.columns:
            final_df['dNatts_1'] = final_df[dNatts_1]
            cols_to_remove.append(dNatts_1)
        if dNatts_2 and dNatts_2 in final_df.columns:
            final_df['dNatts_2'] = final_df[dNatts_2]
            cols_to_remove.append(dNatts_2)

        final_df.drop(columns=cols_to_remove, errors='ignore', inplace=True)

        # --- Drop extra columns if present ---
        extra_columns_to_remove = [
            "dTOpens", "dBOpens", "DealDate", "dDia", "dCert", "diaF", "diaT",
            "TotHeight", "Cert_No", "dRap1", "dTinch", "dRatio", "M_QLTY",
            "rowid", "fSize", "tSize", "tot_Value", "Back_Disc", "Back_Rap",
            "Back_Value", "Memo_P", "Memo_V", "IsDemand", "DemandCust", "dLuster"
        ]
        final_df.drop(columns=extra_columns_to_remove, errors='ignore', inplace=True)

        # --- Move "Stock Id" to 2nd column if it exists ---
        if "Stock Id" in final_df.columns:
            cols = list(final_df.columns)
            cols.remove("Stock Id")
            cols.insert(1, "Stock Id")
            final_df = final_df[cols]

        # --- Ensure output directory exists ---
        out_dir = os.path.dirname(output_file_path)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        # --- Save ---
        final_df.to_excel(output_file_path, index=False)
        print(f"Data processing completed for {output_file_path}.")

    except Exception as e:
        print(f"Error: {e}")

def main():
    # Ask for folders
    deal_folder_path = select_folder("Select the folder containing the DEAL files:")
    org_folder_path  = select_folder("Select the folder containing the ORG files:")

    # Handle cancel
    if not deal_folder_path or not org_folder_path:
        print("Folder selection cancelled. Exiting.")
        return

    # Validate directories
    if not os.path.isdir(deal_folder_path) or not os.path.isdir(org_folder_path):
        print("One or both selected paths are not directories. Exiting.")
        return

    # Check if folders are empty
    if not os.listdir(deal_folder_path) or not os.listdir(org_folder_path):
        print("No files found in one or both of the selected folders. Exiting.")
        return

    paladiya_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 4,
        'org_merge_col': 'Stone Id',
        'stock_id': '',        # leave blank if not needed
        'dNatts_1': 'TB',
        'dNatts_2': 'SB',
    }

    params_list = [
        {'file_prefix': 'DEAL_paladiya', 'org_prefix': 'PALADIYA', 'params': paladiya_params},
    ]

    # Output directory
    output_directory = r"G:\Certificate\Python Output\VLOOKUP Output"
    os.makedirs(output_directory, exist_ok=True)

    # Loop through each set of parameters
    for item in params_list:
        deal_file_path = find_latest_file(deal_folder_path, item['file_prefix'])
        org_file_path  = find_latest_file(org_folder_path,  item['org_prefix'])

        if deal_file_path is None or org_file_path is None:
            continue

        # Output file name
        deal_file_name = os.path.splitext(os.path.basename(deal_file_path))[0]
        # If you want a dated filename, uncomment:
        # date_str = datetime.now().strftime('%d-%m-%Y')
        output_file_path = os.path.join(output_directory, f"{deal_file_name}.xlsx")

        # Process and merge
        process_excel_files(
            org_file_path     = org_file_path,
            deal_file_path    = deal_file_path,
            output_file_path  = output_file_path,
            org_skiprows      = item['params'].get('skip_rows'),
            org_merge_col     = item['params']['org_merge_col'],
            deal_merge_col    = item['params']['deal_merge_col'],
            dNatts_1          = item['params']['dNatts_1'] or None,
            dNatts_2          = item['params']['dNatts_2'] or None,
            stock_id          = item['params']['stock_id'] or None,
        )

        # Open the result on Windows (safe-guard if file exists)
        if os.path.exists(output_file_path):
            try:
                os.startfile(output_file_path)
            except Exception as e:
                print(f"Could not open '{output_file_path}': {e}")

if __name__ == "__main__":
    main()
