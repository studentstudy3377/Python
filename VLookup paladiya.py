import pandas as pd
import os
from datetime import datetime
from tkinter import Tk, filedialog

def select_folder(prompt):
    root = Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title=prompt)
    root.destroy()   # Ensure the window is closed after selection
    return folder_path

def find_latest_file(folder_path, file_prefix):
   
    files = [
        f for f in os.listdir(folder_path)
        if f.startswith(file_prefix) and f.endswith(('.xls', '.xlsx'))
    ]
    if not files:
        print(f"Warning: No files with prefix '{file_prefix}' found in {folder_path}. Skipping this item.")
        return None

    # Sort by last modified time (descending order)
    files.sort(key=lambda f: os.path.getmtime(os.path.join(folder_path, f)), reverse=True)
    return os.path.join(folder_path, files[0])

def normalize_org_df(org_df, org_merge_col, disc_column):
   
    # Remove leading/trailing whitespace in column names
    org_df.columns = org_df.columns.str.strip()

    # Fill any missing values (if needed) with appropriate defaults
    org_df.fillna({org_merge_col: 0, disc_column: 0}, inplace=True)

    # Convert to numeric
    org_df[org_merge_col] = pd.to_numeric(org_df[org_merge_col], errors='coerce')
    org_df[disc_column] = pd.to_numeric(org_df[disc_column], errors='coerce')

    # Drop rows where the merge or discount columns are missing
    org_df.dropna(subset=[org_merge_col, disc_column], inplace=True)

    return org_df

def process_excel_files(org_file_path,deal_file_path,output_file_path,org_skiprows=None,org_merge_col=None,deal_merge_col=None,disc_column=None,is_disc_negative=False):
    try:
        # Load deal and org Excel files
        deal_df = pd.read_excel(deal_file_path)
        if org_skiprows is not None:
            org_df = pd.read_excel(org_file_path, skiprows=org_skiprows)
        else:
            org_df = pd.read_excel(org_file_path)


        if deal_merge_col in deal_df.columns:
            deal_df[deal_merge_col] = (
                deal_df[deal_merge_col]
                .astype(str)
                .str.replace(r'[^\w\s]', '', regex=True)
            )
        if org_merge_col in org_df.columns:
            org_df[org_merge_col] = (
                org_df[org_merge_col]
                .astype(str)
                .str.replace(r'[^\w\s]', '', regex=True)
            )
        merged_df = pd.merge(deal_df,org_df[[org_merge_col, disc_column]],left_on=deal_merge_col,right_on=org_merge_col,how='left')

        # Get today's date for naming the discount column
        today_date = datetime.today().strftime('%d-%m-%Y')

        # Remove existing column with today's date if it exists
        if today_date in merged_df.columns:
            merged_df.drop(columns=[today_date], inplace=True)

        # Rename discount column to today's date
        merged_df.rename(columns={disc_column: today_date}, inplace=True)

        # If PRICE column already exists, remove it
        if 'PRICE' in merged_df.columns:
            merged_df.drop(columns=['PRICE'], inplace=True)

        # Calculate new PRICE
        merged_df['PRICE'] = merged_df['dRap'] + (merged_df['dRap'] * merged_df[today_date] / 100)

        # If VALUE column already exists, remove it
        if 'VALUE' in merged_df.columns:
            merged_df.drop(columns=['VALUE'], inplace=True)

        # Calculate new VALUE
        merged_df['VALUE'] = merged_df['PRICE'] * merged_df['dWt']

        # If DIFF column already exists, remove it
        if 'DIFF' in merged_df.columns:
            merged_df.drop(columns=['DIFF'], inplace=True)

        # Calculate DIFF
        merged_df['DIFF'] = merged_df['dDisc1'] - merged_df[today_date]

        # Select all columns except org_merge_col for final output
        output_columns = [col for col in merged_df.columns if col != org_merge_col]
        final_df = merged_df[output_columns]

        # Trim final DataFrame to match the non-null count of deal_merge_col in deal_df
        dlot_length = deal_df[deal_merge_col].notna().sum()
        final_df = final_df.head(dlot_length)

        # Save the processed DataFrame to an Excel file
        final_df.to_excel(output_file_path, index=False)
        print(f"Data processing completed successfully for {output_file_path}.")

    except Exception as e:
        print(f"Error: {e}")

def main():
    # Prompt user for folder paths
    deal_folder_path = select_folder("Select the folder DEAL files:")
    org_folder_path = select_folder("Delect the folder ORG files:")

    if not os.listdir(deal_folder_path) or not os.listdir(org_folder_path):
        print("No files found in one or both of the selected folders. Exiting.")
        return

    # Define your parameter dictionaries
    paladiya_params = {
        'skip_rows': 4,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Stone Id',
        'disc_column': 'Disc %',
    }

    params_list = [
        {'file_prefix': 'DEAL_paladiya','org_prefix': 'PALADIYA','params': paladiya_params},
    ]

    # Output directory
    output_directory = r"G:\Certificate\Python Output\VLOOKUP Output"

    # Loop through each parameter set
    for item in params_list:
        deal_file_path = find_latest_file(deal_folder_path, item['file_prefix'])
        org_file_path = find_latest_file(org_folder_path, item['org_prefix'])
        if deal_file_path is None or org_file_path is None:
            continue

        # Extract file name (without extension) from the deal file path
        deal_file_name = os.path.splitext(os.path.basename(deal_file_path))[0]

        # Construct the output file path
        output_file_path = os.path.join(output_directory, f"{deal_file_name}.xlsx")

        # Process the files
        process_excel_files(
            org_file_path=org_file_path,
            deal_file_path=deal_file_path,
            output_file_path=output_file_path,
            org_skiprows=item['params'].get('skip_rows'),
            org_merge_col=item['params']['org_merge_col'],
            deal_merge_col=item['params']['deal_merge_col'],
            disc_column=item['params']['disc_column'],
            is_disc_negative=item['params'].get('is_disc_negative', False)
        )

if __name__ == "__main__":
    main()
