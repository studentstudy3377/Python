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


def process_excel_files(org_file_path,deal_file_path,output_file_path,org_skiprows=None,org_merge_col=None,deal_merge_col=None,dNatts_1=None,dNatts_2=None,stock_id=None):

    try:
        
        deal_df = pd.read_excel(deal_file_path)
        if org_skiprows is not None:
            org_df = pd.read_excel(org_file_path, skiprows=org_skiprows)
        else:
            org_df = pd.read_excel(org_file_path)


        deal_df[deal_merge_col] = pd.to_numeric(deal_df[deal_merge_col], errors='coerce')
        org_df[org_merge_col]   = pd.to_numeric(org_df[org_merge_col], errors='coerce')

        # Build the list of columns to merge from org_df
        merge_columns = [org_merge_col]
        if dNatts_1:
            merge_columns.append(dNatts_1) 
        if dNatts_2:
            merge_columns.append(dNatts_2) 
        if stock_id:
            merge_columns.append(stock_id)  

        # Merge the dataframes on the merge columns
        merged_df = pd.merge(deal_df,org_df[merge_columns],left_on=deal_merge_col,right_on=org_merge_col,how='left')

        # Drop the extra merge column if not needed
        final_df = merged_df.drop(columns=[org_merge_col])

        # Optionally limit rows based on non-null in deal_df's merge column
        dlot_length = deal_df[deal_merge_col].notna().sum()
        final_df = final_df.head(dlot_length)

        if stock_id and stock_id in final_df.columns:
            final_df['Stock Id'] = final_df[stock_id]
        if dNatts_1 and dNatts_1 in final_df.columns:
            final_df['dNatts_1'] = final_df[dNatts_1]
        if dNatts_2 and dNatts_2 in final_df.columns:
            final_df['dNatts_2'] = final_df[dNatts_2]

        cols_to_remove = []
        if stock_id:
            cols_to_remove.append(stock_id)   
        if dNatts_1:
            cols_to_remove.append(dNatts_1)   
        if dNatts_2:
            cols_to_remove.append(dNatts_2)   

        final_df.drop(columns=cols_to_remove, errors='ignore', inplace=True)

        extra_columns_to_remove = [
            "dTOpens", "dBOpens", "DealDate", "dDia", "dCert", "diaF", "diaT",
            "TotHeight", "Cert_No", "dRap1", "dTinch", "dRatio", "M_QLTY",
            "rowid", "fSize", "tSize", "tot_Value", "Back_Disc", "Back_Rap",
            "Back_Value", "Memo_P", "Memo_V", "IsDemand", "DemandCust", "dLuster"
        ]
        final_df.drop(columns=extra_columns_to_remove, errors='ignore', inplace=True)

        if "Stock Id" in final_df.columns:
            cols = list(final_df.columns)
            cols.remove("Stock Id")
            cols.insert(1, "Stock Id")
            final_df = final_df[cols]

        # Save the final DataFrame to an Excel file
        final_df.to_excel(output_file_path, index=False)
        print(f"Data processing completed for {output_file_path}.")

    except Exception as e:
        print(f"Error: {e}")

def main():
    # Prompt user for folder paths
    deal_folder_path = select_folder("Select the folder DEAL Files")
    org_folder_path  = select_folder("Select the folder ORG Files")

    # Check if folders are empty
    if not os.listdir(deal_folder_path) or not os.listdir(org_folder_path):
        print("No files found in one or both of the selected folders. Exiting.")
        return
    
    akarsh_params = {
        'sheet_name': 'Stock',
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'stock_id': 'Packet No',    
        'dNatts_1': 'Black Incl.',     
        'dNatts_2': 'Table Incl.',
    }
    
    params_list = [  
        {'file_prefix': 'DEAL_akarsh', 'org_prefix': 'AKARSH', 'params': akarsh_params},
    ]
    
    # Output directory
    output_directory = r"G:\Certificate\Python Output\VLOOKUP Output"

    # Loop through each set of parameters
    for item in params_list:
        try:
            deal_file_path = find_latest_file(deal_folder_path, item['file_prefix'])
            org_file_path  = find_latest_file(org_folder_path,  item['org_prefix'])

            if deal_file_path is None or org_file_path is None:
                continue

            deal_file_name = os.path.splitext(os.path.basename(deal_file_path))[0]
            date_str = datetime.now().strftime('%d-%m-%Y')
            output_file_path = os.path.join(output_directory, f"{deal_file_name}.xlsx")
            
            # Process and merge 
            process_excel_files(
                org_file_path     = org_file_path,
                deal_file_path    = deal_file_path,
                output_file_path  = output_file_path,
                org_skiprows      = item['params'].get('skip_rows'),
                org_merge_col     = item['params']['org_merge_col'],
                deal_merge_col    = item['params']['deal_merge_col'],
                dNatts_1          = item['params']['dNatts_1'],
                dNatts_2          = item['params']['dNatts_2'],
                stock_id          = item['params']['stock_id'],
            )

            # Only try to open if file exists
            if os.path.exists(output_file_path):
                os.startfile(output_file_path)
            else:
                print(f"Output file {output_file_path} was not created. Skipping opening.")

        except Exception as e:
            print(f"Error processing files with prefixes {item['file_prefix']} / {item['org_prefix']}: {e}")
            # Continue to the next set!
            continue

if __name__ == "__main__":
    main()
