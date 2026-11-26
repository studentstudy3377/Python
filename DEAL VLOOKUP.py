import pandas as pd
import os
from datetime import datetime
from tkinter import Tk, filedialog

def select_folder(prompt):
    """ Creates a GUI to select a folder path. """
    root = Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title=prompt)
    root.destroy()  # Ensure the window is closed after selection
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

def normalize_org_df(org_df, org_merge_col, disc_column):

    # Remove leading/trailing whitespace in column names
    org_df.columns = org_df.columns.str.strip()

    # Fill any missing values (if needed) with appropriate default values
    org_df.fillna({org_merge_col: 0, disc_column: 0}, inplace=True)

    # Ensure that all values in the merge and discount columns are numeric
    org_df[org_merge_col] = pd.to_numeric(org_df[org_merge_col], errors='coerce')
    org_df[disc_column] = pd.to_numeric(org_df[disc_column], errors='coerce')

    # Optionally drop any rows where critical information is missing
    org_df.dropna(subset=[org_merge_col, disc_column], inplace=True)

    return org_df

def process_excel_files(org_file_path, deal_file_path, output_file_path, org_skiprows=None, org_merge_col=None, deal_merge_col=None, disc_column=None, is_disc_negative=False):
    """ Processes Excel files by merging and applying discount calculations. """
    try:
        # Load deal and organization Excel files
        deal_df = pd.read_excel(deal_file_path)
        org_df = pd.read_excel(org_file_path, skiprows=org_skiprows) if org_skiprows is not None else pd.read_excel(org_file_path)

        # Convert the columns into numeric, coercing errors to NaN
        org_df[org_merge_col] = pd.to_numeric(org_df[org_merge_col], errors='coerce')
        org_df[disc_column] = pd.to_numeric(org_df[disc_column], errors='coerce')
        org_df[disc_column] = org_df[disc_column].apply(lambda x: x * -1 if x > 0 else x)

        deal_df[deal_merge_col] = pd.to_numeric(deal_df[deal_merge_col], errors='coerce')

        # Merge the dataframes on specified columns
        merged_df = pd.merge(deal_df, org_df[[org_merge_col, disc_column]], left_on=deal_merge_col, right_on=org_merge_col, how='left')

        # Calculate new prices and values based on the discount
        today_date = datetime.today().strftime('%d-%m-%Y')
        if today_date in merged_df.columns:
            merged_df.drop(columns=[today_date], inplace=True)
        merged_df.rename(columns={disc_column: today_date}, inplace=True)

        if 'PRICE' in merged_df.columns:
            merged_df.drop(columns=['PRICE'], inplace=True)
        merged_df['PRICE'] = merged_df['dRap'] + (merged_df['dRap'] * merged_df[today_date] / 100)

        if 'VALUE' in merged_df.columns:
            merged_df.drop(columns=(['VALUE']), inplace=True)
        merged_df['VALUE'] = merged_df['PRICE'] * merged_df['dWt']

        # Calculate the difference between existing and new discounts
        if 'DIFF' in merged_df.columns:
            merged_df.drop(columns=['DIFF'], inplace=True)
        merged_df['DIFF'] = merged_df['dDisc1'] - merged_df[today_date]

        # Select all columns except org_merge_col for final output
        output_columns = [col for col in merged_df.columns if col != org_merge_col]
        final_df = merged_df[output_columns]

        # Trim the DataFrame to the length of 'dLot' column in deal_df
        dlot_length = deal_df[deal_merge_col].notna().sum()  # Calculate non-null count in dLot
        final_df = final_df.head(dlot_length)  # Keep only rows up to the length of 'dLot'

        # Save the processed DataFrame to an Excel file
        final_df.to_excel(output_file_path, index=False)
        print(f"Data processing completed successfully for {output_file_path}.")
    except Exception as e:
        print(f"Error: {e}")
        
def main():
    deal_folder_path = select_folder("Please select the folder containing the DEAL files:")
    org_folder_path = select_folder("Please select the folder containing the ORG files:")

    if not os.listdir(deal_folder_path) or not os.listdir(org_folder_path):
        print("No files found in one or both of the selected folders. Exiting.")
        return

    # Define all the parameter dictionaries for both sets of parameters related to the 'DEAL_kb gems' files
    akarsh_params = {
        'sheet_name': 'Stock',
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    ankit_params = {
        'sheet_name': 'Sheet1',
        'deal_merge_col': 'dLot',
        'skip_rows': 10,
        'org_merge_col': 'Certificate No',
        'disc_column': 'RAP DIS',
    }
    asian_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 7,
        'org_merge_col': 'Report No.',
        'disc_column': 'Disc',
    }
    finestar_params = {
        'sheet_name': 'diamonds',
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Packet No',
        'disc_column': 'Disc %',
    }
    glowstar_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 4,
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    gnyan_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'disc_column': 'Net %',
    }
    hiradias_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    hk_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Stock No',
        'disc_column': 'Rap%',
    }
    hrv_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    jodhani_params = {
        'org_merge_col': 'Report No',
        'deal_merge_col': 'dLot',
        'disc_column': 'Disc(%)',
        'is_disc_negative': True,
    }
    jewlex_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    komal_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc%',
    }
    kiran_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 10,
        'org_merge_col': 'Rep No',
        'disc_column': 'Disc%',
    }
    kg_params = {
        'org_merge_col': 'REPORT NO.',
        'deal_merge_col': 'dLot',
        'disc_column': 'DISC',
    }
    kg_params_web = {
        'skip_rows': 1,
        'org_merge_col': 'REPORT NO.',
        'deal_merge_col': 'dLot',
        'disc_column': 'DISC',
    }
    kb_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 2,
        'org_merge_col': 'CertNo',
        'disc_column': 'Disc%',
    }
    kb_params_web = {
        'deal_merge_col': 'dLot',
        'skip_rows': 2,
        'org_merge_col': 'CertNo',
        'disc_column': 'Disc%',
    }
    kbs_params = {
        'sheet_name': 'stock',
        'deal_merge_col': 'dLot',
        'skip_rows': 2,
        'org_merge_col': 'CertNo',
        'disc_column': 'DISC',
    }
    kc_params = {
        'sheet_name': 0,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Cert No',
        'disc_column': 'Dis %',
    }
    kgk_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 6,
        'org_merge_col': 'Cert #',
        'disc_column': 'Disc',    
    }
    laxmi_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 5,
        'org_merge_col': 'Certificate',
        'disc_column': 'Disc(%)',
    }     
    lakhi_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report #',
        'disc_column': 'Rapnet  Discount %',
    }
    mb_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Cert No.',
        'disc_column': 'Disc',
    }
    mb_params_web = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Cert No.',
        'disc_column': 'Disc',
    }
    m_d_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',  
    }
    msuresh_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'disc_column': 'Online Back',  
    }
    navin_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'disc_column': 'Disc. (%)',
    }
    narola_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'ReportNo',
        'disc_column': 'Disc%',
    }
    pragati_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }  
    pramukh_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc%',
    }
    ratnakala_params = {
        'skip_rows': 4,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    rays_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Cert.No.',
        'disc_column': 'Disc%',
    }
    red_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report#',
        'disc_column': 'Disc',
    }
    rajgir_params = {
        'sheet_name': 'Rajgir Gems STOCK LIST',
        'deal_merge_col': 'dLot',
        'skip_rows': 3,
        'org_merge_col': 'REPORT NO',
        'disc_column': 'DISC %',
    }
    rajesh_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 5,
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    rajesh_params_web = {
        'deal_merge_col': 'dLot',
        'skip_rows': 5,
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    r_n_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Certificate',
        'disc_column': 'DISC',
    }
    shairu_params = {
        'skip_rows': 5,
        'org_merge_col': 'Certi No.',
        'deal_merge_col': 'dLot',
        'disc_column': 'Disc(%)',
    }
    shivam_params = {
        'org_merge_col': 'Certificate No',
        'deal_merge_col': 'dLot',
        'disc_column': 'Dis%',
    }
    starrays_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 3,
        'org_merge_col': 'Cert.No',
        'disc_column': 'Rap.%',
    }
    Sheetal_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 8,
        'org_merge_col': 'Certificate',
        'disc_column': 'Disc%',
    }
    sanghvi_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 5,
        'org_merge_col': 'CERTNO',
        'disc_column': 'RAP DISC%',
    } 
    ske_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc%',
    } 
    snj_params = {
        'skip_rows': 5,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Certificate No',
        'disc_column': 'Final Disc%',
    } 
    sunrays_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    } 
    shree_krishna_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': 'Disc%',
    }  
    shyam_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report #',
        'disc_column': 'Disc.',
    }  
    srd_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 4,
        'org_merge_col': 'Cert No',
        'disc_column': 'Rap %',
    }
    starlight_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Cert.No.',
        'disc_column': 'Back',  
    }
    unique_params = {
        'skip_rows': 4,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'disc_column': ' Disc %',
    }
    vaibhav_params = {
        'sheet_name': 'Stock',
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'disc_column': 'Disc %',
    }
    veer_params = {
        'sheet_name': 'Stock',
        'deal_merge_col': 'dLot',
        'skip_rows': 3,
        'org_merge_col': 'Cert#',
        'disc_column': 'Disc',
    }
    vdglobal_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 4,
        'org_merge_col': 'ReportNo',
        'disc_column': 'Disc%',  
    }

    # List of dictionaries for file prefixes and parameters
    params_list = [
        {'file_prefix': 'DEAL_akarsh', 'org_prefix': 'akarsh', 'params':akarsh_params},
        {'file_prefix': 'DEAL_ankit', 'org_prefix': 'ankit', 'params': ankit_params},
        {'file_prefix': 'DEAL_asian', 'org_prefix': 'asian star', 'params': asian_params},   
        {'file_prefix': 'DEAL_finestar', 'org_prefix': 'finestar', 'params': finestar_params},
        {'file_prefix': 'DEAL_glowstar', 'org_prefix': 'glowstar', 'params': glowstar_params},
        {'file_prefix': 'DEAL_gnyan', 'org_prefix': 'gnyan', 'params': gnyan_params},
        {'file_prefix': 'DEAL_haridias', 'org_prefix': 'hiradias', 'params': hiradias_params},
        {'file_prefix': 'DEAL_hk', 'org_prefix': 'hk', 'params': hk_params},
        {'file_prefix': 'DEAL_hrv', 'org_prefix': 'hrv', 'params': hrv_params},
        {'file_prefix': 'DEAL_jodhani', 'org_prefix': 'Jodhani', 'params': jodhani_params},
        {'file_prefix': 'DEAL_jewelex', 'org_prefix': 'Jewelex', 'params': jewlex_params},
        {'file_prefix': 'DEAL_kiran', 'org_prefix': 'Kiran', 'params': kiran_params},
        {'file_prefix': 'DEAL_kb gems', 'org_prefix': 'kb', 'params': kb_params},
        {'file_prefix': 'DEAL_kb gems', 'org_prefix': 'kb', 'params': kb_params_web},
        {'file_prefix': 'DEAL_kbs', 'org_prefix': 'kbs', 'params': kbs_params},
        {'file_prefix': 'DEAL_komal', 'org_prefix': 'komal', 'params': komal_params},
        {'file_prefix': 'DEAL_kg ', 'org_prefix': 'kg', 'params': kg_params},
        {'file_prefix': 'DEAL_kg ', 'org_prefix': 'kg', 'params': kg_params_web},
        {'file_prefix': 'DEAL_kc', 'org_prefix': 'kc', 'params': kc_params},
        {'file_prefix': 'DEAL_kgk', 'org_prefix': 'kgk', 'params': kgk_params},
        {'file_prefix': 'DEAL_laxmi', 'org_prefix': 'laxmi', 'params': laxmi_params},
        {'file_prefix': 'DEAL_lakhi', 'org_prefix': 'lakhi', 'params': lakhi_params},
        {'file_prefix': 'DEAL_m suresh', 'org_prefix': 'm suresh', 'params': msuresh_params},
        {'file_prefix': 'DEAL_m d', 'org_prefix': 'md', 'params': m_d_params},
        {'file_prefix': 'DEAL_mb', 'org_prefix': 'mb', 'params': mb_params},
        {'file_prefix': 'DEAL_mb', 'org_prefix': 'mb', 'params': mb_params_web},
        {'file_prefix': 'DEAL_mahendra', 'org_prefix': 'mb', 'params': mb_params},
        {'file_prefix': 'DEAL_d navin', 'org_prefix': 'd navin', 'params': navin_params},
        {'file_prefix': 'DEAL_narola', 'org_prefix': 'narola', 'params': narola_params},
        {'file_prefix': 'DEAL_pragati', 'org_prefix': 'pragati', 'params': pragati_params},
        {'file_prefix': 'DEAL_pramukh', 'org_prefix': 'pramukh', 'params': pramukh_params},
        {'file_prefix': 'DEAL_red', 'org_prefix': 'red', 'params': red_params},
        {'file_prefix': 'DEAL_rays', 'org_prefix': 'rays', 'params': rays_params},
        {'file_prefix': 'DEAL_ratnakal', 'org_prefix': 'ratnakala', 'params': ratnakala_params},
        {'file_prefix': 'DEAL_rajgir', 'org_prefix': 'rajgir gems', 'params': rajgir_params},
        {'file_prefix': 'DEAL_r.rajesh', 'org_prefix': 'r.rajesh', 'params': rajesh_params},
        {'file_prefix': 'DEAL_r rajesh', 'org_prefix': 'r rajesh', 'params': rajesh_params_web},
        {'file_prefix': 'DEAL_r n', 'org_prefix': 'r n', 'params': r_n_params},
        {'file_prefix': 'DEAL_shivam', 'org_prefix': 'shivam', 'params': shivam_params},
        {'file_prefix': 'DEAL_shairu', 'org_prefix': 'Shairu', 'params': shairu_params},
        {'file_prefix': 'DEAL_starrays', 'org_prefix': 'starrays', 'params': starrays_params},
        {'file_prefix': 'DEAL_starlight', 'org_prefix': 'Starlight', 'params': starlight_params},
        {'file_prefix': 'DEAL_shree krishna', 'org_prefix': 'shree krishna', 'params': shree_krishna_params},
        {'file_prefix': 'DEAL_shyam', 'org_prefix': 'shyam', 'params': shyam_params},
        {'file_prefix': 'DEAL_rsd', 'org_prefix': 'srd', 'params': srd_params},
        {'file_prefix': 'DEAL_starrays', 'org_prefix': 'starrays RAYS', 'params': starrays_params},
        {'file_prefix': 'DEAL_starrays', 'org_prefix': 'STAR RAYS', 'params': starrays_params},
        {'file_prefix': 'DEAL_sheetal', 'org_prefix': 'sheetal', 'params': Sheetal_params},
        {'file_prefix': 'DEAL_sanghvi', 'org_prefix': 'sanghvi & sons', 'params': sanghvi_params},
        {'file_prefix': 'DEAL_ske', 'org_prefix': 'ske', 'params': ske_params},
        {'file_prefix': 'DEAL_snj', 'org_prefix': 'snj', 'params': snj_params},
        {'file_prefix': 'DEAL_sunrays', 'org_prefix': 'sunrays', 'params': sunrays_params},
        {'file_prefix': 'DEAL_unique', 'org_prefix': 'unique', 'params': unique_params},
        {'file_prefix': 'DEAL_vd', 'org_prefix': 'Vd', 'params': vdglobal_params}, 
        {'file_prefix': 'DEAL_vd', 'org_prefix': 'v d global', 'params': vdglobal_params}, 
        {'file_prefix': 'DEAL_vaibhav', 'org_prefix': 'vaibhav', 'params': vaibhav_params},      
        {'file_prefix': 'DEAL_veer', 'org_prefix': 'veer', 'params': veer_params},     
    ]
    # Output directory for the processed files
    output_directory = r"G:\Certificate\Python Output\VLOOKUP Output"

    # Loop through params_list and process the files
    for item in params_list:
        deal_file_path = find_latest_file(deal_folder_path, item['file_prefix'])
        org_file_path = find_latest_file(org_folder_path, item['org_prefix'])

        if deal_file_path is None or org_file_path is None:
            print(f"Skipping due to missing file(s): {item['file_prefix']} and {item['org_prefix']}")
            continue

        # Extract the file name (without extension) from the deal file path
        deal_file_name = os.path.splitext(os.path.basename(deal_file_path))[0]
        
        # Set the output file path using the deal file name
        output_file_path = os.path.join(output_directory, f"{deal_file_name}.xlsx")

        # Call the process function with the appropriate parameters
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
