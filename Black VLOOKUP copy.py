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
    ankit_params = {
        'sheet_name': 'Sheet1',
        'deal_merge_col': 'dLot',
        'skip_rows': 10,
        'org_merge_col': 'Certificate No',
        'stock_id': 'Stock ID',    
        'dNatts_1': 'Natts',     
        'dNatts_2': '', 
    }
    asian_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 7,
        'org_merge_col': '',
        'stock_id': 'Report No.',    
        'dNatts_1': 'BIT',     
        'dNatts_2': 'BIS', 
    }
    
    
    
    diaclust_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 0,
        'org_merge_col': 'Report No',
        'stock_id': 'StoneNo',    
        'dNatts_1': 'BIC',     
        'dNatts_2': 'BIS',
    }

    finestar_params = {
        'sheet_name': 'diamonds',
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Packet No',
        'stock_id': '',    
        'dNatts_1': 'CN',     
        'dNatts_2': 'SN', 
    }
    gemstar_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 0,
        'org_merge_col': 'Report No.',
        'stock_id': 'Stock Id',    
        'dNatts_1': '',     
        'dNatts_2': '',         
    }
    
    glowstar_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 4,
        'org_merge_col': 'Report No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Natts',     
        'dNatts_2': '',         
    }
    
    gnyan_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'stock_id': 'Stone ID',    
        'dNatts_1': '',     
        'dNatts_2': '',         
    }
    
    hk_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Stock No',
        'stock_id': '',    
        'dNatts_1': 'BT',     
        'dNatts_2': 'BC',
    }
    hrv_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Natts',     
        'dNatts_2': '',
    }
    hrv_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Stone No',    
        'dNatts_1': '',     
        'dNatts_2': '',
    }
    haridas_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Stock Id',    
        'dNatts_1': '',     
        'dNatts_2': '',
    }
    jodhani_params = {
        'org_merge_col': 'Report No',
        'deal_merge_col': 'dLot',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',
    }
    jewlex_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',
    }
    kbs_params = {
        'sheet_name': 'stock',
        'deal_merge_col': 'dLot',
        'skip_rows': 2,
        'org_merge_col': 'CertNo',
        'stock_id': 'StoneID',    
        'dNatts_1': 'BIC',     
        'dNatts_2': 'BIS',
    }
    komal_params= {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Stock ID',    
        'dNatts_1': '',     
        'dNatts_2': '',
    }
    kb_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 2,
        'org_merge_col': 'CertNo',
        'stock_id': 'Ref No',    
        'dNatts_1': 'Black Table Inclusion',     
        'dNatts_2': 'Black Side Inclusion',
    }
    kc_params = {
        'sheet_name': 0,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Cert No',
        'stock_id': 'Lot No.',    
        'dNatts_1': '',     
        'dNatts_2': '',
    }
    # kg_params = {
    #     'skip_rows': 1,
    #     'org_merge_col': 'REPORT NO.',
    #     'deal_merge_col': 'dLot',
    #     'stock_id': 'Stone No',    
    #     'dNatts_1': 'Natts',     
    #     'dNatts_2': '',
    # }
    kg_params = {
        'skip_rows': 1,
        'org_merge_col': 'REPORT NO.',
        'deal_merge_col': 'dLot',
        'stock_id': 'STONE ID',    
        'dNatts_1': 'SIDE BLACK INCLUSION ',     
        'dNatts_2': 'TABLE BLACK INCLUSION',
    }
    # kgk_params = {
    #     'deal_merge_col': 'dLot',
    #     'skip_rows': 6,
    #     'org_merge_col': 'Cert #',
    #     'stock_id': 'Stone No',    
    #     'dNatts_1': 'Natts',     
    #     'dNatts_2': '',   
    # }
    
    kgk_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Cert #',
        'stock_id': '',    
        'dNatts_1': 'TBL BLK',     
        'dNatts_2': 'Side BLK',
    }
    kiran_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 10,
        'org_merge_col': 'Rep No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'BIC',     
        'dNatts_2': 'BIS',
    }
    laxmi_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 5,
        'org_merge_col': 'Certificate',
        'stock_id': 'Stone ID',    
        'dNatts_1': 'Black Inc',     
        'dNatts_2': '',
    }
    lakhi_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report #',
        'stock_id': 'Stock #',    
        'dNatts_1': 'Black Inclusion   ',     
        'dNatts_2': '',
    }
    mb_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Cert No.',
        'stock_id': 'Stock ID',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',
    }
    mahendra_params = {
        'skip_rows': 1,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Cert No.',
        'stock_id': 'Stock ID',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',
    }
    msuresh_params = {
        'skip_rows': 1,
        'deal_merge_col': 'dLot',   
        'org_merge_col': 'Report No',
        'stock_id': 'Packet No',    
        'dNatts_1': 'Tab_BLK',     
        'dNatts_2': 'Side_BLK',
    }
    navin_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'stock_id': 'Packet No',    
        'dNatts_1': 'Centre Black',     
        'dNatts_2': 'Crown Black',
    }
    parishi_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Lot no.',
        'stock_id': '',    
        'dNatts_1': 'BIT',     
        'dNatts_2': 'BIC',
    }
    pramukh_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Stone ID',    
        'dNatts_1': 'T.BLK',     
        'dNatts_2': 'S.BLK',
    }
    # pragati_params = {
    #     'deal_merge_col': 'dLot',
    #     'org_merge_col': 'Report No',
    #     'stock_id': 'Stone No',    
    #     'dNatts_1': 'Natts',     
    #     'dNatts_2': '',
    # }
    pragati_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Ref No',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',
    }
    rays_params = {
        'skip_rows': 1,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Packet No',    
        'dNatts_1': 'BLACK',     
        'dNatts_2': '',
    }
    rajgir_params = {
        'sheet_name': 'Rajgir Gems STOCK LIST',
        'deal_merge_col': 'dLot',
        'skip_rows': 3,
        'org_merge_col': 'REPORT NO',
        'stock_id': 'Stone ID',    
        'dNatts_1': 'TB',     
        'dNatts_2': 'SB',

    }
    rajesh_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 5,
        'org_merge_col': 'Report No',
        'stock_id': 'Stone ID',    
        'dNatts_1': 'TB',     
        'dNatts_2': 'SB',  
    }
    ratnakala_params = {
        'sheet_name': 'Stone List',
        'skip_rows': 4,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',  
    }
    ratnakala_mail_param = {
        'sheet_name': 'Stone List',
        'skip_rows': 4,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Stone Id',    
        'dNatts_1': 'TB',     
        'dNatts_2': 'SB',  
    }
    rsd_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Cert No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Natts',     
        'dNatts_2': '',   
    }
    rds_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Cert No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Natts',     
        'dNatts_2': '',  
    }
    sanghvi_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 8,
        'org_merge_col': 'CERTNO',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Natts',     
        'dNatts_2': '',   
    }
    sanghvi_param = {
        'deal_merge_col': 'dLot',
        'skip_rows': 5,
        'org_merge_col': 'CERTNO',
        'stock_id': 'REFNO',    
        'dNatts_1': '',     
        'dNatts_2': '',   
    }
    shairu_params = {
        'skip_rows': 5,
        'org_merge_col': 'Certi No.',
        'deal_merge_col': 'dLot',
        'stock_id': 'Stock Id',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Crown Black',
    }
    starrays_params = {
        'skip_rows': 3,
        'org_merge_col': 'Cert.No',
        'deal_merge_col': 'dLot',
        'stock_id': '',    
        'dNatts_1': 'BIT',     
        'dNatts_2': 'BIC',
    }

    starlight_params = {
        'org_merge_col': 'Cert.No.',
        'deal_merge_col': 'dLot',
         'stock_id': 'StockID',    
        'dNatts_1': 'Black In Table',     
        'dNatts_2': 'Side Black',
    }
    Sheetal_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 8,
        'org_merge_col': 'Certificate',
        'stock_id': 'StoneID',    
        'dNatts_1': 'BIT',     
        'dNatts_2': 'BIC',
    }
    shree_krishna_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',
    }
    starrays_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 3,
        'org_merge_col': 'Cert.No',
        'stock_id': '',    
        'dNatts_1': 'BIT',     
        'dNatts_2': 'BIC',
    }
    srd_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 4,
        'org_merge_col': 'Cert No',
        'stock_id': 'Stone No',    
        'dNatts_1': 'Natts',     
        'dNatts_2': '',
    }
    snj_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 5,
        'org_merge_col': 'Certificate No',
        'stock_id': 'Stone Id',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',
    }
    snj_param = {
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Certificate No',
        'stock_id': 'Stone Id',    
        'dNatts_1': 'Table Black',     
        'dNatts_2': 'Side Black',
    }
    sagar_params = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'CertNo',
        'stock_id': 'StockID',    
        'dNatts_1': '',     
        'dNatts_2': '',
    }
    unique_params = {
        'skip_rows': 4,
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Packet No',    
        'dNatts_1': 'TB',     
        'dNatts_2': 'SB',
    }
    unique_param = {
        'deal_merge_col': 'dLot',
        'org_merge_col': 'Report No',
        'stock_id': 'Packet No',    
        'dNatts_1': 'TB',     
        'dNatts_2': 'SB',
    }
    vdglobal_params = {
        'deal_merge_col': 'dLot',
        'skip_rows': 4,
        'org_merge_col': 'ReportNo',
        'stock_id': 'Stock No',    
        'dNatts_1': 'BT',     
        'dNatts_2': '',  
    }
    vaibhav_params = {
        'sheet_name': 'Stock',
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'Report No',
        'stock_id': 'Packet No',    
        'dNatts_1': 'CB',     
        'dNatts_2': 'SB',
    }   
    vinay_params = {
        'sheet_name': 'GIA',
        'deal_merge_col': 'dLot',
        'skip_rows': 1,
        'org_merge_col': 'ReportNumber',
        'stock_id': 'StoneID',    
        'dNatts_1': 'Inclusion',     
        'dNatts_2': '',
    }   
    params_list = [  
        {'file_prefix': 'DEAL_akarsh', 'org_prefix': 'AKARSH', 'params': akarsh_params},
        {'file_prefix': 'DEAL_ankit', 'org_prefix': 'ANKIT', 'params': ankit_params},
        {'file_prefix': 'DEAL_asian', 'org_prefix': 'ASIAN', 'params': asian_params},
        {'file_prefix': 'DEAL_d navin', 'org_prefix': 'D NAVIN', 'params': navin_params},
        {'file_prefix': 'DEAL_diaclust', 'org_prefix': 'Diaclust', 'params': diaclust_params},
        {'file_prefix': 'DEAL_finestar', 'org_prefix': 'FINESTAR', 'params': finestar_params},
        {'file_prefix': 'DEAL_gemstar', 'org_prefix': 'GEMSTAR', 'params': gemstar_params},
        {'file_prefix': 'DEAL_glowstar', 'org_prefix': 'GLOWSTAR', 'params': glowstar_params},
        {'file_prefix': 'DEAL_gnyan', 'org_prefix': 'GNYAN', 'params': gnyan_params},
        {'file_prefix': 'DEAL_hk', 'org_prefix': 'HK', 'params': hk_params},
        {'file_prefix': 'DEAL_hrv', 'org_prefix': 'HRV', 'params': hrv_params},
        {'file_prefix': 'DEAL_haridias', 'org_prefix': 'HIRADIAS', 'params': haridas_params},
        {'file_prefix': 'DEAL_jodhani', 'org_prefix': 'Jodhani', 'params': jodhani_params},
        {'file_prefix': 'DEAL_jewelex', 'org_prefix': 'Jewelex', 'params': jewlex_params},
        {'file_prefix': 'DEAL_kiran', 'org_prefix': 'Kiran', 'params': kiran_params},
        {'file_prefix': 'DEAL_kb gems ', 'org_prefix': 'KB GEMS', 'params': kb_params},
        {'file_prefix': 'DEAL_kbs', 'org_prefix': 'KBS', 'params': kbs_params},
        {'file_prefix': 'DEAL_komal', 'org_prefix': 'KOMAL', 'params': komal_params},
        {'file_prefix': 'DEAL_kg ', 'org_prefix': 'KG', 'params': kg_params},
        {'file_prefix': 'DEAL_kc', 'org_prefix': 'KC', 'params': kc_params},
        {'file_prefix': 'DEAL_kgk', 'org_prefix': 'KGK', 'params': kgk_params},
        {'file_prefix': 'DEAL_laxmi', 'org_prefix': 'LAXMI', 'params': laxmi_params},
        {'file_prefix': 'DEAL_lakhi', 'org_prefix': 'LAKHI', 'params': lakhi_params},
        {'file_prefix': 'DEAL_m suresh', 'org_prefix': 'M SURESH', 'params': msuresh_params},
        {'file_prefix': 'DEAL_mahendra', 'org_prefix': 'MAHENDRA', 'params': mahendra_params},
        {'file_prefix': 'DEAL_mb', 'org_prefix': 'MB', 'params': mb_params},
        {'file_prefix': 'DEAL_parishi', 'org_prefix': 'PARISHI', 'params': parishi_params},
        {'file_prefix': 'DEAL_pragati', 'org_prefix': 'PRAGATI', 'params': pragati_params},
        {'file_prefix': 'DEAL_pramukh', 'org_prefix': 'PRAMUKH', 'params': pramukh_params},
        {'file_prefix': 'DEAL_rays', 'org_prefix': 'RAYS', 'params': rays_params},
        {'file_prefix': 'DEAL_ratnakal', 'org_prefix': 'RATNAKALA', 'params': ratnakala_params},
        # {'file_prefix': 'DEAL_ratnakal', 'org_prefix': 'RATNAKALA', 'params': ratnakala_mail_param},
        {'file_prefix': 'DEAL_rajgir', 'org_prefix': 'RAJGIR GEMS', 'params': rajgir_params},
        {'file_prefix': 'DEAL_r.rajesh', 'org_prefix': 'R RAJESH', 'params': rajesh_params},
        {'file_prefix': 'DEAL_r rajesh', 'org_prefix': 'R.RAJESH', 'params': rajesh_params},    
        {'file_prefix': 'DEAL_rsd', 'org_prefix': 'RSD', 'params': rsd_params}, 
        {'file_prefix': 'DEAL_rds', 'org_prefix': 'RDS', 'params': rds_params},
        {'file_prefix': 'DEAL_shairu', 'org_prefix': 'SHAIRU', 'params': shairu_params},
        {'file_prefix': 'DEAL_starlight', 'org_prefix': 'STARLIGHTDIAMOND', 'params': starlight_params},
        {'file_prefix': 'DEAL_shree krishna', 'org_prefix': 'SHREE KRISHNA', 'params': shree_krishna_params},
        {'file_prefix': 'DEAL_rsd', 'org_prefix': 'SRD', 'params': srd_params},
        {'file_prefix': 'DEAL_snj', 'org_prefix': 'SNJ', 'params': snj_params}, 
        {'file_prefix': 'DEAL_snj', 'org_prefix': 'SNJ', 'params': snj_param}, 
        {'file_prefix': 'DEAL_sagar', 'org_prefix': 'SAGAR', 'params': sagar_params}, 
        {'file_prefix': 'DEAL_starrays', 'org_prefix': 'STAR RAYS', 'params': starrays_params},
        {'file_prefix': 'DEAL_sheetal', 'org_prefix': 'SHEETAL', 'params': Sheetal_params},
        # {'file_prefix': 'DEAL_sanghvi', 'org_prefix': 'SANGHVI', 'params': sanghvi_params},
        {'file_prefix': 'DEAL_sanghvi', 'org_prefix': 'SANGHVI', 'params': sanghvi_param},
        {'file_prefix': 'DEAL_unique', 'org_prefix': 'UNIQUE', 'params': unique_params},
        {'file_prefix': 'DEAL_unique', 'org_prefix': 'UNIQUE', 'params': unique_param},
        {'file_prefix': 'DEAL_vd', 'org_prefix': 'VD GLOBAL', 'params': vdglobal_params}, 
        {'file_prefix': 'DEAL_vaibhav', 'org_prefix': 'VAIBHAV', 'params': vaibhav_params},      
        {'file_prefix': 'DEAL_vinay diamonds', 'org_prefix': 'VINAY', 'params': vinay_params},      
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
