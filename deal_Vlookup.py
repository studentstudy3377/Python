import os
import zipfile
from io import BytesIO
from datetime import datetime
import pandas as pd
import streamlit as st
import time

# ---------------- Utility helpers ---------------- #

def normalize_prefix(s: str) -> str:
    return (s or "").strip().upper()

def match_uploaded_file(files, prefix: str):
    """Return the most recent uploaded file whose name starts with prefix (case-insensitive, trimmed)."""
    if not files:
        return None
    norm_prefix = normalize_prefix(prefix)
    matches = [f for f in files if normalize_prefix(f.name).startswith(norm_prefix)]
    return matches[-1] if matches else None

def _norm_join_series(s: pd.Series) -> pd.Series:
    """Normalize join keys as uppercase, trimmed strings; preserve leading zeros."""
    return s.astype("string").str.strip().str.upper()

def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Convert DataFrame to Excel bytes."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ---------------- Core merge logic ---------------- #

def process_excel_files(
    org_file_obj,
    deal_file_obj,
    org_skiprows=None,
    org_merge_col=None,
    deal_merge_col=None,
    disc_column=None,
    is_disc_negative=False,
    sheet_name=None,
):
    """Reads DEAL and ORG Excel files, merges discount data, and computes PRICE/VALUE/DIFF."""
    # ---- Read DEAL ----
    deal_dtype = {}
    if deal_merge_col:
        deal_dtype[deal_merge_col] = "string"
    deal_df = pd.read_excel(deal_file_obj, dtype=deal_dtype)

    # ---- Read ORG ----
    read_kwargs = {}
    if sheet_name is not None:
        read_kwargs["sheet_name"] = sheet_name
    if org_skiprows is not None:
        read_kwargs["skiprows"] = org_skiprows
    org_dtype = {}
    if org_merge_col:
        org_dtype[org_merge_col] = "string"
    if disc_column:
        org_dtype[disc_column] = "string"
    org_df = pd.read_excel(org_file_obj, dtype=org_dtype, **read_kwargs)

    # ---- Normalize keys ----
    if deal_merge_col in deal_df.columns:
        deal_df[deal_merge_col] = _norm_join_series(deal_df[deal_merge_col])
    if org_merge_col in org_df.columns:
        org_df[org_merge_col] = _norm_join_series(org_df[org_merge_col])

    # ---- Numeric conversions ----
    for col in ["dRap", "dWt", "dDisc1"]:
        if col in deal_df.columns:
            deal_df[col] = pd.to_numeric(deal_df[col], errors="coerce")

    if disc_column not in org_df.columns:
        org_df[disc_column] = pd.NA

    org_df[disc_column] = pd.to_numeric(org_df[disc_column], errors="coerce")
    if not is_disc_negative:
        org_df[disc_column] = org_df[disc_column].apply(
            lambda x: (x * -1) if pd.notna(x) and x > 0 else x
        )

    org_slice_cols = [c for c in [org_merge_col, disc_column] if c and c in org_df.columns]
    org_slice = org_df[org_slice_cols].copy() if org_slice_cols else org_df.copy()

    # ---- Merge ----
    can_merge = (
        deal_merge_col and (deal_merge_col in deal_df.columns)
        and org_merge_col and (org_merge_col in org_slice.columns)
    )
    if can_merge:
        merged_df = pd.merge(
            deal_df,
            org_slice,
            left_on=deal_merge_col,
            right_on=org_merge_col,
            how="left",
        )
    else:
        merged_df = deal_df.copy()

    # ---- Rename discount column to today's date ----
    today_col = datetime.now().strftime("%d-%m-%Y")
    if disc_column in merged_df.columns:
        merged_df.rename(columns={disc_column: today_col}, inplace=True, errors="ignore")

    # ---- Compute PRICE/VALUE/DIFF ----
    if "dRap" in merged_df.columns and today_col in merged_df.columns:
        merged_df["PRICE"] = merged_df["dRap"] + (merged_df["dRap"] * merged_df[today_col] / 100)
    if "PRICE" in merged_df.columns and "dWt" in merged_df.columns:
        merged_df["VALUE"] = merged_df["PRICE"] * merged_df["dWt"]
    if "dDisc1" in merged_df.columns and today_col in merged_df.columns:
        merged_df["DIFF"] = merged_df["dDisc1"] - merged_df[today_col]

    if deal_merge_col and deal_merge_col in deal_df.columns:
        keep_len = int(deal_df[deal_merge_col].notna().sum())
        merged_df = merged_df.head(keep_len)

    merged_df.drop(columns=[org_merge_col], errors="ignore", inplace=True)
    return merged_df, today_col

def run_discount_vlookup():
    st.title("Discount VLOOKUP")
    # Uploads
    st.sidebar.header("Upload Excel Files")
    deal_files = st.sidebar.file_uploader(
        "Upload DEAL Excel files", type=["xls", "xlsx"], accept_multiple_files=True
    )
    org_files = st.sidebar.file_uploader(
        "Upload ORG Excel files", type=["xls", "xlsx"], accept_multiple_files=True
    )

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

    if not deal_files or not org_files:
        st.info("Please upload both DEAL and ORG files.")
        return

    st.success("Processing started …")
    total = len(params_list)
    progress = st.progress(0.0)  # Initialize the progress bar
    completed = 0

    # ✅ Create ZIP file in memory
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        added_files = set()

        for item in params_list:
            file_prefix = item["file_prefix"]
            org_prefix = item["org_prefix"]
            cfg = item["params"]

            deal_file = match_uploaded_file(deal_files, file_prefix)
            org_file = match_uploaded_file(org_files, org_prefix)
            completed += 1
            try:
                # Process the files and get merged dataframe
                df, today_col = process_excel_files(
                    org_file_obj=org_file,
                    deal_file_obj=deal_file,
                    org_skiprows=cfg.get("skip_rows"),
                    org_merge_col=cfg.get("org_merge_col"),
                    deal_merge_col=cfg.get("deal_merge_col"),
                    disc_column=cfg.get("disc_column"),
                    is_disc_negative=cfg.get("is_disc_negative", False),
                    sheet_name=cfg.get("sheet_name"),
                )

                # Check if any of the required columns (PRICE, VALUE, DIFF) are missing
                if df is None or not all(col in df.columns for col in ["PRICE", "VALUE", "DIFF"]):
                    # Skip this file as the necessary columns are missing
                    continue

                # Convert dataframe to Excel bytes
                xlsx_bytes = df_to_xlsx_bytes(df)
                safe_name = os.path.splitext(deal_file.name)[0]
                fname = f"{safe_name}.xlsx"
                
                # Ensure unique file name
                if fname in added_files:
                    fname = f"{safe_name}.xlsx"
                
                added_files.add(fname)
                zipf.writestr(fname, xlsx_bytes)

            except Exception as e:
                continue  # If error occurs, just skip this file

            # Update progress incrementally
            progress.progress(completed / total)  # Update progress bar in real-time

        progress.progress(1.0)

    st.success("🎉 All processing completed successfully!")

    # Add a download button for the ZIP file (from memory)
    zip_buffer.seek(0)  # Move to the beginning of the buffer
    st.download_button(
        label="Download File",
        data=zip_buffer,
        file_name="Discount Output.zip",
        mime="application/zip",
    )

if __name__ == "__main__":
    run_discount_vlookup()