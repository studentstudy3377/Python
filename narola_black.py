import os
import pandas as pd
import streamlit as st
from io import BytesIO

# ========== CORE FUNCTIONS ==========

def match_uploaded_file(file, prefix: str):
    """
    Return the uploaded file whose name starts with the given prefix (case-insensitive).
    """
    if not file:
        return None
    
    # As we are uploading a single file, we don't need to check for `name` attribute.
    # However, if you want to check the prefix, you can handle it based on filename.
    # Here we just assume the `file` is the correct one, as we are no longer accepting multiple files.
    if file.name.upper().startswith(prefix.upper()):
        return file
    return None


def process_excel_files(
    org_file_obj,
    deal_file_obj,
    org_skiprows: int | None = None,
    org_merge_col: str | None = None,
    deal_merge_col: str | None = None,
    dNatts_1: str | None = None,
    dNatts_2: str | None = None,
    stock_id: str | None = None,
    sheet_name: str | None = None,
) -> pd.DataFrame | None:
    """
    Reads DEAL and ORG Excel files (uploaded file-like objects), merges selected columns,
    and returns the processed file as a DataFrame.
    """
    try:
        # Read DEAL (default first sheet)
        deal_df = pd.read_excel(deal_file_obj)

        # Read ORG with optional sheet_name and skiprows
        read_kwargs = {}
        if sheet_name:
            read_kwargs["sheet_name"] = sheet_name
        if org_skiprows is not None:
            read_kwargs["skiprows"] = org_skiprows

        org_df = pd.read_excel(org_file_obj, **read_kwargs)

        # Ensure numeric join keys for robust matching
        if deal_merge_col and deal_merge_col in deal_df.columns:
            deal_df[deal_merge_col] = pd.to_numeric(deal_df[deal_merge_col], errors="coerce")
        if org_merge_col and org_merge_col in org_df.columns:
            org_df[org_merge_col] = pd.to_numeric(org_df[org_merge_col], errors="coerce")

        # Build list of columns to bring from ORG
        merge_columns = [org_merge_col] if org_merge_col else []
        if dNatts_1:
            merge_columns.append(dNatts_1)
        if dNatts_2:
            merge_columns.append(dNatts_2)
        if stock_id:
            merge_columns.append(stock_id)

        # If some requested columns may be missing, keep only those that exist to avoid KeyErrors
        if merge_columns:
            existing = [c for c in merge_columns if c in org_df.columns]
            # Ensure the join key is included if available
            if org_merge_col and org_merge_col in org_df.columns and org_merge_col not in existing:
                existing.insert(0, org_merge_col)
            org_slice = org_df[existing] if existing else org_df
        else:
            org_slice = org_df

        # Merge
        merged_df = pd.merge(
            deal_df,
            org_slice,
            left_on=deal_merge_col,
            right_on=org_merge_col if org_merge_col in org_slice.columns else None,
            how="left",
        )

        # Drop duplicate join column from ORG if present
        final_df = merged_df.drop(columns=[org_merge_col], errors="ignore")

        # Trim to non-null dLot length from DEAL so extra rows (if any) are not written
        if deal_merge_col and deal_merge_col in deal_df.columns:
            dlot_length = deal_df[deal_merge_col].notna().sum()
            final_df = final_df.head(int(dlot_length))

        # Map/rename the pulled columns into desired output names
        if stock_id and stock_id in final_df.columns:
            final_df["Stock Id"] = final_df[stock_id]
        if dNatts_1 and dNatts_1 in final_df.columns:
            final_df["dNatts_1"] = final_df[dNatts_1]
        if dNatts_2 and dNatts_2 in final_df.columns:
            final_df["dNatts_2"] = final_df[dNatts_2]

        # Remove original pulled columns if present
        cols_to_remove = [c for c in [stock_id, dNatts_1, dNatts_2] if c]
        final_df.drop(columns=cols_to_remove, errors="ignore", inplace=True)

        # Drop extra noisy columns if they exist
        columns = ['dLot', 'Stock Id', 'dSize', 'dWt', 'Lab', 'dShape', 'dColor',
                    'dClarity', 'dFlrc', 'cutName', 'dCut', 'dPol', 'dSym', 'dTable',
                    'dDepth', 'dDisc','dRap','Rap value', 'dValue', 'dDisc1', 'dValue1', 'M_DAYS',
                    'M_TREND', 'tot_Pcs', 'dSoldPrice', 'dSoldDisc', 'dNatts_1','dNatts_2']

        for col in columns:
            if col not in final_df.columns:
                final_df[col] = None
        final_df = final_df[columns]

        # Move "Stock Id" to 2nd column for readability, if present
        if "Stock Id" in final_df.columns:
            cols = list(final_df.columns)
            cols.remove("Stock Id")
            cols.insert(1, "Stock Id")
            final_df = final_df[cols]

        return final_df

    except Exception as e:  # noqa: BLE001
        return None

# ========== STREAMLIT UI ==========

def run_narola_black_mod():
    st.title("Narola Black VLOOKUP")

    st.sidebar.header("Upload Excel Files")

    deal_file = st.sidebar.file_uploader("Upload DEAL Excel file", type=["xls", "xlsx"])
    org_file = st.sidebar.file_uploader("Upload ORG Excel file", type=["xls", "xlsx"])

    # filename configuration mapping
    params_dict = {
            "NAROLA": {
                "file_prefix": "DEAL_narola",
                "org_prefix": "Narola",
                "params": {
                    "deal_merge_col": "dLot",
                    "org_merge_col": "ReportNo",
                    "stock_id": "StoneId",
                    "dNatts_1": "TableBlack",
                    "dNatts_2": "SideBlack",
                    "skip_rows": 0,
                },
            }
        }

    if not deal_file or not org_file:
        st.error("Please upload both DEAL and ORG Excel files.")
    else:
        st.success("Starting processing...")

        progress_bar = st.progress(0)
        total = len(params_dict)
        completed = 0

        # Process and display the processed dataframe
        for filename, config in params_dict.items():
            deal_file_obj = match_uploaded_file(deal_file, config["file_prefix"])
            org_file_obj = match_uploaded_file(org_file, config["org_prefix"])

            if not deal_file_obj or not org_file_obj:
                completed += 1
                progress_bar.progress(completed / total)
                continue

            # Process files and get the processed DataFrame
            processed_df = process_excel_files(
                org_file_obj=org_file_obj,
                deal_file_obj=deal_file_obj,
                org_skiprows=config["params"].get("skip_rows"),
                org_merge_col=config["params"]["org_merge_col"],
                deal_merge_col=config["params"]["deal_merge_col"],
                dNatts_1=config["params"]["dNatts_1"],
                dNatts_2=config["params"]["dNatts_2"],
                stock_id=config["params"]["stock_id"],
                sheet_name=config["params"].get("sheet_name"),
            )

            if processed_df is not None:
                st.dataframe(processed_df)

            completed += 1
            progress_bar.progress(completed / total)

            # -------- DOWNLOAD BUTTON --------
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                processed_df.to_excel(writer, index=False, sheet_name="ProcessedData")
            buffer.seek(0)

            out_name = os.path.basename(deal_file_obj.name)
            with st.sidebar:
                st.download_button(
                    label="Generate Excel",
                    data=buffer.getvalue(),
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ======== MAIN GUARD (as requested) ========

if __name__ == "__main__":
    run_narola_black_mod()