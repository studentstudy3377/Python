import os
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

def normalize_filename(value: str) -> str:
   
    base = os.path.splitext(os.path.basename(str(value)))[0]
    return "".join(ch for ch in base.upper() if ch.isalnum())

def match_uploaded_file(files, prefix: str):
   
    if not files:
        return None

    prefix_normalized = normalize_filename(prefix)

    matches = [
        f
        for f in files
        if normalize_filename(f.name).startswith(prefix_normalized)
    ]

    return matches[-1] if matches else None

def read_uploaded_table(file_obj,skiprows: int | None = None,sheet_name: str | None = None,):
    
    if file_obj is None:
        raise ValueError("No file provided.")

    if hasattr(file_obj, "seek"):
        file_obj.seek(0)

    file_name = getattr(file_obj, "name", "")
    ext = os.path.splitext(file_name)[1].lower()

    if ext == ".csv":
        read_kwargs = {}

        if skiprows is not None:
            read_kwargs["skiprows"] = skiprows

        try:
            return pd.read_csv(file_obj,encoding="utf-8-sig",**read_kwargs,)
              
        except UnicodeDecodeError:
            if hasattr(file_obj, "seek"):
                file_obj.seek(0)

            return pd.read_csv(file_obj,encoding="latin1",**read_kwargs,)
            
    if ext in {".xls", ".xlsx", ".xlsm"}:
        read_kwargs = {}

        if sheet_name:
            read_kwargs["sheet_name"] = sheet_name

        if skiprows is not None:
            read_kwargs["skiprows"] = skiprows

        return pd.read_excel( file_obj,**read_kwargs, )
           
    raise ValueError(f"Unsupported file type: {ext or 'unknown'}")
      
def uploaded_file_bytes(file_obj) -> bytes:
    """Return uploaded file bytes without depending on the current pointer."""
    if hasattr(file_obj, "getvalue"):
        return file_obj.getvalue()

    if hasattr(file_obj, "seek"):
        file_obj.seek(0)

    data = file_obj.read()

    if hasattr(file_obj, "seek"):
        file_obj.seek(0)

    return data


def normalize_sheet_text(value: str) -> str:
    """Normalize worksheet text for case-insensitive prefix matching."""
    return " ".join(
        str(value).replace("\xa0", " ").split()
    ).casefold()

def get_deal_sheet_names(deal_file_obj) -> list[str]:
    """Return all worksheet names from the one uploaded DEAL workbook."""
    deal_name = getattr(deal_file_obj, "name", "")
    ext = os.path.splitext(deal_name)[1].lower()

    if ext not in {".xlsx", ".xlsm"}:
        raise ValueError("DEAL must be an .xlsx or .xlsm workbook.")
            
    workbook = load_workbook(BytesIO(uploaded_file_bytes(deal_file_obj)),read_only=True,keep_vba=(ext == ".xlsm"))
   
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()

def find_sheet_starting_with(deal_file_obj,sheet_prefix: str,) -> str:
    
    if not sheet_prefix:
        raise ValueError("Sheet prefix is required.")

    sheet_names = get_deal_sheet_names(deal_file_obj)
    prefix = normalize_sheet_text(sheet_prefix)

    matches = [
        name
        for name in sheet_names
        if normalize_sheet_text(name).startswith(prefix)
    ]

    if not matches:
        available = ", ".join(sheet_names) if sheet_names else "(none)"

        raise ValueError( f"No DEAL sheet starts with '{sheet_prefix}'. "f"Available sheets: {available}" )
         
    # Prefer exact name if available.
    for name in matches:
        if normalize_sheet_text(name) == prefix:
            return name

    return matches[0]


def excel_safe_value(value):
    """Convert pandas/numpy missing/scalar values to openpyxl-safe values."""
    if pd.isna(value):
        return None

    if hasattr(value, "item"):
        try:
            return value.item()
        except (ValueError, AttributeError):
            pass

    return value

def build_processed_dataframe(
    org_file_obj,
    deal_file_obj,
    org_skiprows: int | None = None,
    org_merge_col: str | None = None,
    deal_merge_col: str | None = None,
    dNatts_1: str | None = None,
    dNatts_2: str | None = None,
    stock_id: str | None = None,
    org_sheet_name: str | None = None,
    deal_sheet_prefix: str | None = None,
) -> tuple[str, pd.DataFrame]:
    if not deal_sheet_prefix:
        raise ValueError("DEAL sheet prefix is required.")

    actual_deal_sheet_name = find_sheet_starting_with(
        deal_file_obj,
        deal_sheet_prefix,
    )

    # Read the matching sheet from the ONE DEAL workbook.
    deal_df = read_uploaded_table(
        deal_file_obj,
        sheet_name=actual_deal_sheet_name,
    )

    # Read the supplier ORG file.
    org_df = read_uploaded_table(
        org_file_obj,
        skiprows=org_skiprows,
        sheet_name=org_sheet_name,
    )

    if not deal_merge_col or deal_merge_col not in deal_df.columns:
        raise ValueError(
            f"DEAL merge column '{deal_merge_col}' was not found "
            f"in sheet '{actual_deal_sheet_name}'."
        )

    if not org_merge_col or org_merge_col not in org_df.columns:
        raise ValueError(
            f"ORG merge column '{org_merge_col}' was not found "
            f"in file '{getattr(org_file_obj, 'name', '')}'."
        )

    # Convert matching columns to numeric for consistent joins.
    deal_df[deal_merge_col] = pd.to_numeric(
        deal_df[deal_merge_col],
        errors="coerce",
    )

    org_df[org_merge_col] = pd.to_numeric(
        org_df[org_merge_col],
        errors="coerce",
    )

    # Columns required from ORG.
    merge_columns = [org_merge_col]

    for col in (dNatts_1, dNatts_2, stock_id):
        if col and col not in merge_columns:
            merge_columns.append(col)

    missing_org_columns = [
        col
        for col in merge_columns
        if col not in org_df.columns
    ]

    if missing_org_columns:
        raise ValueError(
            f"ORG column(s) not found in "
            f"'{getattr(org_file_obj, 'name', '')}': "
            + ", ".join(missing_org_columns)
        )

    org_slice = org_df[merge_columns].copy()

    merged_df = pd.merge(
        deal_df,
        org_slice,
        left_on=deal_merge_col,
        right_on=org_merge_col,
        how="left",
    )

    final_df = merged_df.drop(
        columns=[org_merge_col],
        errors="ignore",
    )

    # Keep same behavior as your existing code:
    # output only the rows corresponding to non-empty dLot values.
    dlot_length = deal_df[deal_merge_col].notna().sum()
    final_df = final_df.head(int(dlot_length))

    # Map supplier columns to DEAL destination columns.
    if stock_id and stock_id in final_df.columns:
        if stock_id != "Stock Id":
            final_df["Stock Id"] = final_df[stock_id]

    if dNatts_1 and dNatts_1 in final_df.columns:
        if dNatts_1 != "dNatts_1":
            final_df["dNatts_1"] = final_df[dNatts_1]

    if dNatts_2 and dNatts_2 in final_df.columns:
        if dNatts_2 != "dNatts_2":
            final_df["dNatts_2"] = final_df[dNatts_2]

    # Remove supplier source field names after mapping.
    cols_to_remove = []

    if stock_id and stock_id != "Stock Id":
        cols_to_remove.append(stock_id)

    if dNatts_1 and dNatts_1 != "dNatts_1":
        cols_to_remove.append(dNatts_1)

    if dNatts_2 and dNatts_2 != "dNatts_2":
        cols_to_remove.append(dNatts_2)

    final_df.drop(columns=cols_to_remove,errors="ignore",inplace=True)
    
    output_columns = ['dLot', 'Stock Id', 'dSize', 'dWt', 'Lab', 'dShape', 'dColor',
                        'dClarity', 'dFlrc', 'cutName', 'dCut', 'dPol', 'dSym', 'dTable',
                        'dDepth', 'dDisc','dRap','Rap value', 'dValue', 'dDisc1', 'dValue1', 'M_DAYS',
                        'M_TREND', 'tot_Pcs', 'dSoldPrice', 'dSoldDisc', 'dNatts_1','dNatts_2']

    for col in output_columns:
        if col not in final_df.columns:
            final_df[col] = None

    final_df = final_df[output_columns]

    return actual_deal_sheet_name, final_df

def write_all_processed_sheets(deal_file_obj,processed_sheets: list[tuple[str, pd.DataFrame]],) -> BytesIO:
    
    deal_name = getattr(deal_file_obj, "name", "")
    ext = os.path.splitext(deal_name)[1].lower()

    if ext not in {".xlsx", ".xlsm"}:
        raise ValueError("To preserve all DEAL worksheets, upload DEAL as .xlsx or .xlsm.")
          
    keep_vba = ext == ".xlsm"

    workbook = load_workbook(BytesIO(uploaded_file_bytes(deal_file_obj)),keep_vba=keep_vba,)
       
    try:
        for target_sheet, final_df in processed_sheets:

            if target_sheet not in workbook.sheetnames:
                raise ValueError(f"Sheet '{target_sheet}' was not found in DEAL workbook.")
                    
            ws = workbook[target_sheet]

            old_max_row = ws.max_row
            old_max_col = ws.max_column

            for row in ws.iter_rows(min_row=1,max_row=old_max_row,min_col=1,max_col=old_max_col,):
                
                for cell in row:
                    cell.value = None

            # Headers.
            for col_idx, column_name in enumerate(final_df.columns,start=1,):
                
                ws.cell(row=1,column=col_idx,value=str(column_name),)
               
            # Data.
            for row_idx, values in enumerate( final_df.itertuples(index=False,name=None,),start=2,):
               
                for col_idx, value in enumerate(values,start=1,):
                    
                    ws.cell(row=row_idx,column=col_idx,value=excel_safe_value(value),)
                        
        output = BytesIO()
        workbook.save(output)
        output.seek(0)

        return output

    finally:
        workbook.close()

def run_black_sheet_mod():
    st.title("Black VLOOKUP")

    st.sidebar.header("Upload Excel Files")

    # IMPORTANT:
    # There is only ONE DEAL workbook.
    deal_file = st.sidebar.file_uploader(
        "Upload ONE DEAL Excel file",
        type=["xlsx", "xlsm"],
        accept_multiple_files=False,
    )

    # Supplier / ORG files remain separate.
    org_files = st.sidebar.file_uploader(
        "Upload ORG Excel/CSV files",
        type=["xls", "xlsx", "xlsm", "csv"],
        accept_multiple_files=True,
    )

    params_dict = {
        "AKARSH": {
            "org_prefix": "AKARSH",
            "params": {
                "org_sheet_name": "Stock",
                "deal_sheet_prefix": "AKARSH",
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "Black Incl.",
                "dNatts_2": "Table Incl.",
                "skip_rows": 1,
            },
        },
       "ANKIT": {
            "file_prefix": "DEAL_ankit",
            "org_prefix": "ANKIT",
            "params": {
                # "org_sheet_name": "Sheet1",
                "deal_sheet_prefix": "ANKIT",
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate No",
                "stock_id": "Stock ID",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 10,
            },
        },
        "ASIAN": {
            "file_prefix": "DEAL_asian",
            "org_prefix": "ASIAN",
            "params": {
                "deal_sheet_prefix": "ASIAN",
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No.",
                "stock_id": "Lot No.",
                "dNatts_1": "BIT",
                "dNatts_2": "BIS",
                "skip_rows": 7,
            },
        },
        "D NAVIN": {
            "file_prefix": "DEAL_d navin",
            "org_prefix": "D NAVIN",
            "params": {
                "deal_sheet_prefix": "D NAVIN",
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "Centre Black",
                "dNatts_2": "Crown Black",
                "skip_rows": 1,
            },
        },
        "DIACLUST": {
            "file_prefix": "DEAL_diaclust",
            "org_prefix": "Diaclust",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "DIACLUST",
                "org_merge_col": "Report No",
                "stock_id": "StoneNo",
                "dNatts_1": "BIC",
                "dNatts_2": "BIS",
                "skip_rows": 0,
            },
        },
        "FINESTAR": {
            "file_prefix": "DEAL_finestar",
            "org_prefix": "FINESTAR",
            "params": {
                "sheet_name": "diamonds",
                "deal_sheet_prefix": "FINESTAR",
                "deal_merge_col": "dLot",
                "org_merge_col": "Packet No",
                "stock_id": "",
                "dNatts_1": "CN",
                "dNatts_2": "SN",
                "skip_rows": 1,
            },
        },
        "GEMSTAR": {
            "file_prefix": "DEAL_gemstar",
            "org_prefix": "GEMSTAR",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "GEMSTAR",
                "org_merge_col": "Report No.",
                "stock_id": "Stock Id",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 0,
            },
        },
        "GLOWSTAR": {
            "file_prefix": "DEAL_glowstar",
            "org_prefix": "GLOWSTAR",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "GLOWSTAR",
                "org_merge_col": "Report No",
                "stock_id": "Stone No",
                "dNatts_1": "",
                "dNatts_2": "",
                "skip_rows": 4,
            },
        },
        "GNYAN": {
            "file_prefix": "DEAL_gnyan",
            "org_prefix": "GNYAN",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "GNYAN",
                "org_merge_col": "Report No",
                "stock_id": "Stone ID",
                "dNatts_1": "",
                "dNatts_2": "",
                "skip_rows": 1,
            },
        },
        "HK": {
            "file_prefix": "DEAL_hk",
            "org_prefix": "HK",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "HK",
                "org_merge_col": "Stock No",
                "stock_id": "",
                "dNatts_1": "BT",
                "dNatts_2": "BC",
            },
        },
        "HRV": {
            "file_prefix": "DEAL_hrv",
            "org_prefix": "HRV",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "HRV",
                "org_merge_col": "Report No",
                "stock_id": "Stone No",
                "dNatts_1": "",
                "dNatts_2": "",
            },
        },
        "HIRADIAS": {
            "file_prefix": "DEAL_haridias",
            "org_prefix": "HIRADIAS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "HIRADIAS",
                "org_merge_col": "Report No",
                "stock_id": "Stock Id",
                "dNatts_1": "",
                "dNatts_2": "",
            },
        },
        "JODHANI": {
            "file_prefix": "DEAL_jodhani",
            "org_prefix": "Jodhani",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "deal_sheet_prefix": "JODHANI",
                "stock_id": "Stone No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
            },
        },
        "JEWELEX": {
            "file_prefix": "DEAL_jewelex",
            "org_prefix": "Jewelex",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "deal_sheet_prefix": "JEWELEX",
                "stock_id": "Stone No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 1,
            },
        },
        "KIRAN": {
            "file_prefix": "DEAL_kiran",
            "org_prefix": "Kiran",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KIRAN",
                "org_merge_col": "Rep No",
                "stock_id": "Stone No",
                "dNatts_1": "BIC",
                "dNatts_2": "BIS",
                "skip_rows": 10,
            },
        },
        "KB GEMS": {
            "file_prefix": "DEAL_kb gems ",
            "org_prefix": "KB GEMS",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "CertNo",
                "deal_sheet_prefix": "KB GEMS",
                "stock_id": "Ref No",
                "dNatts_1": "Black Table Inclusion",
                "dNatts_2": "Black Side Inclusion",
                "skip_rows": 2,
            },
        },
        "KBS": {
            "file_prefix": "DEAL_kbs",
            "org_prefix": "KBS",
            "params": {
                "sheet_name": "stock",
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KBS",
                "org_merge_col": "CertNo",
                "stock_id": "StoneID",
                "dNatts_1": "BIC",
                "dNatts_2": "BIS",
                "skip_rows": 2,
            },
        },
        "KOMAL": {
            "file_prefix": "DEAL_komal",
            "org_prefix": "KOMAL",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KOMAL",
                "org_merge_col": "Report No",
                "stock_id": "Stock ID",
                "dNatts_1": "",
                "dNatts_2": "",
            },
        },
        "KG": {
            "file_prefix": "DEAL_kg ",
            "org_prefix": "KG",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "REPORT NO.",
                "deal_sheet_prefix": "KG",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 1,
            },
        },
        "KG_ALT": {
            "file_prefix": "DEAL_kg ",
            "org_prefix": "KG",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "REPORT NO.",
                "deal_sheet_prefix": "KG",
                "stock_id": "STONE ID",
                "dNatts_1": "SIDE BLACK INCLUSION ",
                "dNatts_2": "TABLE BLACK INCLUSION",
                "skip_rows": 0,
            },
        },
        "KC": {
            "file_prefix": "DEAL_kc",
            "org_prefix": "KC",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KC",
                "org_merge_col": "Cert No",
                "stock_id": "Lot No.",
                "dNatts_1": "",
                "dNatts_2": "",
            },
        },
        "KGK": {
            "file_prefix": "DEAL_kgk",
            "org_prefix": "KGK",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KGK",
                "org_merge_col": "Cert #",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 6,
            },
        },
        "KGK_ALT": {
            "file_prefix": "DEAL_kgk",
            "org_prefix": "KGK",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KGK",
                "org_merge_col": "Cert #",
                "stock_id": "",
                "dNatts_1": "TBL BLK",
                "dNatts_2": "Side BLK",
                "skip_rows": 1,
            },
        },
        "LAXMI": {
            "file_prefix": "DEAL_laxmi",
            "org_prefix": "LAXMI",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate",
                "deal_sheet_prefix": "LAXMI",
                "stock_id": "Stone ID",
                "dNatts_1": "Black Inc",
                "dNatts_2": "",
                "skip_rows": 5,
            },
        },
        "LAKHI": {
            "file_prefix": "DEAL_lakhi",
            "org_prefix": "LAKHI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "LAKHI",
                "org_merge_col": "Report #",
                "stock_id": "Stock #",
                "dNatts_1": "Black Inclusion   ",
                "dNatts_2": "",
            },
        },
        "MB": {
            "file_prefix": "DEAL_mb",
            "org_prefix": "MB",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "MB",
                "org_merge_col": "Cert No.",
                "stock_id": "Stock ID",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 0,
            },
        },
        "MAHENDRA": {
            "file_prefix": "DEAL_mahendra",
            "org_prefix": "MAHENDRA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "MAHENDRA",
                "org_merge_col": "Cert No.",
                "stock_id": "Stock ID",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 1,
            },
        },
        "M.SURESH": {
            "file_prefix": "DEAL_m suresh",
            "org_prefix": "M.SURESH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "M SURESH",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "Tab_BLK",
                "dNatts_2": "Side_BLK",
                "skip_rows": 1,
            },
        },
        "M SURESH": {
            "file_prefix": "DEAL_m suresh",
            "org_prefix": "M SURESH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "M SURESH",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "Tab_BLK",
                "dNatts_2": "Side_BLK",
                "skip_rows": 1,
            },
        },
        "PARISHI": {
            "file_prefix": "DEAL_parishi",
            "org_prefix": "PARISHI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "PARISHI",
                "org_merge_col": "Lot no.",
                "stock_id": "",
                "dNatts_1": "BIT",
                "dNatts_2": "BIC",
            },
        },
        "PRAGATI": {
            "file_prefix": "DEAL_pragati",
            "org_prefix": "PRAGATI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "PRAGATI",
                "org_merge_col": "Report No",
                "stock_id": "Ref No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
            }, 
        },
        "PRAMUKH": {
            "file_prefix": "DEAL_pramukh",
            "org_prefix": "PRAMUKH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "PARAMUKH",
                "org_merge_col": "Report No",
                "stock_id": "Stone ID",
                "dNatts_1": "T.BLK",
                "dNatts_2": "S.BLK",
            },
        },
        "RAYS": {
            "file_prefix": "DEAL_rays",
            "org_prefix": "RAYS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "RAYS",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "BLACK",
                "dNatts_2": "",
                "skip_rows": 1,
            },
        },
        "RATNAKALA": {
            "file_prefix": "DEAL_ratnakal",
            "org_prefix": "RATNAKALA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "RATNAKALA",
                "org_merge_col": "Report No",
                "stock_id": "Stone No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 4,
                },
            },
        
        "RATNAKALA": {
            "file_prefix": "DEAL_ratnakal",
            "org_prefix": "RATNAKALA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "RATNAKALA",
                "org_merge_col": "Report No",
                "stock_id": "Stone Id",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 4,
            },
        },  
        "RAJGIR GEMS": {
            "file_prefix": "DEAL_rajgir",
            "org_prefix": "RAJGIR GEMS",
            "params": {
                "sheet_name": "Rajgir Gems STOCK LIST",
                "deal_sheet_prefix": "RAJGIR GEMS",
                "deal_merge_col": "dLot",
                "org_merge_col": "REPORT NO",
                "stock_id": "Stone ID",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 3,
            },
        },
        "R RAJESH": {
            "file_prefix": "DEAL_r.rajesh",
            "org_prefix": "R RAJESH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "R RAJESH",
                "org_merge_col": "Report No",
                "stock_id": "Stone ID",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 5,
            },
        },
        "R.RAJESH": {
            "file_prefix": "DEAL_r rajesh",
            "org_prefix": "R.RAJESH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "R RAJESH",
                "org_merge_col": "Report No",
                "stock_id": "Stone ID",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 5,
            },
        },
        "R N": {
            "file_prefix": "DEAL_rn",
            "org_prefix": "R N",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate",
                "deal_sheet_prefix": "R N",
                "stock_id": "STOCK",
                "dNatts_1": "Black Incl.",
                "dNatts_2": "Table Incl.",
                "skip_rows": 4,
            },
        },
        "RSD": {
            "file_prefix": "DEAL_rsd",
            "org_prefix": "RSD",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "RSD",
                "org_merge_col": "Cert No",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
            },
        },
        "RAYS": {
            "file_prefix": "DEAL_rays",
            "org_prefix": "RAYS",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "deal_sheet_prefix": "RAYS",
                "stock_id": "Packet No",
                "dNatts_1": "BLACK",
                "dNatts_2": "",
                "skip_rows": 1,
            },
        },
        "RED EXIM": {
            "file_prefix": "DEAL_red exim",
            "org_prefix": "RED EXIM",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Report#",
                "deal_sheet_prefix": "RED EXIM",
                "stock_id": "Stock#",
                "dNatts_1": "",
                "dNatts_2": "", 
            },
        },
        "SHAIRU": {
            "file_prefix": "DEAL_shairu",
            "org_prefix": "SHAIRU",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SHAIRU",
                "org_merge_col": "Certi No.",
                "stock_id": "Stock Id",
                "dNatts_1": "Table Black",
                "dNatts_2": "Crown Black",
                "skip_rows": 5,
            },
        },
        "SAGAR": {
            "file_prefix": "DEAL_sagar",
            "org_prefix": "SAGAR",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SAGAR",
                "org_merge_col": "CertNo",
                "stock_id": "StockID",
                "dNatts_1": "",
                "dNatts_2": "",
                "skip_rows": 4,
            },
        },
        "STARLIGHTDIAMOND": {
            "file_prefix": "DEAL_starlight",
            "org_prefix": "STARLIGHTDIAMOND",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "STARLIGHTDIAMOND",
                "org_merge_col": "Cert.No.",
                "stock_id": "StockID",
                "dNatts_1": "Black In Table",
                "dNatts_2": "Side Black",
            },
        },
        "SHREE KRISHNA": {
            "file_prefix": "DEAL_shree krishna",
            "org_prefix": "SHREE KRISHNA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SHREE KRISHNA",
                "org_merge_col": "Report No",
                "stock_id": "Stone No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
            },
        },
        "SRD": {
            "file_prefix": "DEAL_rsd",
            "org_prefix": "SRD",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Cert No",
                "deal_sheet_prefix": "SRD",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 4,
            },
        },
        "SNJ": {
            "file_prefix": "DEAL_snj",
            "org_prefix": "SNJ",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate No",
                "deal_sheet_prefix": "SNJ",
                "stock_id": "Stone Id",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 5,
            },
        },
        "SNJ_ALT": {
            "file_prefix": "DEAL_snj",
            "org_prefix": "SNJ",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate No",
                "deal_sheet_prefix": "SNJ",
                "stock_id": "Stone Id",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 1,
            },
        },
        "SAGAR": {
            "file_prefix": "DEAL_sagar",
            "org_prefix": "SAGAR",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SAGAR",
                "org_merge_col": "CertNo",
                "stock_id": "StockID",
                "dNatts_1": "",
                "dNatts_2": "",
            },
        },
        "STAR RAYS": {
            "file_prefix": "DEAL_starrays",
            "org_prefix": "STAR RAYS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "STAR RAYS",
                "org_merge_col": "Cert.No",
                "stock_id": "",
                "dNatts_1": "BIT",
                "dNatts_2": "BIC",
                "skip_rows": 3,
            },
        },
        "SHEETAL": {
            "file_prefix": "DEAL_sheetal",
            "org_prefix": "SHEETAL",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SHEETAL",
                "org_merge_col": "Certificate",
                "stock_id": "StoneID",
                "dNatts_1": "BIT",
                "dNatts_2": "BIC",
                "skip_rows": 8,
            },
        },
        "SANGHVI": {
            "file_prefix": "DEAL_sanghvi",
            "org_prefix": "SANGHVI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SANGHVI",
                "org_merge_col": "CERTNO",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 8,
            },
        },
        "SANGHVI_ALT": {
            "file_prefix": "DEAL_sanghvi",
            "org_prefix": "SANGHVI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SANGHVI",
                "org_merge_col": "CERTNO",
                "stock_id": "REFNO",
                "dNatts_1": "",
                "dNatts_2": "",
                "skip_rows": 5,
            },
        },
        "SUNRAYS": {
            "file_prefix": "DEAL_sunrays",
            "org_prefix": "SUNRAYS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SUNRAYS",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "B I",
                "dNatts_2": "T I",
                "skiprows":  4
            },
        },
        "UNIQUE": {
            "file_prefix": "DEAL_unique",
            "org_prefix": "UNIQUE",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "UNIQUE",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 4,
            },
        },
        "VD GLOBAL": {
            "file_prefix": "DEAL_vd",
            "org_prefix": "VD GLOBAL",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "VD GLOBAL",
                "org_merge_col": "ReportNo",
                "stock_id": "Stock No",
                "dNatts_1": "BT",
                "dNatts_2": "",
                "skip_rows": 4,
            },
        },
        "VEER GEMS": {
            "file_prefix": "DEAL_veer",
            "org_prefix": "VEER GEMS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "VEER GEMS",
                "org_merge_col": "Cert#",
                "stock_id": "Ref#",
                "dNatts_1": "CTB",
                "dNatts_2": "STB",
                "skip_rows": 4,
            },
        },
        "VAIBHAV": {
            "file_prefix": "DEAL_vaibhav",
            "org_prefix": "VAIBHAV",
            "params": {
                "sheet_name": "Stock",
                "deal_sheet_prefix": "VAIBHAV",
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "CB",
                "dNatts_2": "SB",
                "skip_rows": 1,
            },
        },
        "VINAY": {
            "file_prefix": "DEAL_vinay diamonds",
            "org_prefix": "VINAY",
            "params": {
                "sheet_name": "GIA",
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "VINAY",
                "org_merge_col": "ReportNumber",
                "stock_id": "StoneID",
                "dNatts_1": "Inclusion",
                "dNatts_2": "",
                "skip_rows": 1,
            },
        },
    }


    if not deal_file:
        st.info("Please upload the DEAL workbook.")
        return

    if not org_files:
        st.info("Please upload the ORG files.")
        return

    processed_sheets = []
    display_results = []

    with st.spinner("Processing DEAL workbook..."):

        for supplier_name, config in params_dict.items():

            org_file = match_uploaded_file(
                org_files,
                config["org_prefix"],
            )

            if not org_file:
                uploaded_org_names = ", ".join(
                    f.name for f in org_files
                )

                display_results.append(
                    {
                        "supplier": supplier_name,
                        "warning": (
                            f"{supplier_name}: ORG file not found using "
                            f"prefix '{config['org_prefix']}'. "
                            f"Uploaded ORG files: "
                            f"{uploaded_org_names or '(none)'}"
                        ),
                    }
                )
                continue

            params = config["params"]

            try:
                actual_sheet_name, output_df = build_processed_dataframe(org_file_obj=org_file,deal_file_obj=deal_file,org_skiprows=params.get("skip_rows",params.get("skiprows"),),
                    
                    org_merge_col=params["org_merge_col"],
                    deal_merge_col=params["deal_merge_col"],
                    dNatts_1=params["dNatts_1"],
                    dNatts_2=params["dNatts_2"],
                    stock_id=params["stock_id"],
                    org_sheet_name=params.get("org_sheet_name"),
                    deal_sheet_prefix=params.get(
                        "deal_sheet_prefix"
                    ),
                )

                processed_sheets.append(
                    (
                        actual_sheet_name,
                        output_df,
                    )
                )

                display_results.append(
                    {
                        "supplier": supplier_name,
                        "sheet_name": actual_sheet_name,
                        "output_df": output_df,
                    }
                )

            except Exception as e:  # noqa: BLE001
                error_text = str(e)

                if error_text.startswith("No DEAL sheet starts with"):
                    continue

                # Show other genuine processing errors.
                display_results.append(
                    {
                        "supplier": supplier_name,
                        "error": error_text,
                    }
                )

        # Create ONE output DEAL workbook containing all successful updates.
        output_data = None

        if processed_sheets:
            try:
                output_data = write_all_processed_sheets(deal_file_obj=deal_file,processed_sheets=processed_sheets)
             
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not create final DEAL workbook: {e}")
         
    for item in display_results:

        # Skip warning/error/non-success items silently.
        if "sheet_name" not in item or "output_df" not in item:
            continue

        supplier = item["supplier"]

        st.success( f"{supplier} processed in DEAL sheet "f"'{item['sheet_name']}'.")
           
        st.dataframe(item["output_df"])
            
    if output_data is not None:

        output_filename = os.path.basename(deal_file.name)

        ext = os.path.splitext(output_filename)[1].lower()
    
        mime_type = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if ext == ".xlsm"
            else
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        output_data.seek(0)

        with st.sidebar:
            st.download_button(
                label=f"Download {output_filename}",
                data=output_data.getvalue(),
                file_name=output_filename,
                mime=mime_type,
                key="download_final_deal",
            )

if __name__ == "__main__":
    run_black_sheet_mod()
