import io
import os
import re
import datetime as dt
from typing import Optional, Dict, List, Tuple
from decimal import Decimal, InvalidOperation, ROUND_DOWN
import warnings
# Hide harmless openpyxl notice for Excel files exported by other software
warnings.filterwarnings("ignore", message=".*Workbook contains no default style.*")
import pandas as pd
import streamlit as st

# You can tweak this if you later want partial previews. Not used for preview now.
PREVIEW_ROWS = 100

# ------------------------- Small helpers (self-contained) ----------------------
def _arrow_safe_df(df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    """Make DataFrame safe for Streamlit/Arrow display."""
    if df is None or not isinstance(df, pd.DataFrame):
        return df

    out = df.copy()

    # remove Unnamed columns
    cols = pd.Index([str(c) for c in out.columns])
    out = out.loc[:, ~cols.str.match(r"^Unnamed(:\s*\d+)?$")]
    out.columns = [str(c) for c in out.columns]

    # force every display column to string to avoid Arrow mixed-type errors
    for c in out.columns:
        out[c] = (
            out[c]
            .astype("object")
            .where(out[c].notna(), "")
            .replace({
                "nan": "",
                "NaN": "",
                "NaT": "",
                "None": "",
                "<NA>": "",
            })
            .astype(str)
        )

    return out

def _excel_engine_for_name(lower_name: str):
    if lower_name.endswith(".xls"):
        return "xlrd"
    if lower_name.endswith(".xlsx") or lower_name.endswith(".xlsm") or lower_name.endswith(".xltx") or lower_name.endswith(".xltm"):
        return "openpyxl"
    return None

def _coerce_skiprows(skiprows):
    """Ensure skiprows is an int when user provided as string; pass lists/callables through."""
    if isinstance(skiprows, str):
        s = skiprows.strip()
        return int(s) if s.isdigit() else 0
    return skiprows

def _read_excel_generic(
    content_bytes: bytes,
    *,
    file_name: str,
    skiprows: int = 0,
    sheet_name=0,
    nrows: Optional[int] = None,
) -> pd.DataFrame:
    """Robust Excel reader that picks the correct engine; falls back to pandas default if needed."""
    skiprows = _coerce_skiprows(skiprows)
    lower = file_name.lower() if file_name else ""
    engine = _excel_engine_for_name(lower)
    buf = io.BytesIO(content_bytes)
    try:
        return pd.read_excel(
            buf,
            skiprows=skiprows,
            nrows=nrows,
            sheet_name=sheet_name,
            engine=engine,
            dtype=str,
            keep_default_na=False,
        )
    except Exception:
        try:
            buf.seek(0)
            return pd.read_excel(
                buf,
                skiprows=skiprows,
                nrows=nrows,
                sheet_name=sheet_name,
                dtype=str,
                keep_default_na=False,
            )
        except Exception as e2:
            if lower.endswith(".xls"):
                st.error("Failed to read .xls. Install xlrd: `pip install xlrd`")
            elif lower.endswith(".xlsx") or lower.endswith(".xlsm"):
                st.error("Failed to read .xlsx/.xlsm. Install openpyxl: `pip install openpyxl`")
            raise e2

def _read_csv_full(content_bytes: bytes, skiprows: int = 0) -> pd.DataFrame:
    """Read full CSV with best-effort fallbacks."""
    skiprows = _coerce_skiprows(skiprows)
    buf = io.BytesIO(content_bytes)
    try:
        return pd.read_csv(buf, skiprows=skiprows, dtype=str, keep_default_na=False)
    except Exception:
        buf.seek(0)
        return pd.read_csv(buf, skiprows=skiprows, dtype=str, keep_default_na=False, low_memory=False)

def _read_excel_full(content_bytes: bytes, file_name: str, skiprows: int = 0, sheet_name=0) -> pd.DataFrame:
    """Read full Excel with proper engine selection."""
    skiprows = _coerce_skiprows(skiprows)
    return _read_excel_generic(content_bytes, file_name=file_name, skiprows=skiprows, nrows=None, sheet_name=sheet_name)

def _file_to_bytes(uploaded):
    return uploaded.getvalue() if uploaded else None

def _read_full_best_effort(uploaded, skiprows: int = 0) -> pd.DataFrame:
    """
    Read FULL content of an uploaded file with the correct engine.
    Supports CSV, XLS, XLSX/XLSM/XLT*.
    """
    skiprows = _coerce_skiprows(skiprows)
    data = uploaded.getvalue()
    name = uploaded.name.lower()
    bio = io.BytesIO(data)

    if name.endswith(".csv"):
        # full CSV
        try:
            return pd.read_csv(bio, skiprows=skiprows, dtype=str, keep_default_na=False)
        except Exception:
            bio.seek(0)
            return pd.read_csv(bio, skiprows=skiprows, dtype=str, keep_default_na=False, low_memory=False)

    if name.endswith(".xls"):
        # old Excel: xlrd
        bio.seek(0)
        return pd.read_excel(bio, skiprows=skiprows, dtype=str, keep_default_na=False, engine="xlrd")

    # Modern Excel: openpyxl
    bio.seek(0)
    return pd.read_excel(bio, skiprows=skiprows, dtype=str, keep_default_na=False, engine="openpyxl")

# ---------------------------- Your original functions -------------------------
def parse_skiprows_map(text) -> Dict[str, int]:
    """Parse optional key=value lines like FILEA=2; tolerant of non-strings."""
    if not isinstance(text, str):
        return {}
    mapping: Dict[str, int] = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or "=" not in line:
            continue
        key, val = line.split("=", 1)
        key = key.strip().upper()
        try:
            mapping[key] = int(val.strip())
        except ValueError:
            pass
    return mapping

def select_and_load_excel_files_streamlit(
    uploads: List,
    skiprows_map: Dict[str, int],
    global_skiprows: int = 0,
) -> Dict[str, pd.DataFrame]:
    """
    Streamlit version: takes uploaded files and skiprows_map,
    returns {filename: DataFrame} using best-effort skiprows.
    Filename map overrides the global skip.
    """
    dfs: Dict[str, pd.DataFrame] = {}
    for up in uploads:
        filename = up.name
        data = _file_to_bytes(up)
        ext = os.path.splitext(filename)[1].lower()
        skiprows = next((v for k, v in skiprows_map.items() if k in filename.upper()), global_skiprows)
        skiprows = _coerce_skiprows(skiprows)
        try:
            if ext in [".xlsx", ".xls", ".xlsm", ".xltx", ".xltm"]:
                df = _read_excel_full(data, file_name=filename, skiprows=skiprows)
            elif ext == ".csv":
                df = _read_csv_full(data, skiprows=skiprows)
            else:
                st.warning(f"Unsupported file type for {filename}")
                continue
            dfs[filename] = df
        except Exception as e:
            st.error(f"❌ Failed to read file: {filename} ({e})")
    return dfs

def concatenate_dataframes(dfs: Dict[str, pd.DataFrame], column_names: List[str]) -> pd.Series:
    concatenated = []
    for df in dfs.values():
        for column_name in column_names:
            if column_name in df.columns:
                concatenated.append(df[column_name])
                break
        else:
            concatenated.append(pd.Series([""] * len(df)))
    return pd.concat(concatenated, ignore_index=True) if concatenated else pd.Series(dtype="str")

def add_columns(df: pd.DataFrame, lab_columns: List[str], report_no_columns: List[str]) -> pd.DataFrame:
    report_no_col = next((col for col in report_no_columns if col in df.columns), None)
    lab_col = next((col for col in lab_columns if col in df.columns), None)

    inscription_columns = [
        "Inscription",
        "Inscription(s)",
        "Laser Inscription",
        "Inscription No",
        "Inscription Number",
    ]
    inscription_col = next((col for col in inscription_columns if col in df.columns), None)

    def _cast_report(x):
        try:
            if isinstance(x, float) and x.is_integer():
                return str(int(x))
            return str(x).strip()
        except Exception:
            return str(x).strip()

    def _guess_lab():
        # For files like GIA_Reports where a Lab column is missing
        if report_no_col and "GIA" in report_no_col.upper():
            return "GIA"
        if report_no_col and "IGI" in report_no_col.upper():
            return "IGI"
        if inscription_col and "GIA" in inscription_col.upper():
            return "GIA"
        if inscription_col and "IGI" in inscription_col.upper():
            return "IGI"
        return ""

    # CASE 1: Inscription / Inscription(s) exists like "GIA 5546238669"
    if inscription_col is not None:
        ins = df[inscription_col].astype(str).str.strip()

        # First word is lab, remaining text is report number
        split_data = ins.str.extract(r"^([A-Za-z]+)\s+(.+)$")

        df["lab"] = split_data[0].fillna("").str.strip()
        df["Report No"] = split_data[1].fillna("").str.strip()

        # fallback lab from original lab column
        if lab_col is not None:
            df["lab"] = df["lab"].where(
                df["lab"].astype(str).str.strip() != "",
                df[lab_col].astype(str).str.strip(),
            )

        # fallback lab by report column name, e.g. GIA Report Number
        guessed_lab = _guess_lab()
        if guessed_lab:
            df["lab"] = df["lab"].where(
                df["lab"].astype(str).str.strip() != "",
                guessed_lab,
            )

        # fallback report no from original report column
        if report_no_col is not None:
            df["Report No"] = df["Report No"].where(
                df["Report No"].astype(str).str.strip() != "",
                df[report_no_col].apply(_cast_report),
            )

        df["Inscription"] = (
            df["lab"].astype(str).str.strip() + " " + df["Report No"].astype(str).str.strip()
        ).str.strip()
        return df

    # CASE 2: No inscription column, create Inscription from Lab + Report No
    if report_no_col is None:
        raise KeyError(f"Missing Report No columns: {report_no_columns}")

    df["Report No"] = df[report_no_col].apply(_cast_report)

    if lab_col is not None:
        df["lab"] = df[lab_col].astype(str).str.strip()
    else:
        guessed_lab = _guess_lab()
        if guessed_lab:
            df["lab"] = guessed_lab
        else:
            raise KeyError(f"Missing Lab columns: {lab_columns}")

    df["Inscription"] = (
        df["lab"].astype(str).str.strip() + " " + df["Report No"].astype(str).str.strip()
    ).str.strip()
    return df

def process_measurement(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
    if column_name not in df.columns:
        return df

    def extract_values(measurement) -> Tuple[Optional[float], Optional[float], Optional[float]]:
        numbers = re.findall(r"\d*\.\d+|\d+", str(measurement))
        if len(numbers) == 3:
            return float(numbers[0]), float(numbers[1]), float(numbers[2])
        return None, None, None

    df[["Length", "Width", "Hights Measurement"]] = df.apply(
        lambda row: pd.Series(extract_values(row[column_name])),
        axis=1,
    )
    return df

def process_cut_polish_symmetry(v):
    v = str(v).strip().upper()
    return (
        "GD" if v in {"G", "GOOD"} else
        "FAIR" if v == "FR" else
        "EX" if v in {"EXCELLENT", "3EX", "EXCL", "EXC"} else
        "VG" if v in {"VERY GOOD", "VGOOD", "VERY"} else v
    )

def replace_comma_with_hyphen(v):
    return re.sub(r"\s*,\s*", "-", re.sub(r"\(s\)", "", str(v))).strip() if pd.notna(v) else v

def remove_sign(v):
    return re.sub(r"\s*[+-]\s*", "", str(v)) if pd.notna(v) else v

def fluorescence_intensity(v):
    lookup = {
        "NN": "NONE", "NON": "NONE", "N": "NONE", "NIL": "NONE",
        "F": "FAINT", "FNT": "FAINT", "FA": "FAINT",
        "MED":"MEDIUM","MD": "MED. BLUE",
        "MED. BLUE": "MED. BLUE",
        "SLIGHT": "SLIGHT", "SL": "SLIGHT", "STG": "STRONG",
        "V.STRONG": "V.STRONG", "VERY SLIGHT": "VERY SLIGHT",
    }
    return lookup.get(str(v).strip().upper(), str(v).strip().upper() if pd.notna(v) else "NONE")

def culet(v):
    if pd.isna(v):
        return v

    text = str(v).strip().upper()

    mapping = {
        "N": "NON",
        "NONE": "NON",
        "MIN": "MIN",
        "V": "VSM",
        "POINTED": "POINTED",
    }

    if text in mapping:
        return mapping[text]

    # Example: "0.7 % NATURAL" -> "NATURAL"
    match = re.search(r"%\s*([A-Z ]+)$", text)
    if match:
        return match.group(1).strip()

    return text

def girdle(v):
    if pd.isna(v):
        return v

    value = str(v).strip().upper()

    value = re.sub(r"\s*\([^)]*\)\s*", "", value).strip()

    value = re.sub(r"\s*,\s*(FACETED|POLISHED)\s*$","",value,flags=re.IGNORECASE).strip()
     
    mapping = {
        "MEDIUM TO SLIGHTLY THICK": "MED TO STK",
        "MEDIUM TO SLIGHTLY THICK ": "MED TO STK",
        "THICK TO EXTREMELY THICK": "THK TO ETK",
        "SLIGHTLY THICK TO THICK": "STK TO THK",
        "THIN TO SLIGHTLY THICK": "THN TO STK",
        "THIN TO MEDIUM": "THN TO MED",
        "SLIGHTLY THICK TO VERY THICK" : "STK TO VTK",
        "THICK TO VERY THICK":"THK TO VTK",
        "STK - THK": "MED TO STK",
        "M.S": "MED TO STK",
        "THN - STK": "THN TO STK",
        "SLIGHTLY THICK TO THICK": "STK TO THK",
        "MEDIUM TO THICK": "MED TO THK",
        "MEDIUM-SLIGHTLY THICK": "MED TO STK",
        "MED - THK": "MED TO THK",
        "MED-THK": "MED TO THK",
        "MEDIUM TO VERY THICK": "MED TO VTK",
        "SLIGHTLY THICK": "STK",
        "SLIGHT THICK": "STK",
        "ST": "STK",
        "MEDIUM": "MED",
        "VERY THICK": "VTK",
        "VERY THIK": "VTK",
        "VERY THIN": "VTN",
        "TK": "THK",
        "THICK,": "THK",
        "THICK": "THK",
        "EXTREMLY THICK": "ETK",
        "EXTREMLY THIN": "ETN",
        "THIN TO THICK": "THK",
    }

 # Check longer keys first to avoid partial/wrong matches
    for key, mapped_value in sorted(
        mapping.items(),
        key=lambda x: len(x[0]),
        reverse=True
    ):
        if value.startswith(key):
            return mapped_value

    return value

def shape(v):
    if pd.isna(v): return v
    val = str(v).strip().upper()
    return (
        "ROUND" if val in {"BR", "RBC", "RD", "ROUND BRILLIANT","BRILLIANT","RB"} else
        "PEAR" if val in {"PR", "PB","PEAR BRILLIANT","Pear Brilliant"} else
        "OVAL" if val in {"OL", "OV","OVAL BRILLIANT","Oval Brilliant"} else
        "MARQUIES" if val in{"MQ","MARQUIES BRILLIANT","Marquies Brilliant"}else
        "HEART" if val in{"HEART BRILLIANT","Heart Brilliant"}else
        "EMRALD" if val in {"EM", "EMD","EMRALD BRILLIANT","EMERALD BRILLIANT","EMERALD CUT"} else
        "CUSHION CUT" if val in {"CUS CUT", "CUSH","CUSHION BRILLIANT"} else val
    )

def Crn_Ag(v):
    if pd.isna(v):
        return "0"
    value = re.sub(r"[^0-9.]", "", str(v).strip())
    if value == "":
        return "0"
    return value

def top_open(v): return "N/A" if str(v).strip() == "" else v
def eyeclean(v): return "YES" if str(v).strip() == "" else v
def cost(v): return "1" if str(v).strip() == "" else v
def milky(v): return "NORMAL" if str(v).strip() == "" else v

def only_number(v):
    if pd.isna(v):
        return ""
    value = re.sub(r"[^0-9.]", "", str(v))

    # Keep only the first decimal point
    if value.count(".") > 1:
        first, rest = value.split(".", 1)
        value = first + "." + rest.replace(".", "")
        return f"{v:.2f}"
    return value

def girdel_condition(v):
    v = str(v).strip().lower()

    if 'faceted' in v:
        return 'Faceted'
    elif 'polished' in v:
        return 'Polished'

    return 'Faceted'
    
def reorder_length_width_if_round(row):
    if str(row.get("Shape", "")).upper() == "ROUND":
        try:
            l, w = float(row["Length"]), float(row["Width"])
            return min(l, w), max(l, w)
        except Exception:
            pass
    return row.get("Length"), row.get("Width")

def format_report_date(v):
    if pd.isna(v) or str(v).strip() == "":
        return ""

    dt_value = pd.to_datetime(
        str(v).strip(),
        errors="coerce",
        dayfirst=True
    )

    if pd.isna(dt_value):
        return ""

    # output like 23/6/2025
    return f"{dt_value.day}/{dt_value.month}/{dt_value.year}"

def decimal_to_percent_format(v):
    if pd.isna(v) or str(v).strip() == "":
        return ""

    value = re.sub(r"[^0-9.]", "", str(v))

    if value == "":
        return ""

    try:
        num = float(value)

        # If value is decimal percent, convert to normal percent
        if 0 < num < 1:
            num = num * 100

        return f"{num:.2f}"
    except Exception:
        return ""

def format_weight(v):
    if pd.isna(v) or str(v).strip() == "":
        return ""

    match = re.search(
        r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)",
        str(v).replace(",", "")
    )

    if not match:
        return ""

    try:
        value = Decimal(match.group())
        truncated = value.quantize(Decimal("0.00"), rounding=ROUND_DOWN)
        return f"{truncated:.2f}"
    except (InvalidOperation, ValueError, TypeError):
        return ""
# ------------------------ Static lists / mapping ------------------------------
LAB_COLUMNS = ['lab', 'CERTI','Lab', 'LABNAME', 'Lab', 'Cert', 'LAB', 'Cert Lab', 'Crt', 'Cert.']
MEASUREMENT_COLUMNS = ['Diam.', 'Dia', 'MM', 'L*W/Diam', 'Measurements(L*W*D)', 'MEASUREMENTS',
                       'Measurements', 'Measurement', 'L/W Or Dia', 'Meas MM', 'Meas', 'Measurement']
REPORT_NO_COLUMNS = ['ReportNo','Certificate No','CERT NO','Cert','CertiNo',' REF NO.', 'Cert No', 'Report.No',
                     'IGI Report Number', 'cert', 'Cert NO', 'Report', 'Report No', 'CERTNO',
                     'Report Number', ' REF NO.', 'Certificate Number', 'CERT NUMBER', 'Lab Report No',
                     'Report no ', 'REPORT NUM', 'Certi Number', 'GIA No', 'REPORT', 'REPORT #',
                     'Certificate ', 'CertNo', 'CertificateNo', 'Certificate', 'Certificate No', 'Id',
                     'Report #', 'ReportNo', 'Cert.No', 'REPORT NO.', 'Rep No', 'Cert No.', 'Certi No.',
                     'Report No', 'Cert.No.', 'Certificate #', 'Report No.', 'GIA Report Number']

COLUMNS_MAPPING = {
    'Job No': ['Job No'],
    'Control No': ['Control No'],
    'Diamond Dossier': ['Diamond Dossier'],
    'Report No': REPORT_NO_COLUMNS,
    'Report Dt': ['Valid From', 'Date', 'Report Dt', 'Certificate Date', 'Report Date','Date of Issue'],
    'Client Ref': REPORT_NO_COLUMNS,
    'Memo No': ['Memo No'],
    'Shape': ['SHP','shape', 'ShapeName', 'Shp.', 'SHAPE', 'Shape', 'Shp', 'Shape', 'shp', 'Shape and Cutting Style','Shape & Cutting Style','Shape and Cut'],
    'Length': ['Length','Len','Diam / LW Min', 'Max', 'Length', '', 'L', 'length', ''],
    'Width': ['Width','Wid','Diam / LW Max', 'Min', 'Width', 'W', 'witdh'],
    'Depth': ['Total Depth MM','Hei','Height', 'Height CAvg', 'Hgt','Depth mm', 'H',  'Hights Measurement', 'D'],
    'Weight': ['CTS','WT','Ct.','Weight R', 'weight', 'SIZE', 'Carat Weight', 'CARAT', 'carat', 'WEIGHT', 'Cts.', 'Weight',
               'cts', 'Cts', 'CRTWT', 'Size', 'CTS', 'Cts', 'Carats', 'Weight', 'Cts', 'Carat', 'Wgt'],
    'Color': ['COL','Colour Grade', 'color', 'Color Grade', 'ColorName', 'col', 'Colour', 'COLOR', 'Color', 'DispColor',
              'Color', 'Col', 'Color', 'COL'],
    'Color Descriptions': ['Color Descriptions'],
    'Clarity': ['CLA','Clrt','Quality', 'purity', 'CTY', 'CLAR', 'Clarity Grade', 'ClarityName', 'Clr', 'Clarity',
                'CLARITY', 'Purity', 'CLR', 'DispClarity', 'Clr', 'Cl.', 'Clarity', 'Cla', 'Clar'],
    'Clarity Status': ['Clarity Status'],
    'Final Cut': ['C','Sh.Ap. / Cut','Proportions', 'Cut Grade', 'Cut', 'Cut Grade', 'Final Cut', 'CUT'],
    'Polish': ['P','POL','polish', 'Polish', 'POLISH', 'PolishName', 'Pol', 'POL.', 'POL', 'Pl', 'Po', 'Pol.'],
    'Symmetry': ['S','SYM','Symm.','sym', 'Symmetry', 'SymmetryName', 'Sym', 'SYM.', 'SYM', 'Symm', 'Sym.', 'Sy'],
    'Fluorescence Intensity': ['FLUO','FLS','LW-fluo', 'fl', 'FLR', 'Fluorescence', 'Flu', 'FLIntName', 'Flor', 'FLO.',
                               'Fluorescence ', 'FlrIntens', 'Fluorescence Intensity', 'Fluor', 'Fluor.', 'Flur',
                               'Flour', 'Fluor', 'Flo', 'Fl', 'Flor', 'FLUOR', 'Flr', 'Fluo Int', 'Fluo', 'Fluorescent',
                               'Fluorescence', 'Fls', 'FL'],
    'Fluorescence Color': ['Fluorescence Color'],
    'Girdle': ['Girdle ID','Grdl','GirdleThick','Gridle','Girdle_Concatenate', 'Girdle Desc', 'Girdle description',
               'Girdle Thick', 'GIRDLE', 'Girdle Name', 'Girdle', 'Girdle Thickness', 'Girdle Name', 'Girdle Desc',
               'Gr', 'Girdle Thin', 'Gridle'],
    'Girdle Condition': ['Girdle Thickness','Grdl.Cond.','Girdle nature','Girdle Condition', 'GIRDLE CONDITION', 'girdlecondition',
                         'GirdleCondition', 'Girdle Cond',],
    'Culet Size': ['Culet ID','culetsize','Culet nature','Culet', 'Culet Size','Culet'],
    'Depth %': ['Total Depth','Depth','TD','T.Dpt','Total Depth Per.','Depth(%)','Dep%','Depth%','Tot.Depth %','Tot Depth','T.Dep.', 
                'Dep', 'Total depth%', 'TD%', 'TD', 'Depth %', 'T.D.%', 'Depth%'],          
    'Table %': ['Table Size','TABLE','TA','Tbl','Table Size Per.','Table width %', 'TAB%', 'TA%', 'Tableprct', 'TBL', 'Tab',
                'Table', 'Tab%', 'Tab%', 'Table Depth', 'Tbl', 'Table%', 'TA', 'Table(%)', 'Table %'],
    'Crn Ag': ['Crn Ang','Cr Ang','Cr.Ang','Crown Angel','Cr Ang', 'Cr Angle', 'CA', 'C Angle', 'CRAng',
               'Crown angle°', 'Crn Ag', 'Crown Angle', 'Cr >', 'C.A.', 'C.Angle', 'C/A', 'Cr.Ang', 'Crown Angle', 'CA'],
    'Crn Ht': ['C.H.','Crn Hgt','CrownHeight','Crown Height Per.','Cr Ht','Cr.Heig%','Cr Hgt', 'CH', 'CHeight', 'CRHgt',
               'Cr. Height%', 'Crn Dp', 'Crn Ht', 'Crown Height', 'Cr Hgt%', 'Crown Height', 'C.H', 'C.Height', 'C/H',
               'Cr.Heig %', 'Crown Height', 'CH'],
    'Pav Ag': ['Pavi. Angle','Pav Ang','Pav.Ang','Pav Angel','PA', 'P Angle', 'PVAng', 'Pavillion angle°', 'Pav Ag',
               'Pav Angle', 'Pav >', ' Pavilion Angle ', 'P.A.', 'Pav. Angle', 'P/A', 'Pav.Ang', 'Pavilion Angle', 'PA'],
    'Pav Dp': ['P.H.','Pav Dpt','PavillionHeight','Pavi. Depth Per.','Pav Ht','Pav.Dep','Pav Hgt', 'PH', 'PHeight',
               'PVHgt', 'Pav. depth%', 'Pavilion Height', 'Pav Dp', 'Pav Hgt%', 'Pavilion Depth', 'P.D.', 'Pav Height',
               'Pav. Height', 'P/H', 'Pav.Per', 'Pavilion Depth', 'PH'],
    'Str Ln': [],
    'Lr Half': [],
    'Painting': ['Painting'],
    'Proportion': ['Proportion'],
    'Paint Comm': ['Paint Comm'],
    'Key to Symbols': ['KET TO SYMBOLS','KeyToSym','Clarity Characteristics', 'Key to Symbols', 'KeyToSymbol',
                       'keytosymbols', 'KeyToSymbols', 'KeyToSymbole', 'Key to Sym', 'KTSVIEW', 'Key to symbols',
                       'Comment', 'Key to Symbols', 'Key To Symbol', 'KeytoSymbols', 'Key Symbol', 'KEY TO SYMBOL',
                       'K T S', 'Keytosymbols', 'Key To Symbols', 'Key To Sym', 'Key to symbols','Clarity Characteristics','Identification Features'],
    'Report Comments': ['Report Comment', 'Report Comments', 'Comments', 'Cert comment','Comments'],
    'Inscription': ['Inscription', 'Inscription(s)'],
    'Synthetic Indicator': ['Synthetic Indicator'],
    'Girdle %': ['Grdl%','Gridle Thickness Per.','Girdle %','Girdle%','GIRDLE%', 'Girdle Per', 'Girdle Size',
                 'Girdle%', 'Girdle', 'Girdle%', 'Girdle size %', 'Girdle size %','Girdle Thickness'],
    'Polish Features': ['Polish Features'],
    'Symmetry Features': ['Symmetry Features'],
    'Shape Description': [''],
    'Report Type': [''],
    'Sorting': [''],
    'Basket Status': [''],
    'Black1':['Natts','Black Inc','TB','BIT','BT','CN','BlkC','T.B'],
    'Black2':['SB','BIS','BIC','BC','SN','sideNatts','BlkS','S.B'],
    'natts': ['natts'],
    'milky': ['milky'],
    'top opens': [],
    'bot opens': [],
    'tinch': ['tinch'],
    'eyeclean': ['eyeclean'],
    'ha': ['ha'],
    'cost': ['cost'],
    'STOCK ID': ['STOCK ID'],
    'origin': ['origin'],
    'lab': LAB_COLUMNS,
}

# ------------------------ NATTS from Black1 / Black2 ---------------------------
# (Black1, Black2, natts) - copied from Book1.xlsx.
# "A/B" means A or B ("CN2/3" -> CN2 or CN3). "-" is treated as blank.
# Combined values in one cell ("BC0 - BT1", "TB01,SB1") match in any order.
# Exact rows win over rows expanded from "A/B".
NATTS_RULES = [
    ('B0', '', 'NO BLK'),
    ('B1', '', 'CENTER-L/SIDE - L'),
    ('B2', '', 'HEAVY BLACK'),
    ('BPP', '', 'PINPOINT BLACK'),
    ('NO', '', 'NO BLK'),
    ('YES', '', 'SIDE OR CENTER LIGHT'),
    ('YT', '', 'SIDE OR CENTER LIGHT'),
    ('YT1', '', 'SIDE OR CENTER LIGHT'),
    ('YT2', '', 'HEAVY BLACK'),
    ('NONE', 'NONE', 'NO BLK'),
    ('NONE', 'PINPOINT', 'SIDE-L'),
    ('PINPOINT', 'NONE', 'CENTER-L'),
    ('PINPOINT', 'PINPOINT', 'S & C-L'),
    ('MINOR/MAJOR', 'NONE', 'CENTER - H'),
    ('', 'MINOR/MAJOR', 'SIDE - H'),
    ('MINOR/MAJOR', 'MINOR/MAJOR', 'S & C-H'),
    ('', '', 'NO BLK'),
    ('B1/B2', '-', 'CENTER -L'),
    ('-', 'B1/B2', 'SIDE -L'),
    ('B3/B4', '', 'CENTER -H'),
    ('', 'B3/B4', 'SIDE - H'),
    ('B3/B4', 'B3/B4', 'S & C-H'),
    ('NN', 'NN', 'NO BLK'),
    ('B0/B1', '', 'CENTER -L'),
    ('', 'B0/B1', 'SIDE -L'),
    ('B2/B3/B4', '', 'CENTER -H'),
    ('', 'B2/B3/B4', 'SIDE - H'),
    ('B2/B3/B4', 'B2/B3/B4', 'S & C-H'),
    ('', '', 'NO BLK'),
    ('CN0', '', 'CENTER-L'),
    ('', 'SN0', 'SIDE-L'),
    ('CN1', 'SN1', 'CSL'),
    ('CN2/3', '', 'CENTER - H'),
    ('', 'SN2/3', 'SIDE - H'),
    ('CN2/3', 'SN2/3', 'S & C-H'),
    ('NO', 'NO', 'NO BLK'),
    ('MINOR', '', 'CENTER-L'),
    ('', 'MINOR', 'SIDE-L'),
    ('MEDIUM', '', 'CENTER - H'),
    ('', 'MEDIUM', 'SIDE - H'),
    ('MINOR', 'MINOR', 'S & C-L'),
    ('MEDIUM', 'MEDIUM', 'S & C-H'),
    ('-', '-', 'NO BLK'),
    ('NB', 'NB', 'NO BLK'),
    ('TB0/TB1', '', 'CENTER-L'),
    ('NB', 'TB0/TB1', 'SIDE-L'),
    ('TB0/TB1', 'TB0/TB1', 'S & C-L'),
    ('TB2/TB3', 'NB', 'CENTER -H'),
    ('NB', 'TB2/TB3', 'SIDE - H'),
    ('TB2/TB3', 'TB2/TB3', 'S & C-H'),
    ('BT0', 'BC0', 'NO BLK'),
    ('BT1', 'BC0', 'CENTER -L'),
    ('BT0', 'BC1', 'SIDE -L'),
    ('BT1', 'BC1', 'S & C-L'),
    ('BT2', 'BC0', 'CENTER -H'),
    ('BT0/BT1', 'BC2', 'SIDE - H'),
    ('BT2', 'BC2', 'S & C-H'),
    ('BT3', 'BC3', 'HEAVY BLACK'),
    ('BC0 - BT0', '', 'NO BLK'),
    ('BC0 - BT1', '', 'CENTER-L'),
    ('BC0 - BT2', '', 'CENTER-H'),
    ('BC0 - BT3', '', 'CENTER-H'),
    ('BC1 - BT0', '', 'SIDE-L'),
    ('BC1 - BT1', '', 'S & C-L'),
    ('BC1 - BT2', '', 'S & C-H'),
    ('BC1 - BT3', '', 'S & C-H'),
    ('BC2 - BT1', '', 'S & C-H'),
    ('BC2 - BT0', '', 'SIDE - H'),
    ('BC2 - BT2', '', 'S & C-H'),
    ('BC2 - BT3', '', 'S & C-H'),
    ('BC3 - BT0', '', 'HEAVY'),
    ('BC3 - BT1', '', 'HEAVY'),
    ('BC3 - BT2', '', 'HEAVY'),
    ('BC3 - BT3', '', 'HEAVY'),
    ('BT0', 'BC0', 'NO BLK'),
    ('BT1', 'BC0', 'CENTER-L'),
    ('BT0', 'BC1', 'SIDE-L'),
    ('BT1', 'BC1', 'S & C-L'),
    ('BT2/3', 'BC0', 'CENTER -H'),
    ('BT0', 'BC2/3', 'SIDE - H'),
    ('BT2/3', 'BC2/3', 'S & C-H'),
    ('BT0', 'BC0', 'NO BLK'),
    ('BT1', 'BC0', 'CENTER -L'),
    ('BT0', 'BC1', 'SIDE -L'),
    ('BT1', 'BC1', 'S & C-L'),
    ('BT2', 'BC0', 'CENTER -H'),
    ('BT0/BT1', 'BC2', 'SIDE - H'),
    ('BT2', 'BC2', 'S & C-H'),
    ('BT3', 'BC3', 'HEAVY BLACK'),
    ('NO', '', 'NO BLK'),
    ('TB01/TB1', '', 'CENTER - L'),
    ('SB01/SB1', '', 'SIDE - L'),
    ('TB01,SB01', '', 'S & C-L'),
    ('SB1,TB1', '', 'S & C-L'),
    ('TB01,SB1', '', 'S & C-L'),
    ('SB1,TB2', '', 'HEAVY'),
    ('SB01,TB1', '', 'S & C-L'),
    ('SB01,TB2', '', 'HEAVY'),
    ('TB2', '', 'CENTER -H'),
    ('TB1,SB2', '', 'SIDE-H'),
    ('TB01,SB2', '', 'SIDE-H'),
    ('TB1,SB3', '', 'SIDE-H'),
    ('SB2,TB2', '', 'S & C-H'),
    ('SB2', '', 'SIDE-H'),
    ('SB01,TB3', '', 'S & C-H'),
    ('TB2,SB3', '', 'S & C-H'),
    ('TB01,SB3', '', 'S & C-H'),
    ('NONE', 'N0', 'NO BLK'),
    ('LBCS/LBC', 'N2', 'CENTER - H'),
    ('VLBTS', 'N1', 'S & C-L'),
    ('BIS', 'N3', 'HEAVY(CENTER)'),
    ('N', 'N', 'NO BLK'),
    ('NV', 'BC0', 'NO BLK'),
    ('NV', 'NV', 'NO BLK'),
]

def _norm_black(v) -> str:
    s = re.sub(r"\s+", " ", str(v if v is not None else "")).strip().upper()
    if s in {"-", "NAN", "<NA>"}:
        return ""
    # combined cell like "BC0 - BT1" or "TB01,SB1" -> order-independent key
    parts = [p.strip() for p in re.split(r"\s*,\s*|\s+-\s+", s) if p.strip()]
    if len(parts) > 1:
        return ",".join(sorted(parts))
    return s

def _expand_black(v: str) -> List[str]:
    """'B1/B2' -> ['B1','B2'];  'CN2/3' -> ['CN2','CN3'];  'BT2/3' -> ['BT2','BT3']."""
    s = str(v).strip()
    if "/" not in s:
        return [s]
    parts = [p.strip() for p in s.split("/") if p.strip()]
    out = []
    prefix = re.match(r"^([A-Za-z]*)", parts[0]).group(1) if parts else ""
    for p in parts:
        # short form like "3" in "CN2/3" -> reuse prefix of first part
        out.append(prefix + p if prefix and p.isdigit() else p)
    return out

def _build_natts_lookup() -> Dict[Tuple[str, str], str]:
    lookup: Dict[Tuple[str, str], str] = {}
    # 1) exact values first (first row wins)
    for b1, b2, n in NATTS_RULES:
        lookup.setdefault((_norm_black(b1), _norm_black(b2)), n)
    # 2) then expanded "A/B" values, only where not already defined
    for b1, b2, n in NATTS_RULES:
        for x in _expand_black(b1):
            for y in _expand_black(b2):
                lookup.setdefault((_norm_black(x), _norm_black(y)), n)
    return lookup

NATTS_LOOKUP = _build_natts_lookup()

def natts_from_black(b1, b2, current="") -> str:
    """natts from Black1/Black2; if no rule matches keep current natts, else default."""
    key = (_norm_black(b1), _norm_black(b2))
    if key in NATTS_LOOKUP:
        # both blank: keep natts given in the file, if any
        if key == ("", "") and str(current).strip():
            return str(current).strip()
        return NATTS_LOOKUP[key]
    cur = str(current).strip()
    return cur if cur else "NO BLK"

# ------------------------ Auto header-row detection ---------------------------
HEADER_SCAN_ROWS = 60  # look for the header in the first N rows

def _norm_hdr(v) -> str:
    return re.sub(r"\s+", " ", str(v)).strip().lower()

def _known_header_names() -> set:
    names = set()
    for keys in COLUMNS_MAPPING.values():
        names.update(_norm_hdr(k) for k in keys if str(k).strip())
    for lst in (LAB_COLUMNS, MEASUREMENT_COLUMNS, REPORT_NO_COLUMNS):
        names.update(_norm_hdr(k) for k in lst if str(k).strip())
    return names

KNOWN_HEADERS = _known_header_names()
REPORT_HEADERS = {_norm_hdr(k) for k in REPORT_NO_COLUMNS if str(k).strip()}

def _read_raw(content_bytes: bytes, file_name: str) -> pd.DataFrame:
    """Read the sheet with NO header so every row (incl. title rows) is visible."""
    lower = file_name.lower()
    bio = io.BytesIO(content_bytes)
    if lower.endswith(".csv"):
        try:
            return pd.read_csv(bio, header=None, dtype=str, keep_default_na=False)
        except Exception:
            bio.seek(0)
            return pd.read_csv(bio, header=None, dtype=str, keep_default_na=False,
                               low_memory=False, on_bad_lines="skip")
    return pd.read_excel(bio, header=None, dtype=str, keep_default_na=False,
                         engine=_excel_engine_for_name(lower))

def detect_header_row(raw: pd.DataFrame) -> int:
    """Row index whose cells match the most known column names (report no column counts extra)."""
    best_row, best_score = 0, 0
    for i in range(min(HEADER_SCAN_ROWS, len(raw))):
        cells = {_norm_hdr(c) for c in raw.iloc[i].tolist() if str(c).strip()}
        score = len(cells & KNOWN_HEADERS) + (3 if cells & REPORT_HEADERS else 0)
        if score > best_score:
            best_row, best_score = i, score
    return best_row if best_score >= 3 else 0

def frame_from_header_row(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    """Use raw row `header_row` as column names and everything below as data."""
    if raw.empty or header_row >= len(raw):
        return pd.DataFrame()
    cols, seen = [], {}
    for j, c in enumerate(raw.iloc[header_row].tolist()):
        name = str(c).strip() or f"Unnamed: {j}"
        if name in seen:                      # duplicate header -> Name.1, Name.2
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        cols.append(name)
    df = raw.iloc[header_row + 1:].copy()
    df.columns = cols
    # drop fully blank rows
    df = df[df.apply(lambda r: any(str(x).strip() for x in r), axis=1)]
    return df.reset_index(drop=True)

# ------------------------------ UI entrypoint ---------------------------------
def run_mpb_natts_mod():
    st.markdown("---")
    # Sidebar controls
    with st.sidebar:
        st.header("📤 Feading")
        uploads = st.file_uploader(
            "Select multiple CSV/XLSX/XLS files",
            type=["csv", "xlsx", "xls", "xlsm", "xltx", "xltm"],
            accept_multiple_files=True,
            key="fead_files",
        )
        # Download button sits right below the uploader (filled in after processing)
        download_slot = st.empty()
    if not uploads:
        st.info("Upload Files ")
        return

    # Read every file raw once and auto-detect its header row
    raws: Dict[str, pd.DataFrame] = {}
    detected: Dict[str, int] = {}
    for up in uploads:
        try:
            raws[up.name] = _read_raw(up.getvalue(), up.name)
            detected[up.name] = detect_header_row(raws[up.name])
        except Exception as e:
            st.error(f"❌ Failed to read file: {up.name} ({e})")

    # Per-file skip rows shown next to each file's preview (pre-filled, editable)
    skips: Dict[str, int] = {}
    dfs_dict: Dict[str, pd.DataFrame] = {}
    with st.expander("Files", expanded=True):
        for name, raw in raws.items():
            c1, c2 = st.columns([4, 1])
            with c2:
                skips[name] = int(st.number_input(
                    "Skip rows", 0, 100000, detected[name], 1,
                    key=f"skip_{name}_{detected[name]}", format="%d"))
               
            with c1:
                st.markdown(f"### File: {name}")
            dfs_dict[name] = frame_from_header_row(raw, skips[name])
            st.dataframe(_arrow_safe_df(dfs_dict[name]), height=250)

    try:
        if not dfs_dict:
            st.error("No readable files.")
            return

        # Enrich each df (add_columns + measurement)
        skipped: List[str] = []
        for file_name, df in list(dfs_dict.items()):
            try:
                dfs_dict[file_name] = add_columns(df, LAB_COLUMNS, REPORT_NO_COLUMNS)
            except KeyError as e:
                st.warning(f"Skipping {file_name}: {e}")
                skipped.append(file_name)
                continue

            for col in MEASUREMENT_COLUMNS:
                if col in df.columns:
                    dfs_dict[file_name] = process_measurement(dfs_dict[file_name], col)
                    break

            if "GIRDLE THIN" in df.columns and "GIRDLE THICK" in df.columns:
                df["Girdle_Concatenate"] = df["GIRDLE THIN"].fillna("") + " - " + df["GIRDLE THICK"].fillna("")

        for file_name in skipped:
            dfs_dict.pop(file_name, None)
        if not dfs_dict:
            st.error("No file has a Report No column - check skip rows.")
            return

        # Build final consolidated table
        data = pd.DataFrame()
        data["PARTY NAME"] = pd.concat(
            [pd.Series([name.split(".")[0]] * len(df)) for name, df in dfs_dict.items()],
            ignore_index=True,
        )

        for column, keys in COLUMNS_MAPPING.items():
            data[column] = concatenate_dataframes(dfs_dict, keys)
            if column in ["Final Cut", "Polish", "Symmetry"]:
                data[column] = data[column].apply(process_cut_polish_symmetry)
            if column == "Final Cut":
                data[column] = (data[column].replace(r"^[\s\-–—]*$", pd.NA, regex=True).fillna("VG"))
            elif column in ["Key to Symbols", "Report Comments"]:
                data[column] = data[column].apply(replace_comma_with_hyphen)
            elif column == "Fluorescence Intensity":
                data[column] = data[column].apply(fluorescence_intensity)
            elif column == "Girdle":
                data[column] = data[column].apply(girdle)
            elif column == "Shape":
                data[column] = data[column].apply(shape)
            elif column in ["Color", "Clarity"]:
                data[column] = data[column].apply(lambda x: remove_sign(str(x)).upper() if pd.notna(x) else x)
            elif column in ["top opens", "bot opens", "ha", "origin","tinch"]:
                data[column] = data[column].apply(top_open)
            elif column == "eyeclean":
                data[column] = data[column].apply(eyeclean)
            elif column == "natts":
                # natts decided by Black1 + Black2 (see NATTS_RULES)
                data[column] = [
                    natts_from_black(b1, b2, cur)
                    for b1, b2, cur in zip(data["Black1"], data["Black2"], data[column])
                ]
            elif column == "cost":
                data[column] = data[column].apply(cost)
            elif column == "milky":
                data[column] = data[column].apply(milky)
            elif column in ["Depth %", "Table %", "Crn Ht", "Pav Dp", "Girdle %"]:
                data[column] = data[column].apply(decimal_to_percent_format)
            elif column == 'Girdle Condition':
                data[column] = data[column].apply(girdel_condition)
            elif column == "Culet Size":
                data[column] = data[column].apply(culet)
            elif column == "Weight":
                # Extract values such as "0.50 carat" and preserve two decimals in CSV.
                data[column] = data[column].apply(format_weight)
            elif column in ["Length", "Width","Girdle %","Report No"]:
                data[column] = data[column].apply(only_number)
            elif column in ["Girdle %","Crn Ag","Crn Ht","Pav Ag","Pav Dp","Str Ln","Lr Half"]:
                data[column] = data[column].apply(Crn_Ag)
            elif column == "Report Dt":
                data[column] = data[column].apply(format_report_date)

        # If shape + L/W exist, reorder for ROUND
        if all(x in data.columns for x in ["Shape", "Length", "Width"]):
            data[["Length", "Width"]] = data.apply(reorder_length_width_if_round, axis=1, result_type="expand")

        # Drop empty/null report numbers
        data = data[~data["Report No"].astype(str).str.lower().isin(["", "nan", "none", "null"])].reset_index(drop=True)
        data.drop(['Black1', 'Black2'], axis=1, inplace=True)
        # Output & full inline preview
        out_csv = io.StringIO()
        data.to_csv(out_csv, index=False)
        out_csv.seek(0)

        st.success("Done")
        st.caption("Table View")
        st.dataframe(_arrow_safe_df(data), height=250)  # full table inline

        fname = f"Feeding {dt.datetime.now().strftime('%d-%m-%Y')}.csv"
        download_slot.download_button(
            "⬇️ Download Feeding CSV",
            data=out_csv.getvalue(),
            file_name=fname,
            mime="text/csv",
        )
    except Exception as e:
        st.error(f"Feeding failed: {e}")

if __name__ == "__main__":
    run_mpb_natts_mod()