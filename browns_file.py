import re
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os

st.set_page_config(page_title="Brown File",page_icon="📄",layout="wide",)

def reset_file_pointer(file_path):
    """Reset Streamlit UploadedFile pointer before every pandas read."""
    if hasattr(file_path, "seek"):
        try:
            file_path.seek(0)
        except Exception:
            pass

def load_data(file_path, skip_rows=0):
    """Load Excel or CSV files uploaded in Streamlit."""
    reset_file_pointer(file_path)
    file_name = str(getattr(file_path, "name", file_path)).lower()

    if file_name.endswith(".csv"):
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                reset_file_pointer(file_path)
                return pd.read_csv(file_path, skiprows=skip_rows, encoding=encoding, keep_default_na=False)
            except UnicodeDecodeError as exc:
                last_error = exc
                continue
        raise last_error

    return pd.read_excel(file_path, skiprows=skip_rows, keep_default_na=False)

def apply_filters(data, filters):
    filtered = data.copy()

    for column, criteria in filters.items():
        if column not in filtered.columns:
            # If the column doesn't exist in the uploaded file, skip this filter.
            continue

        series = filtered[column]

        if isinstance(criteria, list):
            # Keep original matching, plus stripped text matching for columns where suppliers add spaces.
            original_mask = series.isin(criteria)
            text_mask = series.astype(str).str.strip().isin([str(v).strip() for v in criteria])
            mask = original_mask | text_mask
        elif callable(criteria):
            mask = criteria(series)
        else:
            continue

        if not isinstance(mask, pd.Series):
            mask = pd.Series(mask, index=filtered.index)

        mask = mask.reindex(filtered.index).fillna(False).astype(bool)
        filtered = filtered.loc[mask].copy()

    return filtered

def concatenate_dataframes(df, column_names):
    concatenated = []
    for sub_df in df.values():
        for column_name in column_names:
            if sub_df is not None and column_name in sub_df.columns:
                concatenated.append(sub_df[column_name])
                break
    concatenated = [series for series in concatenated if not series.empty]
    return pd.concat(concatenated, ignore_index=True) if concatenated else pd.Series(dtype='str')

def process_measurement(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""

    value = str(value).strip()

    # Standardize multiplication symbols.
    value = re.sub(r"[\*xX×]", " X ", value)

    # Find numeric values.
    numbers = re.findall(r"\d+(?:\.\d+)?", value)

    # Always return: number - number X number
    if len(numbers) >= 3:
        return f"{numbers[0]} - {numbers[1]} X {numbers[2]}"

    # Return cleaned value if 3 numbers are not found.
    value = re.sub(r"\s*-\s*", " - ", value)
    value = re.sub(r"\s+", " ", value).strip()

    return value


def process_cut_polish_symmetry(value):
    if pd.isna(value):
        return value

    value = str(value).strip().upper()

    mappings = {
        "NONE": "",
        "-": "",
        "G": "GD",
        "GOOD": "GD",
        "FR": "FAIR",
        "EXCELLENT": "EX",
        "3EX": "EX",
        "EXCL": "EX",
        "EXC": "EX",
        "V. GOOD": "VG",
        "VERY GOOD": "VG",
        "VGOOD": "VG",
        "VERY": "VG",
    }

    return mappings.get(value, value)


def process_shape(value):
    if pd.isna(value):
        return value

    value = str(value).strip().upper()

    mappings = {
        "BR": "ROUND",
        "RB": "ROUND",
        "RBC": "ROUND",
        "RD": "ROUND",
        "ROUND BRILLIANT": "ROUND",
        "BRILLIANT": "ROUND",
        "PR": "PEAR",
        "PEAR BRILLIANT": "PEAR",
        "PEAR": "PEAR",
        "PE": "PEAR",
        "PEAR MODIFIED BRILLIANT": "PEAR",
        "OL": "OVAL",
        "OV": "OVAL",
        "OVAL": "OVAL",
        "OVAL BRILLIANT": "OVAL",
        "OVAL MODIFIED BRILLIANT": "OVAL",
        "MARQUIES": "MARQUISE",
        "MARQUISE": "MARQUISE",
        "MQ": "MARQUISE",
        "MARQUISE BRILLIANT": "MARQUISE",
        "EM": "EMERALD",
        "EMD": "EMERALD",
        "EMERALD CUT": "EMERALD",
        "CUS CUT": "CUSHION CUT",
        "CUSH": "CUSHION CUT",
        "CUSHION MODIFIED BRILLIANT": "CUSHION CUT",
        "HEART BRILLIANT": "HEART",
        "HEART MODIFIED BRILLIANT": "HEART",
    }

    return mappings.get(value, value)


def process_fluor(value):
    if pd.isna(value):
        return ""

    value = str(value).strip().upper()

    mappings = {
        "FAINT": "FNT",
        "F": "FNT",
        "N": "NONE",
        "NAN": "NONE",
        "NON": "NONE",
        "MED. BLUE": "MED",
        "MEDIUMBLUE": "MED",
        "M": "MED",
        "MB": "MED",
        "MEDIUMYELLOW": "MED",
        "MEDIUMGREEN": "MED",
        "STRONG": "FNT",
        "ST": "FNT",
        "V.STRONG": "VST",
        "VS": "VST",
        "V.STR": "VST",
    }

    return mappings.get(value, value)

def process_culet(value):
    if pd.isna(value) or str(value).strip().lower() == "nan":
        return ""

    value = str(value).strip()
    value_upper = value.upper()

    mappings = {
        "NONE": "None",
        "N": "None",
        "NON": "None",
        "POINTED": "Pointed",
    }

    return mappings.get(value_upper, value)


def replace_comma_with_hyphen(value):
    if pd.isna(value):
        return value

    return re.sub(r"\s*,\s*", "-", str(value)).strip()


def only_number(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""

    text = str(value).strip()

    # Remove everything except digits, decimal point and minus sign.
    cleaned = re.sub(r"[^0-9.\-]", "", text)

    if not cleaned:
        return ""

    try:
        number = float(cleaned)

        if abs(number) <= 1 and "%" not in text:
            number *= 100

        return f"{number:.2f}%"

    except ValueError:
        return ""


def process_discount(value):
    if pd.isna(value) or str(value).strip() == "":
        return None

    text = str(value).strip()

    # Extract first number, positive or negative.
    match = re.search(r"-?\d+(?:\.\d+)?", text)

    if not match:
        return None

    try:
        number = float(match.group())

        # Example: 0.35 / -0.35 -> 35
        if -1 <= number <= 1 and number != 0:
            number *= 100

        # Always make discount negative.
        number = -abs(number)

        return round(number, 2)

    except ValueError:
        return None

def process_girdle(v):
    if pd.isna(v) or str(v).strip() == "":
        return ""

    import html

    # Convert HTML entities such as &nbsp;
    value = html.unescape(str(v))

    # Convert <br>, <br/>, <br /> etc. into spaces
    value = re.sub(r"<br\s*/?>", " ", value, flags=re.IGNORECASE)

    # Remove any remaining HTML tags
    value = re.sub(r"<[^>]+>", " ", value)

    # Normalize non-breaking spaces
    value = value.replace("\xa0", " ")

    # Uppercase and clean
    value = value.upper().strip()

    # Remove finish descriptions
    value = re.sub(
        r"\b(FACETED|POLISHED|BRUTED)\b",
        " ",
        value
    )

    # Remove brackets
    value = re.sub(r"[(),]", " ", value)

    # Remove numeric values
    value = re.sub(r"\b\d+(?:\.\d+)?\b", " ", value)

    # Standardize separators to TO
    value = re.sub(r"\s*-\s*", " TO ", value)
    value = re.sub(r"\s+TO\s+", " TO ", value)

    # Normalize spaces
    value = re.sub(r"\s+", " ", value).strip()

    # Individual girdle mappings
    mapping = {
        "VERY THIN": "VTN",
        "VTN": "VTN",

        "THIN": "THN",
        "TN": "THN",
        "THN": "THN",

        "SLIGHTLY THIN": "STN",
        "SLIGHT THIN": "STN",
        "STN": "STN",

        "MEDIUM": "MED",
        "MED": "MED",
        "M": "MED",

        "SLIGHTLY THICK": "STK",
        "SLIGHT THICK": "STK",
        "ST": "STK",
        "STK": "STK",

        "THICK": "THK",
        "TK": "THK",
        "THK": "THK",

        "VERY THICK": "VTHK",
        "VERY THIK": "VTHK",
        "VTK": "VTHK",
        "VTHK": "VTHK",

        "EXTREMELY THICK": "ETK",
        "EXTREMLY THICK": "ETK",
        "XTK": "ETK",
        "ETK": "ETK",
    }

    # First check single exact value
    if value in mapping:
        return mapping[value]

    # ---------------------------------------------------
    # Handle ranges
    # Example:
    # MEDIUM TO SLIGHTLY THICK
    # THIN TO MEDIUM
    # VTN TO STK
    # ---------------------------------------------------
    if " TO " in value:
        parts = value.split(" TO ", 1)

        left = parts[0].strip()
        right = parts[1].strip()

        left_code = mapping.get(left)
        right_code = mapping.get(right)

        if left_code and right_code:
            return f"{left_code} TO {right_code}"

    # ---------------------------------------------------
    # Sometimes supplier values do not contain "TO"
    # Example: "MED STK"
    # ---------------------------------------------------
    words_mapping = [
        ("VERY THIN", "VTN"),
        ("SLIGHTLY THIN", "STN"),
        ("SLIGHT THIN", "STN"),
        ("MEDIUM", "MED"),
        ("SLIGHTLY THICK", "STK"),
        ("SLIGHT THICK", "STK"),
        ("VERY THICK", "VTHK"),
        ("VERY THIK", "VTHK"),
        ("EXTREMELY THICK", "ETK"),
        ("EXTREMLY THICK", "ETK"),
        ("THIN", "THN"),
        ("THICK", "THK"),
    ]

    converted = value

    for text, code in words_mapping:
        converted = re.sub(
            rf"\b{re.escape(text)}\b",
            code,
            converted
        )

    # Extract standardized codes
    codes = re.findall(
        r"\b(?:VTN|THN|STN|MED|STK|THK|VTHK|ETK)\b",
        converted
    )

    if len(codes) >= 2:
        return f"{codes[0]} TO {codes[1]}"

    if len(codes) == 1:
        return codes[0]

    return ""
def process_rapsize(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""

    try:
        carat = float(value)

        if 0.50 <= carat <= 0.59:
            return "50UP"
        if 0.60 <= carat <= 0.69:
            return "60UP"
        if 0.70 <= carat <= 0.79:
            return "70UP"
        if 0.80 <= carat <= 0.89:
            return "80UP"
        if 0.90 <= carat <= 0.99:
            return "90UP"
        if 1.00 <= carat <= 1.49:
            return "1.00UP"
        if 1.50 <= carat <= 1.99:
            return "1.50UP"
        if 2.00 <= carat <= 2.99:
            return "2.00UP"

        return ""

    except (ValueError, TypeError):
        return ""


def only_numeric_value(value):
    if pd.isna(value) or str(value).strip() == "":
        return ""

    numbers = re.findall(r"-?\d+(?:\.\d+)?", str(value))

    if not numbers:
        return ""

    return f"{float(numbers[0]):.2f}"

def process_black_combined(
    bic="",
    bis="",
    table_b="",
    side_b="",
    bt="",
    bc="",
    tb="",
    sb="",
    table_black="",
    side_black="",
    black_inclusion="",
    natts="",
    inclusion="",
    black_incl="",
    table_incl="",
    ngs_comment="",
    black_existing="",
    side_black_2="",
    black_in_table="",
):
    """Create one BLACK value from the different supplier BLACK columns."""

    def clean(value):
        if pd.isna(value):
            return ""
        value = str(value).strip()
        if value.upper() in {"", "NAN", "NONE", "NULL"}:
            return ""
        return value

    def combine(first, second):
        first = clean(first)
        second = clean(second)

        if first and second:
            return f"{first}-{second}"
        if first:
            return first
        if second:
            return second
        return ""

    # Paired BLACK fields. Use the first available pair.
    pairs = [
        (tb, sb),                      # TB + SB (RATNAKALA)
        (bic, bis),                    # BIC + BIS
        (table_b, side_b),             # T.B + S.B
        (bt, bc),                      # BT + BC
        (table_black, side_black),     # Table Black + Side Black
        (black_incl, table_incl),      # Black Incl. + Table Incl.
        (black_in_table, side_black_2),# Black In Table + Side Black
    ]

    for first, second in pairs:
        result = combine(first, second)
        if result:
            return result

    # Single columns that may already contain the complete BLACK description.
    single_values = [
        black_inclusion,
        natts,
        inclusion,
        ngs_comment,
        black_existing,
    ]

    for value in single_values:
        value = clean(value)
        if value:
            return value

    return ""


def create_excel_file(output_data):
    """Create one Excel file in memory from the combined processed data."""
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        output_data.to_excel(writer, index=False, sheet_name="DEAL DATA")

        # Make the worksheet easier to read.
        worksheet = writer.sheets["DEAL DATA"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    cell_length = len(str(cell.value)) if cell.value is not None else 0
                    max_length = max(max_length, cell_length)
                except Exception:
                    pass
            worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 35)

    excel_buffer.seek(0)
    return excel_buffer.getvalue()

def process_data(file_path, filters, column_mapping, prefix, skip_rows=0):
    data = load_data(file_path, skip_rows=skip_rows)

    black_series = data.apply(
        lambda row: process_black_combined(
            row.get("BIC", ""),
            row.get("BIS", ""),
            row.get("T.B", ""),
            row.get("S.B", ""),
            row.get("BT", ""),
            row.get("BC", ""),
            row.get("TB", ""),
            row.get("SB", ""),
            row.get("Table Black", ""),
            row.get("Side Black", ""),
            row.get("Black Inclusion   ", ""),
            row.get("Natts", ""),
            row.get("Inclusion", ""),
            row.get("Black Incl.", ""),
            row.get("Table Incl.", ""),
            row.get("NGS Comment", ""),
            row.get("BLACK", ""),
            row.get("Side Black", ""),
            row.get("Black In Table", ""),
        ),
        axis=1,
    )
    data["BLACK"] = black_series

    if all(col in data.columns for col in ["Max", "Min", "Hgt"]):
        data["Measurement"] = data.apply(
            lambda row: f"{row['Max']} - {row['Min']} X {row['Hgt']}",
            axis=1,
        )

    if all(col in data.columns for col in ["Length", "Width", "Depth"]):
        data["Measurement"] = data.apply(
            lambda row: f"{row['Length']} - {row['Width']} X {row['Depth']}",
            axis=1,
        )

    if all(col in data.columns for col in ["M1", "M2", "M3"]):
        data["Measurement"] = data.apply(
            lambda row: f"{row['M1']} - {row['M2']} X {row['M3']}",
            axis=1,
        )

    if all(col in data.columns for col in ["Length", "Width", "Height"]):
        data["Measurement"] = data.apply(
            lambda row: f"{row['Length']} - {row['Width']} X {row['Height']}",
            axis=1,
        )

    if all(col in data.columns for col in ["Meas1", "Meas2", "Meas3"]):
        data["Measurement"] = data.apply(
            lambda row: f"{row['Meas1']} - {row['Meas2']} X {row['Meas3']}",
            axis=1,
        )

    # Build GIRDLE when thin/thick girdle values are supplied separately.
    if all(col in data.columns for col in ["Girdle Thin", "Girdle Thick"]):
        data["GIRDLE"] = data.apply(
            lambda row: f"{row['Girdle Thin']} TO {row['Girdle Thick']}",
            axis=1,
        )

    filtered_data = apply_filters(data, filters).copy()

    # Build final output directly here. No map_columns() function is used.
    output_data = pd.DataFrame(index=range(len(filtered_data)))

    for column, keys in column_mapping.items():

        # Use Seller value for Name when the uploaded file has a Seller column.
        # If Seller is missing or blank, fall back to the supplier/file identifier.
        if column == "Name":
            seller_column = next(
                (
                    source_col
                    for source_col in filtered_data.columns
                    if str(source_col).strip().lower() == "seller"
                ),
                None,
            )

            if seller_column is not None:
                seller_series = filtered_data[seller_column].reset_index(drop=True)

                blank_seller = (
                    seller_series.isna()
                    | seller_series.astype(str).str.strip().eq("")
                    | seller_series.astype(str).str.strip().str.lower().eq("nan")
                )

                output_data[column] = seller_series.mask(blank_seller, prefix)
            else:
                output_data[column] = prefix

            continue

        # Columns with no source mapping stay blank.
        if not keys:
            output_data[column] = np.nan
            continue

        # Create the output column FIRST from the first matching source column.
        source_series = concatenate_dataframes({0: filtered_data}, keys).reset_index(drop=True)

        source_series = source_series.reindex(range(len(filtered_data)))
        output_data[column] = source_series

        # Apply cleanup only AFTER the column exists.
        if column in ["CUT", "POL", "SYM"]:
            output_data[column] = output_data[column].apply(process_cut_polish_symmetry)

        elif column == "FLUOR":
            output_data[column] = output_data[column].apply(process_fluor)

        elif column == "MEASUERMENT":
            output_data[column] = output_data[column].apply(process_measurement)

        elif column == "SHAPE":
            output_data[column] = output_data[column].apply(process_shape)

        elif column in ["KEY TO SYMBOL", "REPORT COMMENT"]:
            output_data[column] = output_data[column].apply(replace_comma_with_hyphen)

        elif column == "CULET":
            output_data[column] = output_data[column].apply(process_culet)

        elif column == "GIRDLE":
            output_data[column] = output_data[column].apply(process_girdle)

        elif column == "DISC":
            output_data[column] = output_data[column].apply(process_discount)

    # Calculate RAPSIZE from the final mapped CARAT column.
    if "CARAT" in output_data.columns:
        output_data["RAPSIZE"] = output_data["CARAT"].apply(process_rapsize)
            
        return len(output_data), output_data

def run_browns_mod():
    st.title("Brown File")
    st.caption("AKARSH-1, GLOWSTAR-4, HK-4 , KIRAN-10, LAKHI-0,NAROLA-0 RAPNET-0, RATNAKALA-4, RAYS-1,STARLIGHT-0  ")

    # File uploader
    with st.sidebar:
        files = st.file_uploader("Upload Excel or CSV Files",type=["xlsx", "xls", "csv"],accept_multiple_files=True, key="deal_uploader")
            
    current_datetime = datetime.now().strftime("%d.%m.%Y")

    # Store each processed DataFrame so all successful files can be merged later.
    if "deal_results" not in st.session_state:
        st.session_state.deal_results = {}

    if not files:
        st.warning("No files selected.")
        return

    # Keep output list synced with the currently uploaded files.
    st.session_state.deal_results = {}

    # Column Mapping
    columns_mapping = {
        "Name": ["Name"],
        "STOCKID": ["Packet No","Stone Id","Stone No","Stone ID","StockID","Vendor Stock number","Vendor","Stock #","Stock No",'Stone Id'],
        "Lot No. ": ["Diamond ID"],
        "SHAPE": ["Shape", "Shp",],
        "RAPSIZE": [],
        "CERT": ["Lab", "LAB", "Cert","Certi"],
        "CARAT": ["Weight", "Size", "Carat", "Cts", "Carats",'Wgt'],
        "COLOR": ["Color", "Col"],
        "CLARITY": ["Clarity", "Clr","Purity"],
        "CUT": ["Cut"],
        "POL": ["Polish", "Pol"],
        "SYM": ["Symmetry", "Sym","Symm"],
        "FLUOR": ["Fls","Fluorescence Intensity","Fluorescence","Flour","Flr","Flo", "FL",'Fluorescent'],
        "MEASUERMENT": ["Measurements", "Measurement", "Diam."],
        "GIRDLE": ["GIRDLE", "Girdle"],
        "RAPPRICE": ["New Rap", "Rap Rate", "Rap", "Rap $/CT", "RapList", "Rapa"],
        "RAP V": ["Rap $", "Rap", "New Total Rap", "Rate", "Rap Value"],
        "DISC": ["Back","Rapnet  Discount %","%Rap","Rap%","Disc%","Base Off %","Disc %","Disc(%)"],
        "PRICE": ["Rate", "New $/ct", "$/ct", "Pr/Ct", "Rap $", "Base Rate", "Price/Ct"],
        "VALUE": ["New Total", "Total", "Amt $", "Amount", "value", "Amount","Net Value"],
        "CERTNO": ["Report #", "Rep No", "Report No", "Cert.No."],
        "DEPTH": ["T.Dep.", "Depth %", "TD%", "Total Depth", "Depth",'Depth%'],
        "TABLE": ["Table %", "Table", "Tab%", "Tbl",'Table%'],
        "RATIO": ["Ratio", "D/R", "L:W"],
        "BLACK": ["BLACK"],
        "TINCH": ['Shade ',"Tinge","Shade", "shade", "Shd.", "Col Tinge", "Color Shade",'Shade'],
        "LUSTER": ["Luster", "milky", "Lust", "Lus", "MILKY",'TB','Milky'],
        "CULET": ["Culet", "CULET", "Culet Size","Culet"],
        "KEY TO SYMBOL": ["Key To Symbols", "Key To Symbol", "K T S", "Key to symbols",'Key To Symbol'],
        "Member Comment": ['Comments','BGM Comment'],
        "Lab Comment": ["Rep Com", "Cert comment", "Lab Comment",'Report Comments'],
        "STATUS": ["Status"],
        "STONE LOCATION": ["Location", "Loc", "Country"],
        "City": ["City"],
    }

    filters_akarsh = {}
     
    filters_glowstar = {}
     
    filters_hk = {}
     
    filters_kiran = {}
     
    filters_lakhi = { }
    
    filters_narola = {}
    
    filters_ratnakala = {}
    
    filters_rays = {}
               
    filters_srk= {}

    filters_starlight= {}
              

    prefix_info = {
        "AKARSH": {"filters": filters_akarsh, "skip_rows": 1},
        "GLOWSTAR": {"filters": filters_glowstar, "skip_rows": 4},
        "HK": {"filters": filters_hk, "skip_rows": 4},
        "KIRAN": {"filters": filters_kiran, "skip_rows": 10},
        "LAKHI": {"filters": filters_lakhi, "skip_rows": 0},
        "NAROLA": {"filters": filters_narola, "skip_rows": 0},
        "RATNAKALA": {"filters": filters_ratnakala, "skip_rows": 4},
        "RAYS": {"filters": filters_rays, "skip_rows": 1},
        "SRK": {"filters": filters_srk, "skip_rows": 0},
        "STARLIGHTDIAMONDS": {"filters": filters_starlight, "skip_rows": 0},
        }

  
    with st.spinner("Processing files..."):
        for file in files:
            file_name = file.name
            file_name_upper = file_name.upper()
            base_name, _ = os.path.splitext(file_name)
            base_name_upper = base_name.upper()

            matched = False

            # ---------------------------------------------------------
            # 1. PROCESS KNOWN SUPPLIERS
            # ---------------------------------------------------------
            for prefix, configurations in prefix_info.items():
                if prefix in file_name_upper:
                    matched = True

                    if isinstance(configurations, list):
                        # If configurations is a list, process each configuration separately
                        for config in configurations:
                            try:
                                row_count, output_data = process_data(
                                    file_path=file,
                                    filters=config["filters"],
                                    column_mapping=columns_mapping,
                                    prefix=prefix,
                                    skip_rows=config["skip_rows"],
                                )

                                if row_count > 0:
                                    st.session_state.deal_results[(file_name, prefix)] = {
                                        "row_count": row_count,
                                        "df": output_data,
                                    }
                                    # Stop after the first successful config with data.
                                    break

                            except Exception as e:
                                st.error(
                                    f"Error processing file {file_name} with prefix '{prefix}' using configuration: {e}"
                                )
                                continue
                    else:
                        # Single configuration
                        try:
                            row_count, output_data = process_data(
                                file_path=file,
                                filters=configurations["filters"],
                                column_mapping=columns_mapping,
                                prefix=prefix,
                                skip_rows=configurations["skip_rows"],
                            )

                            if row_count > 0:
                                st.session_state.deal_results[(file_name, prefix)] = {
                                    "row_count": row_count,
                                    "df": output_data,
                                }

                        except Exception as e:
                            st.error(
                                f"Error processing file {file_name} with prefix '{prefix}' using configuration: {e}"
                            )

                    # A file should only use the first matching known supplier prefix.
                    break

            # ---------------------------------------------------------
            # 2. PROCESS ANY FILE NOT PRESENT IN prefix_info
            # ---------------------------------------------------------
            if not matched:
                try:
                    row_count, output_data = process_data(
                        file_path=file,
                        filters={},
                        column_mapping=columns_mapping,
                        prefix=base_name_upper,
                        skip_rows=0,
                    )

                    if row_count > 0:
                        st.session_state.deal_results[(file_name, base_name_upper)] = {
                            "row_count": row_count,
                            "df": output_data,
                        }
                    else:
                        st.warning(
                            f"{file_name}: mapped successfully but no rows passed the common filters."
                        )

                except Exception as e:
                    st.error(
                        f"Error processing unknown file {file_name}: {e}"
                    )

        # ---- Show every processed file inside an expander ----
        st.markdown("---")

        if not st.session_state.deal_results:
            st.warning("No matching file format.")
            return

        total_files = len(st.session_state.deal_results)
        total_rows = sum(info.get("row_count", 0) for info in st.session_state.deal_results.values())

        # Merge every processed output vertically into one table.
        combined_frames = [
            info["df"]
            for info in st.session_state.deal_results.values()
            if info.get("df") is not None and not info["df"].empty
        ]

        if not combined_frames:
            st.warning("No rows are available to create the Excel file.")
            return

        combined_df = pd.concat(combined_frames, ignore_index=True)

        # Remove duplicate STOCKID values from the FINAL combined file.
        # Keep only the first occurrence across all uploaded supplier files.
        # ---------------------------------------------------------
        # REMOVE DUPLICATE STOCKID FROM FINAL COMBINED DATA
        # ---------------------------------------------------------

        if "STOCKID" in combined_df.columns:

            stockid_clean = ( combined_df["STOCKID"].astype(str).str.strip().str.upper().str.replace(r"\.0$", "", regex=True))
            
            blank_mask = stockid_clean.isin(["", "NAN", "NONE", "NULL"] )
                
            duplicate_mask = (stockid_clean.duplicated(keep="first")& ~blank_mask )
             
            combined_df = (combined_df.loc[~duplicate_mask].reset_index(drop=True))

        excel_data = create_excel_file(combined_df)
        excel_name = f"BROWN File {current_datetime}.xlsx"

        st.success(f"Combined {total_files} processed file(s) into one Excel file with {len(combined_df)} rows.")

        st.download_button(
            label="Download File",
            data=excel_data,
            file_name=excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_combined_excel",
        )

        # Normal interactive Streamlit preview.
        with st.expander(
            f"Combined Data Preview ({total_files} files | {total_rows} rows)",
            expanded=True,
        ):
            preview_df = combined_df.head(200).copy()
            preview_df.columns = [str(col) for col in preview_df.columns]

            st.dataframe(
                preview_df,
                hide_index=True,
                width="stretch",
                height=500,
            )

if __name__ == "__main__":
    run_browns_mod()
