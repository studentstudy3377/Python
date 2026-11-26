import base64
import os
from pathlib import Path
import platform
import re
import datetime as dt
import csv
from io import BytesIO, StringIO
from zipfile import ZipFile, ZIP_DEFLATED, BadZipFile
import pandas as pd
import streamlit as st

def select_files():
  files = st.file_uploader(
    "Select Files",
    accept_multiple_files=True,
  )
  return files

def open_file(file_path):
  if platform.system() == "Windows":
    os.startfile(file_path)

def process_name(df, filename):
  party_name = os.path.splitext(os.path.basename(filename))[0]
  seller_column = None
  for col in df.columns:
    if str(col).strip().lower() == "seller":
      seller_column = col
      break

  if seller_column is None:
    return pd.Series(
      [party_name] * len(df),
      index=df.index,
      dtype="object",
    )

  def get_name(value):
    if pd.isna(value):
      return party_name
    text = str(value).strip()
    if text == "" or text.lower() in {"nan", "none"}:
      return party_name
    return text
  return df[seller_column].apply(get_name)
def process_files(uploaded_files):
  dfs = load_excel_files(uploaded_files)
  converted = []
  for filename, df in dfs:
    df = df.copy()
    df["Name"] = process_name(df, filename)
    converted_df = convert_dataframe(df)
    converted.append(converted_df)
  if not converted:
    return pd.DataFrame()
  return pd.concat(converted, ignore_index=True)

def process_measurement(value):
  if pd.isna(value) or str(value).strip() == "":
    return ""
  value = str(value).strip()
  value = re.sub(r"[\\*xX×]", " X ", value)
  numbers = re.findall(r"\d+(?:\.\d+)?", value)

  if len(numbers) >= 3:
    return f"{numbers[0]} - {numbers[1]} X {numbers[2]}"

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
    "None":"",
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
    "PE": "PEAR",
    "PEAR MODIFIED BRILLIANT": "PEAR",
    "OL": "OVAL",
    "OV": "OVAL",
    "OVAL BRILLIANT": "OVAL",
    "OVAL MODIFIED BRILLIANT": "OVAL",
    "MARQUIES": "MARQUISE",
    "MQ": "MARQUISE",
    "MARQUISE BRILLIANT": "MARQUIES",
    "EM": "EMERALD",
    "EMD": "EMERALD",
    "EMERALD CUT": "EMERALD",
    "CUS CUT": "CUSHION CUT",
    "CUSH": "CUSHION CUT",
    "CUSHION MODIFIED BRILLIANT": "CUSHION CUT",
    "HEART BRILLIANT": "HEART",
    "HEART MODIFIED BRILLIANT": "HEART",
    "HT": "HEART",
  }

  return mappings.get(value, value)


def process_fluor(value):
  if pd.isna(value):
    return ""

  value = str(value).strip().upper()

  mappings = {
    "N": "NONE",
    "F": "FNT",
    "NAN": "NONE",
    "NON": "NONE",
    "MED. BLUE": "MED",
    "MEDIUMBLUE": "MED",
    "MEDIUMWHITE": "MED",
    "M": "MED",
    "MEDIUMYELLOW": "MED",
    "MEDIUMGREEN": "MED",
    "S": "STG",
    "STRONG": "STG",
    "V.STRONG": "VST",
  }

  return mappings.get(value, value)


def girdle_process(value):
  if pd.isna(value) or str(value).strip() == "":
    return ""

  value = str(value).strip().upper()

  # Remove numeric girdle percentages, for example 3.5 or 3.5%
  value = re.sub(r"\b\d+(?:\.\d+)?%?\b", "", value)

  # Remove FACETED text because it does not affect the required output grade.
  value = re.sub(r"\bFACETED\b", "", value, flags=re.IGNORECASE)

  # Clean punctuation/spacing left after removing percentages or FACETED.
  value = re.sub(r"\s*,\s*", " ", value)
  value = re.sub(r"\s+", " ", value).strip(" ,- ")

  mapping = {
    "MEDIUM TO SLIGHTLY THICK": "MED TO STK",
    "THICK TO EXTREMELY THICK": "THK TO ETK",
    "SLIGHTLY THICK TO THICK": "STK TO THK",
    "THIN TO SLIGHTLY THICK": "THN TO STK",
    "THIN TO MEDIUM": "THN TO MED",
    "STK - THK": "MED TO STK",
    "M.S": "MED TO STK",
    "THN - STK": "THN TO STK",
    "MEDIUM TO THICK": "MED TO THK",
    "MEDIUM-SLIGHTLY THICK": "MED TO STK",
    "MED - THK": "MED TO THK",
    "MED-THK": "MED TO THK",
    "MEDIUM TO VERY THICK": "MED TO VTHK",
    "SLIGHTLY THICK": "STK",
    "SLIGHT THICK": "STK",
    "ST": "STK",
    "MEDIUM": "MED",
    "VERY THICK": "VTHK",
    "VERY THIK": "VTHK",
    "VERY THIN": "VTN",
    "TK": "THK",
    "THICK,": "THK",
    "THICK": "THK",
    "EXTREMLY THICK": "ETK",
    "EXTREMLY THIN": "ETN",
    "THIN TO THICK": "THK",
  }

  for key, mapped_value in sorted(mapping.items(), key=lambda x: len(x[0]), reverse=True):
    if value.startswith(key):
      return mapped_value

  return value


def process_tinch(value):
  value_upper = value.upper()
  mappings = {
    "YELLOW": "NONE",
    "WH": "EX",
    "N": "NONE",
    "MT1": "B-1",
  }

  return mappings.get(value_upper,value)


def process_luster(value):
  value = str(value).strip().upper()
  mappings = {
    "NON": "EX",
    "NONE": "EX",
    "VG":"M-1",
    "VERY LIGHT MILKY":"M-1",
    "LIGHT MILKY":"M-2",
  }

  return mappings.get(value,value)


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


def process_location(value):
  if pd.isna(value) or str(value).strip().upper() == "NAN":
    return ""

  value = str(value).strip()

  mappings = {
    "MUMBAI": "INDIA",
  }

  return mappings.get(value.upper(), value.upper())


def replace_comma_with_hyphen(value):
  if pd.isna(value):
    return value

  return re.sub(r"\s*,\s*", "-", str(value)).strip()


def only_number(value):
  if pd.isna(value) or str(value).strip() == "":
    return ""

  text = str(value).strip()
  # Remove everything except digits, decimal point and minus sign
  cleaned = re.sub(r"[^0-9.\\-]", "", text)

  if not cleaned:
    return ""

  try:
    number = float(cleaned)

    if abs(number) <= 1 and "%" not in text:
      number *= 100

    return f"{number:.2f}%"

  except ValueError:
    return ""


def get_ca_ph(value, position):
  if pd.isna(value) or str(value).strip() == "":
    return ""

  original_value = str(value).strip()
  numbers = re.findall(r"\d+(?:\.\d+)?", original_value)

  if len(numbers) > position:
    return f"{float(numbers[position]):.2f}"

  # If requested position is not available, keep the original value.
  return original_value


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


def only_numeric_value(value):
  if pd.isna(value) or str(value).strip() == "":
    return ""

  numbers = re.findall(r"-?\d+(?:\.\d+)?", str(value))

  if not numbers:
    return ""

  return f"{float(numbers[0]):.2f}"


MINIMAL_EXCEL_STYLES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="1">
        <font>
            <sz val="11"/>
            <name val="Calibri"/>
            <family val="2"/>
        </font>
    </fonts>
    <fills count="2">
        <fill><patternFill/></fill>
        <fill><patternFill patternType="gray125"/></fill>
    </fills>
    <borders count="1">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <tableStyles
        count="0"
        defaultTableStyle="TableStyleMedium9"
        defaultPivotStyle="PivotStyleLight16"
    />
</styleSheet>
"""


OPENXML_EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}

SUPPORTED_DATA_EXTENSIONS = {
  ".xlsx", ".xlsm", ".xltx", ".xltm",
  ".xls", ".xlsb",
  ".ods", ".odf",
  ".csv", ".tsv", ".txt", ".dat",
  ".json", ".jsonl", ".ndjson",
  ".xml",
  ".html", ".htm",
  ".parquet", ".pq",
  ".feather",
}


OPTIONAL_DEPENDENCY_HELP = {
  ".xls": "pip install xlrd",
  ".xlsb": "pip install pyxlsb",
  ".ods": "pip install odfpy",
  ".odf": "pip install odfpy",
  ".parquet": "pip install pyarrow",
  ".pq": "pip install pyarrow",
  ".feather": "pip install pyarrow",
  ".html": "pip install lxml",
  ".htm": "pip install lxml",
  ".xml": "pip install lxml",
}


# Common vendor-stock headers used to locate the real header row when a
# workbook has title/summary rows above the table (as in the HK files).
HEADER_HINT_KEYS = {
  "stockno", "stockid", "stoneid", "vendorstocknumber",
  "shape", "shp", "carat", "carats", "weight", "cts", "size",
  "clarity", "clr", "color", "col", "cut", "pol", "polish",
  "sym", "symmetry", "flo", "flr", "fluorescence",
  "lab", "cert", "rap", "rapv", "rappct", "prct", "amount",
  "measurement", "measurements", "ratio", "location", "girdle",
  "status", "bt", "bc", "milky", "luster", "keytosymbol",
  "giacomments", "pvang", "pvhgt", "crang", "crhgt", "gtpct",
}

CORE_HEADER_KEYS = {
  "stockno", "stockid", "stoneid", "shape", "carat", "weight",
  "clarity", "color",
}


def column_key(value):
  """Normalize a column/header name for tolerant matching."""

  if value is None:
    return ""

  text = str(value).replace("\xa0", " ").strip().lower()

  if text in {"", "nan", "none"}:
    return ""

  # Keep useful Excel-header meaning before removing punctuation.
  text = text.replace("%", " pct ")
  text = text.replace("#", " no ")
  text = text.replace("$", " dollar ")

  return re.sub(r"[^a-z0-9]+", "", text)


def clean_column_name(value, fallback):
  """Return a readable, whitespace-normalized column name."""

  if value is None:
    return fallback

  text = str(value).replace("\xa0", " ").strip()
  text = re.sub(r"\s+", " ", text)

  if text == "" or text.lower() in {"nan", "none"}:
    return fallback

  return text


def make_unique_columns(values):
  """Build unique column names while keeping vendor header text readable."""

  result = []
  counts = {}

  for position, value in enumerate(values, start=1):
    base = clean_column_name(value, f"Column {position}")
    key = base.lower()
    counts[key] = counts.get(key, 0) + 1

    if counts[key] == 1:
      result.append(base)
    else:
      result.append(f"{base}__{counts[key]}")

  return result


def detect_stock_header_row(raw_df, max_rows=40):
  """
  Detect a vendor stock-table header row inside a raw DataFrame.

  Many supplier workbooks place totals, notes or report titles above the
  actual table. pandas' normal header=0 behavior then treats those report
  rows as column names, which makes all later column mappings fail.
  """

  if raw_df is None or raw_df.empty:
    return 0

  best_index = 0
  best_score = -1
  rows_to_check = min(len(raw_df), max_rows)

  for index in range(rows_to_check):
    row = raw_df.iloc[index].tolist()
    keys = {column_key(value) for value in row}
    keys.discard("")

    hint_matches = keys & HEADER_HINT_KEYS
    core_matches = keys & CORE_HEADER_KEYS

    # Core identifiers are weighted more heavily than optional fields.
    score = (len(core_matches) * 4) + len(hint_matches)

    if score > best_score:
      best_score = score
      best_index = index

  # Require enough stock-related evidence before skipping report/title rows.
  # If not found, fall back to the first row to preserve normal Excel behavior.
  if best_score < 8:
    return 0

  return best_index


def promote_detected_header(raw_df):
  """Promote the detected table-header row and remove report rows above it."""

  if raw_df is None or raw_df.empty:
    return raw_df

  header_index = detect_stock_header_row(raw_df)
  headers = make_unique_columns(raw_df.iloc[header_index].tolist())

  df = raw_df.iloc[header_index + 1:].copy()
  df.columns = headers
  df.reset_index(drop=True, inplace=True)

  # Remove fully blank lines that sometimes exist below/inside vendor reports.
  nonblank = df.apply(
    lambda row: any(str(value).strip() not in {"", "nan", "None"} for value in row),
    axis=1,
  )
  df = df.loc[nonblank].reset_index(drop=True)

  return df


def normalize_dataframe_columns(df):
  """Strip extra spaces/NBSPs and make duplicate input headings unique."""

  if df is None:
    return df

  df = df.copy()
  df.columns = make_unique_columns(df.columns)
  return df


def read_excel_with_header_detection(raw_bytes, engine, sheet_name=0):
  """Read an Excel-like file as raw rows, then locate its true table header."""

  raw_df = pd.read_excel(
    BytesIO(raw_bytes),
    engine=engine,
    sheet_name=sheet_name,
    header=None,
    keep_default_na=False,
  )

  return promote_detected_header(raw_df)


def is_stylesheet_error(exc):
  """Return True only when the exception chain points to Excel style XML."""

  messages = []
  current = exc
  seen = set()

  while current is not None and id(current) not in seen:
    seen.add(id(current))
    messages.append(str(current).lower())

    traceback_obj = current.__traceback__
    while traceback_obj is not None:
      frame_path = traceback_obj.tb_frame.f_code.co_filename.lower().replace("\\\\", "/")
      function_name = traceback_obj.tb_frame.f_code.co_name.lower()

      if (
        "/openpyxl/styles/" in frame_path
        or frame_path.endswith("/openpyxl/styles/stylesheet.py")
        or function_name == "apply_stylesheet"
      ):
        return True

      traceback_obj = traceback_obj.tb_next

    current = current.__cause__ or current.__context__

  error_text = " ".join(messages)

  return any(
    marker in error_text
    for marker in (
      "could not read stylesheet",
      "stylesheet",
      "styles.xml",
    )
  )


def remove_xlsx_styles(xlsx_bytes):
  """
  Build an in-memory copy of an OpenXML Excel workbook without source styles.

  This fallback is intended for XLSX/XLSM/XLTX/XLTM workbooks whose
  xl/styles.xml is invalid. Worksheet values/formulas are preserved;
  formatting is intentionally lost in the temporary copy only.
  """

  source = BytesIO(xlsx_bytes)
  output = BytesIO()
  styles_found = False

  try:
    with ZipFile(source, "r") as zin:
      with ZipFile(output, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
          data = zin.read(item.filename)

          if item.filename == "xl/styles.xml":
            styles_found = True
            data = MINIMAL_EXCEL_STYLES

          elif (
            item.filename.startswith("xl/worksheets/")
            and item.filename.endswith(".xml")
          ):
            # Remove cell style IDs, e.g. <c r="A1" s="15" ...>.
            data = re.sub(
              rb'(<c\b[^>]*?)\s+s="\d+"',
              rb"\1",
              data,
            )

            # Remove row style IDs and custom-format flags.
            data = re.sub(
              rb'(<row\b[^>]*?)\s+s="\d+"',
              rb"\1",
              data,
            )
            data = re.sub(
              rb'\s+customFormat="[^"]*"',
              b"",
              data,
            )

            # Remove column style IDs, e.g. style="4".
            data = re.sub(
              rb'(<col\b[^>]*?)\s+style="\d+"',
              rb"\1",
              data,
            )

          zout.writestr(item, data)

        if not styles_found:
          zout.writestr("xl/styles.xml", MINIMAL_EXCEL_STYLES)

  except BadZipFile as exc:
    raise ValueError(
      "The uploaded file is not a valid OpenXML Excel/ZIP workbook."
    ) from exc

  output.seek(0)
  return output


def read_openxml_excel_safely(raw_bytes, sheet_name=0):
  """
  Read XLSX/XLSM/XLTX/XLTM data with openpyxl.

  The workbook is first read without assuming row 1 is the header. The real
  stock-table header is detected automatically, so files with totals/notes
  above the table (including the HK workbooks) are handled correctly.

  If openpyxl fails specifically on corrupt formatting, retry from an
  in-memory formatting-free copy. Other workbook errors are re-raised.

  Returns:
      (DataFrame, repaired_styles)
  """

  try:
    df = read_excel_with_header_detection(raw_bytes, "openpyxl", sheet_name)
    return normalize_dataframe_columns(df), False

  except Exception as exc:
    if not is_stylesheet_error(exc):
      raise

    repaired_file = remove_xlsx_styles(raw_bytes)
    repaired_bytes = repaired_file.getvalue()

    df = read_excel_with_header_detection(repaired_bytes, "openpyxl", sheet_name)
    return normalize_dataframe_columns(df), True


def read_delimited_text(raw_bytes, extension):
  """Read CSV/TSV/TXT/DAT with delimiter and encoding fallbacks."""

  encodings = ("utf-8-sig", "utf-8", "cp1252", "latin1")
  last_error = None

  for encoding in encodings:
    try:
      text = raw_bytes.decode(encoding)
    except UnicodeDecodeError as exc:
      last_error = exc
      continue

    if extension == ".tsv":
      delimiter = "\t"
    else:
      sample = text[:65536]
      try:
        delimiter = csv.Sniffer().sniff(
          sample,
          delimiters=",\t;|",
        ).delimiter
      except csv.Error:
        # Choose the delimiter that appears most often in the header line.
        header = sample.splitlines()[0] if sample.splitlines() else ""
        candidates = [",", "\t", ";", "|"]
        delimiter = max(candidates, key=header.count)

        # No delimiter found: treat whitespace as the separator.
        if header.count(delimiter) == 0:
          delimiter = r"\s+"

    try:
      return pd.read_csv(
        StringIO(text),
        sep=delimiter,
        engine="python",
        keep_default_na=False,
      )
    except pd.errors.ParserError as exc:
      last_error = exc
      continue

  if last_error is not None:
    raise last_error

  raise ValueError("Unable to detect the text file encoding/delimiter.")


def read_json_safely(raw_bytes, extension):
  """Read regular JSON arrays/objects and line-delimited JSON."""

  lines_first = extension in {".jsonl", ".ndjson"}
  attempts = (True, False) if lines_first else (False, True)
  last_error = None

  for lines in attempts:
    try:
      data = pd.read_json(BytesIO(raw_bytes), lines=lines)

      # Convert a Series-like result to a one-column DataFrame if necessary.
      if isinstance(data, pd.Series):
        data = data.to_frame()

      return data
    except (ValueError, TypeError) as exc:
      last_error = exc

  raise ValueError(f"Unable to read JSON data: {last_error}")


def read_html_safely(raw_bytes):
  """Read the first HTML table from a file."""

  tables = pd.read_html(BytesIO(raw_bytes), keep_default_na=False)
  if not tables:
    raise ValueError("No HTML table was found in the file.")
  return tables[0]


def read_single_data_file(filename, raw_bytes):
  """
  Read one supported tabular file from bytes.

  Returns:
      (DataFrame, repaired_styles)
  """

  extension = os.path.splitext(filename)[1].lower()

  try:
    if extension in OPENXML_EXCEL_EXTENSIONS:
      return read_openxml_excel_safely(raw_bytes)

    if extension == ".xls":
      return (
        normalize_dataframe_columns(
          read_excel_with_header_detection(raw_bytes, "xlrd")
        ),
        False,
      )

    if extension == ".xlsb":
      return (
        normalize_dataframe_columns(
          read_excel_with_header_detection(raw_bytes, "pyxlsb")
        ),
        False,
      )

    if extension in {".ods", ".odf"}:
      return (
        normalize_dataframe_columns(
          read_excel_with_header_detection(raw_bytes, "odf")
        ),
        False,
      )

    if extension in {".csv", ".tsv", ".txt", ".dat"}:
      return read_delimited_text(raw_bytes, extension), False

    if extension in {".json", ".jsonl", ".ndjson"}:
      return read_json_safely(raw_bytes, extension), False

    if extension == ".xml":
      return pd.read_xml(BytesIO(raw_bytes)), False

    if extension in {".html", ".htm"}:
      return read_html_safely(raw_bytes), False

    if extension in {".parquet", ".pq"}:
      return pd.read_parquet(BytesIO(raw_bytes)), False

    if extension == ".feather":
      return pd.read_feather(BytesIO(raw_bytes)), False

    raise ValueError(
      f"Unsupported file type: {extension or '(no extension)'}. "
      "This converter reads tabular/data files, not arbitrary documents, "
      "images, executables or media files."
    )

  except ImportError as exc:
    install_hint = OPTIONAL_DEPENDENCY_HELP.get(extension)
    if install_hint:
      raise ValueError(
        f"Reading '{extension}' files needs an optional Python package. "
        f"Install it with: {install_hint}"
      ) from exc
    raise


def read_zip_archive(filename, raw_bytes):
  """Read every supported tabular file stored inside a ZIP archive."""

  loaded = []
  repaired = []
  skipped = []

  try:
    with ZipFile(BytesIO(raw_bytes), "r") as archive:
      members = [m for m in archive.infolist() if not m.is_dir()]

      # Prevent accidental extraction of extremely large archive contents.
      max_uncompressed_bytes = 500 * 1024 * 1024
      total_size = sum(member.file_size for member in members)
      if total_size > max_uncompressed_bytes:
        raise ValueError(
          "ZIP archive is larger than 500 MB after extraction. "
          "Please split it into smaller archives."
        )

      for member in members:
        inner_name = os.path.basename(member.filename)
        if not inner_name or inner_name.startswith("~$"):
          continue

        inner_ext = os.path.splitext(inner_name)[1].lower()
        if inner_ext not in SUPPORTED_DATA_EXTENSIONS:
          skipped.append(inner_name)
          continue

        inner_bytes = archive.read(member)
        df, repaired_styles = read_single_data_file(inner_name, inner_bytes)
        loaded.append((inner_name, df))

        if repaired_styles:
          repaired.append(f"{filename} :: {inner_name}")

  except BadZipFile as exc:
    raise ValueError(f"'{filename}' is not a valid ZIP archive.") from exc

  if not loaded:
    raise ValueError(
      f"ZIP file '{filename}' does not contain any supported tabular/data files."
    )

  return loaded, repaired, skipped


def load_excel_files(uploaded_files):
  """
  Load common tabular/data file formats while preserving source filenames.

  This function keeps the historical name `load_excel_files` so the rest of
  the existing converter code does not need to change.
  """

  dataframes = []
  repaired_files = []
  skipped_archive_files = []

  for uploaded_file in uploaded_files:
    uploaded_file.seek(0)

    filename = uploaded_file.name
    extension = os.path.splitext(filename)[1].lower()
    raw_bytes = uploaded_file.read()

    if extension == ".zip":
      try:
        zip_dataframes, zip_repaired, zip_skipped = read_zip_archive(
          filename,
          raw_bytes,
        )
      except Exception as exc:
        raise ValueError(
          f"Unable to read ZIP file '{filename}': {exc}"
        ) from exc

      dataframes.extend(zip_dataframes)
      repaired_files.extend(zip_repaired)
      skipped_archive_files.extend(
        f"{filename} :: {inner_name}" for inner_name in zip_skipped
      )
      continue

    try:
      df, repaired_styles = read_single_data_file(filename, raw_bytes)
    except Exception as exc:
      raise ValueError(
        f"Unable to read file '{filename}': {exc}"
      ) from exc

    if repaired_styles:
      repaired_files.append(filename)

    df = normalize_dataframe_columns(df)
    dataframes.append((filename, df))

  if repaired_files:
    st.info(
      f"{len(repaired_files)} Excel file(s) had invalid formatting, "
      "but their worksheet data was loaded successfully."
    )

  if skipped_archive_files:
    st.info(
      f"Skipped {len(skipped_archive_files)} unsupported file(s) inside ZIP "
      "archive(s). Supported tabular files were still processed."
    )

  return dataframes


def get_source_series(df, possible_columns):
  """
  Return the first matching source column using tolerant header matching.

  Matching ignores capitalization, repeated spaces, non-breaking spaces and
  punctuation differences such as ``Stock #`` versus ``Stock No``.
  """

  lookup = {}
  for actual_column in df.columns:
    key = column_key(actual_column)
    if key and key not in lookup:
      lookup[key] = actual_column

  for source_column in possible_columns:
    key = column_key(source_column)
    if key and key in lookup:
      return df[lookup[key]].copy()

  return pd.Series([""] * len(df), index=df.index, dtype="object")


def process_rapsize(value):
  if pd.isna(value) or str(value).strip() == "":
    return ""

  try:
    carat = float(value)

    if 0.18 <= carat <= 0.20:
      return "18UP"
    if 0.20 <= carat <= 0.22:
      return "20UP"
    if 0.23 <= carat <= 0.29:
      return "23UP"
    if 0.23 <= carat <= 0.29:
      return "23UP"
    if 0.30 <= carat <= 0.39:
      return "30UP"
    if 0.40 <= carat <= 0.49:
      return "40UP"
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
    if 3.00 <= carat <= 3.99:
      return "3.00UP"
    if 4.00 <= carat <= 4.99:
      return "4.00UP"
    if 5.00 <= carat <= 5.99:
      return "5.00UP"
    if 6.00 <= carat <= 6.99:
      return "6.00UP"
    if 7.00 <= carat <= 7.99:
      return "7.00UP"
    if 8.00 <= carat <= 8.99:
      return "8.00UP"
    if 9.00 <= carat <= 9.99:
      return "9.00UP"
    if 10.00 <= carat <= 10.99:
      return "10.00UP"

    return ""

  except (ValueError, TypeError):
    return ""


def process_black(value):
  if pd.isna(value) or str(value).strip() == "":
    return "NO BLK"

  value = str(value).strip().upper()

  if value.startswith("FN"):
    return "CENTER - L"

  if value.startswith("N1"):
    return "CENTER - L"

  if value.startswith("N2"):
    return "CENTER - H"

  if value.startswith("N3"):
    return "CENTER - H"

  if value.endswith("BT1"):
    return "CENTER - L"

  if value.endswith("BT2"):
    return "CENTER - H"

  if value.endswith("BT3"):
    return "CENTER - H"

  mappings = {
    "N": "NO BLK",
    "BC0 - BT0": "NO BLK",
    "BC1 - BT0": "SIDE - L",
    "BC2 - BT0": "SIDE - H",
    "BC3 - BT0": "SIDE - H",
    "NONE": "NO BLK",
    "BT0": "NO BLK",
    "BT1": "CENTER - L",
    "B1": "CENTER - L",
    "B2": "CENTER - H",
    "BT2": "CENTER - H",
    "B3": "CENTER - H",
    "BPP": "CENTER - L",
    "NA":"NO BLK",
    "NV": "NO BLK",
    "BT0": "NO BLK",
    "BT1": "CENTER - L",
    "BT2": "CENTER - H",
  }

  return mappings.get(value, value)


def process_black_combined(bic, bis, tb, sb,bt,bc,tabbl,sidbl,blin,natt,incl):
  bic_value = "" if pd.isna(bic) else str(bic).strip().upper()
  bis_value = "" if pd.isna(bis) else str(bis).strip().upper()
  tb_value = "" if pd.isna(tb) else str(tb).strip().upper()
  sb_value = "" if pd.isna(sb) else str(sb).strip().upper()
  bt_value = "" if pd.isna(bt) else str(bt).strip().upper()
  bc_value = "" if pd.isna(bc) else str(bc).strip().upper()
  tabbl_value = "" if pd.isna(tabbl) else str(tabbl).strip().upper()
  sidbl_value = "" if pd.isna(sidbl) else str(sidbl).strip().upper()
  blin_value = "" if pd.isna(blin) else str(blin).strip().upper()
  natt_value = "" if pd.isna(natt) else str(natt).strip().upper()
  incl_value = "" if pd.isna(incl) else str(incl).strip().upper()

  # T.B + S.B rules
  if tb_value == "N" and sb_value == "BPP":
    return "SIDE - L"

  # BIC + BIS rules
  if bic_value == "" and bis_value == "B1":
    return "SIDE - L"

  if bic_value == "" and bis_value == "B2":
    return "SIDE - H"

  if bic_value == "" and bis_value == "B3":
    return "SIDE - H"
  if bt_value == "BT1" and bc_value == "BC2":
    return "NO BLK"
  if bt_value == "BT2" and bc_value == "BC1":
    return "SIDE - L"
  if tabbl_value == "MEDIUM" and sidbl_value == "MEDIUM":
    return "CENTER - H"
  if tabbl_value == "MEDIUM" and sidbl_value == "MINOR":
    return "CENTER - H"
  if tabbl_value == "MEDIUM" and sidbl_value == "NONE":
    return "CENTER - H"
  if tabbl_value == "MEDIUM" and sidbl_value == "PINPOINT":
    return "CENTER - H"
  if tabbl_value == "MINOR" and sidbl_value == "MEDIUM":
    return "CENTER - L"
  if tabbl_value == "MINOR" and sidbl_value == "MINOR":
    return "CENTER - L"
  if tabbl_value == "MINOR" and sidbl_value == "NONE":
    return "CENTER - L"
  if tabbl_value == "MINOR" and sidbl_value == "PINPOINT":
    return "CENTER - L"
  if tabbl_value == "PINPOINT" and sidbl_value == "MINOR":
    return "CENTER - L"
  if tabbl_value == "PINPOINT" and sidbl_value == "NONE":
    return "CENTER - L"
  if tabbl_value == "PINPOINT" and sidbl_value == "PINPOINT":
    return "CENTER - L"
  if tabbl_value == "NONE" and sidbl_value == "MEDIUM":
    return "SIDE - H"
  if tabbl_value == "NONE" and sidbl_value == "MINOR":
    return "SIDE - L"
  if tabbl_value == "NONE" and sidbl_value == "PINPOINT":
    return "SIDE - L"
  if tabbl_value == "NONE" and sidbl_value == "NONE":
    return "NO BLK"
  if bic_value:
    return process_black(bic_value)

  if tb_value:
    return process_black(tb_value)
  if blin_value:
    return process_black(blin_value)
  if natt_value:
    return process_black(natt_value)
  if incl_value:
    return process_black(incl_value)
  return ""

# Core output columns that identify a real stock file. If a converted file
# has ANY one of these columns entirely blank, the file is treated as not a
# valid stock file: it is skipped, and the remaining files are still processed
# and renamed.
REQUIRED_STOCK_COLUMNS = [
  "SHAPE", "RAPSIZE", "CERT", "CARAT", "COLOR", "CLARITY",
  "POL", "SYM", "FLUOR", "STOCKID", "LOTNO",
  "DISC", "DEPTH", "TABLE",
]

def missing_stock_columns(converted_df):
  """Return the required columns that are entirely blank in this file.

  A column counts as blank when every value in it is empty/NaN/None. If the
  returned list is non-empty, the file should be skipped.

  Note: "none" is intentionally NOT treated as missing here. Several required
  columns (e.g. FLUOR, CULET) legitimately produce the value NONE for real
  diamonds, so a column full of NONE is valid data, not a blank column.
  """
  if converted_df is None or converted_df.empty:
    return list(REQUIRED_STOCK_COLUMNS)

  missing = []
  for column in REQUIRED_STOCK_COLUMNS:
    if column not in converted_df.columns:
      missing.append(column)
      continue

    has_value = False
    for value in converted_df[column]:
      if value is None:
        continue
      if isinstance(value, float) and pd.isna(value):
        continue
      if str(value).strip().lower() not in {"", "nan"}:
        has_value = True
        break

    if not has_value:
      missing.append(column)

  return missing

def convert_dataframe(df):
  df = df.copy()
  if all(col in df.columns for col in ["Max", "Min", "Hgt"]):
    df["Measurement"] = df.apply(
      lambda row: f"{row['Max']} - {row['Min']} X {row['Hgt']}",
      axis=1,
    )

  if all(col in df.columns for col in ["Length", "Width", "Depth"]):
    df["Measurement"] = df.apply(
      lambda row: f"{row['Length']} - {row['Width']} X {row['Depth']}",
      axis=1,
    )

  if all(col in df.columns for col in ["M1", "M2", "M3"]):
    df["Measurement"] = df.apply(
      lambda row: f"{row['M1']} - {row['M2']} X {row['M3']}",
      axis=1,
    )

  if all(col in df.columns for col in ["Girdle Thin", "Girdle Thick"]):
    df["GIRDLE"] = df.apply(
      lambda row: f"{row['Girdle Thin']} TO {row['Girdle Thick']}",
      axis=1,
    )

  columns_mapping = {
    "Name":['Name'],
    "STOCKID": ['Report No','Stone ID','StockID','Vendor Stock number','Vendor','Stock #', "Stock No", "Stone No"],
    'LOTNO': ['Diamond ID','CERTNO','Stone Id','Rep No', 'Cert#', 'Report No', 'Cert No', 'Certificate No', 'Report No.',
                      'Packet No', 'CertNo', 'Cert No.', 'REPORT NO.', 'Report #', 'Control No', 'ReportNo',
                      'Report', 'Certificate', 'Cert.No.', 'Report#', 'Certi No.', 'Diamond ID', 'Stone ID','Stock No', 'Cert.No'],

    "SHAPE": ["Shape", "Shp"],
    "RAPSIZE": [],
    "CERT": ['Lab', 'LAB', 'Cert', 'Cert Lab','Certi', 'Cert.'],
    "CARAT": ['Ct.','Cts', 'Wgt', 'CRTWT', 'Weight', 'Carats', 'Size', 'Carat', 'SIZE', 'CTS', 'Ct','WEIGHT', 'Cts.','Carat Weight'],
    "COLOR": ['Col.','Color', 'COL', 'Col', 'COLOR'],
    "CLARITY": ['Purity', 'Cla', 'CLR', 'Clar', 'Clarity', 'CLARITY', 'Clr', 'Cts','WEIGHT', 'CLAR', 'Cla.'],
    "CUT": ['Cut ','Cut', 'CUT', 'Final Cut', 'CutGrade','Cut '],
    "POL": ['Pol.',' Pol','Polish', 'POL', 'Pol'],
    "SYM": ['Symm', 'SYM', 'Sym', 'Sym.', 'Symmetry', 'SYMM', 'Sy'],
    "FLUOR": ['Fluorescence','Fluorescent','Fl.','Fls', 'Flur', 'Fluor', 'Fluor.', 'Flour', 'Fluo', 'Flo', 'FLR', 'FLUOR',
                 'FLUO', 'Fluorescence Intensity', 'Fluo Int', 'Flor.', 'FL', 'Fluorescence', 'Fl.Int.',
                 'Flour.', 'Flr', 'Fluorescence'],

    "MEASUERMENT": ["Measurements", "Measurement",'Diam.'],
    "RAPPRICE": ['New Rap',"Rap Rate", "Rap","Rap $/CT",'RapList','Rapa'],
    "RAP V": ["Rap V", "Rap $",'Rap','New Total Rap','Rate','Rap Value'],
    "DISC":  ['Discount %','Disc. (%)','Rapnet Back',' Disc %','Disc %', 'Rap %', 'RAP DIS', 'DISC.', 'Disc', 'Net %',
                     'Rap%', 'DISC', 'Dis %', 'BACK', 'Rapnet  Discount %', 'Disc(%)', 'Disc%', 'Dis', 'Back%','Base Off %',
                     'Dis%', 'Disc. ', 'Final Disc%', 'Online Back', '%Rap', '% Rap', 'RAP DISC%', 'Rap.%','Cash: %Rap','% Discount'],
    "PRICE": ['Price','Rate','New $/ct',"$/ct", "Pr/Ct","Base Rate",'Price/Ct','Cash: $/ct','Buy PPC'],
    "VALUE": ['New Total',"Total", "Amt $","Amount",'value','Amount','Cash: Total','Buy Price'],
    "CERTNO": ['Report #',"Rep No",'Report No','Cert.No.'],
    'DEPTH': ['Depth %', 'DPL', 'Depth%', 'TD%', 'T Depth', 'DEPTH %', 'Depth Percentage', 'Depth(%)',
                      'Depth', 'DEPTH', 'T.Dep.', 'T.Dpt','TD', 'T.D.%','Total Depth'],
    'TABLE': ['TB%','Table %', 'TBL', 'Table%', 'Tab%', 'TA%', 'T.Size', 'TABLE %', 'Table', 'Table(%)', 'Tabl',
                      'TAB%', 'Tbl'],
    "RATIO": ["Ratio","D/R",'L:W'],
    "BLACK": [],
    "TINCH": ['Shade','shade',"Shd.","Col Tinge", "Shade", "Color Shade", "CS"],
    "LUSTER": ['Luster','milky',"Lust","Lus", "MILKY"],
    "CA": ["Crown Angle","C.A","CA", "Crown", "CR Ang",'Cr.Angle'],
    "CH": ["Crown Height","C.H","CH", "Crown", "CR Hgt",'Cr.Height'],
    "PA": ['Pavilion Depth',"Pav Angle","P.A","PA", "Pavilion", "PV ANG"],
    "PH": ['PH','Pavilion Angle',"Pav Height","P.H","PD", "Pavilion", "PV HGT"],
    "GIRDLE": ['Girdle','Girdle Thickness','GIRDLE',"Girdle",'Girdle Desc','Girdle'],
    "GIRDLE%": ['Girdle Size','Girdle %',"Girdle%", "GT%", "Girdle"],
    "CULET": ["Culet", "CULET","Culet Size", "Cul"],
    "KEY TO SYMBOL": ["Key To Symbols", "Key To Symbol", "K T S",'Key to symbols'],
    "REPORT COMMENT": ["Rep Com",'Cert comment','Lab Comment', "GIA Comments", "Comments",'Report Comments'],
    "Member Comment":['Member Comment'],
    "STATUS": ["Status"],
    "SECTION NO": [],
    "ROUGH ORIGIN": [],
    "DIGITAL CERT": ["Certificate Link", "Cert Link", "Digital Cert"],
    "IMAGES/VIDEO": ["Video Link", "Image Link", "Images/Video"],
    "V360": [],
    "STONE LOCATION": ["Location", "Loc",'City'],
    "TRACR HYPERLINK": [],
    "DETAILS LINK": ["Details Link"],
  }
  output = pd.DataFrame(index=df.index)
  for column, source_columns in columns_mapping.items():
    if column == "BLACK":
      continue
    output[column] = get_source_series(df, source_columns)
    if column in ["CUT", "POL", "SYM"]:
      output[column] = output[column].apply(process_cut_polish_symmetry)
    elif column == "FLUOR":
      output[column] = output[column].apply(process_fluor)
    elif column == "MEASUERMENT":
      output[column] = output[column].apply(process_measurement)
    elif column == "SHAPE":
      output[column] = output[column].apply(process_shape)
    elif column in ["KEY TO SYMBOL", "REPORT COMMENT"]:
      output[column] = output[column].apply(replace_comma_with_hyphen)
    elif column in ["GIRDLE%"]:
      output[column] = output[column].apply(only_number)
    elif column in ["CA", "PA","CH", "PH"]:
      output[column] = output[column].apply(only_numeric_value)
    elif column == "STONE LOCATION":
      output[column] = output[column].apply(process_location)
    elif column == "TINCH":
      output[column] = output[column].apply(process_tinch)
    elif column == "LUSTER":
      output[column] = output[column].apply(process_luster)
    elif column == "CULET":
      output[column] = output[column].apply(process_culet)
    elif column == "GIRDLE":
      output[column] = output[column].apply(girdle_process)
    elif column == "DISC":
      output[column] = output[column].apply(process_discount)
    elif column in ["CA", "PA"]:
      output[column] = output[column].apply(lambda x: get_ca_ph(x, 0))
    elif column in ["CH", "PH"]:
      output[column] = output[column].apply(lambda x: get_ca_ph(x, 1))

  output["RAPSIZE"] = output["CARAT"].apply(process_rapsize)

  black_series = df.apply(
    lambda row: process_black_combined(
      row.get("BIC", ""),
      row.get("BIS", ""),
      row.get("T.B", ""),
      row.get("S.B", ""),
      row.get("BT", ""),
      row.get("BC", ""),
      row.get("Table Black", ""),
      row.get("Side Black", ""),
      row.get("Black Inclusion  ", ""),
      row.get("Natts", ""),
      row.get("Inclusion", ""),
    ),
    axis=1,
  )
  output["BLACK"] = black_series.reindex(output.index)

  output = output[list(columns_mapping.keys())]

  return output.reset_index(drop=True)

# Upload modes: independently uploaded vendor files, or chosen sheets in one workbook.
EXCEL_SHEET_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods", ".odf"}


def workbook_sheet_names(filename, raw_bytes):
  extension = os.path.splitext(filename)[1].lower()
  engines = {".xlsx": "openpyxl", ".xlsm": "openpyxl",
             ".xls": "xlrd", ".xlsb": "pyxlsb",
             ".ods": "odf", ".odf": "odf"}
  engine = engines[extension]
  try:
    with pd.ExcelFile(BytesIO(raw_bytes), engine=engine) as excel:
      return excel.sheet_names, raw_bytes
  except Exception as exc:
    if extension not in {".xlsx", ".xlsm"} or not is_stylesheet_error(exc):
      raise
    fixed = remove_xlsx_styles(raw_bytes).getvalue()
    with pd.ExcelFile(BytesIO(fixed), engine=engine) as excel:
      return excel.sheet_names, fixed


def load_chosen_sheets(uploaded_file, chosen_sheets):
  raw = uploaded_file.getvalue()
  extension = os.path.splitext(uploaded_file.name)[1].lower()
  engines = {".xlsx": "openpyxl", ".xlsm": "openpyxl",
             ".xls": "xlrd", ".xlsb": "pyxlsb",
             ".ods": "odf", ".odf": "odf"}
  names, valid_bytes = workbook_sheet_names(uploaded_file.name, raw)
  loaded = []
  for sheet_name in chosen_sheets:
    if sheet_name not in names:
      raise ValueError(f"Unknown sheet: {sheet_name}")
    if extension in OPENXML_EXCEL_EXTENSIONS:
      df, _ = read_openxml_excel_safely(valid_bytes, sheet_name=sheet_name)
    else:
      df = normalize_dataframe_columns(
        read_excel_with_header_detection(valid_bytes, engines[extension], sheet_name))
    loaded.append((sheet_name, df))
  return loaded


def convert_named_sources(named_sources):
  frames = []
  skipped = []
  for name, frame in named_sources:
    frame = frame.copy()
    frame["Name"] = process_name(frame, name)
    converted = convert_dataframe(frame)

    missing = missing_stock_columns(converted)
    if missing:
      skipped.append((name, missing))
      continue

    frames.append((name, converted))

  if skipped:
    lines = "\n".join(
      f"- {name}  (blank: {', '.join(cols)})" for name, cols in skipped
    )
    st.warning(
      f"Skipped {len(skipped)} file(s)/sheet(s) because a required column "
      f"was blank:\n{lines}"
    )

  return frames

def make_excel_download(converted_sheets):
  """Single combined sheet with duplicate STOCKIDs removed."""
  output = BytesIO()
  combined = [frame for _, frame in converted_sheets]
  final = pd.concat(combined, ignore_index=True) if combined else pd.DataFrame()
  if "STOCKID" in final.columns:
    # Consistent with the original converter's duplicate-removal rule.
    final = final.drop_duplicates(subset=["STOCKID"], keep="first")
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    final.to_excel(writer, index=False, sheet_name="Combined Stock")
    from openpyxl.styles import PatternFill, Font
    for sheet in writer.book.worksheets:
      sheet.auto_filter.ref = sheet.dimensions
      for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(bold=True, color="FFFFFF")
      for column_cells in sheet.columns:
        letter = column_cells[0].column_letter
        max_length = 0
        for cell in column_cells:
          if cell.value is not None:
            length = len(str(cell.value))
            if length > max_length:
              max_length = length
        sheet.column_dimensions[letter].width = min(max(max_length + 2, 10), 50)
  return output.getvalue(), final


def prepare_preview_dataframe(frame):
  """Make mixed-type stock columns Arrow-compatible for Streamlit preview.

  Keep the original converted DataFrame unchanged for Excel downloads.
  """
  preview = frame.copy()
  for column in preview.columns:
    if preview[column].dtype == "object":
      preview[column] = preview[column].map(
        lambda value: "" if value is None or (not isinstance(value, (list, dict, tuple)) and pd.isna(value)) else str(value)
      ).astype("string")
  return preview

def _find_sample(filename: str):
    """Locate a file inside a 'Sample Files' folder, searching upward from this
    script. Works on Linux (Streamlit Cloud) and Windows regardless of the
    current working directory or where this page sits in a multipage app."""
    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        candidate = parent / "Sample Files" / filename
        if candidate.exists():
            return candidate
    return None

def run_convert_stock_mod():
  st.title("Convert Stock File")
  with st.sidebar:
    mode = st.radio("Upload option", ["Files", "Sheets"], horizontal=True)

    converted = []
    if mode == "Files":
      uploaded_files = select_files()
      if not uploaded_files:

        return
      try:
        with st.spinner("Converting files..."):
          loaded = load_excel_files(uploaded_files)
          converted = convert_named_sources(loaded)
      except Exception as exc:
        st.error(f"Unable to convert files: {exc}")
        return
    else:
      uploaded_file = st.file_uploader(
        "Select Excel workbook", type=["xlsx", "xlsm", "xls", "xlsb", "ods", "odf"],
        accept_multiple_files=False, key="sheet_workbook")
      if uploaded_file is None:
        return
      try:
        with st.spinner("Reading workbook..."):
          sheet_names, _ = workbook_sheet_names(uploaded_file.name, uploaded_file.getvalue())
      except Exception as exc:
        st.error(f"Unable to read workbook sheet names: {exc}")
        return
      chosen = st.multiselect("Sheets to convert", sheet_names, default=sheet_names)
      if not chosen:
        st.info("Select at least one sheet.")
        return
      try:
        with st.spinner("Converting selected sheets..."):
          loaded = load_chosen_sheets(uploaded_file, chosen)
          converted = convert_named_sources(loaded)
      except Exception as exc:
        st.error(f"Unable to convert sheets: {exc}")
        return

  if not converted:
    st.warning("No data was processed.")
    return
  with st.spinner("Building Excel file..."):
    output_bytes, combined = make_excel_download(converted)
 
  st.dataframe(prepare_preview_dataframe(combined), width="stretch")

  with st.sidebar:
    st.download_button("Download",
                      data=output_bytes,
                      file_name=f"Converted Stock {dt.datetime.now():%d-%m-%Y}.xlsx",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
  run_convert_stock_mod()