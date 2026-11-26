from __future__ import annotations

import argparse
import json
import os
import re
import time
from io import BytesIO
from typing import Any, Dict, Iterable, List
from urllib.parse import quote_plus

import pandas as pd
import requests
import pdfplumber
import undetected_chromedriver as uc

from selenium.common.exceptions import SessionNotCreatedException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

try:
    from tkinter.filedialog import askopenfilename
except Exception:
    askopenfilename = None

DEFAULT_OUTPUT = fr"D:\Gaurav\Download\Certificate_{pd.Timestamp.now():%S}.xlsx"
REPORT_COLUMN_NAME = "Report Number"

IGI_VERIFY_URL_TEMPLATE = "https://www.igi.org/Verify-Your-Report/?r={report_no}"
IGI_DIAGNOSIS_URL_TEMPLATE = "https://www.igi.org/API-IGI/report-diagnosis.php?r={report_no}"


OUTPUT_COLUMNS = [
    "Input Report Number",
    "IGI Report Number",
    "Report Date",
    "Shape and Cut",
    "Carat Weight",
    "Color Grade",
    "Clarity Grade",
    "Cut Grade",
    "Polish",
    "Symmetry",
    "Measurements",
    "Table Size",
    "Crn Ag",
    "Crn Ht",
    "Pav Ag",
    "Pav Dp",
    "Girdle Thickness",
    "Culet",
    "Total Depth",
    "Fluorescence",
    "Inscription(s)",
    "Identification Features",
    "Comments"
]


LABEL_VARIANTS: Dict[str, List[str]] = {
    "IGI Report Number": [
        "REPORT NUMBER",
        "Report Number",
        "Report No",
        "Report No.",
        "Report #",
        "Certificate Number",
        "Certificate No",
    ],
    "Report Date": ["REPORT DATE", "Issue Date", "Date of Issue", "Report Date"],
    "Description": ["DESCRIPTION", "Description"],
    "Shape and Cut": [
        "SHAPE AND CUT",
        "Shape and Cut",
        "Shape and Cutting Style",
        "Shape & Cut",
        "Shape/cut",
        "Shape Cut",
    ],
    "Carat Weight": ["CARAT WEIGHT", "Carat Weight", "Weight"],
    "Color Grade": ["COLOR GRADE", "Color Grade", "Color"],
    "Clarity Grade": ["CLARITY GRADE", "Clarity Grade", "Clarity"],
    "Cut Grade": ["CUT GRADE", "Cut Grade", "Cut"],
    "Polish": ["POLISH", "Polish"],
    "Symmetry": ["SYMMETRY", "Symmetry"],
    "Measurements": ["Measurements", "MEASUREMENTS", "Measurement"],
    "Table Size": ["Table Size", "TABLE SIZE", "Table", "TABLE"],
    "Crown Height": ["Crown Height", "CROWN HEIGHT", "Crown Height - Angle"],
    "Pavilion Depth": [
        "Pavilion Depth",
        "PAVILION DEPTH",
        "Pavillion Depth",
        "PAVILLION DEPTH",
        "Pavilion Depth - Angle",
        "Pavillion Depth - Angle",
    ],
    "Girdle Thickness": ["Girdle Thickness", "GIRDLE THICKNESS", "Girdle"],
    "Culet": ["Culet", "CULET"],
    "Total Depth": ["Total Depth", "TOTAL DEPTH", "Depth"],
    "Fluorescence": ["FLUORESCENCE", "Fluorescence"],
    "Inscription(s)": [
        "Inscription(s)",
        "INSCRIPTION(S)",
        "Inscriptions",
        "Inscription",
        "Laser Inscription",
    ],
    "Identification Features": [
        "Identification Features",
        "IDENTIFICATION FEATURES",
        "Key To Symbols",
        "KEY TO SYMBOLS",
    ],
    "Comments": ["COMMENTS", "Comments", "Comment"],
}


_ORIGINAL_UC_CHROME_DEL = uc.Chrome.__del__


def _quiet_uc_chrome_del(self) -> None:
    try:
        _ORIGINAL_UC_CHROME_DEL(self)
    except OSError as exc:
        if getattr(exc, "winerror", None) != 6:
            raise
    except Exception:
        pass


uc.Chrome.__del__ = _quiet_uc_chrome_del


def safe_quit_driver(driver) -> None:
    if driver is None:
        return
    try:
        driver.quit()
    except OSError as exc:
        if getattr(exc, "winerror", None) != 6:
            raise
    except Exception:
        pass


def clean_report_number(value) -> str:
    if pd.isna(value):
        return ""

    report_no = str(value).strip()

    if report_no.endswith(".0"):
        report_no = report_no[:-2]

    return re.sub(r"\s+", "", report_no)


def read_report_numbers(path: str, column: str) -> List[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")

    df = pd.read_excel(path, dtype=str)

    if column not in df.columns:
        raise ValueError(
            f"Column '{column}' not found. Available columns: {list(df.columns)}"
        )

    numbers = df[column].apply(clean_report_number).tolist()
    numbers = [n for n in numbers if n]

    if not numbers:
        raise ValueError(f"No report numbers found in column '{column}'.")

    return list(dict.fromkeys(numbers))


def create_driver(headless: bool = False) -> uc.Chrome:
    options = uc.ChromeOptions()

    if headless:
        options.add_argument("--headless=new")

    options.add_argument("--start-maximized")
    options.add_argument("--disable-features=IsolateOrigins,site-per-process")
    options.add_argument("--disable-blink-features=AutomationControlled")

    try:
        return uc.Chrome(options=options)
    except SessionNotCreatedException as exc:
        message = str(exc)
        match = re.search(r"Current browser version is (\d+)\.", message)

        if match:
            version_main = int(match.group(1))
            retry_options = uc.ChromeOptions()

            if headless:
                retry_options.add_argument("--headless=new")

            retry_options.add_argument("--start-maximized")
            retry_options.add_argument("--disable-features=IsolateOrigins,site-per-process")
            retry_options.add_argument("--disable-blink-features=AutomationControlled")

            return uc.Chrome(options=retry_options, version_main=version_main)

        raise


def wait_for_page_body(driver: uc.Chrome, timeout: int = 30) -> bool:
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        return True
    except TimeoutException:
        return False


def body_text(driver: uc.Chrome) -> str:
    try:
        return driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""


def normalize_text(text: str) -> str:
    text = (text or "").replace("\xa0", " ")
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def clean_value(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, float) and pd.isna(value):
        return ""

    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t\f\v]+", " ", text)
    text = re.sub(r"\s*\n+\s*", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip(" :|/")

def canonical_label(label: str) -> str:
    cleaned = re.sub(r"\s+", " ", str(label or "").strip()).lower()

    for canonical, variants in LABEL_VARIANTS.items():
        for variant in variants:
            if cleaned == re.sub(r"\s+", " ", variant.strip()).lower():
                return canonical

    return re.sub(r"\s+", " ", str(label or "").strip())


def add_unique_key(data: Dict[str, str], label: str, value: Any) -> None:
    key = canonical_label(label)
    val = clean_value(value)

    if not key or not val:
        return

    if key in OUTPUT_COLUMNS and data.get(key):
        return

    original_key = key
    counter = 2

    while key in data and data[key] != val:
        key = f"{original_key} {counter}"
        counter += 1

    data[key] = val


def split_percent_angle(value: Any) -> tuple[str, str]:
    text = clean_value(value)

    if not text:
        return "", ""

    percent_match = re.search(r"\d+(?:\.\d+)?\s*%", text)
    angle_match = re.search(r"\d+(?:\.\d+)?\s*(?:°|deg(?:ree)?s?)", text, flags=re.I)

    percent = percent_match.group(0).replace(" ", "") if percent_match else ""
    angle = angle_match.group(0).replace(" ", "") if angle_match else ""

    if not angle and "-" in text:
        parts = [part.strip() for part in re.split(r"\s*-\s*", text, maxsplit=1)]
        if len(parts) == 2:
            possible_angle = parts[1].strip()

            if re.fullmatch(r"\d+(?:\.\d+)?", possible_angle):
                angle = f"{possible_angle}°"
            elif possible_angle:
                angle = possible_angle.replace(" ", "")

    return percent, angle


def apply_crown_pavilion_result(result: Dict[str, str]) -> None:
    crown_height, crown_angle = split_percent_angle(result.get("Crown Height", ""))
    pavilion_depth, pavilion_angle = split_percent_angle(result.get("Pavilion Depth", ""))

    if not result.get("Crn Ht"):
        result["Crn Ht"] = crown_height

    if not result.get("Crn Ag"):
        result["Crn Ag"] = crown_angle

    if not result.get("Pav Dp"):
        result["Pav Dp"] = pavilion_depth

    if not result.get("Pav Ag"):
        result["Pav Ag"] = pavilion_angle


def repair_json_string_literals(candidate: str) -> str:
    out: List[str] = []
    in_string = False
    escape = False

    for ch in candidate:
        if in_string:
            if escape:
                out.append(ch)
                escape = False
            elif ch == "\\":
                out.append(ch)
                escape = True
            elif ch == '"':
                out.append(ch)
                in_string = False
            elif ch in "\r\n":
                out.append("\\n")
            else:
                out.append(ch)
        else:
            out.append(ch)
            if ch == '"':
                in_string = True

    return "".join(out)


def iter_json_array_candidates(text: str) -> Iterable[str]:
    text = text or ""
    start = None
    depth = 0
    in_string = False
    escape = False

    for idx, ch in enumerate(text):
        if start is None:
            if ch == "[":
                start = idx
                depth = 1
                in_string = False
                escape = False
            continue

        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
        else:
            if ch == '"':
                in_string = True
            elif ch == "[":
                depth += 1
            elif ch == "]":
                depth -= 1
                if depth == 0 and start is not None:
                    yield text[start : idx + 1]
                    start = None


def extract_report_json_records(text: str) -> List[Dict[str, Any]]:
    for candidate in iter_json_array_candidates(text):
        fixed = repair_json_string_literals(candidate)

        try:
            decoded = json.loads(fixed)
        except json.JSONDecodeError:
            continue

        if isinstance(decoded, list) and decoded and isinstance(decoded[0], dict):
            return decoded

    return []


def extract_pdf_url(driver: uc.Chrome, page_text: str, report_no: str) -> str:
    try:
        anchors = driver.find_elements(By.XPATH, "//a[@href]")

        for anchor in anchors:
            href = (anchor.get_attribute("href") or "").strip()

            if not href:
                continue

            href_lower = href.lower()

            if "pdf.igi.org" in href_lower or href_lower.endswith(".pdf"):
                return href
    except Exception:
        pass

    match = re.search(r"https?://[^\s]+\.pdf", page_text or "", flags=re.I)

    if match:
        return match.group(0).strip()

    if report_no:
        return f"https://pdf.igi.org/ID{quote_plus(report_no)}.pdf"

    return ""


def normalize_report_record(record: Dict[str, Any]) -> Dict[str, str]:
    data: Dict[str, str] = {}

    for key, value in record.items():
        add_unique_key(data, key, value)

    return data


def extract_text_from_pdf_url(pdf_url: str) -> str:
    try:
        response = requests.get(pdf_url, timeout=30)
        response.raise_for_status()

        with pdfplumber.open(BytesIO(response.content)) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)

        return text
    except Exception:
        return ""


def clean_identification_features(value: str) -> str:
    """Return only the actual IGI Identification Features value.

    IGI PDF/API text sometimes includes the color/clarity scale, report type,
    measurements, or OCR-garbled words after the feature list. This function
    removes that noise and keeps only recognized identification feature terms.
    """
    text = clean_value(value).upper()

    if not text:
        return ""

    # Stop before common IGI text that appears after the real feature value.
    stop_patterns = [
        r"\bCOLORLESS\b",
        r"\bNEAR\s+COLORLESS\b",
        r"\bSLIGHTLY\s+TINTED\b",
        r"\bVERY\s+LIGHT\b",
        r"\bLIGHT\s+COLOR\b",
        r"\bCARAT\s+WEIGHT\b",
        r"\b\d+(?:\.\d+)?\s*CARAT\b",
        r"\bNATURAL\s+DIAMOND\b",
        r"\bLABORATORY\s+GROWN\s+DIAMOND\b",
        r"\bLAB\s+GROWN\s+DIAMOND\b",
        r"\bROUND\s+BRILLIANT\b",
        r"\bDIAMOND\s+REPORT\b",
        r"\bIGI\s+REPORT\s+NUMBER\b",
        r"\bREPORT\s+NUMBER\b",
        r"\bINSCRIPTION\b",
        r"\bMEASUREMENTS\b",
        r"\b\d+(?:\.\d+)?\s*-\s*\d+(?:\.\d+)?\s*X\s*\d+(?:\.\d+)?\s*MM\b",
    ]

    for pattern in stop_patterns:
        match = re.search(pattern, text, flags=re.I)
        if match:
            text = text[: match.start()]

    # Normalize the known OCR-garbled value seen in some IGI PDFs.
    text = re.sub(r"INDE\s*F\s*N\s*EA\s*TE\s*TH\s*D\s*E\s*N\s*R\s*A", "INDENTED NATURAL, FEATHER", text)
    text = re.sub(r"T\s*C\s*U\s*A\s*R\s*V\s*A\s*IT\s*L\s*Y", "CAVITY", text)

    # Keep recognized feature names in the order they appear.
    known_features = [
        "INDENTED NATURAL",
        "NATURAL",
        "FEATHER",
        "CRYSTAL",
        "CLOUD",
        "CAVITY",
        "NEEDLE",
        "PINPOINT",
        "CHIP",
        "KNOT",
        "TWINNING WISP",
        "SURFACE GRAINING",
        "INTERNAL GRAINING",
        "EXTRA FACET",
        "BRUISE",
        "BEARDING",
    ]

    found: List[tuple[int, str]] = []
    for feature in known_features:
        for match in re.finditer(r"\b" + re.escape(feature) + r"\b", text, flags=re.I):
            found.append((match.start(), feature))

    if found:
        found.sort(key=lambda item: item[0])
        result: List[str] = []
        for _, feature in found:
            # Do not output plain NATURAL when it is part of INDENTED NATURAL.
            if feature == "NATURAL" and result and result[-1] == "INDENTED NATURAL":
                continue
            if feature not in result:
                result.append(feature)
        return ", ".join(result)

    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s*,\s*", ", ", text)
    return text.strip(" ,:|/")


def extract_identification_features_from_pdf_text(pdf_text: str) -> str:
    """Extract exact Identification Features from IGI PDF text.

    The correct value is usually between "Identification Features" and
    "Inscription(s)". A final cleaner is applied as a safety net for PDFs whose
    extracted text order is imperfect.
    """
    if not pdf_text:
        return ""

    text = pdf_text.replace("\xa0", " ")

    # Best case: capture the text directly between the two printed labels.
    match = re.search(
        r"Identification\s+Features\s*(.*?)\s*Inscription\s*\(?s?\)?",
        text,
        flags=re.I | re.S,
    )

    if match:
        value = match.group(1)
        value = value.replace("\r", "\n")
        value = re.sub(r"[ \t\f\v]+", " ", value)
        value = re.sub(r"\s*\n\s*", " ", value)
        value = re.sub(r"\s+", " ", value)
        value = value.strip(" ,:|/")
        cleaned = clean_identification_features(value)
        if cleaned:
            return cleaned

    # Fallback: use the first line starting with the label and following lines
    # until the next label.
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for i, line in enumerate(lines):
        if re.match(r"^Identification\s+Features\b", line, flags=re.I):
            value = re.sub(r"^Identification\s+Features\b", "", line, flags=re.I).strip(" :")
            j = i + 1
            while j < len(lines):
                if re.match(r"^(Inscription|DIAMOND REPORT|IGI Report Number|Carat Weight|Color Grade|Clarity Grade)\b", lines[j], flags=re.I):
                    break
                value += " " + lines[j]
                j += 1
            return clean_identification_features(value)

    return ""


def enrich_result_from_pdf(result: Dict[str, str]) -> None:
    pdf_url = result.get("PDF URL", "")

    if not pdf_url:
        return

    pdf_text = extract_text_from_pdf_url(pdf_url)

    if not pdf_text:
        return

    identification_features = extract_identification_features_from_pdf_text(pdf_text)

    if identification_features:
        result["Identification Features"] = identification_features


def fetch_report(report_no: str, driver: uc.Chrome) -> Dict[str, str]:
    result: Dict[str, str] = {column: "" for column in OUTPUT_COLUMNS}
    result["Input Report Number"] = report_no
    result["Report URL"] = IGI_VERIFY_URL_TEMPLATE.format(report_no=quote_plus(report_no))

    diagnosis_url = IGI_DIAGNOSIS_URL_TEMPLATE.format(report_no=quote_plus(report_no))

    try:
        driver.get(diagnosis_url)

        if not wait_for_page_body(driver, timeout=30):
            result["Status"] = "Failed"
            result["Error"] = "Timed out waiting for IGI diagnosis page body."
            return result

        time.sleep(1)

        text = body_text(driver)
        normalized_text = normalize_text(text)

        records = extract_report_json_records(normalized_text)

        if not records:
            result["Status"] = "Failed"
            result["Error"] = "No IGI report JSON data found."
            result["PDF URL"] = extract_pdf_url(driver, normalized_text, report_no)
            enrich_result_from_pdf(result)
            return result

        merged: Dict[str, str] = {}

        for record in records:
            for key, value in normalize_report_record(record).items():
                if key not in merged or not merged[key]:
                    merged[key] = value

        result.update(merged)

        apply_crown_pavilion_result(result)

        result["PDF URL"] = extract_pdf_url(driver, normalized_text, report_no)

        enrich_result_from_pdf(result)

        result["Status"] = "Success"

    except Exception as exc:
        result["Status"] = "Failed"
        result["Error"] = str(exc)

    return result


def order_columns(df: pd.DataFrame) -> pd.DataFrame:
    for column in OUTPUT_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    extra_columns = [column for column in df.columns if column not in OUTPUT_COLUMNS]

    return df[OUTPUT_COLUMNS + extra_columns]


def save_results(results: List[Dict[str, str]], output_path: str) -> None:
    df_new = pd.DataFrame(results)
    df_new.fillna("", inplace=True)

    # Keep only required output columns
    for column in OUTPUT_COLUMNS:
        if column not in df_new.columns:
            df_new[column] = ""

    df_new = df_new[OUTPUT_COLUMNS]

    df_new.to_excel(output_path, index=False)

def choose_input_file() -> str:
    if askopenfilename is None:
        return ""

    return askopenfilename(
        title="Select the Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape IGI report data from API and PDF."
    )

    parser.add_argument(
        "--input",
        "-i",
        default="",
        help="Path to Excel file containing report numbers.",
    )

    parser.add_argument(
        "--output",
        "-o",
        default=DEFAULT_OUTPUT,
        help=f"Output Excel file path. Default: {DEFAULT_OUTPUT}",
    )

    parser.add_argument(
        "--column",
        "-c",
        default=REPORT_COLUMN_NAME,
        help=f"Column containing report numbers. Default: {REPORT_COLUMN_NAME}",
    )

    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run Chrome in headless mode.",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    input_path = args.input or choose_input_file()
    output_path = args.output

    if not input_path:
        print("No input Excel file selected.")
        return

    try:
        report_numbers = read_report_numbers(input_path, args.column)
    except Exception as exc:
        print(f"Error reading report numbers: {exc}")
        return

    print(f"Loaded {len(report_numbers)} report numbers from {input_path}")

    driver = create_driver(headless=args.headless)

    results: List[Dict[str, str]] = []

    try:
        for idx, report_no in enumerate(report_numbers, start=1):
            print(f"[{idx}/{len(report_numbers)}] Fetching IGI report {report_no}...")

            result = fetch_report(report_no, driver)
            results.append(result)

            status = result.get("Status", "")
            error = result.get("Error", "")

            print(f"    {status}: {report_no}" + (f" - {error}" if error else ""))

            save_results(results, output_path)

            time.sleep(1)

    finally:
        safe_quit_driver(driver)

    save_results(results, output_path)

    print(f"Results saved to {output_path}")

    try:
        os.startfile(output_path)
    except Exception:
        pass


if __name__ == "__main__":
    main()