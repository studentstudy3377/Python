from datetime import datetime
import re
import os
import pandas as pd
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    return text

def extract_key_to_symbols_and_identification(text, flags=re.IGNORECASE):
    key_to_symbols_text = ""

    # Extract paragraph under "KEY TO SYMBOLS"
    match = re.search(r"KEY TO SYMBOLS\s+([\s\S]+?)(?=\n\s*\n|ADDITIONAL GRADING INFORMATION)", text, flags)
    if match:
        key_to_symbols_text = match.group(1).strip().replace("\n", " ")

    # Also get "Identification Features"
    ident = re.search(r"Identification Features\s*([A-Z ,]+)", text, flags)
    if ident:
        idents = ident.group(1).strip()
        if key_to_symbols_text:
            key_to_symbols_text += " | Identification Features: " + idents
        else:
            key_to_symbols_text = idents

    return key_to_symbols_text

def extract_culet_type(text):
    """
    Extracts the Culet Type by searching for known keywords and returning the last occurrence.
    This helps avoid picking up earlier matches (like 'None' from the Fluorescence field) that appear before
    the actual culet type in the document.
    """
    matches = re.findall(r"\b(Pointed|None|Very Small|Small|Medium|Large|Very Large|Extremely Large)\b", text, re.IGNORECASE)
    if matches:
        return matches[-1].strip().title()
    return ""

def assign_sorted_percentage_values(percents):
    """
    Takes a list of percentage strings (e.g. ["13.5%", "41.5%", "59.2%", "57%"]),
    sorts them in descending order (by numeric value) and returns a dictionary
    mapping the sorted values to the keys: Depth %, Table %, Pav Dp, Crn Dp.
    The returned values are numeric strings without the '%' symbol.
    """
    try:
        # Remove the '%' and convert to float
        float_values = [(float(p.replace("%", "")), p) for p in percents]
        # Sort by numeric value in descending order
        sorted_values = sorted(float_values, key=lambda x: x[0], reverse=True)
        
        # Format values: if the float is whole, remove decimal part; otherwise, keep as is.
        def format_val(val):
            return str(int(val)) if val.is_integer() else str(val)
        
        sorted_percents = [format_val(x[0]) for x in sorted_values]
        keys = ["Depth %", "Table %", "Pav Dp", "Crn Dp"]
        return dict(zip(keys, sorted_percents))
    except Exception as e:
        print(f"Error in assign_sorted_percentage_values: {e}")
        return {key: "" for key in ["Depth %", "Table %", "Pav Dp", "Crn Dp"]}

def extract_value_percentages(text):
    # Normalize the text
    clean_text = re.sub(r"\([^)]*\)", "", text)  # remove any text within parentheses
    clean_text = re.sub(r"\s+", " ", clean_text)

    # Find all percentages in the entire document
    percents = re.findall(r"\d{1,3}\.?\d*%", clean_text)

    result = {}
    if len(percents) >= 4:
        # Use the first four percentages found and sort them descending without '%' symbol
        first_four = percents[:4]
        sorted_columns = assign_sorted_percentage_values(first_four)
        result.update(sorted_columns)
    return result

def extract_girdle(text):
    """
    Extracts the Girdle information by normalizing the text and capturing a phrase
    that consists only of letters and spaces in the format "Word To Word" before "(Faceted)".
    For example, from "582354717 Medium To Thick (Faceted)" it returns "Medium To Thick".
    """
    text_clean = re.sub(r'\s+', ' ', text)
    match = re.search(r"([A-Za-z\s]+To\s+[A-Za-z\s]+)\s*\(Faceted\)", text_clean, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""

def extract_standard_fields(text):
    flags = re.IGNORECASE
    data = {}

    rn = re.search(r"IGI\s+Report\s+Number\s*[:\s]*([\d]+)", text, flags)
    data["IGI Report Number"] = rn.group(1) if rn else ""
    dt = re.search(r"([A-Za-z]+\s+\d{1,2},\s+\d{4})", text)
    data["Date"] = dt.group(1) if dt else ""

    # Modified extraction: remove "BRILLIANT" from the Shape and Cutting Style value.
    shape = re.search(r"Shape\s+and\s+Cutting\s+Style\s+([A-Z ]+)", text, flags)
    if shape:
        shape_val = shape.group(1).strip()
        shape_val = shape_val.replace("BRILLIANT", "").strip()
        data["Shape and Cutting Style"] = shape_val
    else:
        data["Shape and Cutting Style"] = ""

    text_clean = re.sub(r'\s+', ' ', text.strip())
    match_meas = re.search(r'(?i)Measurements\s+([\d\.\sxX-]+m{1,2})(?=\s|$)', text_clean)
    data["Measurements"] = match_meas.group(1).strip() if match_meas else ""

    carat = re.search(r"Carat\s+Weight\s*[:\s]*([\d\.]+)", text, flags)
    data["Carat Weight"] = carat.group(1) if carat else ""
    color = re.search(r"Color\s+Grade\s*[:\s]*([A-Z])", text, flags)
    data["Color Grade"] = color.group(1) if color else ""
    clarity = re.search(r"Clarity\s+Grade\s*[:\s]*([A-Z0-9 ]+)", text, flags)
    data["Clarity Grade"] = clarity.group(1).strip() if clarity else ""
    cut = re.search(r"Cut\s+Grade\s*[:\s]*([A-Z ]+)", text, flags)
    data["Cut Grade"] = cut.group(1).strip() if cut else ""
    polish = re.search(r"Polish\s+([A-Z]+)", text, flags)
    data["Polish"] = polish.group(1).strip() if polish else ""
    symmetry = re.search(r"Symmetry\s+([A-Z]+)", text, flags)
    data["Symmetry"] = symmetry.group(1).strip() if symmetry else ""
    fluoro = re.search(r"Fluorescence\s+([A-Z ]+)", text, flags)
    data["Fluorescence"] = fluoro.group(1).strip() if fluoro else ""

    comments = re.search(r"Comments:\s*(.+)", text, flags)
    data["Comments"] = comments.group(1).strip() if comments else ""
    data["Lab"] = "IGI"
    data["Key To Symbols"] = extract_key_to_symbols_and_identification(text, flags)
    data["Culet Type"] = extract_culet_type(text)
    # New field: Girdle using the updated extraction (only alphabetic words)
    data["Girdle"] = extract_girdle(text)

    return data

def extract_diamond_data(text):
    data = extract_standard_fields(text)
    value_data = extract_value_percentages(text)
    data.update(value_data)
    return data

def select_pdf_files():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF Files", "*.pdf")])
    return list(file_paths)

def select_pdf_folder():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Select folder containing PDF files")
    if not folder:
        print("No folder selected.")
        return []
    pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith(".pdf")]
    return pdf_files

def main():
    pdf_files = select_pdf_folder()

    if not pdf_files:
        print("No files found.")
        return

    all_data = []
    for file_path in pdf_files:
        print(f"Processing {os.path.basename(file_path)} ...")
        text = extract_text_from_pdf(file_path)
        data = extract_diamond_data(text)
        all_data.append(data)

    df = pd.DataFrame(all_data)
    date = datetime.now()
    output_filename = f'G:\\Certificate\\Python Output\\PDF to Excel\\Fancy IGI {date.strftime("%d-%m-%Y")}.xlsx'
    df.to_excel(output_filename, index=False)
    os.startfile(output_filename)

if __name__ == "__main__":
    main()
