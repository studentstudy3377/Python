import os
import re
import pandas as pd
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog

def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main window.
    folder_path = filedialog.askdirectory(title="Select Folder Containing PDF Files")
    return folder_path

def extract_text_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text

def parse_data(text):
    # Normalize line breaks for consistent matching.
    text_clean = re.sub(r'\r\n', '\n', text)
    text_clean = re.sub(r'\n+', '\n', text_clean.strip())
    
    data = {}

    # Extract HRD Antwerp Report Number
    match_hrd = re.search(r"N°\s*(?:\*+)?\s*(\d+)", text_clean, re.IGNORECASE)
    data["Report No"] = match_hrd.group(1) if match_hrd else ""
    
    # Assign Lab directly.
    data["Lab"] = "HRD"
    
    # Define regex patterns for other fields.
    patterns = {
        "Colour Grade": r"^Colour Grade\s+(.+?)(?:\s+grading|\n)",
        "Report Date": r"^(\w+\s\d{2},\s\d{4})",
        "Shape": r"^Shape\s+(\w+)",
        "Carat Weight": r"^Carat \(weight\)\s+([\d\.]+)\s*ct",
        "Fluorescence": r"^Fluorescence\s+(\w+)", 
        "Clarity Grade": r"^Clarity Grade\s+(\w+)",
        "Cut": r"(?:Proportions\s*(?:[:\-]?)\s*(excellent|very good|good|fair|poor))|Cut\s+\(Prop\./Pol\./Symm\.\)\s+(?P<cut_grade>\w+)",
        "Polish": r"^Polish\s+(very good|excellent|good|fair|poor)",
        "Symmetry": r"^Symmetry\s+(very good|excellent|good|fair|poor)",
        "Measurements": r"^Measurements\s+([\d\.\s\-\wx]+mm)",
        "Girdle": r"^Girdle\s+([a-zA-Z\s]+?)\s+\d+\.?\d*\s*%\s*faceted",
        'Girdle %': r"^Girdle\s+\w+\s+([\d\.]+\s*%)",
        "Culet": r"^Culet\s+(\w+)",
        "Depth %": r"^Total Depth\s+([\d\.]+\s*%)",
        "Table %": r"^Table Width\s+([\d\.]+\s*%)",
        "Crown Height": r"^Crown Height\s+\(.*?\)\s+([\d\.]+\s*%)",
        "Crown Angle": r"^Crown Height\s+\(.*?\)\s+[\d\.]+\s*%\s*\(\s*([\d\.]+)",
        "Pavilion Depth": r"^Pavilion Depth\s+\(.*?\)\s+([\d\.]+\s*%)",
        "Pavilion Angle": r"^Pavilion Depth\s+\(.*?\)\s+[\d\.]+\s*%\s*\(\s*([\d\.]+)",
        "Length Halves Crown": r"^Length Halves Crown\s+([\d\.]+\s*%)",
        "Length Halves Pavilion": r"^Length Halves Pavilion\s+([\d\.]+\s*%)",
    }

    for field, pattern in patterns.items():
        match = re.search(pattern, text_clean, re.MULTILINE)
        if match:
            if field == "Cut":
                extracted_value = match.group(1) or match.groupdict().get("cut_grade", "")
                extracted_value = extracted_value.strip()
            else:
                extracted_value = match.group(1).strip()

            if field == "Colour Grade":
                bracket_match = re.search(r"\(\s*([A-Za-z-]+)\s*\)", extracted_value)
                extracted_value = bracket_match.group(1) if bracket_match else ""
            elif field == "Girdle":
            # Remove the old 'split()[0]' so it keeps the full phrase like "very thick"
                    extracted_value = extracted_value.lower().strip()


            data[field] = extracted_value
        else:
            data[field] = ""

    return data  # <-- This is the missing piece!


def main():
    folder = select_folder()
    if not folder:
        print("No folder selected.")
        return
    
    # Get all PDF files in the selected folder.
    pdf_files = [os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print("No PDF files found in the selected folder.")
        return
    
    all_data = []
    for pdf in pdf_files:
        print(f"Processing: {pdf}")
        text = extract_text_from_pdf(pdf)
        data = parse_data(text)
        # Include the filename for reference if desired.
        all_data.append(data)
    
    # Create a DataFrame with all extracted data.
    df = pd.DataFrame(all_data)
    # Save the combined data to one Excel file.
    output_dir = r"G:\\Certificate\\Python Output\\PDF to Excel"
    output_excel = os.path.join(output_dir, "HRD.xlsx")
    df.to_excel(output_excel, index=False)
    
    print(f"All data extracted and saved to {output_excel}")
    os.startfile(output_excel)

if __name__ == '__main__':
    main()
