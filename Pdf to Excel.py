import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import re

def parse_rap_pdf_text(text):
    title_pattern = re.compile(r'([0-9.]+ TO [0-9.]+ .{3,6})')
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    tables = []
    i = 0
    while i < len(lines):
        match = title_pattern.match(lines[i])
        if match:
            title = match.group(1)
            # Find header (must contain 'Color')
            j = i + 1
            while j < len(lines) and ('Color' not in lines[j]):
                j += 1
            if j >= len(lines): break
            headers_line = lines[j].replace(' ',' ').replace('\t',' ')
            headers = re.split(r'\s+', headers_line.strip())
            # Data follows
            j += 1
            data = []
            color_re = re.compile(r'^[D-M]')
            while j < len(lines):
                if color_re.match(lines[j]):
                    row = re.split(r'\s+', lines[j].strip())
                    while len(row) < len(headers):
                        row.append("")
                    data.append(row)
                    j += 1
                else:
                    break
            tables.append({'title': title, 'headers': headers, 'rows': data})
            i = j
        else:
            i += 1
    return tables

def try_number(val):
    # Remove commas for thousands, handle empty cells, and try conversion
    if val is None or val == "":
        return ""
    try:
        return int(val.replace(',', ''))
    except Exception:
        try:
            return float(val.replace(',', ''))
        except Exception:
            return val

def write_tables_to_single_sheet(tables, excel_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "All_Tables"
    bold = Font(bold=True)
    big = Font(size=14, bold=True)
    center = Alignment(horizontal='center', vertical='center')
    thin = Side(border_style="thin", color="000000")

    row_cursor = 1
    max_cols = 0
    for t in tables:
        ncols = len(t['headers'])
        max_cols = max(max_cols, ncols)
        # Write title (merged)
        ws.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=ncols)
        title_cell = ws.cell(row=row_cursor, column=1, value=t['title'])
        title_cell.font = big
        title_cell.alignment = center
        row_cursor += 1

        # Write headers
        for ci, h in enumerate(t['headers'], 1):
            cell = ws.cell(row=row_cursor, column=ci, value=h)
            cell.font = bold
            cell.alignment = center
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        row_cursor += 1

        # Write data rows with numeric conversion
        for row in t['rows']:
            for ci, val in enumerate(row, 1):
                v = try_number(val)
                cell = ws.cell(row=row_cursor, column=ci, value=v)
                cell.alignment = center
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
            row_cursor += 1

        # Blank row between tables
        row_cursor += 1

    # Auto-fit column width for all used columns
    for idx in range(1, max_cols + 1):
        max_len = 0
        for row in ws.iter_rows(min_col=idx, max_col=idx, min_row=1, max_row=ws.max_row):
            for cell in row:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(idx)].width = max_len + 2

    wb.save(excel_path)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    pdf_file = filedialog.askopenfilename(
        title="Select PDF to convert",
        filetypes=[("PDF files","*.pdf")]
    )
    if not pdf_file:
        messagebox.showinfo("Cancelled", "No file selected, exiting.")
        exit()

    base, _ = os.path.splitext(pdf_file)
    xlsx_file = base + "_ALLTABLES.xlsx"

    try:
        import pdfplumber
        with pdfplumber.open(pdf_file) as pdf:
            text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
        tables = parse_rap_pdf_text(text)
        if not tables:
            raise Exception("No tables found in the PDF.")
        write_tables_to_single_sheet(tables, xlsx_file)
        os.startfile(xlsx_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed:\n{e}")
