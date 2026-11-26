import streamlit as st
import os
import re
import openpyxl
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
            color_re = re.compile(r'^[D-M]')  # Matches rows with color code
            while j < len(lines):
                if color_re.match(lines[j]):
                    row = re.split(r'\s+', lines[j].strip())
                    while len(row) < len(headers):
                        row.append("")  # Fill missing columns with empty string
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
            return str(val)  # Convert all non-numeric values to string

def write_tables_to_excel(tables):
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

    # Save the file to a BytesIO object and return it
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit UI
def rap_report_pdf():
    st.title("Rap Report  PDF To Excel")
    with st.sidebar:
        pdf_file = st.file_uploader("Select PDF File ", type=["pdf"])

    if pdf_file is not None:
        try:
            # Show a loader while processing
            with st.spinner('Processing your PDF...'):
                with pdfplumber.open(pdf_file) as pdf:
                    text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
                tables = parse_rap_pdf_text(text)

                if not tables:
                    st.error("No tables found in the PDF.")
                else:
                    st.success("Converting to Excel...")

                    # Write the tables to Excel
                    excel_output = write_tables_to_excel(tables)

                    # Read the Excel file into a pandas DataFrame for preview
                    excel_df = pd.read_excel(excel_output, sheet_name="All_Tables")

                    # Convert all values to strings using applymap replacement
                    excel_df = excel_df.apply(lambda x: x.astype(str))

                    # Preview the first few rows of the Excel file
                    st.subheader("Preview of Excel File")
                    st.dataframe(excel_df.head())  # Show the first few rows

                    # Download button
                    st.download_button(
                        label="Download Excel file",
                        data=excel_output,
                        file_name="Rap Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    rap_report_pdf()
