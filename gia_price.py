import streamlit as st
import openpyxl
import datetime
from openpyxl.styles import PatternFill, Font
from io import BytesIO

# Function to read and process the Excel file
def process_file(uploaded_file):
    # Load Excel file directly from the uploaded file object
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active

    # Keep only one sheet
    for sheet in wb.sheetnames:
        if sheet != ws.title:
            del wb[sheet]

    # Read headers and column indexes
    headers = [cell.value for cell in ws[1]]
    col_idx = {header: idx + 1 for idx, header in enumerate(headers)}

    # Add missing columns
    for col in ['VALUE', 'PRICE', 'DISC']:
        if col not in headers:
            ws.cell(row=1, column=ws.max_column + 1, value=col)
            col_idx[col] = ws.max_column

    # Function to style cell and try numeric conversion
    def clean_and_convert(cell):
        # Preserve hyperlink style
        if cell.hyperlink:
            cell.font = Font(color="0000FF", underline="single")
        else:
            cell.font = Font(color="000000")
        cell.fill = PatternFill(fill_type=None)
        
        # Try numeric conversion if cell has no hyperlink and looks like a number
        if not cell.hyperlink and isinstance(cell.value, str):
            try:
                val = cell.value.strip()
                if val.replace('.', '', 1).isdigit():
                    cell.value = float(val) if '.' in val else int(val)
            except:
                pass

    # Apply style and numeric conversion to all cells
    for row in ws.iter_rows():
        for cell in row:
            clean_and_convert(cell)

    # Main logic to modify the 'VALUE', 'PRICE', and 'DISC' columns
    for row in ws.iter_rows(min_row=2, values_only=False):
        if row[col_idx['CERT'] - 1].value != 'GIA':
            continue

        carat = row[col_idx['CARAT'] - 1].value
        value = row[col_idx['VALUE'] - 1].value

        # Carat addition logic
        if carat is not None:
            if 0.15 <= carat <= 0.22: addition = 24
            elif 0.23 <= carat <= 0.29: addition = 25.6
            elif 0.30 <= carat <= 0.39: addition = 31.2
            elif 0.40 <= carat <= 0.49: addition = 32.8
            elif 0.50 <= carat <= 0.69: addition = 33.6
            elif 0.70 <= carat <= 0.99: addition = 40.8
            elif 1.00 <= carat <= 1.49: addition = 68
            elif 1.50 <= carat <= 1.99: addition = 88
            elif 2.00 <= carat <= 2.99: addition = 127
            else: addition = 0
        else:
            addition = 0

        new_value = (value or 0) + round(addition, 2)
        row[col_idx['VALUE'] - 1].value = new_value

        # PRICE
        price = round(new_value / carat, 2) if carat else 0
        row[col_idx['PRICE'] - 1].value = price

        # DISC
        rapprice = row[col_idx['RAPPRICE'] - 1].value
        disc = round((price - rapprice) / rapprice * 100, 2) if rapprice else 0
        row[col_idx['DISC'] - 1].value = disc

    # Convert the updated workbook to a BytesIO object
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit app layout
def run_gia_price():
    st.title("GIA Price")
    
    with st.sidebar:
        st.markdown("### Upload File")
        uploaded_file = st.file_uploader("Select  File", type=["xlsx", "xls"])

    if not uploaded_file:
        st.info("⬆️ Please upload a file to continue.")
        st.stop()
    
    with st.spinner('Process file...'):
        # Process the uploaded file
        processed_file = process_file(uploaded_file)

    # Provide the processed file for download
    st.success("File processed successfully!")
    today_str = datetime.datetime.now().strftime("%d-%m-%Y")
    
    with st.sidebar:
        st.download_button(
            label="⬇️ Download Excel",
            data=processed_file,
            file_name=f"GIA Price {today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":

    run_gia_price()
