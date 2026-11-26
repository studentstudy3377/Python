import base64
import io
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter


NEEDED_COLUMNS = [
    "QualityCode",
    "FOR",
    "Size",
    "Pcs",
    "Carats",
    "Price",
    "Cost",
    "Disc 1%",
    "Disc 2%",
    "Add %",
    "Remark",
    "stoneid",
    "PARTY NAME",
    "BANK$",
    "BROKER NAME",
]


def read_excel(uploaded_file):
    uploaded_file.seek(0)
    wb = load_workbook(filename=io.BytesIO(uploaded_file.read()), data_only=True)
    return wb.active


def get_columns_to_keep(sheet):
    header_row = list(sheet.iter_rows(values_only=True))[0]
    name_to_idx = {
        name: idx
        for idx, name in enumerate(header_row)
        if name in NEEDED_COLUMNS
    }
    return name_to_idx


def extract_data(sheet, columns_indexes):
    data = []
    for row in sheet.iter_rows(values_only=True, min_row=2):
        filtered = {}
        for col_name, idx in columns_indexes.items():
            filtered[col_name] = row[idx] if idx is not None and idx < len(row) else None
        data.append(filtered)
    return data


def process_data(data):
    columns_to_clear = [
        "Cost",
        "Disc 1%",
        "Disc 2%",
        "Add %",
        "FOR",
        "Size",
        "Pcs",
        "Remark",
    ]

    for row in data:
        for k in NEEDED_COLUMNS:
            row.setdefault(k, None)

        if row.get("Cost") is not None:
            row["Price"] = row["Cost"]

        for col in columns_to_clear:
            row[col] = None

    return data


def create_party_csvs_in_memory(data):
    files = {}
    party_frames = {}

    parties = sorted(
        {
            str(row.get("PARTY NAME")).strip()
            for row in data
            if row.get("PARTY NAME") is not None and str(row.get("PARTY NAME")).strip() != ""
        }
    )

    csv_columns = [
        "QualityCode",
        "FOR",
        "Size",
        "Pcs",
        "Carats",
        "Price",
        "Cost",
        "Disc 1%",
        "Disc 2%",
        "Add %",
        "Remark",
        "stoneid",
    ]

    for party in parties:
        party_rows = [
            row for row in data
            if str(row.get("PARTY NAME")).strip() == party
        ]
        if not party_rows:
            continue

        df = pd.DataFrame(party_rows)
        party_frames[party] = df.copy()

        df_csv = df[[c for c in csv_columns if c in df.columns]].copy()

        csv_buf = io.StringIO()
        df_csv.to_csv(csv_buf, index=False, encoding="utf-8")

        safe_party = "".join(c if c.isalnum() or c in (" ", "_", "-") else "_" for c in party).strip()
        if not safe_party:
            safe_party = "party"

        files[f"{safe_party}.csv"] = csv_buf.getvalue().encode("utf-8")

    return files, party_frames


def create_excel_from_partyframes(party_frames: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    alignment_center = Alignment(horizontal="center", vertical="center")
    alignment_left = Alignment(horizontal="left", vertical="center")

    today_date = datetime.today().strftime("%d/%m/%Y")

    start_column = 0
    start_row = 1
    invoices_per_column = 4
    row_offset = 12
    column_offset = 4

    def looks_numeric(x) -> bool:
        if x is None:
            return False
        s = str(x).strip()
        if s == "":
            return False
        try:
            float(s)
            return True
        except ValueError:
            return False

    for index, (party, party_data) in enumerate(party_frames.items()):
        carats_value = pd.to_numeric(party_data.get("Carats"), errors="coerce").fillna(0).sum()
        carats = f"{carats_value:.2f}"

        if "BANK$" in party_data.columns:
            bank_values = (
                party_data["BANK$"]
                .dropna()
                .astype(str)
                .str.strip()
                .unique()
            )
            total_bank = bank_values[0] if len(bank_values) == 1 else ", ".join(bank_values)
        else:
            total_bank = ""

        if "BROKER NAME" in party_data.columns and not party_data["BROKER NAME"].isna().all():
            broker_name = str(party_data["BROKER NAME"].iloc[0] or "").strip()
        else:
            broker_name = ""

        stone_id = (
            party_data["stoneid"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )
        total_stone_id = stone_id[0] if len(stone_id) == 1 else "✔"

        bill_type = "RS BILL" if looks_numeric(total_bank) else "$ BILL"

        if index % invoices_per_column == 0 and index > 0:
            start_column += column_offset
            start_row = 1

        data_template = [
            ["DATE :-", today_date, "", "", "", ""],
            ["PARTY :-", party, "", "", "", ""],
            ["THRU:-", broker_name, "", "", "", ""],
            ["QUALITY", "CARAT", bill_type, "", "", ""],
            ["", carats, "", "", "", ""],
            ["", "", "", "", "", ""],
            ["LESS (-)", "NET", "", "", "", ""],
            ["X", total_bank, "", "", "", ""],
            ["TERM", "  Days", "", "", "", ""],
            ["REMARK", total_stone_id, "", "", "", ""],
            ["BUYING PERSON", "SAURABH SHAH", "", "", "", ""],
        ]

        merge_rows = [
            start_row,
            start_row + 1,
            start_row + 2,
            start_row + 6,
            start_row + 7,
            start_row + 8,
            start_row + 9,
            start_row + 10,
        ]
        for r in merge_rows:
            ws.merge_cells(
                start_row=r,
                start_column=2 + start_column,
                end_row=r,
                end_column=3 + start_column,
            )

        for row_idx, row_vals in enumerate(data_template, start=start_row):
            for col_idx, value in enumerate(row_vals[:3], start=1):
                cell = ws.cell(row=row_idx, column=col_idx + start_column)

                if not isinstance(cell, MergedCell):
                    cell.value = value

                cell.alignment = alignment_left if row_vals[0] == "LESS (-)" else alignment_center
                cell.border = thin_border

        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 17

        start_row += row_offset

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def trigger_auto_downloads(csv_files, run_key="default"):
    if not csv_files:
        return

    js_parts = []
    for i, (filename, file_bytes) in enumerate(csv_files.items()):
        b64 = base64.b64encode(file_bytes).decode()
        js_parts.append(
            f"""
            setTimeout(function() {{
                const a = document.createElement('a');
                a.href = "data:text/csv;base64,{b64}";
                a.download = {filename!r};
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }}, {i * 400});
            """
        )

    html = f"""
    <!DOCTYPE html>
    <html>
    <body>
    <script>
    {''.join(js_parts)}
    </script>
    </body>
    </html>
    """
    st.components.v1.html(html, height=0, scrolling=False)


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

def run_mpb_bill():
    st.title("MBP Bills")

    sample_path = _find_sample("MPB Bills.xlsx")
    
    with st.sidebar:
    
        if sample_path is not None:
            with open(sample_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            href = (
                f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" '
                f'download="{sample_path.name}">Sample File</a>'
            )
            st.markdown(href, unsafe_allow_html=True)
        else:
            st.warning("Sample file not found on server.")

        st.markdown("### Upload File")
        uploaded_file = st.file_uploader(
            "Select Excel File",
            type=["xlsx", "xls"],
            key="mpb_bill_uploader",
        )

    if not uploaded_file:
        st.info("Upload an Excel file to begin.")
        return

    sheet = read_excel(uploaded_file)
    columns_indexes = get_columns_to_keep(sheet)

    if "PARTY NAME" not in columns_indexes:
        st.error("The uploaded file does not contain a 'PARTY NAME' column in the header row.")
        return

    data = extract_data(sheet, columns_indexes)
    processed = process_data(data)

    csv_files, party_frames = create_party_csvs_in_memory(processed)

    if not csv_files:
        st.warning("No parties were found to generate CSV files.")
    else:
        st.success(f"{len(csv_files)} CSV file(s) generated successfully.")

        if "csv_auto_downloaded" not in st.session_state:
            st.session_state.csv_auto_downloaded = False

        if not st.session_state.csv_auto_downloaded:
            trigger_auto_downloads(csv_files, run_key=str(len(csv_files)))
            st.session_state.csv_auto_downloaded = True

        with st.expander("Mahesh Sir File", expanded=True):
            for fname, bytes_csv in csv_files.items():
                st.download_button(
                    label=fname,
                    data=bytes_csv,
                    file_name=fname,
                    mime="text/csv",
                    key=f"dl_{fname}",
                )

        if st.button("Download CSV Files"):
            trigger_auto_downloads(csv_files, run_key=f"retry_{datetime.now().timestamp()}")
            st.info("Tried downloading all CSV files again.")

    if party_frames:
        excel_bytes = create_excel_from_partyframes(party_frames)
        current_datetime = datetime.now().strftime("%d.%m.%Y")
        xlsx_name = f"MPB Bill {current_datetime}.xlsx"

        st.subheader("Bill Excel")
        st.download_button(
            label=f"Download {xlsx_name}",
            data=excel_bytes,
            file_name=xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="bill_excel_download",
        )


if __name__ == "__main__":
    run_mpb_bill()
