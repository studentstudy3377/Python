import os
from io import BytesIO
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
import streamlit as st


# ---------- File I/O helpers ----------

def read_uploaded(uploaded, skiprows=0):
    """
    Read a Streamlit UploadedFile and return (filename, df or None).
    Works for .xlsx, .xls, .csv read from memory (no temp files needed).
    """
    filename = uploaded.name
    ext = os.path.splitext(filename)[1].lower()
    try:
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(uploaded, skiprows=skiprows)
        elif ext == ".csv":
            df = pd.read_csv(uploaded, skiprows=skiprows)
        else:
            return (filename, None)
        return (filename, df)
    except Exception as e:
        print(f"❌ Failed to read file: {filename} -> {e}")
        return (filename, None)


def load_files_parallel(uploaded_files, skiprows=0):
    """
    uploaded_files: list[UploadedFile]
    Returns a dict: {filename: DataFrame}
    """
    dfs = {}
    with ThreadPoolExecutor(max_workers=4) as executor:
        results = list(executor.map(lambda f: read_uploaded(f, skiprows), uploaded_files))
    for filename, df in results:
        if df is not None:
            dfs[filename] = df
    return dfs


# ---------- Column mapping & value normalization ----------

def map_column(df, old_names, new_name):
    """
    Return first matching column as Series; else empty Series of length len(df).
    """
    for col in old_names:
        if col in df.columns:
            return df[col]
    return pd.Series([""] * len(df))


def normalize_text_series(series):
    """Return an Arrow-safe pandas string Series from mixed spreadsheet values."""
    def normalize_value(value):
        if value is None or value is pd.NA:
            return ""
        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass

        if isinstance(value, (bytes, bytearray)):
            return bytes(value).decode("utf-8", errors="replace").strip()
        return str(value).strip()

    return series.map(normalize_value).astype("string")


def process_cut_polish_symmetry(v):
    """Normalize Cut/Polish/Symmetry values."""
    v = str(v).strip().upper()
    if v in {"", "-", "NAN", "NONE"}:
        return ""
    if v in {"G", "GOOD"}:
        return "Good"
    if v == "FR":
        return "Fear"
    if v in {"EXCELLENT", "3EX", "EXCL", "EXC", "EX", "3EXCELLENT"}:
        return "EX"
    if v in {"VERY GOOD", "VGOOD", "VERY"}:
        return "VG"
    return v


def fluorescence(v):
    lookup = {
        "NN": "NONE", "NON": "NONE", "N": "NONE", "NIL": "NONE", "NONE": "NONE", "None": "NONE",
        "F": "FAINT", "FNT": "FAINT", "FA": "FAINT",
        "MED. BLUE": "MED", "MEDIUM": "MED", "MD": "MED", "M": "MED",
        "SLIGHT": "SLIGHT", "SL": "SLIGHT",
        "STG": "STRONG",
        "V.STRONG": "V.STRONG",
        "VERY SLIGHT": "VERY SLIGHT"
    }
    s = str(v).strip()
    u = s.upper()
    if not s:
        return "NONE"
    return lookup.get(u, u)


def culet(v):
    lookup = {"NN": "None", "NON": "None", "N": "None", "NIL": "None", "NONE": "None"}
    s = str(v).strip()
    u = s.upper()
    if not s:
        return "None"
    return lookup.get(u, s)


def location(v):
    lookup = {"MUMBAI": "Mumbai", "NEWYORK": "New York", "NEW YORK": "New York"}
    s = str(v).strip()
    u = s.upper()
    if not s:
        return "None"
    return lookup.get(u, s)


def combine_cols(df, col1, col2):
    """Combine two columns with ' - ' separator, handling missing columns and NaNs."""
    if col1 in df.columns:
        val1 = df[col1].astype(str).str.strip().replace("nan", "", regex=False)
    else:
        val1 = pd.Series([""] * len(df))

    if col2 in df.columns:
        val2 = df[col2].astype(str).str.strip().replace("nan", "", regex=False)
    else:
        val2 = pd.Series([""] * len(df))

    combined = (val1.fillna("") + " - " + val2.fillna("")).str.replace(r"^\s*-\s*$", "", regex=True)
    combined = combined.str.replace(r"^ - ", "", regex=True).str.replace(r" - $", "", regex=True)
    return combined


def build_inclusion_series(dfs, col1a="T.B", col1b="S.B", col2a="BT", col2b="BC"):
    """
    Build Inclusion using this priority for each row:
    1. Inclusion
    2. NGS Comment
    3. T.B / S.B
    4. BT / BC
    """

    def has_value(value):
        if value is None or value is pd.NA:
            return False
        try:
            if pd.isna(value):
                return False
        except (TypeError, ValueError):
            pass
        return str(value).strip().upper() not in {"", "NAN", "NONE", "<NA>"}

    result_series = []

    for df in dfs.values():
        inclusion = normalize_text_series(
            map_column(df, ["Inclusion"], "Inclusion")
        )
        ngs_comment = normalize_text_series(
            map_column(df, ["NGS Comment"], "NGS Comment")
        )

        # Use NGS Comment row-by-row when Inclusion is blank.
        source_comment = inclusion.combine(
            ngs_comment,
            lambda current, fallback: current if has_value(current) else fallback,
        )

        tb_sb = combine_cols(df, col1a, col1b)
        bt_bc = combine_cols(df, col2a, col2b)

        combined = source_comment.combine(
            tb_sb,
            lambda comment, fallback: comment if has_value(comment) else fallback,
        )
        combined = combined.combine(
            bt_bc,
            lambda current, fallback: current if has_value(current) else fallback,
        )

        result_series.append(combined)

    return pd.concat(result_series, ignore_index=True)


# ---------- Excel helper ----------

def to_excel(df):
    """Convert a DataFrame to Excel bytes in memory (no temp file)."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


# ---------- Streamlit App ----------

def run_rapnet_format_mod():
    st.title("Rapnet Format")
    with st.sidebar:
        st.markdown("### Upload Files")
        files = st.file_uploader(
            "Select one or more files",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
        )

    if not files:
        st.info("Upload at least one .xlsx/.xls/.csv to proceed.")
        st.stop()

    # Format Meas1, Meas2 and Meas3 into one Measurement value
    def format_measurement(length, width, depth):
        values = []
        for value in (length, width, depth):
            if pd.isna(value):
                return ""
            try:
                values.append(float(str(value).replace(",", "").strip()))
            except (TypeError, ValueError):
                return ""

        length_num, width_num, depth_num = values
        return f"{length_num:.2f} - {width_num:.2f} X {depth_num:.2f}"

    # 🔹 Loader only: show a spinner during processing (no other UI changes)
    with st.spinner("⏳ Processing files..."):
        dfs = load_files_parallel(files)
        if not dfs:
            st.error("No readable data found in the uploaded files.")
            st.stop()

        # Build Measurement for every uploaded file that has all three columns.
        for source_df in dfs.values():
            required_cols = {"Meas1", "Meas2", "Meas3"}
            if required_cols.issubset(source_df.columns):
                source_df["Measurement"] = source_df.apply(
                    lambda row: format_measurement(
                        row["Meas1"],
                        row["Meas2"],
                        row["Meas3"],
                    ),
                    axis=1,
                )

        # Column candidates
        lab_columns = ["CERTI", "Lab", "LABNAME", "Cert", "LAB", "Cert Lab", "Crt", "Cert.","Certi",'CERT']
        measurement_columns = [
            "Diam.", "Dia", "MM", "L*W/Diam", "Measurements(L*W*D)","MEASUREMENT",
            "MEASUREMENTS", "Measurements", "Measurement", "L/W Or Dia", "Meas MM", "Meas"
        ]
        report_no_columns = ["Stone ID", "Stock No", "Vendor Stock number","Stone Id","CERTNO"]

        columns_mapping = {
            "Stone ID": report_no_columns,
            "Shape": ["SHP", "shape", "ShapeName", "Shp.", "SHAPE", "Shape", "Shp", "shp", "Shape and Cutting Style"],
            "Size": ["Carat", "Size","CARAT"],
            "Color": ["COL", "Colour Grade", "color", "Color Grade", "ColorName", "col", "Colour", "COLOR", "Color", "DispColor", "Col"],
            "Clarity": ["CLA", "Clrt", "Quality", "purity", "CTY", "CLAR", "Clarity Grade", "ClarityName", "Clr", "Clarity", "CLARITY", "Purity", "CLR", "DispClarity", "Cl.", "Cla", "Clar"],
            "Cut": ["C", "Sh.Ap. / Cut", "Proportions", "Cut Grade", "Cut", "Final Cut", "CUT"],
            "Polish": ["P", "POL", "polish", "Polish", "POLISH", "PolishName", "Pol", "POL.", "Pl", "Po", "Pol."],
            "Symmetry": ["S", "SYM", "Symm.", "sym", "Symmetry", "SymmetryName", "Sym", "SYM.", "Symm", "Sym.", "Sy"],
            "Fluorescence": ["FLUO", "FLS", "LW-fluo", "fl", "FLR", "Fluorescence", "Flu", "FLIntName", "Flor", "FLO.", "Fluorescence ", "FlrIntens", "Fluorescence Intensity", "Fluor", "Fluor.", "Flur", "Flour", "Flo", "Fl", "FLUOR", "Flr", "Fluo Int", "Fluo", "Fluorescent", "Fls", "FL"],
            "Lab": lab_columns,
            "Ratio": ["Ratio", "D/R",'RATIO'],
            "Depth": ["Total Depth", "DEPTH", "Depth", "TD", "T.Dpt", "Total Depth Per.", "Depth(%)", "Dep%", "Depth%", "Tot.Depth %", "Tot Depth", "T.Dep.", "Dep", "Total depth%", "TD%", "Depth %", "T.D.%"],
            "Table": ["TABLE", "TA", "Tbl", "Table Size Per.", "Table width %", "TAB%", "TA%", "Tableprct", "TBL", "Tab", "Table", "Tab%", "Table Depth", "Table%", "Table(%)", "Table %"],
            "Measurement": measurement_columns,
            "Shade": ["Shade", "Shd.", "Color Shade",'TINCH'],
            "Inclusion": ["Inclusion", "NGS Comment"],
            "Key To Symbols": ["Key To Symbol", "Key To Symbols","KEY TO SYMBOLS"],
            "Vendor Stock number": ["Stone Id","STOCKID"],
            "Rap": ["Rap ", "Rap Rate", "Rap $/CT","Rap","RAPPRICE"],
            "$/ct": ["Pr/Ct", "Base Rate", "$/ct","$/ct.","PRICE"],
            "%Rap": ["Rap%", "Base Off %", "%Rap","Disc(%)","DISC",],
            "Total": ["Amount", "Total","Total $.","VALUE"],
            "State": ["State", "Location"],
            "Milky": ["Lust", "MILKY","LUSTER"]
        }

        # Build output DataFrame
        data = pd.DataFrame()

        # Seller from filename (without extension)
        party_names = []
        for name, df in dfs.items():
            party_names.extend([os.path.splitext(name)[0]] * len(df))
        data["Seller"] = pd.Series(party_names, dtype="string")

        # Map all configured columns
        for column, keys in columns_mapping.items():
            series_list = [map_column(df, keys, column) for df in dfs.values()]
            data[column] = pd.concat(series_list, ignore_index=True)

            # Normalizations
            if column in ["Cut", "Polish", "Symmetry"]:
                data[column] = data[column].map(process_cut_polish_symmetry)
            elif column == "Fluorescence":
                data[column] = data[column].map(fluorescence)
            elif column == "Culet":
                data[column] = data[column].map(culet)
            elif column in ["State", "Location"]:
                data[column] = data[column].map(location)

        # Inclusion priority: Inclusion, NGS Comment, T.B/S.B, then BT/BC
        data["Inclusion"] = build_inclusion_series(dfs)

        # ---- Arrow-safe dtypes ----
        # Spreadsheet columns with dtype=object can contain a mixture of strings,
        # bytes, integers, and blanks. Arrow requires one consistent type per column.
        numeric_cols = {"Size", "Ratio", "Depth", "Table", "Rap", "$/ct", "%Rap", "Total"}
        text_cols = [col for col in data.columns if col not in numeric_cols]

        for col in text_cols:
            data[col] = normalize_text_series(data[col])

        # Keep numeric fields numeric where possible; invalid text becomes null.
        for col in numeric_cols:
            if col in data.columns:
                data[col] = pd.to_numeric(data[col], errors="coerce")

        # Convert to Excel inside spinner
        excel_bytes = to_excel(data)

    # After spinner completes
    st.success("✅ File uploaded successfully!")
    st.dataframe(data)  

    with st.sidebar:
        st.download_button(
            label="⬇️ Download Excel",
            data=excel_bytes,
            file_name="Rapnet Format.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

if __name__ == "__main__":
    run_rapnet_format_mod()
