import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# ---------- IO ----------

def read_file_safe(file, skiprows: int = 0) -> pd.DataFrame:
    """Read Excel or CSV safely while preserving Job No as text."""
    file_extension = file.name.rsplit(".", 1)[-1].lower()

    read_options = {
        "skiprows": skiprows,
        "dtype": {"Job No": "string"},
        "keep_default_na": False,
    }

    if file_extension in ["xlsx", "xls", "xlx"]:
        df = pd.read_excel(file, **read_options)

    elif file_extension == "csv":
        df = pd.read_csv(file, **read_options)

    else:
        raise ValueError("Unsupported file type.")

    # Clean column names
    df.columns = [str(column).strip() for column in df.columns]

    return df

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Serialize DataFrame to XLSX."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Merged")
    return output.getvalue()

# ---------- Numeric conversion (cell-by-cell) ----------
RE_PAREN_NEG = re.compile(r'^\((.*)\)$')  # (123) -> -123

def _prep_for_numeric(s: str) -> str:
    """Normalize a single string for numeric parsing."""
    s = s.strip()
    if not s:
        return s
    m = RE_PAREN_NEG.match(s)
    if m:
        s = "-" + m.group(1).strip()
    # remove commas and spaces
    return re.sub(r'[,\s]', '', s)

def convert_cell_numeric_like(val):
    """Convert a single cell to int/float if numeric-like; else return original."""
    if pd.isna(val):
        return val
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        s = _prep_for_numeric(val)
        if s == "":
            return val
        try:
            if "." not in s and "e" not in s.lower():
                return int(s)
        except ValueError:
            pass
        try:
            return float(s)
        except ValueError:
            return val
    return val

def convert_dataframe_cellwise(df: pd.DataFrame) -> pd.DataFrame:
    """Cell-wise conversion: if a cell looks numeric → make it a number; else keep text."""
    out = df.copy()
    for c in out.columns:
        if out[c].dtype in ("string", "object"):
            out[c] = out[c].map(convert_cell_numeric_like)
    return out

# ---------- Arrow-safe display helper ----------
def make_display_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Convert mixed-type columns to string for safe Streamlit display (Arrow)."""
    disp = df.copy()
    for c in disp.columns:
        inferred = pd.api.types.infer_dtype(disp[c], skipna=True)
        if inferred.startswith("mixed") or inferred in {"unknown", "mixed-integer", "mixed-integer-float"}:
            disp[c] = disp[c].astype("string")
    return disp

# ---------- App ----------
def run_merge_mod():
    st.title("🧾 Merge File")
    with st.sidebar:
        st.header("🧾 Merge")
        files = st.file_uploader("Select Files", type=["xlsx", "xls", "csv", "xlx"], accept_multiple_files=True)
        skip = st.number_input("Rows to skip", 0, 1000, 0, 1)

    if not files:
        st.stop()

    with st.spinner("⏳ Processing files..."):
        # Read all files
        dfs = []
        for f in files:
            try:
                df = read_file_safe(f, skiprows=skip)
                dfs.append(df)
            except Exception as e:
                st.error(f"Failed to read {getattr(f, 'name', 'file')}: {e}")
                st.stop()

        if not dfs:
            st.warning("No data could be read.")
            st.stop()

        # Common columns (intersection), preserve order from first file
        common_cols = set(dfs[0].columns)
        for d in dfs[1:]:
            common_cols &= set(d.columns)

        if not common_cols:
            st.error("No common columns across files.")
            st.stop()

        common_cols = [c for c in dfs[0].columns if c in common_cols]
        dfs_common = [d[common_cols] for d in dfs]
        merged = pd.concat(dfs_common, ignore_index=True, join="inner")

        # Cell-wise numeric conversion
        merged = convert_dataframe_cellwise(merged)

        # Display-only copy for Streamlit
        merged_display = make_display_safe(merged)

        # Downloads
        date_tag = datetime.now().strftime("%d-%m-%Y")
        csv_bytes = merged.to_csv(index=False)
        xlsx_bytes = to_excel_bytes(merged)

    # Loader ends here — now show downloads
    with st.sidebar:
        st.download_button(
            "⬇️ Download Excel",
            data=xlsx_bytes,
            file_name=f"Merged {date_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button(
            "⬇️ Download CSV",
            data=csv_bytes,
            file_name=f"Merged {date_tag}.csv",
            mime="csv",
        )

    st.success(f"Merged {len(files)} files • Rows: {len(merged)} • Columns: {len(merged.columns)}")
    st.dataframe(merged_display)

if __name__ == "__main__":
    run_merge_mod()
