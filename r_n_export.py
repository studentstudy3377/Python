# app_rn_export.py
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["Certificate"],
    "CARAT": ["Weight"],
    "SHAPE": ["SHAPE"],
    "COLOR": ["Color"],
    "CLARITY": ["Clarity"],
    "CUT": ["Cut"],
    "POL": ["Polish"],
    "SYM": ["Symm"],
    "FLRC": ["Fluor"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["Table"],
    "DEPTH": ["DEPTH"],
    "LAB": ["LAB"],
    "DISC": ["DISC"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Black Incl.'],
    'Natts2':['Table Incl.'],
    "Stone id":['STOCK']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]

# Input can contain varied spacing/case for cut/fluor; we first filter on raw tokens, then normalize.
cut_raw = ["EX", " EX", "EX ", "VG", " VG", "VG ", "Ex", "Vg"]  # include a few common variants
flour_raw = ["N", "F", " F", "M", " M", " N"]

key_symbol_block = ["Cavity", "Chip", "Cavity, Chip", "Chip, Cavity"]

# ---------- Helpers ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore")
                if isinstance(x, (bytes, bytearray)) else x
            )
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df


def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)


def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df


def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df


# ---------- Filters (R N Export) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # LAB == GIA
    df = equal(df, "LAB", "GIA")

    # SHAPE == RBC, then normalize to ROUND
    if "SHAPE" in df.columns:
        df = df[df["SHAPE"] == "RBC"]
        df["SHAPE"] = df["SHAPE"].astype(str).str.replace("RBC", "ROUND", regex=False)

    # DISC given as fraction -> convert to % string, force negative, then numeric; keep <= -10
    if "DISC" in df.columns:
        # multiply by 100; make positive values negative; keep 2 decimals; then back to numeric
        df["DISC"] = pd.to_numeric(df["DISC"], errors="coerce")
        df["DISC"] = df["DISC"] * 100
        df["DISC"] = df["DISC"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        df["DISC"] = pd.to_numeric(df["DISC"], errors="coerce")
        df = df[df["DISC"] <= -10]

    # Location == IND
    df = equal(df, "LOC", "IND")

    # Weight 0.30–3.99
    df = between(df, "Weight", 0.30, 3.99, inclusive="both")

    # Attribute lists
    if "Color" in df.columns:
        df = df[df["Color"].isin(colors)]
    if "Clarity" in df.columns:
        df = df[df["Clarity"].isin(clarity)]

    # Cut/Polish/Symm: filter on raw variants, then normalize to EX/VG
    for c in ["Cut", "Polish", "Symm"]:
        if c in df.columns:
            df = df[df[c].isin(cut_raw)]
            df.loc[:, c] = (
                df[c]
                .astype(str)
                .str.strip()
                .str.upper()
                .replace({"EX": "EX", "VG": "VG"})
            )

    # Fluor: filter on raw variants, then normalize to NON/FNT/MED
    if "Fluor" in df.columns:
        df = df[df["Fluor"].isin(flour_raw)]
        df.loc[:, "Fluor"] = (
            df["Fluor"]
            .astype(str)
            .str.strip()
            .replace({"F": "FNT", "M": "MED", "N": "NON"})
        )

    # Block Key to Sym
    if "Key to Sym" in df.columns:
        df = df[~df["Key to Sym"].isin(key_symbol_block)]

    # Optional: exclude comments starting with "Clarity" (uncomment if needed)
    # if "COMM" in df.columns:
    #     s = df["COMM"].astype(str)
    #     df = df[~s.str.startswith("Clarity", na=False)]

    return df


def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Rename aliases + ensure all expected columns exist
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""

    # SRNO
    if "SRNO" in df.columns:
        df["SRNO"] = range(1, len(df) + 1)

    # Defaults
    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        if col in df.columns:
            df[col] = 0
    for col in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
        if col in df.columns:
            df[col] = "N/A"
    if "NATTS" in df.columns:
        df["NATTS"] = "CENTER - L"

    # Order
    return df[list(columns_mapping)]


# ---------- App ----------
def run_rn_export_mod() -> None:
    st.header("R N Export")

    with st.sidebar:
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 0, 1)  # default 0 like your script

    if upload is None:
        st.stop()

    try:
        df_raw = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw Input (sanitized)")
    st.dataframe(df_raw)

    df = apply_filters(df_raw)
    df = fix_columns(df)

    st.subheader("Final Output")
    st.dataframe(df)

    # Download CSV
    file_name = f"r n export {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df.to_csv(index=False)  # BOM helps Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )


if __name__ == "__main__":
    run_rn_export_mod()
