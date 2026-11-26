import streamlit as st
import pandas as pd
import numpy as np

# ---------- Config ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["CertNo"],
    "CARAT": ["Cts"],
    "SHAPE": ["Shape"],
    "COLOR": ["Col"],
    "CLARITY": ["Cla"],
    "CUT": ["Cut"],
    "POL": ["Pol"],
    "SYM": ["Sym"],
    "FLRC": ["Fluo"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["TA%"],
    "DEPTH": ["TD"],
    "LAB": ["Lab"],
    "DISC": ["Disc%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Black Table Inclusion'],
    'Natts2':['Black Side Inclusion'],
    "Stone id":['Ref No']
}

# Lists (from your script)
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
flour = ["NONE", "FNT", "MED"]
key_symbol = ["Cavity", "Cavity", "Chip", "Cavity/Chip", "Chip/Cavity"]
luster_ok = ["EX", ""]
shade_ok = ["NONE", ""]
table_spot_ok = ["no black", ""]

# ---------- Utils ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore") if isinstance(x, (bytes, bytearray)) else x
            )
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df

def read_file(uploaded_file, skiprows: int = 0) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)

def numeric(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df

def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df

def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df = numeric(df, col)
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df

def negation(df: pd.DataFrame, col: str, values, negate: bool = False, keep_na: bool = True) -> pd.DataFrame:
    if col in df.columns:
        mark = df[col].isin(values)
        if negate:
            mark = ~mark
        if keep_na:
            mark = mark | df[col].isna()
        return df[mark]
    return df

def fix_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={a: k for k, v in columns_mapping.items() for a in v + [k]})
    for k in columns_mapping:
        if k not in df.columns:
            df[k] = ""
    return df[list(columns_mapping)]

def default_value(df: pd.DataFrame) -> pd.DataFrame:
    # Set numeric-like defaults
    for col in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
        if col in df.columns:
            df[col] = 0
    # Ensure numeric types
    for num_col in ["CARAT", "TABLE", "DEPTH", "DISC", "RATIO", "RAP", "LENGTH", "WIDTH", "HEIGHT"]:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")
    # Text defaults
    for col in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
        if col in df.columns:
            df[col] = "N/A"
    if "SRNO" in df.columns:
        df["SRNO"] = range(1, len(df) + 1)
    if "NATTS" in df.columns:
        df["NATTS"] = "CENTER - L"
    # Output shape as ROUND (input is normalized earlier)
    if "SHAPE" in df.columns:
        df["SHAPE"] = df["SHAPE"].replace({"RD": "ROUND"})
    # Normalize Fluo NONE->NON for output
    if "FLRC" in df.columns:
        df["FLRC"] = df["FLRC"].replace({"NONE": "NON"})
    return df

# ---------- Filters (mirror original logic) ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    # Clean '+' from Col/Cla first
    if "Col" in df.columns:
        df["Col"] = df["Col"].astype(str).str.replace("+", "", regex=False)
    if "Cla" in df.columns:
        df["Cla"] = df["Cla"].astype(str).str.replace("+", "", regex=False)

    df = equal(df, "Lab", "GIA")
    df = equal(df, "Shape", "RD")
    # Map RD -> ROUND for downstream display
    if "Shape" in df.columns:
        df["Shape"] = df["Shape"].replace({"RD": "ROUND"})

    df = between(df, "Cts", 0.30, 3.99, inclusive="both")
    df = equal(df, "Location", "INDIA")

    # Disc% force negative, then keep <= -10
    if "Disc%" in df.columns:
        df = numeric(df, "Disc%")
        df["Disc%"] = df["Disc%"].apply(lambda x: -abs(x) if pd.notna(x) else x)
        df = df[df["Disc%"] <= -10]

    df = negation(df, "Col", colors)
    df = negation(df, "Cla", clarity)
    df = negation(df, "Fluo", flour)

    # Cut/Pol/Sym in EX/VG
    if all(c in df.columns for c in ["Cut", "Pol", "Sym"]):
        df = df[df["Cut"].isin(cut) & df["Pol"].isin(cut) & df["Sym"].isin(cut)]
    else:
        df = negation(df, "Cut", cut)
        df = negation(df, "Pol", cut)
        df = negation(df, "Sym", cut)

    df = negation(df, "SH", shade_ok)  # shade
    df = negation(df, "Black Side Inclusion", table_spot_ok)
    # It appeared twice in your script—once is enough
    df = negation(df, "Key To Symbol", key_symbol, negate=True)

    # Exclude comments starting with "Clarity"
    if "Comment" in df.columns:
        df = df[~df["Comment"].astype(str).str.startswith("Clarity", na=False)]

    return df

# ---------- App ----------
def run_kb_mod() -> None:
    st.header("KB Gems ")

    with st.sidebar:
        upload = st.file_uploader("Select KG Excel", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 2, 1)  # your script used skiprows=2

    if upload is None:
        st.stop()

    try:
        df = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw Data")
    st.dataframe(df)

    # Apply filters and normalize
    df = apply_filters(df)
    df = fix_columns(df)
    df = default_value(df)

    st.subheader("Processed Data")
    st.dataframe(df)

    csv_bytes = df.to_csv(index=False)
    file_name = f"kb gems {pd.Timestamp.now():%d.%m.%Y}.csv"

    with st.sidebar:
        st.download_button(
            label="Generate CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_kb_mod()
