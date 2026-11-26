import base64
from pathlib import Path
import re
import streamlit as st
import pandas as pd
import numpy as np

# ---------- Column Mapping (target -> source in input) ----------
columns_mapping = {
    "SRNO": [],
    "LOTNO": ["CERTNO"],
    "CARAT": ["CTS"],
    "SHAPE": ["SHAPE"],
    "COLOR": ["COL"],
    "CLARITY": ["CLAR"],
    "CUT": ["CUT"],
    "POL": ["POL"],
    "SYM": ["SYM"],
    "FLRC": ["FL"],
    "LENGTH": [],
    "WIDTH": [],
    "HEIGHT": [],
    "TABLE": ["TAB%"],
    "DEPTH": ["T.D.%"],
    "LAB": ["LAB"],
    "DISC": ["RAP DISC%"],
    "TINCH": [],
    "NATTS": [],
    "RATIO": [],
    "RAP": [],
    "LUSTER": [],
    "T OPENS": [],
    "B OPENS": [],
    'Natts1':['Natts'],
    'Natts2':[],
    "Stone id":['Stone No','REFNO']
}

# ---------- Lists (criteria) ----------
colors = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]
clarity = ["IF", "VVS1", "VVS2", "VS1", "VS2", "SI1", "SI2"]
cut = ["EX", "VG"]
flour = ["NON", "FNT", "MED"]
luster_allow = ["EX", "N"]
culet_block = ["VSM", "LNG"]
flcolor_block = ["Green", "Yellow"]
# Will remove these tokens from INCLUSION (word-bounded)
inclusion_codes = ["CRB2", "TB2", "TB3", "TW3"]

# ---------- Helpers ----------
def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    # Trim headers, drop Unnamed, normalize object columns
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(
                lambda x: x.decode("utf-8", "ignore")
                if isinstance(x, (bytes, bytearray)) else x
            )
            # Mixed types -> stringify to avoid .str errors
            if df[col].map(type).nunique() > 1:
                df[col] = df[col].astype(str)
    return df


def read_file(uploaded_file, skiprows: int = 5) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, skiprows=skiprows)
    return _sanitize_df(df)


def equal(df: pd.DataFrame, col: str, value) -> pd.DataFrame:
    if col in df.columns:
        return df[df[col] == value]
    return df


def between(df: pd.DataFrame, col: str, low, high, inclusive: str = "both") -> pd.DataFrame:
    if col in df.columns:
        df = df.copy()
        df[col] = pd.to_numeric(df[col], errors="coerce")
        return df[df[col].between(low, high, inclusive=inclusive)]
    return df


def keep_empty_or(df: pd.DataFrame, col: str, allowed_exact_lower: set[str]) -> pd.DataFrame:
    """Keep rows where col is blank-ish OR in allowed_exact_lower (case-insensitive)."""
    if col not in df.columns:
        return df
    s = df[col].astype("string")
    empty_mask = s.isna() | s.str.strip().eq("")
    in_allowed = s.str.strip().str.casefold().isin(allowed_exact_lower)
    return df[empty_mask | in_allowed]


# ---------- Filters ----------
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()

    # Normalize key text columns to string early
    for col in [
        "LAB","SHAPE","COL","CLAR","CUT","POL","SYM","FL","TINGE","MILKY",
        "LUSTER","CULET","FLCOLOR","KEY TO SYMBOL","REPORT COMMENT","INCLUSION"
    ]:
        if col in d.columns:
            d[col] = d[col].astype(str)

    # Numbers coercion
    for col in ["RAP DISC%","CTS","TAB%","T.D.%"]:
        if col in d.columns:
            d[col] = pd.to_numeric(d[col], errors="coerce")

    # Base filters
    d = equal(d, "LAB", "GIA")
    d = equal(d, "SHAPE", "ROUND")
    if "RAP DISC%" in d.columns:
        d = d[d["RAP DISC%"] <= -10]
    d = between(d, "CTS", 0.30, 3.99, inclusive="both")

    # COL & CLAR cleanup (remove hyphens) + list filters
    if "COL" in d.columns:
        d["COL"] = d["COL"].str.replace("-", "", regex=False)
        d = d[d["COL"].isin(colors)]
    if "CLAR" in d.columns:
        d["CLAR"] = d["CLAR"].str.replace("-", "", regex=False)
        d = d[d["CLAR"].isin(clarity)]

    # Cut/Pol/Sym
    for c in ["CUT","POL","SYM"]:
        if c in d.columns:
            d = d[d[c].isin(cut)]

    # Fluor
    if "FL" in d.columns:
        d = d[d["FL"].isin(flour)]

    # TINGE must be NBG
    if "TINGE" in d.columns:
        d = d[d["TINGE"] == "NBG"]

    # MILKY keep blank or "NONE"
    if "MILKY" in d.columns:
        s = d["MILKY"].astype("string")
        mask = s.isna() | s.str.strip().str.upper().eq("NONE")
        d = d[mask]

    # LUSTER allowed or blank
    if "LUSTER" in d.columns:
        s = d["LUSTER"].astype("string")
        mask = s.isna() | s.str.strip().isin(luster_allow)
        d = d[mask]

    # Exclude CULET values
    if "CULET" in d.columns:
        d = d[~d["CULET"].isin(culet_block)]

    # Exclude FLCOLOR values
    if "FLCOLOR" in d.columns:
        d = d[~d["FLCOLOR"].isin(flcolor_block)]

    # KEY TO SYMBOL: exclude if contains "Cavity" or "Chip" anywhere
    if "KEY TO SYMBOL" in d.columns:
        s = d["KEY TO SYMBOL"].astype("string").fillna("")
        d = d[~s.str.contains(r"\b(?:Cavity|Chip)\b", na=False, regex=True)]

    # REPORT COMMENT: exclude if startswith "Clarity"
    if "REPORT COMMENT" in d.columns:
        s = d["REPORT COMMENT"].astype("string").fillna("")
        d = d[~s.str.startswith("Clarity")]

    # INCLUSION: remove listed codes with word boundaries
    if "INCLUSION" in d.columns and inclusion_codes:
        pattern = r"\b(?:{})\b".format("|".join(map(re.escape, inclusion_codes)))
        d["INCLUSION"] = (
            d["INCLUSION"]
            .fillna("")
            .str.replace(pattern, "", regex=True)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    return d


def fix_columns(filtered: pd.DataFrame) -> pd.DataFrame:
    """Create the exact output schema and fill from source columns."""
    out = pd.DataFrame(index=filtered.index)

    # fill mapped / defaults
    for tgt, src_candidates in columns_mapping.items():
        if tgt in ["LENGTH", "WIDTH", "HEIGHT", "RATIO", "RAP"]:
            out[tgt] = 0
        elif tgt == "SRNO":
            out[tgt] = range(1, len(filtered) + 1)
        elif tgt in ["TINCH", "LUSTER", "T OPENS", "B OPENS"]:
            out[tgt] = "N/A"
        elif tgt == "NATTS":
            out[tgt] = "CENTER - L"
        elif not src_candidates:
            out[tgt] = np.nan
        else:
            # take first available source column
            src = next((c for c in src_candidates if c in filtered.columns), None)
            out[tgt] = filtered[src] if src else np.nan

    # ensure column order
    out = out[list(columns_mapping.keys())].reset_index(drop=True)
    return out

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

# ---------- App ----------
def run_sanghvi_sons_mod() -> None:
    st.header("Sanghvi & Sons")

    sample_path = _find_sample("Sanghvi & Sons.xlsx")
    
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
        upload = st.file_uploader("Select Excel File", type=["xlsx", "xls"])
        skip = st.number_input("Skip Rows", 0, 10000, 5, 1)  # original used skiprows=5

    if upload is None:
        st.stop()

    try:
        df_raw = read_file(upload, skiprows=skip)
    except Exception as e:
        st.error(f"Failed to read file: {e}")
        st.stop()

    st.subheader("Raw Input (sanitized)")
    st.dataframe(df_raw)

    df_filtered = apply_filters(df_raw)

    df_final = fix_columns(df_filtered)
    st.subheader("Final Output (Mapped)")
    st.dataframe(df_final)

    # Download CSV
    file_name = f"sanghvi & sons {pd.Timestamp.now():%d.%m.%Y}.csv"
    csv_bytes = df_final.to_csv(index=False)  # BOM for Excel
    with st.sidebar:
        st.download_button(
            label="Download CSV",
            data=csv_bytes,
            file_name=file_name,
            mime="csv",
            
        )

if __name__ == "__main__":
    run_sanghvi_sons_mod()
