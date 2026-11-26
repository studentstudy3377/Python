from PIL import ImageSequence
import io
import os
import platform
import re
import subprocess
import streamlit as st
import numpy as np
import pandas as pd

def read_file(uploaded_file, skiprows: int = 0):
    """Read CSV/XLS/XLSX and clean obvious whitespace from column headers."""
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, skiprows=skiprows, keep_default_na=False)
    elif name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file,skiprows=skiprows,engine="openpyxl",keep_default_na=False,)

    elif name.endswith(".xls"):
        df = pd.read_excel(
            uploaded_file,
            skiprows=skiprows,
            engine="xlrd",
            keep_default_na=False,
        )
    else:
        raise ValueError("Unsupported file type")

    # Removes leading/trailing spaces, BOM and embedded line breaks in headers.
    df.columns = (
        pd.Index(df.columns)
        .astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.replace("\n", " ", regex=False)
        .str.strip()
    )
    return df


def _column_key(name):
    """Create a case/space/punctuation-insensitive key for a column name."""
    return re.sub(r"[^a-z0-9]+", "", str(name).strip().lower())


def normalize_market_columns(df):
    """Map common Market File header variations to the names used by this app."""
    aliases = {
        "lab": "LAB",
        "laboratory": "LAB",
        "cut": "CUT",
        "cutgrade": "CUT",
        "shape": "Shape",
        "fluorescence": "FLUORESCENCE",
        "fluorescenceintensity": "FLUORESCENCE",
        "fromsize": "FromSize",
        "color": "Color",
        "colour": "Color",
        "clarity": "Clarity",
        "averagediscountpercent": "AverageDiscountPercent",
        "diamondsincategory": "DiamondsInCategory",
    }

    rename_map = {}
    existing = set(df.columns)

    for col in df.columns:
        target = aliases.get(_column_key(col))
        if target and col != target and target not in existing:
            rename_map[col] = target
            existing.add(target)

    if rename_map:
        df = df.rename(columns=rename_map)

    return df


def require_cols(df, cols, df_name="dataframe"):
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise KeyError(
            f"Missing columns in {df_name}: {missing}. "
            f"Available columns: {df.columns.tolist()}"
        )


def normalize_market_size(value):
    """Return the lower/from carat size formatted to 2 decimals."""
    if pd.isna(value):
        return ""

    text = str(value).strip()
    if not text:
        return ""

    # Supports both numeric values (0.50) and ranges such as '0.50 - 0.59'.
    match = re.match(r"^\s*(\d+(?:\.\d+)?)", text)
    if not match:
        return text

    try:
        return f"{float(match.group(1)):.2f}"
    except ValueError:
        return text

def get_new_size(carat):
    try:
        carat = float(carat)
    except (TypeError, ValueError):
        return ""

    if 0.18 <= carat < 0.20:
        return "0.18"
    elif 0.20 <= carat <= 0.22:
        return "0.20"
    elif 0.23 <= carat <= 0.24:
        return "0.23"
    elif 0.25 <= carat <= 0.29:
        return "0.25"
    elif 0.30 <= carat <= 0.36:
        return "0.30"
    elif 0.37 <= carat <= 0.39:
        return "0.37"
    elif 0.40 <= carat <= 0.45:
        return "0.40"
    elif 0.46 <= carat <= 0.49:
        return "0.46"
    elif 0.50 <= carat <= 0.59:
        return "0.50"
    elif 0.60 <= carat <= 0.69:
        return "0.60"
    elif 0.70 <= carat <= 0.79:
        return "0.70"
    elif 0.80 <= carat <= 0.89:
        return "0.80"
    elif 0.90 <= carat <= 0.95:
        return "0.90"
    elif 0.96 <= carat <= 0.99:
        return "0.96"
    elif 1.00 <= carat <= 1.19:
        return "1.00"
    elif 1.20 <= carat <= 1.24:
        return "1.20"
    elif 1.25 <= carat <= 1.34:
        return "1.25"
    elif 1.35 <= carat <= 1.49:
        return "1.35"
    elif 1.50 <= carat <= 1.59:
        return "1.50"
    elif 1.60 <= carat <= 1.69:
        return "1.60"
    elif 1.70 <= carat <= 1.89:
        return "1.70"
    elif 1.90 <= carat <= 1.99:
        return "1.90"
    elif 2.00 <= carat <= 2.19:
        return "2.00"
    elif 2.20 <= carat <= 2.49:
        return "2.20"
    elif 2.50 <= carat <= 2.69:
        return "2.50"
    elif 2.70 <= carat <= 2.89:
        return "2.70"
    elif 2.90 <= carat <= 2.99:
        return "2.90"
    elif 3.00 <= carat <= 3.49:
        return "3.00"
    elif 3.50 <= carat <= 3.99:
        return "3.50"
    elif 4.00 <= carat <= 4.49:
        return "4.00"
    elif 4.50 <= carat <= 4.99:
        return "4.50"
    elif 5.00 <= carat <= 5.49:
        return "5.00"
    elif 5.50 <= carat <= 9.99:
        return "5.50"

    return ""

def get_cut_category(row):

    # Fancy shapes
    if row["Shape"] != "RD":
        return "VG+"

    cut = row["CUT_norm"]
    polish = row["POLISH_norm"]
    sym = row["SYM_norm"]

    # ROUND

    # EX + EX + EX
    if cut == "EX" and polish == "EX" and sym == "EX":
        return "3X"

    # Cut EX, Polish/Sym EX or VG
    if cut == "EX" and polish in ["EX", "VG"] and sym in ["EX", "VG"]:
        return "EX-"

    # Cut EX but one of Polish/Sym is GOOD
    if cut == "EX" and (polish == "GD" or sym == "GD"):
        return "GD+"

    # Cut Very Good
    if cut == "VG":
        return "VG+"

    # Cut Good
    if cut == "GD": 
        return "GD+"

    return ""

def process_files(df, dfs, sold, sold_price,promo,Stock_pds):
    """Process the uploaded files and return the modified dataframe."""
    # Ensure expected columns exist
    df = normalize_market_columns(df)
    require_cols(
        df,
        [
            "LAB",
            "CUT",
            "Shape",
            "FLUORESCENCE",
            "FromSize",
            "Color",
            "Clarity",
            "AverageDiscountPercent",
            "DiamondsInCategory",
        ],
        "market file",
    )

    df = df[df["LAB"].isin(["GIA"])].copy()
    df["FromSize"] = df["FromSize"].apply(normalize_market_size)
    # --- Filter Sold PDS ---
    sold = sold[sold["Cert"].isin(["GIA"])].copy()
    sold = sold[sold["Luster"].isin(["N/A", "NORMAL"])].copy()
    sold = sold[sold["Tinch"].isin(["N/A", ""]) | sold["Tinch"].isna()].copy()
    sold = sold[~sold["Natts No"].str.contains("-H|RED SPOT", na=False)].copy()

    # --- Filter Stock PDS ---
    Stock_pds = Stock_pds[Stock_pds["Cert"].isin(["GIA"])].copy()
    Stock_pds = Stock_pds[Stock_pds["Luster"].isin(["N/A", "NORMAL"])].copy()
    Stock_pds = Stock_pds[Stock_pds["Tinch"].isin(["N/A", ""]) | Stock_pds["Tinch"].isna()].copy()
    Stock_pds = Stock_pds[~Stock_pds["Natts No"].str.contains("-H|RED SPOT", na=False)].copy()

    grade_map = {
        "EXCELLENT": "EX",
        "EX": "EX",
        "V. GOOD": "VG",
        "V.GOOD": "VG",
        "VERY GOOD": "VG",
        "VG": "VG",
        "GOOD": "GD",
        "GD": "GD",
        "": "",
    }
    # Normalize CUT/POLISH/SYMMETRY for Stock PDS
    Stock_pds["CUT_norm"] = (Stock_pds["CUT"].astype(str).str.strip().str.upper().map(grade_map).fillna(""))
    Stock_pds["POLISH_norm"] = (Stock_pds["POLISH"].astype(str).str.strip().str.upper().map(grade_map).fillna(""))
    Stock_pds["SYM_norm"] = (Stock_pds["Symetry"].astype(str).str.strip().str.upper().map(grade_map).fillna(""))

    # Normalize CUT/POLISH/SYMMETRY for Sold PDS in the same format
    sold["CUT_norm"] = (sold["CUT"].astype(str).str.strip().str.upper().map(grade_map).fillna(""))
    sold["POLISH_norm"] = (sold["POLISH"].astype(str).str.strip().str.upper().map(grade_map).fillna(""))
    sold["SYM_norm"] = (sold["Symetry"].astype(str).str.strip().str.upper().map(grade_map).fillna(""))

    # --- Shape ---
    shape = {
        "Round": "RD",
        "Emerald": "EME",
        "Emrald":"EME",
        "Heart": "H",
        "Marquise": "MQ",
        "Marquies": "MQ",
        "Oval": "OV",
        "Princess": "PRI",
        "Radiant": "RDT",
        "Pear": "PRS",
        "Cushion Cut": "CUS",
        "Cushion Modified": "CUS",
        "Cushion": "CUS",
        "Sq Emerald":"SQ EME",
        "Sq. Emerald":"SQ EME",
        "Rd": "RD",
        "Eme": "EME",
        "H": "H",
        "Mq": "MQ",
        "Ov": "OV",
        "Pri": "PRI",
        "Rdt": "RDT",
        "Prs": "PRS",
        "Cus": "CUS",
    }
    shape_col = df["Shape"].astype(str).str.strip().str.title()
    shape_cols = dfs["shape"].astype(str).str.strip().str.title()
    shape_pds = Stock_pds["Shape"].astype(str).str.strip().str.title()
    shape_sold = sold["Shape"].astype(str).str.strip().str.title()
    df["Shape"] = shape_col.map(shape)
    dfs["Shap"] = shape_cols.map(shape)
    Stock_pds["Shape"] = shape_pds.map(shape)
    sold["Shape"] = shape_sold.map(shape)

    # Shape must be converted to RD/EME/... before calculating CUT_NEW.
    Stock_pds["CUT_NEW"] = Stock_pds.apply(get_cut_category, axis=1)
    sold["CUT_NEW"] = sold.apply(get_cut_category, axis=1)

    # --- FLUORESCENCE ---
    fluor_mar = {
        "None": "N",
        "Faint / Slight": "F",
        "Medium": "M",
        "Strong / Very Strong": "S",
        "Faint": "F",
        "Slight": "F",
        "Very Slight": "F",
        "Strong": "S",
        "Very Strong": "S",
        "N": "N",
        "F": "F",
        "M": "M",
        "S": "S",
    }
    fluor_col = df["FLUORESCENCE"].astype(str).str.strip().str.title()
    df["Fluor"] = fluor_col.map(fluor_mar).fillna("N")
    
    fluor = {
        "None": "N",
        "Faint": "F",
        "Medium": "M",
        "Med. Blue": "M",
        "Strong": "S",
        "Very Strong": "S",
        "Slight": "F",
        "Very Slight": "F",
        "V.Strong": "S",
        "N": "N",
        "F": "F",
        "M": "M",
        "S": "S",
    }
    fluor_cols = dfs["flrc"].astype(str).str.strip().str.title()
    fluor_pd = Stock_pds["FLRC"].astype(str).str.strip().str.title()
    fluor_sold = sold["FLRC"].astype(str).str.strip().str.title()
    dfs["Fluor_stock"] = fluor_cols.map(fluor)
    Stock_pds["FLRC"] = fluor_pd.map(fluor)
    sold["FLRC"] = fluor_sold.map(fluor)
    
    size_str = df["FromSize"].astype(str)
    lower = pd.to_numeric(size_str.str.extract(r"^\s*([\d.]+)\s*-\s*")[0], errors="coerce")
    df["Size_mar"] = lower.apply(lambda x: f"{float(x):.2f}" if pd.notna(x) else None)

    size_str = dfs["sizedesc"].astype(str)
    lowers = pd.to_numeric(size_str.str.extract(r"^\s*([\d.]+)\s*TO")[0], errors="coerce")
    dfs["LowerSize"] = lowers.apply(lambda x: f"{float(x):.2f}" if pd.notna(x) else None)

    Stock_pds["Cts"] = pd.to_numeric(Stock_pds["Cts"], errors="coerce")
    Stock_pds["new_size"] = Stock_pds["Cts"].apply(get_new_size)

    # Sold PDS uses the exact same size buckets as Stock PDS
    sold["Cts"] = pd.to_numeric(sold["Cts"], errors="coerce")
    sold["new_size"] = sold["Cts"].apply(get_new_size)
    
    mapping = {
        ("EX", "2EX"): "3X",
        ("EX", "1EX"): "EX-",
        ("EX", "VG"): "EX-",
        ("EX", "2VG"): "EX-",
        ("EX", "GD"): "GD+",
        ("VG", "2EX"): "VG+",
        ("VG", "1EX"): "VG+",
        ("VG", "VG"): "VG+",
        ("VG", "2VG"): "VG+",
        ("VG", "GD"): "VG+",
        ("VG", "2GD"): "VG+",
        ("GD", "GD"): "GD+",
        ("GD", "2GD"): "GD+",
        ("GD", "1EX"): "GD+",
        ("GD", "2EX"): "GD+",
        ("GD", "VG"): "GD+",
        ("GD", "2VG"): "GD+",
    }

    dfs["Sym"] = dfs.apply(lambda row: mapping.get((row["cut"], row["polsym"]))if row["shape"] == "RD" else "VG+",axis=1)
    
    # --- Concatenate columns ---
    dfs["Concatenate_stock"] = (
        dfs["LowerSize"].astype(str)
        + "-"
        + dfs["Shap"].astype(str)
        + "-"
        + dfs["color"].astype(str)
        + "-"
        + dfs["clarity"].astype(str)
        + "-"
        + dfs["Fluor_stock"].astype(str)
        + "-"
        + dfs["Sym"].astype(str)
    )

    df["Concatenate"] = (
        df["FromSize"].astype(str)
        + "-"
        + df["Shape"].astype(str)
        + "-"
        + df["Color"].astype(str)
        + "-"
        + df["Clarity"].astype(str)
        + "-"
        + df["Fluor"].astype(str)
        + "-"
        + df["CUT"].astype(str)
    )
    
    Stock_pds["Concatenate_psd_stock"] = (
        Stock_pds["new_size"].astype(str)
        + "-"
        + Stock_pds["Shape"].astype(str)
        + "-"
        + Stock_pds["COLOR"].astype(str)
        + "-"
        + Stock_pds["Clarity"].astype(str)
        + "-"
        + Stock_pds["FLRC"].astype(str)
        + "-"
        + Stock_pds["CUT_NEW"].astype(str)
    )

    # Sold PDS key: same exact format as Stock PDS
    sold["Concatenate_psd_sold"] = (
        sold["new_size"].astype(str)
        + "-"
        + sold["Shape"].astype(str)
        + "-"
        + sold["COLOR"].astype(str)
        + "-"
        + sold["Clarity"].astype(str)
        + "-"
        + sold["FLRC"].astype(str)
        + "-"
        + sold["CUT_NEW"].astype(str)
    )

    # --- Prepare right table for merge ---
    right_cols = [
        "Concatenate",
        "AverageDiscountPercent",
        "DiamondsInCategory"
    ]

    # Aggregate market data by Concatenate before merge
    market_agg = (
        df[right_cols]
        .groupby("Concatenate", as_index=False)
        .agg({
            "AverageDiscountPercent": "min",
            "DiamondsInCategory": "min",
        })
    )

    # merge aggregated result
    dfs = dfs.merge(market_agg,left_on="Concatenate_stock",right_on="Concatenate",how="left")
    dfs = dfs.drop_duplicates(subset='idno', keep='first')
    
    sold_price = sold_price.merge(sold[["Sr. No","Concatenate_psd_sold","RAPAPORT"]],left_on="Stone Id",right_on="Sr. No",)
    sold_price["Sale Cost"] = pd.to_numeric(sold_price["Sale Cost"].astype(str).str.replace(",", ""),errors="coerce")  
    sold_price["RAPAPORT"] = pd.to_numeric(sold_price["RAPAPORT"].astype(str).str.replace(",", ""),errors="coerce") 
    sold_price["Sold Price Disc"] = ((sold_price["Sale Cost"]-sold_price["RAPAPORT"])/sold_price["RAPAPORT"]*100).round(0)

    dfs = dfs.merge(sold_price[["Concatenate_psd_sold","Sold Price Disc"]],left_on="Concatenate_stock", right_on="Concatenate_psd_sold",how="left",)
    
    Stock_pds = Stock_pds.merge(promo[["STOCKID","PDISC"]],left_on="Sr. No", right_on="STOCKID",how="left",)
    
    dfs = dfs.merge(Stock_pds[["Concatenate_psd_stock","PDISC"]],left_on="Concatenate_stock", right_on="Concatenate_psd_stock",how="left",)

    average_df = dfs.groupby('idno', as_index=False)[['Sold Price Disc']].mean().round(0)

    # Merge the calculated averages back into the original dataframe (if you need them alongside other data)
    dfs = dfs.merge(average_df, on='idno', suffixes=('', ' '))
    
    dfs.rename(columns={"DISC": "Sold Disc"}, inplace=True)

    dfs = dfs.drop_duplicates(subset='idno', keep='first')
    
    return dfs

desired_order = [
    "idno","shape","sizedesc","fromsize","tosize","color",
    "clarity","flrc","cut","polsym","qlty","cutmap",
    "curdisc","newdisc","curtrend","diff","discchgdays",
    "Sold Price Disc ","PDISC","AvgRapListPercent10","TotalRows"]
 
def run_market_disc():
    st.title("Market Disc")
    with st.spinner("Process.."):
        with st.sidebar:
            st.header("📂 Upload Files")
            file = st.file_uploader("Market File", type=["xlsx", "xls", "csv"])
            files = st.file_uploader("Stock File", type=["xlsx", "xls", "csv"])
            Stock_pds = st.file_uploader("Stock List PDS", type=["xlsx", "xls", "csv"])
            Sold = st.file_uploader("Sold List PDS", type=["xlsx", "xls", "csv"])
            sold_price = st.file_uploader("Sold Price E Skpi - 4", type=["xlsx", "xls", "csv"])
            promo_file = st.file_uploader("Promo File", type=["xlsx", "xls", "csv"])
            
        if file is None or files is None or Sold is None or sold_price is None or promo_file is None or Stock_pds is None:
            st.info("Files")
            st.stop()

        df = read_file(file)   # Market file
        dfs = read_file(files) # Stock file
        sold = read_file(Sold) # Stock file
        sold_price = read_file(sold_price,skiprows=4)
        promo = read_file(promo_file)
        Stock_pds = read_file(Stock_pds)
        
        try:
            dfs = process_files(df, dfs, sold, sold_price, promo, Stock_pds)
        except KeyError as exc:
            st.error(str(exc))
            st.stop()

        dfs.rename(
                    columns={
                        "AverageDiscountPercent":"AvgRapListPercent10",
                        "DiamondsInCategory":"TotalRows",
                    },
                    inplace=True,
                )

        for col in desired_order:
            if col not in dfs.columns:
                dfs[col] = None
        # Reorder columns
        dfs = dfs[desired_order]

        col = "AvgRapListPercent10"

        dfs[col] = (dfs[col].astype("string").str.replace("%", "", regex=False).str.strip().pipe(pd.to_numeric, errors="coerce").abs().mul(-1))
        dfs[col] = dfs[col].map(lambda x: f"{x:.2f}" if pd.notna(x) else "")
        
        dfs["Sold Price Disc "] = dfs["Sold Price Disc "].map( lambda x: f"{x:.2f}").replace({"nan":""})
        dfs["PDISC"] = dfs["PDISC"].map( lambda x: f"{x:.2f}").replace({"nan":""})
        # Build downloadable Excel in memory
        stock_file_name = os.path.splitext(files.name)[0] or "stock"
        out_name = f"{stock_file_name}.xlsx"
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            dfs.to_excel(writer, index=False, sheet_name="Data")
        buffer.seek(0)
        st.dataframe(dfs)
        st.download_button(
            label="⬇️ Download Excel",
            data=buffer,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
if __name__ == "__main__":
    run_market_disc()
