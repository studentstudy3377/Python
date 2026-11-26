import os
import re
import html
import hashlib
import zipfile
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook


# ----------------------------------------------------------------------------
# Excel engine
# ----------------------------------------------------------------------------

try:
    import python_calamine  # type: ignore # noqa: F401
    _FAST_ENGINE = "calamine"
except ImportError:  # still works, just slower
    _FAST_ENGINE = None


def _read_excel(file_data, **kwargs):
    """pd.read_excel using calamine when available, else pandas default."""
    if _FAST_ENGINE:
        try:
            return pd.read_excel(BytesIO(file_data), engine=_FAST_ENGINE, **kwargs)
        except (ValueError, ImportError, TypeError):
            pass
    return pd.read_excel(BytesIO(file_data), **kwargs)


# ----------------------------------------------------------------------------
# HYPERLINK helpers (unchanged logic)
# ----------------------------------------------------------------------------

# Excel HYPERLINK formulas sometimes have no cached display value.
# Recover the visible label from the source formula, not the link URL.
_HYPERLINK_START = re.compile(r"^\s*=+\s*HYPERLINK\s*\(", re.IGNORECASE)
_HYPERLINK_BYTES = re.compile(rb"hyperlink\s*\(", re.IGNORECASE)


def reset_file_pointer(file_obj):
    """Restore an uploaded stream's cursor before another read."""
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)


def hyperlink_arguments(value):
    """Split =HYPERLINK(link, label) while respecting quoted strings."""
    if not isinstance(value, str):
        return None
    match = _HYPERLINK_START.match(value)
    if match is None:
        return None
    content = value[match.end():].strip()
    if not content.endswith(")"):
        return None
    content = content[:-1]
    arguments = []
    start = 0
    depth = 0
    quoted = False
    i = 0
    while i < len(content):
        char = content[i]
        if char == '"':
            if quoted and i + 1 < len(content) and content[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
        elif not quoted:
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth < 0:
                    return None
            elif char in ",;" and depth == 0:
                arguments.append(content[start:i].strip())
                start = i + 1
        i += 1
    if quoted or depth != 0:
        return None
    arguments.append(content[start:].strip())
    if not arguments or any(arg == "" for arg in arguments):
        return None
    return arguments


def excel_literal(value):
    """Return the text of an Excel string/number literal, not an expression."""
    if not isinstance(value, str):
        return None
    value = value.strip()
    if re.fullmatch(r'"(?:[^"]|"")*"', value, re.DOTALL):
        return value[1:-1].replace('""', '"')
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", value):
        return value
    return None


def hyperlink_display(value, cached_value=None):
    """Return the hyperlink's displayed label (never the full formula)."""
    args = hyperlink_arguments(value)
    if args is None:
        return value
    if len(args) >= 2:
        label = excel_literal(args[1])
        if label is not None:
            return label
    if cached_value is not None and not pd.isna(cached_value):
        if not isinstance(cached_value, str) or not _HYPERLINK_START.match(cached_value):
            return cached_value
    if len(args) == 1:
        return excel_literal(args[0]) or ""
    return ""


def clean_hyperlinks_in_dataframe(data):
    """Clean literal =HYPERLINK(...) strings in any Excel or CSV column."""
    for col_idx in range(data.shape[1]):
        series = data.iloc[:, col_idx]
        if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
            # Fast skip: only touch columns that actually contain HYPERLINK text.
            as_text = series.astype(str)
            if not as_text.str.contains("HYPERLINK", case=False, regex=False).any():
                continue
            data.isetitem(col_idx, series.map(hyperlink_display))
    return data


@st.cache_data(show_spinner=False, max_entries=256)
def _xlsx_has_hyperlink_text(file_key, _file_data):
    """Quick zip scan: does any sheet / shared string contain 'HYPERLINK('?"""
    try:
        with zipfile.ZipFile(BytesIO(_file_data)) as zf:
            for name in zf.namelist():
                lower = name.lower()
                if not (
                    (lower.startswith("xl/worksheets/") and lower.endswith(".xml"))
                    or lower == "xl/sharedstrings.xml"
                ):
                    continue
                with zf.open(name) as fh:
                    tail = b""
                    while True:
                        chunk = fh.read(4 << 20)
                        if not chunk:
                            break
                        buf = tail + chunk
                        if _HYPERLINK_BYTES.search(buf):
                            return True
                        tail = buf[-64:]
    except zipfile.BadZipFile:
        return True  # unknown layout: let the full recovery decide
    return False


def recover_excel_hyperlinks(file_data, data, skip_rows=0, sheet_name=None):
    """Recover formula display labels in ALL columns of a workbook sheet."""
    if data.empty:
        return data

    workbook = load_workbook(BytesIO(file_data), data_only=False, read_only=True)
    try:
        if sheet_name is None:
            sheet = workbook.worksheets[0]
        elif isinstance(sheet_name, int):
            sheet = workbook.worksheets[sheet_name]
        else:
            sheet = workbook[sheet_name]

        header_row = int(skip_rows or 0) + 1
        changes = {}
        rows = sheet.iter_rows(
            min_row=header_row + 1,
            max_row=header_row + len(data),
            min_col=1,
            max_col=data.shape[1],
            values_only=True,
        )
        for row_idx, row in enumerate(rows):
            for col_idx, formula in enumerate(row):
                if not isinstance(formula, str) or "(" not in formula:
                    continue
                if hyperlink_arguments(formula) is None:
                    continue
                cached_value = data.iat[row_idx, col_idx]
                changes.setdefault(col_idx, {})[row_idx] = hyperlink_display(
                    formula, cached_value
                )

        for col_idx, updates in changes.items():
            values = data.iloc[:, col_idx].astype(object).to_numpy(copy=True)
            for row_idx, display in updates.items():
                values[row_idx] = display
            data.isetitem(col_idx, pd.Series(values, index=data.index, dtype=object))
    finally:
        workbook.close()
    return data


# ----------------------------------------------------------------------------
# Column matching (unchanged logic)
# ----------------------------------------------------------------------------

def normalize_column_name(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().casefold())


def resolve_column_name(columns, requested):
    if not requested:
        return None
    columns = list(columns)
    if requested in columns:
        return requested
    requested_key = normalize_column_name(requested)
    matches = [c for c in columns if normalize_column_name(c) == requested_key]
    if len(matches) == 1:
        return matches[0]
    return None


def resolve_required_columns(columns, requested_columns):
    resolved = {}
    for requested in requested_columns:
        if not requested:
            continue
        actual = resolve_column_name(columns, requested)
        if actual is None:
            return None
        resolved[requested] = actual
    return resolved


# ----------------------------------------------------------------------------
# File identity: hash each upload ONCE
# ----------------------------------------------------------------------------

def uploaded_file_bytes(file_obj) -> bytes:
    if hasattr(file_obj, "getvalue"):
        return file_obj.getvalue()
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    data = file_obj.read()
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)
    return data


def file_key(file_obj):
    """(name, size, md5) - computed once and remembered on the upload object."""
    cached = getattr(file_obj, "_bm_file_key", None)
    if cached is not None:
        return cached
    data = uploaded_file_bytes(file_obj)
    key = (
        getattr(file_obj, "name", ""),
        len(data),
        hashlib.md5(data).hexdigest(),
    )
    try:
        file_obj._bm_file_key = key
    except AttributeError:
        pass
    return key


# ----------------------------------------------------------------------------
# Cached readers (hash = small file_key, raw bytes are NOT re-hashed)
# ----------------------------------------------------------------------------

_CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin1")


@st.cache_data(show_spinner=False, max_entries=512)
def _cached_header_columns(key, _file_data, skiprows=0, sheet_name=None):
    """Read only column headers."""
    ext = os.path.splitext(key[0])[1].lower()
    skiprows = int(skiprows or 0)

    if ext == ".csv":
        for encoding in _CSV_ENCODINGS:
            try:
                frame = pd.read_csv(
                    BytesIO(_file_data), encoding=encoding,
                    skiprows=skiprows, nrows=0, keep_default_na=False,
                )
                return tuple(frame.columns)
            except UnicodeDecodeError:
                continue
        return tuple()

    if ext in {".xls", ".xlsx", ".xlsm"}:
        kwargs = {"skiprows": skiprows, "nrows": 0, "keep_default_na": False}
        if sheet_name is not None:
            kwargs["sheet_name"] = sheet_name
        return tuple(_read_excel(_file_data, **kwargs).columns)

    return tuple()


def read_uploaded_headers(file_obj, skiprows=0, sheet_name=None):
    if file_obj is None:
        return []
    return list(
        _cached_header_columns(
            file_key(file_obj), uploaded_file_bytes(file_obj),
            int(skiprows or 0), sheet_name,
        )
    )


@st.cache_data(show_spinner=False, max_entries=256)
def _cached_read_table(key, _file_data, skiprows=0, sheet_name=None):
    """Parse a file once and cache the DataFrame across Streamlit reruns."""
    ext = os.path.splitext(key[0])[1].lower()
    skiprows = int(skiprows or 0)

    if ext == ".csv":
        last_error = None
        for encoding in _CSV_ENCODINGS:
            try:
                data = pd.read_csv(
                    BytesIO(_file_data), encoding=encoding,
                    skiprows=skiprows, keep_default_na=False,
                )
                return clean_hyperlinks_in_dataframe(data)
            except UnicodeDecodeError as exc:
                last_error = exc
        if last_error:
            raise last_error

    if ext in {".xls", ".xlsx", ".xlsm"}:
        kwargs = {"skiprows": skiprows, "keep_default_na": False}
        if sheet_name is not None:
            kwargs["sheet_name"] = sheet_name
        data = _read_excel(_file_data, **kwargs)

        if ext in {".xlsx", ".xlsm"}:
            # Expensive formula pass only if the workbook has HYPERLINK at all.
            if _xlsx_has_hyperlink_text(key, _file_data):
                data = recover_excel_hyperlinks(
                    _file_data, data, skip_rows=skiprows, sheet_name=sheet_name,
                )
            return data

        return clean_hyperlinks_in_dataframe(data)

    raise ValueError(f"Unsupported file type: {ext or 'unknown'}")


@st.cache_data(show_spinner=False, max_entries=32)
def _cached_sheet_names(key, _file_data):
    ext = os.path.splitext(key[0])[1].lower()
    if ext not in {".xlsx", ".xlsm"}:
        raise ValueError("DEAL must be an .xlsx or .xlsm workbook.")
    workbook = load_workbook(BytesIO(_file_data), read_only=True)
    try:
        return tuple(workbook.sheetnames)
    finally:
        workbook.close()


def read_uploaded_table(file_obj, skiprows=None, sheet_name=None):
    if file_obj is None:
        raise ValueError("No file provided.")
    data = _cached_read_table(
        file_key(file_obj), uploaded_file_bytes(file_obj),
        int(skiprows or 0), sheet_name,
    )
    return data.copy()


# ----------------------------------------------------------------------------
# Sheet / filename helpers
# ----------------------------------------------------------------------------

def normalize_sheet_text(value: str) -> str:
    return " ".join(str(value).replace("\xa0", " ").split()).casefold()


def sheet_name_starting_with(sheet_names, sheet_prefix):
    if not sheet_prefix:
        return None
    prefix = normalize_sheet_text(sheet_prefix)
    matches = [n for n in sheet_names if normalize_sheet_text(n).startswith(prefix)]
    if not matches:
        return None
    for name in matches:
        if normalize_sheet_text(name) == prefix:
            return name
    return matches[0]


def normalize_filename(value: str) -> str:
    base = os.path.splitext(os.path.basename(str(value)))[0]
    return "".join(ch for ch in base.upper() if ch.isalnum())


def get_deal_sheet_names(deal_file_obj) -> list[str]:
    return list(
        _cached_sheet_names(file_key(deal_file_obj), uploaded_file_bytes(deal_file_obj))
    )


def find_sheet_starting_with(deal_file_obj, sheet_prefix: str) -> str:
    if not sheet_prefix:
        raise ValueError("Sheet prefix is required.")
    sheet_names = get_deal_sheet_names(deal_file_obj)
    match = sheet_name_starting_with(sheet_names, sheet_prefix)
    if match is None:
        available = ", ".join(sheet_names) if sheet_names else "(none)"
        raise ValueError(
            f"No DEAL sheet starts with '{sheet_prefix}'. "
            f"Available sheets: {available}"
        )
    return match


def match_uploaded_file(
    files, prefix, required_columns=None, skiprows=None,
    sheet_name=None, fallback_files=None,
):
    """Filename prefix first; otherwise header-only inspection."""
    if not files:
        return None

    prefix_normalized = normalize_filename(prefix)
    filename_matches = [
        f for f in files if normalize_filename(f.name).startswith(prefix_normalized)
    ]
    if filename_matches:
        return filename_matches[-1]

    required_columns = [c for c in (required_columns or []) if c]
    if not required_columns:
        return None

    candidates = list(fallback_files) if fallback_files is not None else list(files)
    data_matches = []
    for file_obj in candidates:
        try:
            configured_skip = int(skiprows or 0)
            skip_attempts = [configured_skip]
            if configured_skip != 0:
                skip_attempts.append(0)
            for header_skip in skip_attempts:
                columns = read_uploaded_headers(
                    file_obj, skiprows=header_skip, sheet_name=sheet_name,
                )
                if resolve_required_columns(columns, required_columns) is not None:
                    data_matches.append(file_obj)
                    break
        except (OSError, ValueError, ImportError, KeyError):
            continue

    return data_matches[0] if len(data_matches) == 1 else None


# ----------------------------------------------------------------------------
# ID normalization (vectorized over unique values)
# ----------------------------------------------------------------------------

_TAG_RE = re.compile(r"<[^>]*>")
_FLOAT_INT_RE = re.compile(r"[+-]?\d+\.0+")


def normalize_stock_id(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    text = html.unescape(str(value)).strip().upper()
    text = _TAG_RE.sub("", text).strip()
    if _FLOAT_INT_RE.fullmatch(text):
        text = text.split(".", 1)[0]
    normalized = "".join(ch for ch in text if ch.isalnum())
    return normalized or None


def normalize_id_series(series: pd.Series) -> pd.Series:
    """normalize_stock_id applied once per distinct value, then broadcast."""
    codes, uniques = pd.factorize(series, use_na_sentinel=True)
    lookup = np.array(
        [normalize_stock_id(u) for u in uniques] + [None], dtype=object
    )
    return pd.Series(lookup[codes], index=series.index, dtype=object)


# ----------------------------------------------------------------------------
# Display / write helpers
# ----------------------------------------------------------------------------

def excel_safe_value(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "item"):
        try:
            return value.item()
        except (ValueError, AttributeError):
            pass
    return value


def _excel_rows(data: pd.DataFrame):
    """All rows as Python lists with NaN/NaT -> None, in one vectorized pass."""
    obj = data.astype(object)
    obj = obj.where(data.notna(), None)
    rows = obj.to_numpy(dtype=object).tolist()
    # Rare leftover numpy scalars (e.g. np.datetime64) -> plain Python.
    for row in rows:
        for i, v in enumerate(row):
            if v is not None and type(v).__module__ == "numpy":
                row[i] = excel_safe_value(v)
    return rows


def streamlit_safe_dataframe(data: pd.DataFrame) -> pd.DataFrame:
    """Display-only copy that PyArrow can serialize (object -> string)."""
    safe = data.copy()
    for column in safe.columns:
        series = safe[column]
        if pd.api.types.is_object_dtype(series.dtype):
            safe[column] = series.map(
                lambda value: (
                    pd.NA
                    if value is None
                    or (
                        not isinstance(value, (list, dict, tuple, set))
                        and pd.isna(value)
                    )
                    else str(value)
                )
            ).astype("string")
    return safe


# ----------------------------------------------------------------------------
# Core merge
# ----------------------------------------------------------------------------

def read_org_table_with_skip_fallback(file_obj, configured_skiprows, merge_column, sheet_name=None):
    configured_skip = int(configured_skiprows or 0)
    skip_attempts = [configured_skip]
    if configured_skip != 0:
        skip_attempts.append(0)

    last_columns = []
    for skip in skip_attempts:
        data = read_uploaded_table(file_obj, skiprows=skip, sheet_name=sheet_name)
        last_columns = list(data.columns)
        actual_merge_column = resolve_column_name(data.columns, merge_column)
        if actual_merge_column is not None:
            return data, skip, actual_merge_column

    attempted = ", ".join(str(v) for v in skip_attempts)
    raise ValueError(
        f"ORG merge column '{merge_column}' was not found "
        f"in file '{getattr(file_obj, 'name', '')}'. "
        f"Tried skip_rows: {attempted}. "
        f"Found columns: {last_columns[:20]}"
    )


OUTPUT_COLUMNS = [
    'dLot', 'Stock Id', 'dSize', 'dWt', 'Lab', 'dShape', 'dColor',
    'dClarity', 'dFlrc', 'cutName', 'dCut', 'dPol', 'dSym', 'dTable',
    'dDepth', 'dDisc', 'dRap', 'Rap value', 'dValue', 'dDisc1',
    'dValue1', 'M_DAYS', 'M_TREND', 'tot_Pcs', 'dSoldPrice',
    'dSoldDisc', 'dNatts_1', 'dNatts_2',
]


def build_processed_dataframe(
    org_file_obj, deal_file_obj, org_skiprows=None, org_merge_col=None,
    deal_merge_col=None, dNatts_1=None, dNatts_2=None, stock_id=None,
    rapvalue=None, deal_sheet_prefix=None, org_sheet_name=None,
):
    if not deal_sheet_prefix:
        raise ValueError("DEAL sheet prefix is required.")

    actual_deal_sheet_name = find_sheet_starting_with(deal_file_obj, deal_sheet_prefix)
    deal_df = read_uploaded_table(deal_file_obj, sheet_name=actual_deal_sheet_name)

    actual_deal_merge_col = resolve_column_name(deal_df.columns, deal_merge_col)
    if actual_deal_merge_col is None:
        raise ValueError(
            f"DEAL merge column '{deal_merge_col}' was not found "
            f"in sheet '{actual_deal_sheet_name}'."
        )

    org_df, _actual_skip, actual_org_merge_col = read_org_table_with_skip_fallback(
        file_obj=org_file_obj, configured_skiprows=org_skiprows,
        merge_column=org_merge_col, sheet_name=org_sheet_name,
    )

    configured_optional = {
        "dNatts_1": dNatts_1,
        "dNatts_2": dNatts_2,
        "Stock Id": stock_id,
        "Rap value": rapvalue,
    }
    actual_optional = {}
    for destination, configured_source in configured_optional.items():
        actual_optional[destination] = (
            resolve_column_name(org_df.columns, configured_source)
            if configured_source else None
        )

    deal_df["_merge_key"] = normalize_id_series(deal_df[actual_deal_merge_col])

    merge_columns = [actual_org_merge_col]
    for actual_source in actual_optional.values():
        if actual_source and actual_source not in merge_columns:
            merge_columns.append(actual_source)

    org_slice = org_df[merge_columns].copy()
    org_slice["_merge_key"] = normalize_id_series(org_df[actual_org_merge_col])

    merged_df = pd.merge(deal_df, org_slice, on="_merge_key", how="left")

    final_df = merged_df.drop(columns=[actual_org_merge_col, "_merge_key"], errors="ignore")

    dlot_length = deal_df[actual_deal_merge_col].replace("", pd.NA).notna().sum()
    final_df = final_df.head(int(dlot_length))

    for destination, actual_source in actual_optional.items():
        if actual_source and actual_source in final_df.columns and actual_source != destination:
            final_df[destination] = final_df[actual_source]

    cols_to_remove = [
        src for dst, src in actual_optional.items() if src and src != dst
    ]
    final_df = final_df.drop(columns=cols_to_remove, errors="ignore")

    for col in OUTPUT_COLUMNS:
        if col not in final_df.columns:
            final_df[col] = None
    final_df = final_df[OUTPUT_COLUMNS]
    final_df["Rap value"] = None  # must always stay blank

    return actual_deal_sheet_name, final_df


def _blank_series(series: pd.Series) -> pd.Series:
    mask = series.isna()
    if pd.api.types.is_object_dtype(series.dtype) or pd.api.types.is_string_dtype(series.dtype):
        text_values = series.astype("string")
        mask = mask | text_values.str.strip().eq("").fillna(False)
    return mask


def _deal_result_index(data: pd.DataFrame) -> pd.MultiIndex:
    row_numbers = pd.Series(np.arange(len(data)), index=data.index, dtype="int64")
    fallback = "__ROW__" + row_numbers.astype(str)
    if "dLot" not in data.columns:
        keys = fallback
    else:
        keys = normalize_id_series(data["dLot"])
        keys = keys.where(keys.notna(), fallback)
    occurrence = keys.groupby(keys, sort=False).cumcount()
    return pd.MultiIndex.from_arrays([keys, occurrence], names=["__dlot_key", "__occurrence"])


def merge_same_deal_sheet_results(existing_df, new_df):
    """Existing nonblank values win; incoming fills blanks (aligned on dLot)."""
    if existing_df is None or existing_df.empty:
        return new_df.copy()
    if new_df is None or new_df.empty:
        return existing_df.copy()

    existing = existing_df.copy()
    incoming = new_df.copy()

    all_columns = list(existing.columns)
    all_columns.extend(c for c in incoming.columns if c not in all_columns)

    existing = existing.reindex(columns=all_columns)
    incoming = incoming.reindex(columns=all_columns)

    existing.index = _deal_result_index(existing)
    incoming.index = _deal_result_index(incoming)

    extra_index = incoming.index.difference(existing.index, sort=False)
    union_index = existing.index.append(extra_index)

    existing = existing.reindex(union_index)
    incoming = incoming.reindex(union_index)

    result = existing.copy()
    for column in all_columns:
        blank = _blank_series(result[column])
        result[column] = result[column].where(~blank, incoming[column])

    return result.reset_index(drop=True)


def add_or_merge_processed_sheet(processed_sheets, target_sheet, new_df):
    normalized_target = normalize_sheet_text(target_sheet)
    for index, (sheet_name, existing_df) in enumerate(processed_sheets):
        if normalize_sheet_text(sheet_name) == normalized_target:
            processed_sheets[index] = (
                sheet_name, merge_same_deal_sheet_results(existing_df, new_df),
            )
            return
    processed_sheets.append((target_sheet, new_df.copy()))


def _drop_stale_cells(ws, new_max_row, new_max_col):
    """Remove old cells outside the new data area (fast replacement for
    ws.delete_rows / ws.delete_cols on trailing rows/columns)."""
    cells = ws._cells  # openpyxl internal {(row, col): Cell}
    stale = [k for k in cells if k[0] > new_max_row or k[1] > new_max_col]
    for k in stale:
        del cells[k]


def _write_all_processed_sheets_openpyxl(deal_file_obj, processed_sheets) -> BytesIO:
    deal_name = getattr(deal_file_obj, "name", "")
    ext = os.path.splitext(deal_name)[1].lower()
    if ext not in {".xlsx", ".xlsm"}:
        raise ValueError("To preserve all DEAL worksheets, upload DEAL as .xlsx or .xlsm.")

    workbook = load_workbook(
        BytesIO(uploaded_file_bytes(deal_file_obj)),
        keep_vba=(ext == ".xlsm"),
    )
    try:
        for target_sheet, final_df in processed_sheets:
            if target_sheet not in workbook.sheetnames:
                raise ValueError(f"Sheet '{target_sheet}' was not found in DEAL workbook.")

            ws = workbook[target_sheet]
            ws.freeze_panes = None

            new_max_row = len(final_df) + 1
            new_max_col = len(final_df.columns)

            cell = ws.cell
            for col_idx, column_name in enumerate(final_df.columns, start=1):
                cell(row=1, column=col_idx, value=str(column_name))

            for row_idx, values in enumerate(_excel_rows(final_df), start=2):
                for col_idx, value in enumerate(values, start=1):
                    cell(row=row_idx, column=col_idx, value=value)

            _drop_stale_cells(ws, new_max_row, new_max_col)

        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        return output
    finally:
        workbook.close()



# ----------------------------------------------------------------------------
# FAST writer: patch only the processed sheets inside the .xlsx zip.
# Every other part of the workbook (other sheets, VBA, images, charts, styles)
# is copied byte-for-byte. Falls back to openpyxl if anything is unusual.
# ----------------------------------------------------------------------------

import datetime as _dt
import math as _math
import posixpath as _posixpath
import xml.etree.ElementTree as _ET

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_CELL_TAG = re.compile(rb"<c\b([^>]*)>|<c\b([^>]*)/>")
_ROW_TAG = re.compile(rb"<row\b([^>]*)>")
_ATTR = re.compile(rb'([A-Za-z:]+)="([^"]*)"')
_KEEP_ROW_ATTRS = (b"ht", b"customHeight", b"hidden", b"s", b"customFormat")


class _FallbackToOpenpyxl(Exception):
    pass


def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _xml_text(value):
    text = _ILLEGAL_XML.sub("", str(value))
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )


def _sheet_part_paths(zf):
    """{sheet name: 'xl/worksheets/sheetN.xml'}"""
    wb = _ET.fromstring(zf.read("xl/workbook.xml"))
    rels = _ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    targets = {}
    for rel in rels.findall(f"{{{_NS_PKG_REL}}}Relationship"):
        target = rel.get("Target", "")
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = _posixpath.normpath(_posixpath.join("xl", target))
        targets[rel.get("Id")] = path
    out = {}
    for sheet in wb.iter(f"{{{_NS_MAIN}}}sheet"):
        out[sheet.get("name")] = targets.get(sheet.get(f"{{{_NS_REL}}}id"))
    return out


def _render_sheet_xml(old_xml: bytes, final_df: pd.DataFrame) -> bytes:
    start = old_xml.find(b"<sheetData")
    if start < 0:
        raise _FallbackToOpenpyxl("sheetData not found (prefixed namespace?)")
    tag_end = old_xml.find(b">", start)
    if old_xml[tag_end - 1:tag_end] == b"/":
        end = tag_end + 1
        old_data = b""
    else:
        close = old_xml.find(b"</sheetData>", tag_end)
        if close < 0:
            raise _FallbackToOpenpyxl("unterminated sheetData")
        end = close + len(b"</sheetData>")
        old_data = old_xml[tag_end + 1:close]

    n_rows = len(final_df) + 1
    n_cols = len(final_df.columns)
    letters = [_col_letter(i) for i in range(1, n_cols + 1)]

    # Keep existing cell styles / row formatting inside the new data area.
    cell_style = {}
    for m in _CELL_TAG.finditer(old_data):
        attrs = dict(_ATTR.findall(m.group(1) or m.group(2)))
        style = attrs.get(b"s")
        ref = attrs.get(b"r")
        if style and style != b"0" and ref:
            cell_style[ref.decode()] = style.decode()
    row_attrs = {}
    for m in _ROW_TAG.finditer(old_data):
        attrs = dict(_ATTR.findall(m.group(1)))
        r = attrs.get(b"r")
        if r is None:
            continue
        kept = " ".join(
            f'{k.decode()}="{attrs[k].decode()}"' for k in _KEEP_ROW_ATTRS if k in attrs
        )
        if kept:
            row_attrs[int(r)] = " " + kept

    parts = ["<sheetData>"]
    append = parts.append

    def emit_row(row_number, values):
        append(f'<row r="{row_number}"{row_attrs.get(row_number, "")}>')
        for letter, value in zip(letters, values):
            ref = f"{letter}{row_number}"
            style = cell_style.get(ref)
            s_attr = f' s="{style}"' if style else ""
            if value is None or value == "":
                if style:
                    append(f'<c r="{ref}"{s_attr}/>')
                continue
            if isinstance(value, bool):
                append(f'<c r="{ref}"{s_attr} t="b"><v>{int(value)}</v></c>')
            elif isinstance(value, int):
                append(f'<c r="{ref}"{s_attr}><v>{value}</v></c>')
            elif isinstance(value, float):
                if _math.isfinite(value):
                    append(f'<c r="{ref}"{s_attr}><v>{repr(value)}</v></c>')
                elif style:
                    append(f'<c r="{ref}"{s_attr}/>')
            elif isinstance(value, str):
                append(
                    f'<c r="{ref}"{s_attr} t="inlineStr"><is>'
                    f'<t xml:space="preserve">{_xml_text(value)}</t></is></c>'
                )
            elif isinstance(value, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
                raise _FallbackToOpenpyxl("date values need openpyxl formatting")
            else:
                append(
                    f'<c r="{ref}"{s_attr} t="inlineStr"><is>'
                    f'<t xml:space="preserve">{_xml_text(value)}</t></is></c>'
                )
        append("</row>")

    emit_row(1, [str(c) for c in final_df.columns])
    for row_number, values in enumerate(_excel_rows(final_df), start=2):
        emit_row(row_number, values)
    append("</sheetData>")

    head = old_xml[:start]
    tail = old_xml[end:]

    # Remove frozen panes (same as ws.freeze_panes = None).
    head = re.sub(rb"<pane\b[^>]*/>", b"", head)
    head = re.sub(rb"<selection\b[^>]*\bpane=\"[^\"]*\"[^>]*/>", b"", head)
    # Correct the used-range dimension.
    last = f"{letters[-1] if letters else 'A'}{n_rows}".encode()
    head = re.sub(rb'<dimension ref="[^"]*"\s*/>', b'<dimension ref="A1:' + last + b'"/>', head)

    return head + "".join(parts).encode("utf-8") + tail


def _write_processed_sheets_fast(file_data: bytes, processed_sheets) -> bytes:
    src = zipfile.ZipFile(BytesIO(file_data))
    sheet_paths = _sheet_part_paths(src)

    replacements = {}
    for target_sheet, final_df in processed_sheets:
        path = sheet_paths.get(target_sheet)
        if not path or path not in src.namelist():
            raise _FallbackToOpenpyxl(f"sheet part for '{target_sheet}' not found")
        replacements[path] = _render_sheet_xml(src.read(path), final_df)

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=5) as dst:
        for info in src.infolist():
            name = info.filename
            # calcChain lists formula cells; it is rebuilt by Excel. openpyxl
            # also drops it - keeping a stale one makes Excel "repair" the file.
            if name == "xl/calcChain.xml":
                continue
            if name in replacements:
                data = replacements[name]
            else:
                data = src.read(name)
                if name == "[Content_Types].xml":
                    data = re.sub(rb'<Override[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', b"", data)
                elif name == "xl/_rels/workbook.xml.rels":
                    data = re.sub(rb'<Relationship[^>]*Target="[^"]*calcChain\.xml"[^>]*/>', b"", data)
            zi = zipfile.ZipInfo(name, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            dst.writestr(zi, data)
    src.close()
    return out.getvalue()


def write_all_processed_sheets(deal_file_obj, processed_sheets) -> BytesIO:
    """Fast zip-patch writer, with the openpyxl writer as a safety net."""
    deal_name = getattr(deal_file_obj, "name", "")
    ext = os.path.splitext(deal_name)[1].lower()
    if ext not in {".xlsx", ".xlsm"}:
        raise ValueError("To preserve all DEAL worksheets, upload DEAL as .xlsx or .xlsm.")
    try:
        return BytesIO(
            _write_processed_sheets_fast(uploaded_file_bytes(deal_file_obj), processed_sheets)
        )
    except (_FallbackToOpenpyxl, KeyError, zipfile.BadZipFile, _ET.ParseError):
        return _write_all_processed_sheets_openpyxl(deal_file_obj, processed_sheets)

# ----------------------------------------------------------------------------
# Streamlit app
# ----------------------------------------------------------------------------

def run_black_mod():
    st.title("Black VLOOKUP")
    st.sidebar.header("Upload Excel Files")

    deal_file = st.sidebar.file_uploader(
        "Upload DEAL Excel file", type=["xlsx", "xlsm"], accept_multiple_files=False,
    )
    org_files = st.sidebar.file_uploader(
        "Upload ORG Excel/CSV files", type=["xls", "xlsx", "xlsm", "csv"],
        accept_multiple_files=True,
    )

    params_dict = {
        "AKARSH": {
            "org_prefix": "AKARSH",
            "params": {
                "deal_sheet_prefix": "AKARSH",
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "Black Incl.",
                "dNatts_2": "Table Incl.",
                "Rap value":"",
                "skip_rows": 1,
            },
        },
       "ANKIT": {
            "org_prefix": "ANKIT",
            "params": {
                "deal_sheet_prefix": "ANKIT",
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate No",
                "stock_id": "Stock ID",
                "dNatts_1": "Black",
                "dNatts_2": "",
                "Rap value":"",
                "skip_rows": 10,
            },
        },
        "ASIAN": {
            "org_prefix": "ASIAN",
            "params": {
                "deal_sheet_prefix": "ASIAN",
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No.",
                "stock_id": "Lot No.",
                "dNatts_1": "BIT",
                "dNatts_2": "BIS",
                "Rap value":"",
                "skip_rows": 7,
            },
        },
        "D NAVIN": {
            "org_prefix": "D NAVIN",
            "params": {
                "deal_sheet_prefix": "D NAVIN",
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "Centre Black",
                "dNatts_2": "Crown Black",
                "Rap value":"",
                "skip_rows": 1,
            },
        },
        "DIACLUST": {
            "org_prefix": "Diaclust",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "DIACLUST",
                "org_merge_col": "Report No",
                "stock_id": "StoneNo",
                "dNatts_1": "BIC",
                "dNatts_2": "BIS",
                "Rap value":"",
                "skip_rows": 0,
            },
        },
        "FINESTAR": {
            "org_prefix": "FINESTAR",
            "params": {
                "sheet_name": "diamonds",
                "deal_sheet_prefix": "FINESTAR",
                "deal_merge_col": "dLot",
                "org_merge_col": "Packet No",
                "stock_id": "",
                "dNatts_1": "CN",
                "dNatts_2": "SN",
                "Rap value":"",
                "skip_rows": 1,
            },
        },
        "GEMSTAR": {
            "org_prefix": "GEMSTAR",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "GEMSTAR",
                "org_merge_col": "Report No.",
                "stock_id": "Stock Id",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 0,
                "Rap value":"",
            },
        },
        "GLOWSTAR": {
            "org_prefix": "GLOWSTAR",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "GLOWSTAR",
                "org_merge_col": "Report No",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "GNYAN": {
            "org_prefix": "GNYAN",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "GNYAN",
                "org_merge_col": "Report No",
                "stock_id": "Stone ID",
                "dNatts_1": "",
                "dNatts_2": "",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "HK": {
            "org_prefix": "HK",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "HK",
                "org_merge_col": "Stock No",
                "stock_id": "",
                "dNatts_1": "BT",
                "dNatts_2": "BC",
                "Rap value":"",
            },
        },
        "HRV": {
            "org_prefix": "HRV",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "HRV",
                "org_merge_col": "Report No",
                "stock_id": "Stone No",
                "dNatts_1": "",
                "dNatts_2": "",
                "Rap value":"",
            },
        },
        "HIRADIAS": {
            "org_prefix": "HIRADIAS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "HIRADIAS",
                "org_merge_col": "Report No",
                "stock_id": "Stock Id",
                "dNatts_1": "",
                "dNatts_2": "",
                "Rap value":"",
            },
        },
        "JODHANI": {
            "org_prefix": "Jodhani",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "deal_sheet_prefix": "JODHANI",
                "stock_id": "Stone No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "Rap value":"",
            },
        },
        "JEWELEX": {
            "org_prefix": "Jewelex",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "deal_sheet_prefix": "JEWELEX",
                "stock_id": "Stone No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "KIRAN": {
            "org_prefix": "Kiran",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KIRAN",
                "org_merge_col": "Rep No",
                "stock_id": "Stone No",
                "dNatts_1": "BIC",
                "dNatts_2": "BIS",
                "skip_rows": 10,
                "Rap value":"",
            },
        },
        "KB GEMS": {
            "org_prefix": "KB GEMS",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "CertNo",
                "deal_sheet_prefix": "KB",
                "stock_id": "Ref No",
                "dNatts_1": "Black Table Inclusion",
                "dNatts_2": "Black Side Inclusion",
                "skip_rows": 2,
                "Rap value":"",
            },
        },
        "KBS": {
            "org_prefix": "KBS",
            "params": {
                "sheet_name": "stock",
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KBS",
                "org_merge_col": "CertNo",
                "stock_id": "StoneID",
                "dNatts_1": "BIC",
                "dNatts_2": "BIS",
                "skip_rows": 2,
                "Rap value":"",
            },
        },
        "KOMAL": {
            "org_prefix": "KOMAL",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KOMAL",
                "org_merge_col": "Report No",
                "stock_id": "Stock ID",
                "dNatts_1": "",
                "dNatts_2": "",
                "Rap value":"",
            },
        },
        "KG": {
            "org_prefix": "KG",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "REPORT NO.",
                "deal_sheet_prefix": "KG",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "KG_ALT": {
            "org_prefix": "KG",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "REPORT NO.",
                "deal_sheet_prefix": "KG",
                "stock_id": "STONE ID",
                "dNatts_1": "SIDE BLACK INCLUSION ",
                "dNatts_2": "TABLE BLACK INCLUSION",
                "skip_rows": 0,
                "Rap value":"",
            },
        },
        "KC": {
            "org_prefix": "KC",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KC",
                "org_merge_col": "Cert No",
                "stock_id": "Lot No.",
                "dNatts_1": "",
                "dNatts_2": "",
                "Rap value":"",
            },
        },
        "KGK": {
            "org_prefix": "KGK",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KGK",
                "org_merge_col": "Cert #",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 6,
                "Rap value":"",
            },
        },
        "KGK_ALT": {
            "org_prefix": "KGK",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "KGK",
                "org_merge_col": "Cert #",
                "stock_id": "",
                "dNatts_1": "TBL BLK",
                "dNatts_2": "Side BLK",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "LAXMI": {
            "org_prefix": "LAXMI",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate",
                "deal_sheet_prefix": "LAXMI",
                "stock_id": "Stone ID",
                "dNatts_1": "Black Inc",
                "dNatts_2": "",
                "skip_rows": 5,
                "Rap value":"",
            },
        },
        "LAKHI": {
            "org_prefix": "LAKHI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "LAKHI",
                "org_merge_col": "Report #",
                "stock_id": "Stock #",
                "dNatts_1": "Black Inclusion   ",
                "dNatts_2": "",
                "Rap value":"",
            },
        },
        "MB": {
            "org_prefix": "MB",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "MB",
                "org_merge_col": "Cert No.",
                "stock_id": "Stock ID",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "MAHENDRA": {
            "org_prefix": "MAHENDRA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "MAHENDRA",
                "org_merge_col": "Cert No.",
                "stock_id": "Stock ID",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "M.SURESH": {
            "org_prefix": "M.SURESH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "M SURESH",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "Tab_BLK",
                "dNatts_2": "Side_BLK",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "M SURESH": {
            "org_prefix": "M SURESH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "M SURESH",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "Tab_BLK",
                "dNatts_2": "Side_BLK",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "NAROLA": {
            "org_prefix": "NAROLA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "NAROLA",
                "org_merge_col": "ReportNo",
                "stock_id": "Stone Id",
                "dNatts_1": "",
                "dNatts_2": "",
                "skip_rows": 0,
                "Rap value":"",
            },
        },
        "PARISHI": {
            "org_prefix": "PARISHI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "PARISHI",
                "org_merge_col": "Lot no.",
                "stock_id": "",
                "dNatts_1": "BIT",
                "dNatts_2": "BIC",
                "Rap value":"",
            },
        },
        "PRAGATI": {
            "org_prefix": "PRAGATI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "PRAGATI",
                "org_merge_col": "Report No",
                "stock_id": "Ref No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "Rap value":"",
            }, 
        },
        "PRAMUKH": {
            "org_prefix": "PRAMUKH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "PARAMUKH",
                "org_merge_col": "Report No",
                "stock_id": "Stone ID",
                "dNatts_1": "T.BLK",
                "dNatts_2": "S.BLK",
                "Rap value":"",
            },
        },
        "PALADIYA":{
            "org_prefix": "PALADIYA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "PALADIYA",
                "org_merge_col": "Stone Id",
                "stock_id": "Stone Id",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "RAYS": {
            "org_prefix": "RAYS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "RAYS",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "BLACK",
                "dNatts_2": "",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "RATNAKALA_ALT": {
            "org_prefix": "RATNAKALA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "RATNAKALA",
                "org_merge_col": "Report No",
                "stock_id": "Stone No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "RATNAKALA": {
            "org_prefix": "RATNAKALA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "RATNAKALA",
                "org_merge_col": "Report No",
                "stock_id": "Stone Id",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "RAJGIR GEMS": {
            "org_prefix": "RAJGIR GEMS",
            "params": {
                "sheet_name": "Rajgir Gems STOCK LIST",
                "deal_sheet_prefix": "RAJGIR GEMS",
                "deal_merge_col": "dLot",
                "org_merge_col": "REPORT NO",
                "stock_id": "Stone ID",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 3,
                "Rap value":"",
            },
        },
        "R RAJESH": {
            "org_prefix": "R RAJESH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "R RAJESH",
                "org_merge_col": "Report No",
                "stock_id": "Stone ID",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 5,
                "Rap value":"",
            },
        },
        "R.RAJESH": {
            "org_prefix": "R.RAJESH",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "R RAJESH",
                "org_merge_col": "Report No",
                "stock_id": "Stone ID",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 5,
                "Rap value":"",
            },
        },
        "R N": {
            "org_prefix": "R N",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate",
                "deal_sheet_prefix": "R N",
                "stock_id": "STOCK",
                "dNatts_1": "Black Incl.",
                "dNatts_2": "Table Incl.",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "RSD": {
            "org_prefix": "RSD",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "RSD",
                "org_merge_col": "Cert No",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "Rap value":"",
            },
        },
        "RAYS": {
            "org_prefix": "RAYS",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "deal_sheet_prefix": "RAYS",
                "stock_id": "Packet No",
                "dNatts_1": "BLACK",
                "dNatts_2": "",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "RED EXIM": {
            "org_prefix": "RED EXIM",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Report#",
                "deal_sheet_prefix": "RED EXIM",
                "stock_id": "Stock#",
                "dNatts_1": "",
                "dNatts_2": "", 
                "Rap value":"",
            },
        },
        "SHAIRU": {
            "org_prefix": "SHAIRU",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SHAIRU",
                "org_merge_col": "Certi No.",
                "stock_id": "Stock Id",
                "dNatts_1": "Table Black",
                "dNatts_2": "Crown Black",
                "skip_rows": 5,
                "Rap value":"",
            },
        },
        "SAGAR": {
            "org_prefix": "SAGAR",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SAGAR",
                "org_merge_col": "CertNo",
                "stock_id": "StockID",
                "dNatts_1": "",
                "dNatts_2": "",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "STARLIGHTDIAMOND": {
            "org_prefix": "STARLIGHTDIAMOND",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "STARLIGHTDIAMOND",
                "org_merge_col": "Cert.No.",
                "stock_id": "StockID",
                "dNatts_1": "Black In Table",
                "dNatts_2": "Side Black",
                "Rap value":"",
            },
        },
        "SHREE KRISHNA": {
            "org_prefix": "SHREE KRISHNA",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SHREE KRISHNA",
                "org_merge_col": "Report No",
                "stock_id": "Stone No",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "Rap value":"",
            },
        },
        "SRD": {
            "org_prefix": "SRD",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Cert No",
                "deal_sheet_prefix": "SRD",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "SNJ": {
            "org_prefix": "SNJ",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate No",
                "deal_sheet_prefix": "SNJ",
                "stock_id": "Stone Id",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "skip_rows": 5,
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "SNJ_ALT": {
            "org_prefix": "SNJ",
            "params": {
                "deal_merge_col": "dLot",
                "org_merge_col": "Certificate No",
                "deal_sheet_prefix": "SNJ",
                "stock_id": "Stone Id",
                "dNatts_1": "Table Black",
                "dNatts_2": "Side Black",
                "Rap value":"", 
            },
        },
        "SAGAR": {
            "org_prefix": "SAGAR",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SAGAR",
                "org_merge_col": "CertNo",
                "stock_id": "StockID",
                "dNatts_1": "",
                "dNatts_2": "",
                "Rap value":"",
            },
        },
        "STAR RAYS": {
            "org_prefix": "STAR RAYS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "STARRAYS",
                "org_merge_col": "Cert.No",
                "stock_id": "",
                "dNatts_1": "BIT",
                "dNatts_2": "BIC",
                "skip_rows": 3,
                "Rap value":"",
            },
        },
        "SHEETAL": {
            "org_prefix": "SHEETAL",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SHEETAL",
                "org_merge_col": "Certificate",
                "stock_id": "StoneID",
                "dNatts_1": "BIT",
                "dNatts_2": "BIC",
                "skip_rows": 8,
                "Rap value":"",
            },
        },
        "SANGHVI": {
            "org_prefix": "SANGHVI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SANGHVI",
                "org_merge_col": "CERTNO",
                "stock_id": "Stone No",
                "dNatts_1": "Natts",
                "dNatts_2": "",
                "skip_rows": 0,
                "Rap value":"",
            },
        },
        "SANGHVI_ALT": {
            "org_prefix": "SANGHVI",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SANGHVI",
                "org_merge_col": "CERTNO",
                "stock_id": "REFNO",
                "dNatts_1": "",
                "dNatts_2": "",
                "skip_rows": 0,
                "Rap value":"",
            },
        },
        "SUNRAYS": {
            "org_prefix": "SUNRAYS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "SUNRAYS",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "B I",
                "dNatts_2": "T I",
                "skiprows":  4,
                "Rap value":"",
            },
        },
        "UNIQUE": {
            "org_prefix": "UNIQUE",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "UNIQUE",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "TB",
                "dNatts_2": "SB",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "VD GLOBAL": {
            "org_prefix": "VD GLOBAL",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "VD GLOBAL",
                "org_merge_col": "ReportNo",
                "stock_id": "Stock No",
                "dNatts_1": "BT",
                "dNatts_2": "",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "VEER GEMS": {
            "org_prefix": "VEER GEMS",
            "params": {
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "VEER GEMS",
                "org_merge_col": "Cert#",
                "stock_id": "Ref#",
                "dNatts_1": "CTB",
                "dNatts_2": "STB",
                "skip_rows": 4,
                "Rap value":"",
            },
        },
        "VAIBHAV": {
            "org_prefix": "VAIBHAV",
            "params": {
                "sheet_name": "Stock",
                "deal_sheet_prefix": "VAIBHAV",
                "deal_merge_col": "dLot",
                "org_merge_col": "Report No",
                "stock_id": "Packet No",
                "dNatts_1": "CB",
                "dNatts_2": "SB",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
        "VINAY": {
            "org_prefix": "VINAY",
            "params": {
                "sheet_name": "GIA",
                "deal_merge_col": "dLot",
                "deal_sheet_prefix": "VINAY",
                "org_merge_col": "ReportNumber",
                "stock_id": "StoneID",
                "dNatts_1": "Inclusion",
                "dNatts_2": "",
                "skip_rows": 1,
                "Rap value":"",
            },
        },
    }

    if not deal_file:
        st.info("Please upload the DEAL workbook.")
        return
    if not org_files:
        st.info("Please upload the ORG files.")
        return

    # ------------------------------------------------------------------
    # Re-use the previous result when the uploaded files are unchanged.
    # Without this, clicking Download (or any widget) re-runs everything.
    # ------------------------------------------------------------------
    run_signature = (
        file_key(deal_file),
        tuple(file_key(f) for f in org_files),
    )
    cached_run = st.session_state.get("_black_mod_result")

    if cached_run is not None and cached_run["signature"] == run_signature:
        processed_sheets = cached_run["processed_sheets"]
        display_results = cached_run["display_results"]
        output_bytes = cached_run["output_bytes"]
        write_error = cached_run["write_error"]
    else:
        processed_sheets, display_results, output_bytes, write_error = _process_all(
            deal_file, org_files, params_dict,
        )
        st.session_state["_black_mod_result"] = {
            "signature": run_signature,
            "processed_sheets": processed_sheets,
            "display_results": display_results,
            "output_bytes": output_bytes,
            "write_error": write_error,
        }

    if write_error:
        st.error(f"Could not create final DEAL workbook: {write_error}")

    for item in display_results:
        supplier = item.get("supplier", "Supplier")
        if "warning" in item:
            st.warning(item["warning"])
        elif "error" in item:
            st.error(f"{supplier}: {item['error']}")

    preview_rows = 100
    for sheet_name, final_df in processed_sheets:
        st.success(f"Processed DEAL sheet '{sheet_name}'.")
        st.dataframe(
            streamlit_safe_dataframe(final_df.head(preview_rows)),
        )
        if len(final_df) > preview_rows:
            st.caption(
                f"Showing first {preview_rows:,} of {len(final_df):,} rows. "
                "The downloaded workbook contains all rows."
            )

    if output_bytes is not None:
        output_filename = os.path.basename(deal_file.name)
        ext = os.path.splitext(output_filename)[1].lower()
        mime_type = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if ext == ".xlsm"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        with st.sidebar:
            st.download_button(
                label=f"Download {output_filename}",
                data=output_bytes,
                file_name=output_filename,
                mime=mime_type,
                key="download_final_deal",
            )


def _process_all(deal_file, org_files, params_dict):
    """Run every supplier config once. Returns
    (processed_sheets, display_results, output_bytes, write_error)."""
    processed_sheets = []
    display_results = []

    deal_sheet_names = get_deal_sheet_names(deal_file)

    org_name_keys = [(f, normalize_filename(f.name)) for f in org_files]
    configured_prefixes = {
        normalize_filename(c["org_prefix"])
        for c in params_dict.values() if c.get("org_prefix")
    }
    generic_org_files = [
        f for f, n in org_name_keys
        if not any(n.startswith(p) for p in configured_prefixes)
    ]
    claimed_generic_files = set()
    uploaded_org_names = ", ".join(f.name for f in org_files)

    with st.spinner("Processing DEAL workbook..."):
        for supplier_name, config in params_dict.items():
            params = config["params"]
            skip = params.get("skip_rows", params.get("skiprows"))

            if sheet_name_starting_with(deal_sheet_names, params.get("deal_sheet_prefix")) is None:
                supplier_prefix = normalize_filename(config.get("org_prefix", ""))
                if any(n.startswith(supplier_prefix) for _f, n in org_name_keys):
                    display_results.append({
                        "supplier": supplier_name,
                        "warning": (
                            f"{supplier_name}: DEAL sheet starting with "
                            f"'{params.get('deal_sheet_prefix')}' was not found."
                        ),
                    })
                continue

            org_file = match_uploaded_file(
                org_files,
                config["org_prefix"],
                required_columns=[params.get("org_merge_col")],
                skiprows=skip,
                sheet_name=params.get("sheet_name"),
                fallback_files=[
                    f for f in generic_org_files if id(f) not in claimed_generic_files
                ],
            )

            if not org_file:
                display_results.append({
                    "supplier": supplier_name,
                    "warning": (
                        f"{supplier_name}: ORG file not found using "
                        f"prefix '{config['org_prefix']}'. "
                        f"Uploaded ORG files: {uploaded_org_names or '(none)'}"
                    ),
                })
                continue

            try:
                actual_sheet_name, output_df = build_processed_dataframe(
                    org_file_obj=org_file,
                    deal_file_obj=deal_file,
                    org_skiprows=skip,
                    org_merge_col=params["org_merge_col"],
                    deal_merge_col=params["deal_merge_col"],
                    dNatts_1=params["dNatts_1"],
                    dNatts_2=params["dNatts_2"],
                    stock_id=params["stock_id"],
                    rapvalue=params["Rap value"],
                    deal_sheet_prefix=params.get("deal_sheet_prefix"),
                    org_sheet_name=params.get("sheet_name"),
                )
                add_or_merge_processed_sheet(processed_sheets, actual_sheet_name, output_df)
                if org_file in generic_org_files:
                    claimed_generic_files.add(id(org_file))
            except Exception as e:  # noqa: BLE001
                error_text = str(e)
                if error_text.startswith("No DEAL sheet starts with"):
                    continue
                display_results.append({"supplier": supplier_name, "error": error_text})

        output_bytes = None
        write_error = None
        if processed_sheets:
            try:
                output_bytes = write_all_processed_sheets(
                    deal_file_obj=deal_file, processed_sheets=processed_sheets,
                ).getvalue()
            except Exception as e:  # noqa: BLE001
                write_error = str(e)

    return processed_sheets, display_results, output_bytes, write_error


if __name__ == "__main__":
    run_black_mod()