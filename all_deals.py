import re
import csv

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import numpy as np
from datetime import datetime
import io
import os

st.set_page_config(
    page_title="DEAL Files",
    page_icon="📄",
    layout="wide",
)

def reset_file_pointer(file_path):
    """Reset Streamlit UploadedFile pointer before every pandas read."""
    if hasattr(file_path, "seek"):
        try:
            file_path.seek(0)
        except Exception:
            pass

# Applies to EVERY supplier and EVERY column, not only Asian Star / LOTNO.
_HYPERLINK_START = re.compile(r"^\s*=+\s*HYPERLINK\s*\(", re.IGNORECASE)


def hyperlink_arguments(value):
    """Split an Excel HYPERLINK formula into its arguments.

    Supports =HYPERLINK(...), ==HYPERLINK(...), commas/semicolons,
    quoted commas, escaped quotes and nested formula arguments.
    Returns None when the value is not a HYPERLINK formula.
    """
    if not isinstance(value, str):
        return None
    match = _HYPERLINK_START.match(value)
    if not match:
        return None

    remainder = value[match.end():].strip()
    if not remainder.endswith(')'):
        return None
    inside = remainder[:-1]
    arguments = []
    current = []
    quoted = False
    depth = 0
    i = 0
    while i < len(inside):
        char = inside[i]
        if char == '"':
            current.append(char)
            if quoted and i + 1 < len(inside) and inside[i + 1] == '"':
                current.append('"')
                i += 1
            else:
                quoted = not quoted
        elif not quoted and char == '(':
            depth += 1
            current.append(char)
        elif not quoted and char == ')':
            depth -= 1
            current.append(char)
        elif not quoted and depth == 0 and char in ',;':
            arguments.append(''.join(current).strip())
            current = []
        else:
            current.append(char)
        i += 1

    if quoted or depth != 0:
        return None
    arguments.append(''.join(current).strip())
    return arguments


def excel_literal(argument):
    """Return a literal Excel string/number; do not evaluate other formulas."""
    text = argument.strip()
    if len(text) >= 2 and text[0] == text[-1] == '"':
        return text[1:-1].replace('""', '"')
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", text):
        return text
    return None


def hyperlink_display(value, cached_value=None):
    """Return the HYPERLINK display label, never the formula or target URL.

    If the label is calculated rather than a literal, use the Excel cached
    result when available. Excel formulas are not evaluated by this script.
    """
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

    # A one-argument HYPERLINK shows the URL itself.
    if len(args) == 1:
        link = excel_literal(args[0])
        if link is not None:
            return link
    # Complex calculated labels cannot be reconstructed without a cache.
    return ''


def clean_hyperlinks_in_dataframe(data):
    """Clean literal =HYPERLINK text in ANY column (Excel or CSV)."""
    for col_idx in range(data.shape[1]):
        series = data.iloc[:, col_idx]
        if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
            clean = series.map(hyperlink_display)
            data.isetitem(col_idx, clean)
    return data


def recover_excel_hyperlinks(file_path, data, skip_rows):
    """Recover formula display text in ALL columns of the first Excel sheet.

    read_excel loads cached formula results, which are often missing. Read the
    original formula cells with openpyxl and replace only HYPERLINK formulas.
    Matching is by physical row/column, so data stays aligned after filtering.
    """
    reset_file_pointer(file_path)
    workbook = load_workbook(file_path, data_only=False, read_only=True)
    try:
        sheet = workbook.active  # Same default sheet as pd.read_excel
        header_row = skip_rows + 1
        patches = {}
        rows = sheet.iter_rows(
            min_row=header_row + 1,
            max_row=header_row + len(data),
            min_col=1,
            max_col=data.shape[1],
            values_only=True,
        )
        for row_idx, row in enumerate(rows):
            for col_idx, formula in enumerate(row):
                if hyperlink_arguments(formula) is None:
                    continue
                cached_value = data.iat[row_idx, col_idx]
                display = hyperlink_display(formula, cached_value)
                patches.setdefault(col_idx, {})[row_idx] = display

        # Replace each affected column without coercing identifier strings into
        # numeric types (leading zeroes, long report IDs, etc. stay unchanged).
        for col_idx, updates in patches.items():
            values = data.iloc[:, col_idx].astype(object).to_numpy(copy=True)
            for row_idx, display in updates.items():
                values[row_idx] = display
            data.isetitem(col_idx, pd.Series(values, index=data.index, dtype=object))
    finally:
        workbook.close()
        reset_file_pointer(file_path)
    return data


def load_data(file_path, skip_rows=0, prefix=""):
    """Load all suppliers' Excel/CSV files; fix HYPERLINK in every column."""
    reset_file_pointer(file_path)
    file_name = str(getattr(file_path, "name", file_path)).lower()

    if file_name.endswith(".csv"):
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                reset_file_pointer(file_path)
                data = pd.read_csv(
                    file_path, skiprows=skip_rows, encoding=encoding,
                    keep_default_na=False,
                )
                return clean_hyperlinks_in_dataframe(data)
            except UnicodeDecodeError as exc:
                last_error = exc
                continue
        raise last_error

    data = pd.read_excel(file_path, skiprows=skip_rows, keep_default_na=False)
    if file_name.endswith((".xlsx", ".xlsm")):
        data = recover_excel_hyperlinks(file_path, data, skip_rows)
    # Also handles literal formula text stored in CSV, XLS or Excel cells.
    return clean_hyperlinks_in_dataframe(data)

def apply_filters(data, filters):
    filtered = data.copy()

    for column, criteria in filters.items():
        if column not in filtered.columns:
            # If the column doesn't exist in the uploaded file, skip this filter.
            continue

        series = filtered[column]

        if isinstance(criteria, list):
            # Keep original matching, plus stripped text matching for columns where suppliers add spaces.
            original_mask = series.isin(criteria)
            text_mask = series.astype(str).str.strip().isin([str(v).strip() for v in criteria])
            mask = original_mask | text_mask
        elif callable(criteria):
            mask = criteria(series)
        else:
            continue

        if not isinstance(mask, pd.Series):
            mask = pd.Series(mask, index=filtered.index)

        mask = mask.reindex(filtered.index).fillna(False).astype(bool)
        filtered = filtered.loc[mask].copy()

    return filtered

def concatenate_dataframes(df, column_names):
    concatenated = []
    for sub_df in df.values():
        for column_name in column_names:
            if sub_df is not None and column_name in sub_df.columns:
                concatenated.append(sub_df[column_name])
                break
    concatenated = [series for series in concatenated if not series.empty]
    return pd.concat(concatenated, ignore_index=True) if concatenated else pd.Series(dtype='str')

def map_columns(data, column_mapping):
    output_data = pd.DataFrame()
    for column, keys in column_mapping.items():
        if column in ['LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP']:
            output_data[column] = 0
        elif column == 'SRNO':
            output_data[column] = range(1, len(data) + 1)
        elif column in ['TINCH', 'LUSTER', 'T OPENS', 'B OPENS']:
            output_data[column] = 'N/A'
        elif column == 'SHAPE':
            output_data[column] = 'ROUND'
        elif column == 'DISC':
            output_data[column] = pd.to_numeric(concatenate_dataframes({0: data}, keys),errors='coerce').apply(lambda x: (f"{-abs(x * 100):.2f}" if pd.notna(x) and abs(x) < 1 else f"{-abs(x):g}" if pd.notna(x) else ""))
        elif column == 'NATTS':
            output_data[column] = 'CENTER - L'
        elif not keys:
            output_data[column] = np.nan
        else:
            if column == 'DISC':
                output_data[column] = concatenate_dataframes({0: data}, keys).round(2).apply(
                    lambda x: -x if x > 0 else x
                )
            elif column == 'FLRC':
                output_data[column] = concatenate_dataframes({0: data}, keys).replace(
                    {
                        'F': 'FNT', 'M': 'MED', 'N': 'NON','Faint':'FNT',
                        'FAINT': 'FNT', 'NONE': 'NON', 'MEDIUM': 'MED','Medium':'MED',
                        'Fnt': 'FNT', 'Med': 'MED',
                        ' F': 'FNT', ' N': 'NON', ' M': 'MED',
                        'FA': 'FNT', 'MD': 'MED', 'NN': 'NON', 'FN': 'FNT','None':"NON"
                    }
                )
            elif column in ['CUT', 'POL', 'SYM']:
                output_data[column] = concatenate_dataframes({0: data}, keys).replace(
                    {'Vg': 'VG', 'Ex': 'EX'}
                )
            else:
                output_data[column] = concatenate_dataframes({0: data}, keys)
    return output_data

def save_output(output_data, prefix):
    csv_data = output_data.to_csv(index=False)
    return csv_data

# Output columns that should be real numbers in Excel (not text).
# LOTNO is deliberately excluded so long report numbers keep every digit.
NUMERIC_OUTPUT_COLUMNS = (
    'SRNO', 'CARAT', 'LENGTH', 'WIDTH', 'HEIGHT', 'TABLE', 'DEPTH',
    'DISC', 'RATIO', 'RAP',
)
_INVALID_SHEET_CHARS = re.compile(r'[\[\]:*?/\\]')


def safe_sheet_name(name, used):
    """Excel sheet names: max 31 chars, no []:*?/\\ and unique (case-insensitive)."""
    base = _INVALID_SHEET_CHARS.sub(' ', str(name)).strip().strip("'") or 'Sheet'
    base = base[:31]
    candidate = base
    counter = 2
    while candidate.lower() in used:
        suffix = f" ({counter})"
        candidate = base[:31 - len(suffix)] + suffix
        counter += 1
    used.add(candidate.lower())
    return candidate


def excel_ready(frame):
    """Turn numeric-looking output columns into numbers; keep anything else as-is."""
    frame = frame.copy()
    for col in NUMERIC_OUTPUT_COLUMNS:
        if col in frame.columns:
            numbers = pd.to_numeric(frame[col], errors='coerce')
            frame[col] = numbers.where(numbers.notna(), frame[col])
    return frame


def build_deal_workbook(deal_results):
    """Write every processed supplier into ONE .xlsx, one sheet per supplier."""
    buffer = io.BytesIO()
    used = set()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for (file_name, prefix), info in deal_results.items():
            frame = info.get('df')
            if frame is None:
                frame = pd.read_csv(io.StringIO(info.get('csv', '')))
            sheet = safe_sheet_name(prefix.lower(), used)
            info['sheet_name'] = sheet
            excel_ready(frame).to_excel(writer, sheet_name=sheet, index=False)

            ws = writer.sheets[sheet]
            ws.freeze_panes = 'A2'
            for idx, column in enumerate(frame.columns, start=1):
                values = [str(column)] + [str(v) for v in frame[column].head(200)]
                width = min(max(len(v) for v in values) + 2, 40)
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = width
    return buffer.getvalue()

# Automatic header detection is GLOBAL: it applies to every supplier, not
# only MB, and tries the first 35 physical rows of each uploaded file.
HEADER_SCAN_ROWS = 35


def read_candidate_header_rows(file_path, max_rows=HEADER_SCAN_ROWS):
    """Read potential header rows ONCE without loading the full spreadsheet.

    Return rows as original cell values: skip_rows is always the physical
    zero-based row index, for CSV, XLS, XLSX and XLSM alike.
    """
    reset_file_pointer(file_path)
    name = str(getattr(file_path, "name", file_path)).lower()
    try:
        if name.endswith(".csv"):
            # CSV rows may have different numbers of preamble/header fields;
            # csv.reader tolerates this where pd.read_csv(header=None) may not.
            wrapper = io.TextIOWrapper(file_path, encoding="utf-8-sig", errors="replace", newline="")
            try:
                rows = []
                for row in csv.reader(wrapper):
                    rows.append(row)
                    if len(rows) >= max_rows:
                        break
                return rows
            finally:
                wrapper.detach()  # Leave Streamlit's UploadedFile open.

        if name.endswith((".xlsx", ".xlsm")):
            workbook = load_workbook(file_path, read_only=True, data_only=True)
            try:
                return list(workbook.active.iter_rows(
                    min_row=1, max_row=max_rows, max_col=250, values_only=True
                ))
            finally:
                workbook.close()

        # Legacy .xls: openpyxl cannot read this Excel format.
        preview = pd.read_excel(
            file_path, header=None, nrows=max_rows, keep_default_na=False
        )
        return preview.values.tolist()
    finally:
        reset_file_pointer(file_path)


def score_supplier_headers(header_cells, filters, column_mapping):
    """Score one potential header against a supplier's own expected fields.

    A wrong row is *not* valid merely because loading it creates data rows.
    Exact field names are used because the existing supplier filters and
    column mappings also use exact field names.
    """
    actual = {value for value in header_cells if isinstance(value, str) and value}
    matched_filters = sum(name in actual for name in filters)
    map_fields = ("LOTNO", "CARAT", "COLOR", "CLARITY", "CUT", "POL", "SYM", "FLRC", "TABLE", "DEPTH", "LAB", "DISC")
    matched_mappings = {
        field for field in map_fields
        if any(name in actual for name in column_mapping[field])
    }
    core_hits = len({"CARAT", "COLOR", "CLARITY"} & matched_mappings)

    # Generic acceptance guard for ALL suppliers: avoid files with plausible
    # row counts but no usable header / mapping, as in the MB screenshot.
    if (matched_filters < min(3, len(filters))
            or len(matched_mappings) < 3
            or (core_hits < 2 and len(matched_mappings) < 5)):
        return None

    return matched_filters * 10 + len(matched_mappings) * 4 + core_hits * 5


def score_generic_headers(header_cells, column_mapping):
    """Require real report fields before processing an unconfigured supplier.

    Generic mode is explicitly labeled and never borrows another supplier's
    filter rules. This prevents all-blank/incorrect CSV output.
    """
    names = {str(value).strip().lower() for value in header_cells
             if value is not None and str(value).strip()}
    groups = {}
    for target, aliases in column_mapping.items():
        if aliases and target not in ('SHAPE',):
            groups[target] = any(str(alias).strip().lower() in names for alias in aliases)

    essentials = ('LOTNO', 'CARAT', 'COLOR', 'CLARITY', 'DISC')
    if not all(groups.get(name, False) for name in essentials):
        return None
    return 10 * sum(groups.values())


def supplier_for_file(filename, configurations):
    """Match known prefixes including spaces/punctuation variants.

    Examples: 'V D GLOBAL.xlsx' is the configured 'VD GLOBAL'.
    'StockListKBG.xlsx' becomes an unconfigured supplier named KBG.
    """
    stem = os.path.splitext(os.path.basename(filename))[0]
    upper = stem.upper().strip()
    normalized = re.sub(r'[^A-Z0-9]', '', upper)

    for prefix in sorted(configurations, key=len, reverse=True):
        if re.match(rf'^{re.escape(prefix)}(?:[\s._-]|$)', upper):
            return prefix, True

    # Only accept normalization for the COMPLETE name, or a longer prefix
    # with a date/obvious suffix. Avoid 'KB' misidentifying 'KBG'/'KBS'.
    for prefix in sorted(configurations, key=len, reverse=True):
        key = re.sub(r'[^A-Z0-9]', '', prefix.upper())
        if normalized == key or (len(key) >= 6 and normalized.startswith(key)
                                   and normalized[len(key):].isdigit()):
            return prefix, True

    # Stock lists are commonly named StockList<supplier>, with no separator.
    generic_name = re.sub(r'^(?:STOCK[\s._-]*LIST)[\s._-]*', '', upper)
    generic_name = generic_name.strip(' ._-') or upper
    return generic_name, False


def find_supplier_header_candidates(file_path, configurations, column_mapping,
                                    generic=False):
    """Try ranked headers, then ALWAYS try skip_rows=0 as a last fallback.

    Fallback zero is applied to every supplier, even when the detected header
    gives zero filtered rows. Header validation still prevents garbage CSVs.
    """
    rows = read_candidate_header_rows(file_path)
    layouts = configurations if isinstance(configurations, list) else [configurations]
    candidates = []
    seen = set()

    for priority, config in enumerate(layouts):
        preferred = config.get('skip_rows', 0)
        for offset, header in enumerate(rows):
            key = (id(config['filters']), offset)
            if key in seen:
                continue
            seen.add(key)
            score = (score_generic_headers(header, column_mapping) if generic
                     else score_supplier_headers(header, config['filters'], column_mapping))
            if score is None:
                continue
            candidates.append((score, -abs(offset - preferred), -priority, offset, config))

    candidates.sort(reverse=True, key=lambda row: row[:3])
    result = [(offset, config) for _score, _distance, _priority, offset, config in candidates]

    # Explicit user-requested final attempt for ALL files: zero skipped rows.
    # Every configured layout can potentially have different filter rules.
    for config in layouts:
        if not any(offset == 0 and conf is config for offset, conf in result):
            result.append((0, config))
    return result


def find_mapped_column(columns, aliases):
    """Pick an existing source column by normalized name, preserving input."""
    existing = {str(col).strip().lower(): col for col in columns}
    for alias in aliases:
        if str(alias).strip().lower() in existing:
            return existing[str(alias).strip().lower()]
    return None


def generic_supplier_filter(data, column_mapping):
    """Apply published generic GIA/round/size/discount quality requirements.

    Used ONLY when the supplier has no configured rules. It is visibly labeled
    'generic rules' in the app; exact supplier inclusion rules require setup.
    """
    frame = data.copy()
    attrs = []
    def field(name):
        return find_mapped_column(frame.columns, column_mapping[name])

    criteria = [
        ('CARAT', lambda v: pd.to_numeric(v.astype(str).str.replace(',', '', regex=False), errors='coerce').between(.30, 3.99)),
        ('COLOR', lambda v: v.astype(str).str.strip().str.upper().isin(list('DEFGHIJKLM'))),
        ('CLARITY', lambda v: v.astype(str).str.strip().str.upper().isin(['IF','VVS1','VVS2','VS1','VS2','SI1','SI2'])),
        ('DISC', lambda v: pd.to_numeric(v.astype(str).str.replace('%','',regex=False).str.strip(), errors='coerce').abs().ge(10)),
    ]
    # Do not confuse a 'Cert' (certificate number) column with laboratory.
    lab_col = find_mapped_column(frame.columns, ['Lab', 'LAB', 'Cert Lab', 'Certi Lab'])
    if lab_col is not None:
        frame = frame.loc[frame[lab_col].astype(str).str.strip().str.upper().eq('GIA')].copy()
        attrs.append('LAB')

    for key, predicate in criteria:
        col = field(key)
        if col is not None:
            frame = frame.loc[predicate(frame[col])].copy()
            attrs.append(key)

    # Shape must be a round diamond if the file has a shape field.
    shape_col = find_mapped_column(frame.columns, ['Shape','SHAPE','Shp','Shp.'])
    if shape_col is not None:
        frame = frame.loc[frame[shape_col].astype(str).str.strip().str.upper().isin(
            ('RD', 'ROUND', 'RBC', 'RB', 'BR'))].copy()
        attrs.append('SHAPE')

    frame.attrs['generic_fields_applied'] = attrs
    return frame

def validate_source_headers(data, filters, column_mapping, prefix, skip_rows, generic=False):
    """Prevent generation of empty output for an unrecognized header layout."""
    score = (score_generic_headers(data.columns, column_mapping) if generic
             else score_supplier_headers(data.columns, filters, column_mapping))
    if score is None:
        raise ValueError(
            f"Unrecognized headers for {prefix} (skip_rows={skip_rows}): "
            f"{list(data.columns[:15])}"
        )


# Dedicated Rapnet conversion, based on the user's Rapnet page.
# Unlike unconfigured suppliers, Rapnet has its own ID, discount and quality rules.
RAPNET_OUTPUT_MAP = {
    'SRNO': [],
    'LOTNO': ['Diamond ID', 'Stone ID'],
    'CARAT': ['Size'],
    'SHAPE': ['Shape'],
    'COLOR': ['Color'],
    'CLARITY': ['Clarity'],
    'CUT': ['Cut'],
    'POL': ['Polish'],
    'SYM': ['Symmetry'],
    'FLRC': ['Fluorescence'],
    'LENGTH': [], 'WIDTH': [], 'HEIGHT': [],
    'TABLE': ['Table'], 'DEPTH': ['Depth'], 'LAB': ['Lab'],
    'DISC': ['%Rap', '% Rap', 'Rap%', 'Rap %'],
    'TINCH': [], 'NATTS': [], 'RATIO': [], 'RAP': [],
    'LUSTER': [], 'T OPENS': [], 'B OPENS': [],
}


def rapnet_source_columns(data):
    """Resolve source header variants without changing other suppliers."""
    return {
        target: find_mapped_column(data.columns, aliases)
        for target, aliases in RAPNET_OUTPUT_MAP.items() if aliases
    }


def process_rapnet(data, prefix='RAPNET LIST', skip_rows=0):
    """Apply the standalone Rapnet rules and return the shared DEAL CSV shape.

    %Rap <= 10 is preserved *as written* in the supplied Rapnet code;
    do not silently change it to <= -10 or use generic-supplier filters.
    """
    source = rapnet_source_columns(data)
    required = ('LOTNO', 'CARAT', 'SHAPE', 'COLOR', 'CLARITY',
                'CUT', 'POL', 'SYM', 'FLRC', 'LAB', 'DISC')
    missing = [name for name in required if source.get(name) is None]
    if missing:
        raise ValueError(
            f"Unrecognized Rapnet headers (skip_rows={skip_rows}). "
            f"Missing columns for: {', '.join(missing)}. "
            f"Found: {list(data.columns[:30])}"
        )

    frame = data.copy()
    # Pandas's default-na=False means missing text can be an empty string.
    # Numeric cleanup uses coerce so any invalid Size/%Rap values are dropped.
    size = pd.to_numeric(
        frame[source['CARAT']].astype(str).str.replace(',', '', regex=False).str.strip(),
        errors='coerce',
    )
    discount = pd.to_numeric(
        frame[source['DISC']].astype(str).str.replace('%', '', regex=False)
             .str.replace(',', '', regex=False).str.strip(),
        errors='coerce',
    )
    mask = (
        frame[source['LAB']].astype(str).str.strip().str.upper().eq('GIA')
        & frame[source['SHAPE']].astype(str).str.strip().str.casefold().eq('round')
        & size.between(0.30, 3.99, inclusive='both')
        & discount.le(10)
        & frame[source['COLOR']].astype(str).str.strip().str.upper().isin(
            list('DEFGHIJKLM'))
        & frame[source['CLARITY']].astype(str).str.strip().str.upper().isin(
            ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'])
    )

    fluorescence = (
        frame[source['FLRC']].astype(str).str.strip().str.upper().replace(
            {'FAINT': 'FNT', 'NONE': 'NON', 'MEDIUM': 'MED'}
        )
    )
    mask &= fluorescence.isin(['NON', 'FNT', 'MED'])

    grades = {}
    for target in ('CUT', 'POL', 'SYM'):
        values = (
            frame[source[target]].astype(str).str.strip().str.upper().replace(
                {'VERY GOOD': 'VG', 'EXCELLENT': 'EX'}
            )
        )
        grades[target] = values
        mask &= values.isin(['EX', 'VG'])

    key_column = find_mapped_column(frame.columns, ['Key To Symbols'])
    if key_column is not None:
        mask &= ~frame[key_column].astype(str).str.strip().isin(['Cavity', 'Chip'])

    # Keep the original row indexes until all masks have been applied.
    frame = frame.loc[mask].copy().reset_index(drop=True)
    size = size.loc[mask].reset_index(drop=True)
    discount = discount.loc[mask].reset_index(drop=True)
    fluorescence = fluorescence.loc[mask].reset_index(drop=True)
    grades = {name: values.loc[mask].reset_index(drop=True)
              for name, values in grades.items()}

    result = pd.DataFrame(index=range(len(frame)), columns=RAPNET_OUTPUT_MAP.keys())
    for target, source_col in source.items():
        if source_col is not None:
            result[target] = frame[source_col].reset_index(drop=True).to_numpy()

    result['SRNO'] = range(1, len(frame) + 1)
    result['CARAT'] = size
    result['DISC'] = discount
    result['FLRC'] = fluorescence
    for name in ('CUT', 'POL', 'SYM'):
        result[name] = grades[name]

    for name in ('LENGTH', 'WIDTH', 'HEIGHT', 'RATIO', 'RAP'):
        result[name] = 0
    for name in ('TABLE', 'DEPTH'):
        result[name] = pd.to_numeric(result[name], errors='coerce')
    for name in ('TINCH', 'LUSTER', 'T OPENS', 'B OPENS'):
        result[name] = 'N/A'
    result['NATTS'] = 'CENTER - L'

    # All-supplier formula handling remains active, including Rapnet identifiers.
    result = clean_hyperlinks_in_dataframe(result)
    return save_output(result, prefix), len(result), result


def process_data(file_path, filters, column_mapping, prefix, skip_rows=0, generic=False):
    if prefix.upper() in ('RAPNET LIST', 'RAPNET'):
        data = load_data(file_path, skip_rows=skip_rows, prefix=prefix)
        return process_rapnet(data, prefix=prefix, skip_rows=skip_rows)

    data = load_data(file_path, skip_rows=skip_rows, prefix=prefix)
    validate_source_headers(data, filters, column_mapping, prefix, skip_rows, generic=generic)

    # Clean numeric columns commonly used in supplier filters.
    columns_to_convert = [
        'Disc', 'Rapnet Discount %', 'Disc %', 'Rapnet Back', 'DISC',
        'Disc. (%)', ' Disc %', 'DISC.', 'Disc%', 'Disc(%)', 'Dis %',
        'Dis', 'Dis%', 'BACK', 'Back%', 'Final Disc%', 'Online Back',
        'Rap%', 'Rap %', 'RAP DIS', 'RAP DISC%', 'Rap.%', 'Net %', 'Net%', 'Price',
        'Cts', 'Ct.', 'Ct', 'CRTWT', 'Weight', 'weight', 'Size',
        'Carats', 'Carat', 'SIZE', 'CTS', 'WEIGHT', 'Wgt', 'Wt', 'Cts.'
    ]
    for col in columns_to_convert:
        if col in data.columns:
            data[col] = data[col].replace({',': '', '%': '', ' ': ''}, regex=True)
            data[col] = pd.to_numeric(data[col], errors='coerce')
            data[col] = data[col].fillna(0)

    # Apply the filters to the cleaned data.
    filtered_data = (generic_supplier_filter(data, column_mapping)
                     if generic else apply_filters(data, filters)).copy()

    # Additional logic for 'Size' renaming for KGK prefix (kept from your code).
    if prefix == 'KGK' and 'Size' in filtered_data.columns:
        filtered_data = filtered_data.rename(columns={'Size': 'Updated_Size'})

    # Standardize generic column-name capitalization for the shared mapper.
    if generic:
        renames = {}
        existing = set(filtered_data.columns)
        for aliases in column_mapping.values():
            if not aliases:
                continue
            source_col = find_mapped_column(filtered_data.columns, aliases)
            if source_col is None:
                continue
            # Map only a spelling/case variation to the canonical alias.
            canonical = next((a for a in aliases if str(a).strip().lower()
                              == str(source_col).strip().lower()), aliases[0])
            if source_col != canonical and canonical not in existing:
                renames[source_col] = canonical
                existing.add(canonical)
        filtered_data = filtered_data.rename(columns=renames)

    # Apply column mapping and return the final processed data.
    output_data = map_columns(filtered_data, column_mapping)

    # Asian Star's supplier-specific LOTNO source is still Report No.
    # The HYPERLINK conversion above is universal for EVERY supplier/column.
    if prefix.upper() == "ASIAN STAR":
        report_col = next(
            (name for name in filtered_data.columns
             if str(name).strip().lower() in ("report no.", "report no")),
            None,
        )
        if report_col is None:
            raise ValueError("Asian Star report-number column is missing")
        output_data["LOTNO"] = (
            filtered_data[report_col].reset_index(drop=True)
        )

    # Final safeguard: no literal HYPERLINK formula escapes to any output column.
    output_data = clean_hyperlinks_in_dataframe(output_data)
    csv_data = save_output(output_data, prefix)

    return csv_data, len(output_data), output_data

def all_deal():
    st.title("DEAL Files")

    # File uploader
    with st.sidebar:
        files = st.file_uploader(
            "Upload Excel or CSV Files",
            type=["xlsx", "xlsm", "xls", "csv"],
            accept_multiple_files=True,
            key="deal_uploader"
        )

    current_datetime = datetime.now().strftime("%d.%m.%Y")

    # Initialize session state store for processed CSV data
    # We'll store: (file_name, prefix) -> {"csv": csv_data, "download_name": file_name_for_csv}
    if "deal_results" not in st.session_state:
        st.session_state.deal_results = {}

    if not files:
        st.warning("No files selected.")
        return

    # Keep output list synced with the currently uploaded files.
    st.session_state.deal_results = {}

    # Column Mapping
    column_mapping = {
        'SRNO': [],
        'LOTNO': ['CERTNO','Rep No', 'Cert#', 'Report No', 'Cert No', 'Certificate No', 'Report No.',
                  'Packet No', 'CertNo', 'Cert No.', 'REPORT NO.', 'Report #', 'Control No', 'ReportNo',
                  'Report', 'Certificate', 'Cert.No.', 'Report#', 'Certi No.', 'Diamond ID', 'Stone ID','Stock No', 'Cert.No'],
        'CARAT': ['Ct.','Cts', 'Wgt', 'CRTWT', 'Weight', 'Carats', 'Size', 'Carat', 'SIZE', 'CTS', 'Ct','WEIGHT', 'Cts.'],
        'SHAPE': [],
        'COLOR': ['Col.','Color', 'COL', 'Col', 'COLOR'],
        'CLARITY': ['Purity', 'Cla', 'CLR', 'Clar', 'Clarity', 'CLARITY', 'Clr', 'Cts','WEIGHT', 'CLAR', 'Cla.'],
        'CUT': ['Cut ','Cut', 'CUT', 'Final Cut', 'CutGrade','Cut '],
        'POL': ['Pol.',' Pol','Polish', 'POL', 'Pol'],
        'SYM': ['Symm', 'SYM', 'Sym', 'Sym.', 'Symmetry', 'SYMM', 'Sy'],
        'FLRC': ['Fluorescent','Fl.','Fls', 'Flur', 'Fluor', 'Fluor.', 'Flour', 'Fluo', 'Flo', 'FLR', 'FLUOR',
                 'FLUO', 'Fluorescence Intensity', 'Fluo Int', 'Flor.', 'FL', 'Fluorescence', 'Fl.Int.',
                 'Flour.', 'Flr', 'Fluorescence'],
        'LENGTH': [],
        'WIDTH': [],
        'HEIGHT': [],
        'TABLE': ['TB%','Table %', 'TBL', 'Table%', 'Tab%', 'TA%', 'T.Size', 'TABLE %', 'Table', 'Table(%)', 'Tabl',
                  'TAB%', 'Tbl'],
        'DEPTH': ['Depth %', 'DPL', 'Depth%', 'TD%', 'T Depth', 'DEPTH %', 'Depth Percentage', 'Depth(%)',
                  'Depth', 'DEPTH', 'T.Dep.', 'T.Dpt','TD', 'T.D.%'],
        'LAB': ['Lab', 'LAB', 'Cert', 'Cert Lab','Certi', 'Cert.'],
        'DISC': ['Disc. (%)','Rapnet Back',' Disc %','Disc %', 'Rap %', 'RAP DIS', 'DISC.', 'Disc', 'Net %',
                 'Rap%', 'DISC', 'Dis %', 'BACK', 'Rapnet  Discount %', 'Disc(%)', 'Disc%', 'Dis', 'Back%',
                 'Dis%', 'Disc. ', 'Final Disc%', 'Online Back', '%Rap', '% Rap', 'RAP DISC%', 'Rap.%'],
        'TINCH': [],
        'NATTS': [],
        'RATIO': [],
        'RAP': [],
        'LUSTER': [],
        'T OPENS': [],
        'B OPENS': [],
        'Natts1':['Black Incl.','Black','BIT','Centre Black','CN','Natts','BT','Table Black','Black Table Inclusion','SIDE BLACK INCLUSION ','TBL BLK','Black Inc',
                   'Black Inclusion   ','Tab_BLK','T.BLK','TB','BLACK','Black Incl.','Black In Table','B I','CTB','CB','Inclusion','Center Inclusion','BIC'],
        'Natts 2':['BIC','Table Incl.','BIS','Crown Black','SN','BC','Side Black','Black Side Inclusion','TABLE BLACK INCLUSION','Side BLK','Side Black','Side_BLK','S.BLK',
                   'SB','Table Incl.','T I','STB','Black Inclusion   ','BCR'],
        "Stone id":['Packet No','Stock ID','Lot No.','StoneNo','Stock Id','Stone No','Stone ID','Stock No','Stone Id','Ref No','StoneID','STONE ID','Lot No.','Stock #',
                    'Packet No','STOCK','Stock#','StockID','REFNO','Ref#','StoneID','RefNo'],
    }

    filters_akarsh = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['INDIA'],
        'Disc %': lambda x: x <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Purity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Milky': ['NONE','EX'],
        'Tinge': ['NONE'],
        'Brown': ['NONE'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symm': ['EX', 'VG'],
        'Fls': ['NON', 'FNT', 'MED'],
        'Culet': ['NON'],
        'Green': ['NONE'],
        'Key to symbols': lambda x: ~x.isin(['Cavity', 'Chip']),
        'Black Incl.': lambda x: x.isin(['B1', 'B2',"BPP"]),
        'Table Incl.': lambda x: x.isin(['T0', 'T1'])
    }

    filters_ankit = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'CRTWT': lambda x: (x >= 0.30) & (x <= 3.99),
        'LOCATION': ['MUMBAI'],
        'RAP DIS': lambda x: x <= -10,
        'COL': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLR': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Milky': ['NO', ''],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'Flur': ['NON', 'FNT', 'MED'],
        'Culet': ['NON', 'PTD', ''],
        'OPTA': ['NO',''],
        'KTSVIEW': lambda x: ~x.astype(str).str.strip().str.startswith(('Cavity', 'Chip'), na=False),
        'Natts': lambda x: ~x.isin(['YT2']),
        'Shade': ['NB'],
        'Comments': lambda x: ~x.str.startswith(
            "Clarity grade is based on internal graining that is not shown", na=False
        )
    }

    filters_amora = {
        'LAB': ['GIA'],
        'Shape': ['RBC'],
        'DISC.': lambda x: pd.to_numeric(x, errors='coerce').apply(lambda val: val <= -10),
        'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clar': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fluor': ['NON', 'FNT', 'MED'],
        'Culet Size': lambda x: x != 'VSM',
        'Milkey': lambda x: x.isna() | (x == 'None '),
        'Shade': lambda x: x.isna(),
        'Shape Description': lambda x: x.isna() | (x == 'Round Brilliant'),
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity']),
        'Report Comments': lambda x: ~x.str.startswith('Clarity', na=False)
    }

    filters_asian_star = {
        'Lab': ['GIA'],
        'Shape': ['RD'],
        'Disc': lambda x: x <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['IND'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Final Cut': ['EX', 'VG', '3EX'],
        'Polish': ['EX', 'VG', '3EX'],
        'Sym.': ['EX', 'VG', '3EX'],
        'Fluor.': ['NON', 'FNT', 'MED'],
        'Culet': ['NON', 'NONE', 'POINTED'],
        'Tinge': lambda x: x.isna() | (x == 'NN'),
        'Milky': ['M0'],
        'BIT': lambda x: x != 'B2',
        'BIS': lambda x: x != 'B2',
        'TI': lambda x: x != 'W3',
        'Origin': lambda x: x != 'RUSSIA',
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }
    filters_bhavana = {
        'Lab':['GIA'],
        'Shape':['ROUND'],
        'Location': ['MUMBAI'],
        'Disc%': lambda x: x <= -10,
        'Carats': lambda x: ( x >= 0.30) & ( x <= 3.99),
        'Color' : ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity' : ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym':['EX', 'VG'],
        'Fluor':['NON', 'FNT', 'MED'],
        'Key To Symbols': lambda x : ~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'])

    }

    filters_d_navin = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Ct.': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['INDIA'],
        'Disc. (%)': lambda x: x <= -10,
        'Col.': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Fl.': ['NON', 'FNT', 'MED'],
        'Cut': ['EX', 'VG'],
        'Pol.': ['EX', 'VG'],
        'Sym.': ['EX', 'VG'],
        'BGM': lambda x: x.isna() | x.astype(str).str.strip().eq(''),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip','Brown patch of color, Pinpoint, Needle','Cavity, Chip','Chip, Cavity']),
        'Centre Black': lambda x: ~x.isin(['B2', 'B3']),
        'Crown Black': lambda x: ~x.isin(['B2', 'B3','CrownCentre']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)
    }

    filter_diaclust = {
        'Lab':['GIA'],
        'Shape':['ROUND'],
        'Location': ['MUMBAI'],
        'Disc': lambda x: x <= -10,
        'Wt': lambda x: (x >= 0.30) & (x <= 3.99),
        'Col':['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Cla': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Symm': ['EX', 'VG'],
        'Flr': ['NON', 'FNT', 'MED', 'NONE'],
        'BIC': lambda x: ~x.isna() & ~x.isin(['BC2', 'BC3']),
        'BIS': lambda x: ~x.isna() & ~x.isin(['BC2', 'BC3']),
        'KeyToSymbols': lambda x: ~x.isna() & ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_finstar = {
        'Lab': ['GIA'],
        'Shape': ['RD'],
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['IND'],
        'Disc %': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fls': ['NON', 'FNT', 'MED'],
        'Culet': ['NONE', 'POINTED', ''],
        'Shade': ['WHT'],
        'Brl.': ['EX', ''],
        'CN': lambda x: x != 'CN3',
        'SN': lambda x: x != 'SN3',
        'CW': lambda x: x != 'CW3',
        'SW': lambda x: x != 'SW3',
        'To': lambda x: x != 'Y',
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity, Chip', 'Chip, Cavity']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity", na=False)
    }

    filters_glowstar = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['India'],
        'Disc %': lambda x: pd.to_numeric(x, errors='coerce') <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        # 'Natts': lambda x: ~x.isin(['BC0 - BT2','BC0 - BT3','BC1 - BT3','BC2 - BT0','BC2 - BT1','BC2 - BT2','BC2 - BT3','BC3 - BT0','BC3 - BT2','BC3 - BT3']),
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'CULET': ['NON','Pointed'],
        'Shade': lambda x: x == 'NO',
        'BGM Comment': lambda x: x == 'No BROWN NO MILKY',
        'Comment': lambda x: ~x.str.startswith("Clar", na=False),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'])
    }

    filters_gyanimpex = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Net %': lambda x: -x if x >= 0 else x,
        'Net %': lambda x: x <= -10,
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['MUMBAI'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fluo': ['NON', 'FNT', 'MED'],
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comments': lambda x: ~x.str.startswith('Clarity', na=False)
    }
    
    filter_gemstar = {
        'Lab':['GIA'],
        'Shape':['ROUND'],
        'Location': ['IND'],
        'Price': lambda x:x <= - 10,
        'weight':lambda x: (x >= 0.30) & (x <= 3.99),
        'Color':['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut ': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flo': ['N', 'F', 'M'],
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity']),
    }

    filters_haridas = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Disc %': lambda x: x <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
    }

    filters_hk = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['MUMBAI'],
        'Rap%': lambda x: x <= -10,
        'Color Shade': ['WH'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Flo': ['N', 'F', 'M'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symmetry': ['EX', 'VG'],
        'Color Desc': lambda x: x.isna(),
        'BT': lambda x: ~x.isin(['BT2', 'BT3']),
        'BC': lambda x: ~x.isin(['BC2', 'BC3']),
        'WT': lambda x: ~x.isin(['WT3']),
        'WC': lambda x: ~x.isin(['WC3']),
        'Milky': lambda x: ~x.isin(['M1', 'M2', 'M3', 'NV']),
        'OPPV': lambda x: ~x.isin(['PO2', 'PO3']),
        'OPCR': lambda x: ~x.isin(['CO2', 'CO3']),
        'OPTA': lambda x: ~x.isin(['NV', 'TO1', 'TO2']),
        'Culet': lambda x: ~x.isin(['L', 'VSM', '']),
        'Flo Color': lambda x: ~x.isin(['YL', 'GR']),
        'Luster': lambda x: ~x.isin(['GD', 'VG']),
        'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'GIA Comments': lambda x: ~x.str.startswith("Clarity", na=False),
        'LAB': ['GIA'],
    }

    filters_hrv = {
        'Lab': ['GIA'],
        'Shape': ['BR'],
        'Location': ['MUMBAI'],
        'Disc %': lambda x: x <= -10,   
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'Culet Size': ['NON']
    }

    filters_jewelex = {
        'Lab': ['GIA'],
        'Shape': ['RBC'],
        'Disc %': lambda x: x <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flour': ['NON', 'FNT', 'MED'],
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_jodhani = { 
        'Lab': ['GIA'],
        'Shape': ['Round'],
        'Disc(%)':lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symmetry': ['EX', 'VG'],
        'Fluorescence': ['None','Faint','Medium'],
        'Luster':['EX'],
        'Key Symbol': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_kb_diamond = {
        'Lab': ['GIA'],
        'Shape': ['RD'],
        'DISC': lambda x: x.round(2) <= -10,
        'Location': ['INDIA'],
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        "Col": lambda x: x.fillna("").astype(str).str.strip().str.replace("+", "", regex=False),  # Updated
        "Cla": lambda x: x.fillna("").astype(str).str.strip().str.replace("+", "", regex=False),  # Updated
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Cla': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fluo': ['NONE', 'FNT', 'MED'],
        'SH': lambda x: x.isna() | (x == 'NONE'),
        'Black Side Inclusion': lambda x: x.isna() | (x == 'no black'),
        'Comment': lambda x: ~x.str.startswith("Clarity", na=False),
        'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_kbs = {
        'LAB': ['GIA'],
        'SHAPE': ['Round'],
        'DISC': lambda x: x <= -10,
        'LOC': ['INDIA'],
        'SIZE': lambda x: (x >= 0.30) & (x <= 3.99),
        'WEIGHT': lambda x: (x >= 0.30) & (x <= 3.99),
        'COL': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLR': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'FLR': ['N', 'F', 'M'],
        'Brown': lambda x: x.isna() | (x == 'BR0'),
        'MILKY':lambda x: x.isna() | (x == 'None'),
        'KTS': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_kc = {
        'Cert': ['GIA'],
        'Shape': ['Round'],
        'Dis %': lambda x: x <= -10,
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Carats': lambda x: (x >= 0.3) & (x <= 3.9),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fluo': ['N', 'F', 'MED'],
        'Culet': lambda x: x.isna(),
        'Comments': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity", na=False),
    }
    
    filters_kg = {
        'LAB': ['GIA'],
        'SHAPE': ['ROUND'],
        'CTS': lambda x: (x >= 0.30) & (x <= 3.99),
        'LOC.': ['Mumbai'],
        'DISC':lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'COLOR': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLARITY': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'FLUOR': ['NONE', 'FAINT', 'MEDIUM'],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYMM': ['EX', 'VG'],
        'SHADE': ['NONE', ''],
        'TABLE OPEN INCLUSION': ['NONE', ''],
        'LUSTER': ['EX', ''],
        'LUSTER': lambda x: x.isin(['EX', 'VG']),
        'SHADE': ["NONE"],
        'KEY TO SYMBOL': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity/Chip', 'Chip/Cavity']),
        'REPORT COMMENTS': lambda x: ~x.str.startswith("Clarity", na=False),
    }

    filters_kgk = {
        'Lab': ['GIA'],
        'Shape': ['BR'],
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'BACK': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Milky': ['M0'],
        'FLUO': ['NON', 'FNT', 'MED'],
        'Culet': lambda x: ~x.isin(['Long', 'Very Small']),
        'Color Shade': ['White'],
        'Luster': ['EX'],
        'Table Inclusion': lambda x: x != 'T3',
        'Side Inclusion': lambda x: x != 'S3',
        'Side Black': lambda x: ~x.isin(['SB2', 'SB3']),
        'Table Black': lambda x: ~x.isin(['B2', 'B3']),
        'CUT': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_kiran = {
        'Lab': ['GIA'],
        'Shp': ['RB'],
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Disc%': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clr': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Loc': ['INDIA'],
        'Milky': ['M0'],
        'Flr': ['Non', 'Fnt', 'Med'],
        'Culet': lambda x: ~x.isin(['Long', 'Very Small']),
        'Color Shade': ['White'],
        'Luster': ['EX'],
        'Table Inclusion': lambda x: x != 'T3',
        'Side Inclusion': lambda x: x != 'S3',
        'Side Black': lambda x: ~x.isin(['SB2', 'SB3']),
        'Table Black': lambda x: ~x.isin(['B2', 'B3']),
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'K T S': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
    }

    filters_lakhi = {
        'Lab': ['GIA'],
        'Shape': ['RBC'],
        'Country': ['INDIA'],
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Fluorescence Intensity': ['NON', 'FNT', 'MED'],
        'Culet Size': ['NON'],
        'Shade ': ['NONE'],
        'Milky': ['NONE'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symmetry': ['EX', 'VG'],
        'Rapnet  Discount %': lambda x: (x * 100).apply(lambda y: f"{y if y < 0 else -y:.2f}").astype(float) <= -10,
        'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
        'Key to symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Black Inclusion   ': lambda x: ~x.str.startswith("N3,", na=False),
        'Cert comment': lambda x: ~x.str.startswith("Clarity", na=False)
    }
    
    filters_m_d_brother= {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        'Loc': ['MUMBAI'],
        'Disc %': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
        'Size': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symm': ['EX', 'VG'],
        'Fls': ['NON', 'FNT', 'MED'],
        'Shade': lambda x: (x == 'NONE') | (x.isna()),
        'Milky': lambda x: (x == 'NONE') | (x.isna()),
        'Culet': lambda x: ~x.isin(['Very Small', 'VSM']),
        'Black': lambda x: ~x.isin(['SPPB2', 'CPPB2']),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comments': lambda x: ~x.str.startswith(('Clarity', '* This diamond has been'), na=False),
        'Grading Comments': lambda x: x.str.startswith('NO BGM', na=False)
    }

    filters_m_suresh = {
        'Lab': ['GIA'],
        'Shape': ['Round'],
        'Ct': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['Mumbai'],
        'Online Back': lambda x: x <= -10.00,
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clr': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Fls': ['NONE', 'FA', 'MD'],
        'BGM': lambda x: ~x.isin(['Yes', '']),
        'brown': lambda x: x.isna(),
        'Tab_BLK': lambda x: ~x.isin(['B2-T']),
        'Op_Tab': lambda x: x.isna(),
        'Culet': lambda x: x == 'NONE',
        'Milky/Luster': ['EX'],
        'Side_BLK': lambda x: ~x.isin(['B2-S', 'B3-S']),
        'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Report Comments': lambda x: ~x.str.startswith("Clarity grade is based on internal graining that is not shown.", na=False),
        'Rapnet Back':lambda x:x <= -10,

    }

    filters_mb = {
            'Cert Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['MUM'],
            'Disc': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Fluo': ['NON', 'FNT', 'MED'],
            'Shade': ['NN','NA'],
            'Milky': ['NN','NA'],
            'Table Black': ['NN', 'BT1', 'BT1+'],
            'Side Black': ['NN', 'BC1', 'BC1+'],
            'Open': lambda x: ~x.isin(['OT1']),
            'Culet': lambda x: ~x.isin(['VERY SMALL']),
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }
    
    filters_narola = {
            'Lab': ['GIA'],
            'Certi': ['GIA'],
            'Shape': ['ROUND'],
            'Location': ['INDIA'],
            'Disc%': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Disc(%)': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Flo': ['NON', 'FNT', 'MED'],
            'Shade':['NO BROWN , NO GREEN , NO MILKY'],
            'TableOpen': lambda x: x != 'SMALL-OPEN',
            'KeyToSym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
            'NGS Comment': lambda x: ~x.str.startswith('CULET SMALL', na=False)
        }

    filters_pansurya_impex = {
        'Lab': ['GIA'],
        'Shape': ['RD'],
        'Location': ['MUMBAI'],
        'Dis': lambda x: x <= -10,
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flor.': ['NON', 'FNT', 'MED'],
        'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comment': lambda x: ~x.str.startswith('Clarity', na=False)
    }

    filters_pragati = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Disc %': lambda x: x <= -10,
            'Size': lambda x: (x >= 0.30) & (x <= 3.99),
            'Loc': ['IN'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symm': ['EX', 'VG'],
            'Fls': ['NON', 'FNT', 'MED'],
            'Culet': lambda x: x != 'VSM',
            'Table Black': lambda x: x != 'TB-2',
            'Side Black': lambda x: x != 'SB-2',
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_pramukh = {
        'Lab': ['GIA'],
        'Shape': ['RBC'],
        'Disc%': lambda x: x <= -10,
        'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Lust': lambda x: x != 'VG',
        'T.BLK': lambda x: x != 'B2',
        'T.OP': lambda x: x != 'HL',
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['Ex', 'Vg'],
        'Pol': ['Ex', 'Vg'],
        'Sy': ['Ex', 'Vg'],
        'Flo': ['N', 'Fnt', 'Med'],
        'KeyToSym': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comment': lambda x: ~x.str.startswith('Clarity', na=False)
    }

    filters_rn_export = {
        'LAB': ['GIA'],
        'SHAPE': ['RBC'],
        'DISC': lambda x: (x * 100).apply(lambda y: f"{y if y < 0 else -y:.2f}").astype(float) <= -10,
        'LOC': ['IND'],
        'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', ' EX', 'EX ', 'VG', ' VG', 'VG '],
        'Polish': ['EX', ' EX', 'EX ', 'VG', ' VG', 'VG '],
        'Symm': ['EX', ' EX', 'EX ', 'VG', ' VG', 'VG '],
        'Fluor': ['N', 'F', ' F', 'M', ' M', ' N'],
        'Key to Sym': lambda x: ~x.isin(['Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
    }

    filters_r_rajesh = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['INDIA'],
            'Disc %': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Fls': ['NONE', 'FNT', 'MED'],
            'Tinge': ['NONE', ''],
            'TB': lambda x: ~x.isin(['TBL2', 'TBL3']),
            'SB': lambda x: ~x.isin(['SBL2', 'SBL3']),
            'TW': lambda x: ~x.isin(['CW3']),
            'SW': lambda x: ~x.isin(['SW3']),
            'Culet': lambda x: ~x.isin(['VERY SMALL']),
            'Milky': ['NONE', ''],
            'Open On Table': lambda x: ~x.isin(['MEDIUM TABLE OPEN', 'SMALL TABLE OPEN']),
            'Report Comments': lambda x: ~x.str.startswith('Clarity', na=False)
        }

    filters_rays = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Disc%': lambda x: x <= -10,
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'FL': ['NONE', 'FNT', 'MED'],
            'Shade': lambda x: (x == 'NONE') | (x.isna()),
            'Milky': lambda x: (x == 'NONE') | (x.isna()),
            'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),

            'Wgt': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['MUMBAI'],
            'Disc %': lambda x: x <= -10,
            'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Fls': ['NONE', 'FNT', 'MED'],
            'BLACK': ['T1','T2','C1','C2','C2','T2','T2C1','T2C2','T3C3','NO BLACK'],
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }
    
    filters_red_exim = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Disc': lambda x: x.apply(lambda y: -y if y >= 0 else y),
            'Disc': lambda x: x <= -10,
            'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
            'Country': ['INDIA'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'CutGrade': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'Flour': ['NON', 'FNT', 'MED'],
            'CuletSize': ['NONE', 'POINTED'],
            'Milky': lambda x: x.isna() | (x == 'NONE'),
            'Shade': lambda x: x.isna() | (x == 'NONE'),
            'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
            'Comment': lambda x: x.isin(['NO BGM ', 'NO BGM', 'NO BGM  '])
        }
    
    filters_ratnakala_mail = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Disc %': lambda x: x <= -10,
            'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['INDIA'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Flour': ['NON', 'FNT', 'MED'],
            'milky': ["NONE", "NON"],
            'Milky': ["NONE", "NON"],
            'shade': ['NN'],
            'Key To Symbols': lambda x: ~x.isin(["Cavity", "Chip"]),
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'Fluorescent': ['NON', 'FNT', 'MED'],
            'Milky': ["NONE", "NON"],
            'Shade': ['NN'],
            'TB': ['NONE'],
            'SB': ['NONE'],
            'Table Open': ['NONE'],
            'Key To Symbol': lambda x: ~x.isin(["Cavity", "Chip"]),
            'Country': ['INDIA'],
        }
    
    filters_sagar_list = {
        'Lab': ['GIA'],
        'Location': ['Mumbai', 'MUMBAI'],
        'Shape': ['RD'],
        'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
        'Back%':lambda x: x <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'FL': ['NON', 'FNT', 'MED'],
        'KeytoSymbol': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'Comments': lambda x: ~x.str.startswith('Clarity', na=False),
        'Back%': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10
    }

    filters_shairu = {
        'LAB': ['GIA '],
        'SHAPE': ['ROUND'],
        'CTS': lambda x: (x >= 0.30) & (x <= 3.99),
        'RAP DISC%': lambda x: x <= -10,
        'COL': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLAR': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'FL': ['NON', 'FNT', 'MED'],
        'KEY TO SYMBOL': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'REPORT COMMENT': lambda x: ~x.str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)
    }
    
    filters_shairus = {
        'Lab': ['GIA '],
        'Shape': ['ROUND'],
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Disc(%)': lambda x: x <= -10,
        'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish': ['EX', 'VG'],
        'Symm': ['EX', 'VG'],
        'Fls': ['NON', 'FNT', 'MED'],
        'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
        'GIA Comments': lambda x: ~x.str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)
    }
    
    filters_sanghvi_new = {
        'LAB': ['GIA'],
        'SHAPE': ['ROUND'],
        'CTS': lambda x: (x >= 0.30) & (x <= 3.99),
        'RAP DISC%': lambda x: x <= -10,
        'COL': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'CLAR': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'CUT': ['EX', 'VG'],
        'POL': ['EX', 'VG'],
        'SYM': ['EX', 'VG'],
        'FL': ['NON', 'FNT', 'MED'],
        'KEY TO SYMBOL': lambda x: ~x.astype(str).str.contains('Cavity|Chip', case=False, na=False),
        'REPORT COMMENT': lambda x: ~x.astype(str).str.strip().str.startswith("Clarity grade is based on internal graining that is not shown.", na=False)
    }

    filters_sheetal = {
            'Lab': ['GIA'],
            'Shape': ['RBC'],
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['INDIA'],
            'Disc%': lambda x: x <= -10,
            'Hue': ['WH'],
            'Milky': ['NO'],
            'BIT': ['--','BIT0','BIT1'],
            'BIC': ['--','BIC0','BIC1'],
            'BOC': ['--','BOC0','BOC1'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'Fluorescence': ['NON', 'FNT', 'MED'],
            'Keytosymbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_shyam = {
            'Lab': ['GIA'],
            'City': ['MUMBAI'],
            'Shape': ['ROUND'],
            'Disc.': lambda x: -x if x >= 0 else x,
            'Disc.': lambda x: x <= -10,
            'Weight': lambda x: (x >= 0.30) & (x <= 3.99),
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symmetry': ['EX', 'VG'],
            'Fl.Int.': ['NONE', 'FAINT', 'MEDIUM'],
            'Key to Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
            'Cert comment': lambda x: x.str.startswith('NO BGM', na=True) | x.str.startswith(' NO BGM', na=True) | x.str.startswith('NO BMG', na=True) | x.str.startswith(' NO BMG', na=True)
        }

    filters_ske = {
            'Lab': ['GIA'],
            'Shape': ['RBC'],
            'Disc%': lambda x: x <= -10,
            'Carats': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['MUMBAI'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symm': ['EX', 'VG'],
            'Flour.': ['NON', 'FNT', 'MED'],
            'Luster': lambda x: x != 'VG',
            'Table Open': lambda x: x == 'N',
            'Table Black': lambda x: x != 'TB2',
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_snj_list = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Final Disc%': lambda x: x <= -10,
            'Size': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['INDIA'],
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Table Black': lambda x: ~x.isin(['TN2', 'TN3']),
            'Side Black': lambda x: ~x.isin(['SN2', 'SN3']),
            'Luster': lambda x: x != 'VG',
            'Flo': ['NON', 'FNT', 'MED'],
            'Flo Color': lambda x: x != 'Yellow',
            'Table \nOpen': lambda x: (x == 'N') | (x.isna()),
            'Culet Size': lambda x: (x == 'NONE') | (x.isna()),
            'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    filters_srd = {
        'Lab': ['GIA'],
        'Shape': ['ROUND'],
        # 'Rap %': lambda x: x <= 10,
        'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
        'Loc': ['MUMBAI'],
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Cla': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'FL': ['N', 'F', 'M'],
        'CS': ['WH'],
        'BT': lambda x: x != 'BT2',
        'BC': lambda x: x != 'BC2',
        'TO': ['TO0'],
        'FlShd': lambda x: x != 'Yellow',
        'Milky': ['M0'],
        'Luster': ['EX'],
        'Fancy Color Description': lambda x: x.isna(),
        'Key To Sym': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
    }

    filters_sunrays = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            'Disc %': lambda x: -x if +x >= 0 else x,
            'Disc %': lambda x: x <= -10,
            'Size': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['IN'],
            'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Polish': ['EX', 'VG'],
            'Symm': ['EX', 'VG'],
            'Fls': ['NON', 'FNT', 'MED'],
            'Shade': ['WHT'],
            'Milky': lambda x: x != 'M01',
            'B I': lambda x: x != 'B02',
            'T I': lambda x: x != 'T02',
            'Culet': lambda x: ~x.isin(['VSM', 'Very Small']),
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity']),
            'Report Comments': lambda x: ~x.str.startswith('Clarity', na=False)
        }

    filters_starrays = {
            'Cert.': ['GIA'],
            'Shape': ['BR'],
            'Rap.%': lambda x: x <= -10,
            'Cts.': lambda x: (x >= 0.30) & (x <= 3.99),
            'Location': ['India'],
            'Col.': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Cla.': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', 'VG'],
            'Pol.': ['EX', 'VG'],
            'Sym.': ['EX', 'VG'],
            'FL': ['N', 'F', 'M'],
            'Luster': ['EX','VG','VG+'],
            'KeyToSymbols': lambda x: ~x.astype(str).str.contains('Cavity|Chip', case=False, na=False),
            'Comments': lambda x: ~x.astype(str).str.strip().str.startswith('Clarity', na=False)
        }

    filters_vd_global = {
        'LAB': ['GIA'],
        'Shape': ['RD'],
        'CTS': lambda x: (x >= 0.30) & (x <= 3.99),
        'Location': ['India'],
        'Disc%': lambda x: x <= -10,
        'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Cla': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol': ['EX', 'VG'],
        'Sym': ['EX', 'VG'],
        'Flo': ['NON', 'FNT', 'MED'],
        'CULET': lambda x: x != 'VSM',
        'BT': lambda x: ~x.isin(['H']),
        'BC': lambda x: ~x.isin(['H']),
        'BCR': lambda x: ~x.isin(['H']),
        'BP': lambda x: ~x.isin(['H']),
        'WT': lambda x: ~x.isin(['H']),
        'WC': lambda x: ~x.isin(['H']),
        'Key To Symbol': lambda x: ~x.isin(['Cavity', 'Cavity, Chip', 'Chip, Cavity']),
        'GIA Comments': lambda x: ~x.str.startswith("Clarity", na=False)
    }

    filter_vinay_diamond = {
        'LAB':['GIA'],
        'Shape':['RD'],
        'Net%': lambda x : x <= -10,
        'Carats': lambda x : (x >= 0.30 ) & (x <= 3.99),
        'Color' : ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity' : ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Polish':['EX', 'VG'],
        'Sym':['EX', 'VG'],
        'Fls':['NON', 'FNT', 'MED','NONE'],
        'Inclusion':lambda x:~x.isin(['BT0 BC2','BT1 BC2','BT1 BC3','BT2 BC2']),
        'KeytoSymbols':lambda x:~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'])

    }

    filter_vaibhav = {
        'LAB':['GIA'],
        'Lab':['GIA'],
        'Shape':['ROUND'],
        'Disc %': lambda x : x <= -10,
        'Wgt': lambda x : (x >= 0.30 ) & (x <= 3.99),
        'Col' : ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
        'Clarity' : ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
        'Cut': ['EX', 'VG'],
        'Pol':['EX', 'VG'],
        'Sym':['EX', 'VG'],
        'Fls':['NON', 'FNT', 'MED'],
        'Milky':['ML-0'],
        'Shade': lambda x: (x == 'NONE') | (x.isna()),
        'Key To Symbols':lambda x:~x.isin(['Cavity', 'Cavity', 'Chip','Cavity, Chip','Chip, Cavity'])
    }

    filters_veer_gems = {
            'Cert': ['GIA'],
            'Loc': ['Mum'],
            'Shape': ['ROUND'],
            'Disc': lambda x: x.apply(lambda y: -y if y >= 0 else y) <= -10,
            'Cts': lambda x: (x >= 0.30) & (x <= 3.99),
            'Col': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clrt': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut': ['EX', '3EX', 'VG'],
            'Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Flr': ['NN', 'FN', 'MED'],
            'CTB': lambda x: ~x.isin([2, 3]),
            'CTW': lambda x: ~x.isin([3]),
            'STB': lambda x: ~x.isin([2, 3]),
            'STW': lambda x: ~x.isin([3]),
            'CRB': lambda x: ~x.isin([2, 3]),
            'CRW': lambda x: ~x.isin([3]),
            'Culet': lambda x: x != 'OTHER',
            'Key To Symbol': lambda x: ~x.isin(['Chip', 'Cavity, Chip', 'Chip, Cavity', 'Cavity'])
        }
    
    filters_unique = {
            'Lab': ['GIA'],
            'Shape': ['ROUND'],
            ' Disc %': lambda x: x <= -10,
            'Carat': lambda x: (x >= 0.30) & (x <= 3.99),
            'Color': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'],
            'Clarity': ['IF', 'VVS1', 'VVS2', 'VS1', 'VS2', 'SI1', 'SI2'],
            'Cut ': ['EX', 'VG'],
            ' Pol': ['EX', 'VG'],
            'Sym': ['EX', 'VG'],
            'Fls': ['NON', 'FNT', 'MED'],
            'TI': lambda x: x == 'NON',
            'TB': ['MINOR BLACK','NON','PINPOINT '],
            'SB': ['MINOR BLACK','NON','PINPOINT '],
            'Key To Symbols': lambda x: ~x.isin(['Cavity', 'Cavity', 'Chip', 'Cavity, Chip', 'Chip, Cavity'])
        }

    # These header keys participate in the universal automatic header scan;
    # process_rapnet() performs the actual filtering and source validation.
    filters_rapnet = {
        'Lab': ['GIA'], 'Shape': ['Round'], 'Size': [],
        'Color': [], 'Clarity': [], 'Cut': [], 'Polish': [],
        'Symmetry': [], 'Fluorescence': [],
    }

    prefix_info = {
        "RAPNET LIST": {"filters": filters_rapnet, "skip_rows": 0},
        "RAPNET": {"filters": filters_rapnet, "skip_rows": 0},
        "AKARSH": {"filters": filters_akarsh, "skip_rows": 1},
        "ANKIT":  {"filters": filters_ankit,  "skip_rows": 10},   
        "AMORA":  {"filters": filters_amora,  "skip_rows": 0},
        "ASIAN STAR":[ {"filters": filters_asian_star, "skip_rows": 7}],
                    # {"filters": filters_asian_star_web, "skip_rows": 0}],
        "BHAVANA": {"filters":filters_bhavana, "skip_rows":4},
        "D NAVIN": [{"filters": filters_d_navin, "skip_rows": 1}],        
        "DIACLUST" : {"filters":filter_diaclust,"skip_rows":0},
        "FINSTAR": {"filters": filters_finstar, "skip_rows": 1},
        "GLOWSTAR": {"filters": filters_glowstar, "skip_rows": 4},
        "GNYAN IMPEX": {"filters": filters_gyanimpex, "skip_rows": 1},
        "GEMSTAR":{"filters":filter_gemstar,"skip_rows":0},
        "HIRADIAS LIST": {"filters": filters_haridas, "skip_rows": 0},
        "HK": [{"filters": filters_hk, "skip_rows": 0}],
                # {"filters": filters_hk_web, "skip_rows": 1}],
        "HRV": [{"filters": filters_hrv, "skip_rows": 0}],
                # {"filters": filters_hrv_mail, "skip_rows": 1}],
        "JEWELEX": {"filters": filters_jewelex, "skip_rows": 1},
        "JODHANI": {"filters": filters_jodhani, "skip_rows": 0},
        "KB":{"filters": filters_kb_diamond, "skip_rows": 2},
        "KBS":{"filters": filters_kbs, "skip_rows": 2},
        "KC": {"filters": filters_kc, "skip_rows":0},
        "KG": {"filters": filters_kg, "skip_rows": 0}, 
        "KGK": {"filters": filters_kgk, "skip_rows": 1},
        "KIRAN": {"filters": filters_kiran, "skip_rows": 10},
        "LAKHI": [ {"filters": filters_lakhi, "skip_rows": 0}],
                # {"filters": filters_lakhi_mail, "skip_rows": 5}],
        "M D": {"filters": filters_m_d_brother, "skip_rows": 1},
        "M SURESH":[{"filters": filters_m_suresh, "skip_rows": 1}],
                    # {"filters": filters_m_suresh_web, "skip_rows": 1}],
        "MB": {"filters": filters_mb, "skip_rows": 1},
        "NAROLA": {"filters": filters_narola, "skip_rows": 0},
        "PANSURYA": {"filters": filters_pansurya_impex, "skip_rows": 0},
        "PRAGATI": {"filters": filters_pragati, "skip_rows": 0},
        "PRAMUKH": {"filters": filters_pramukh, "skip_rows": 0},
        "R N EXPORT": {"filters": filters_rn_export, "skip_rows": 0},
        "R.RAJESH": {"filters": filters_r_rajesh, "skip_rows": 5},
        "RAYS": [{"filters": filters_rays, "skip_rows": 1}],
                # {"filters": filters_rays_mails, "skip_rows": 1} ],
        "RATNAKALA": {"filters": filters_ratnakala_mail, "skip_rows": 4},                            
        'RED EXIM': {"filters": filters_red_exim, "skip_rows": 0},
        'SAGRA LIST':{"filters": filters_sagar_list, "skip_rows": 0},
        'SHAIRU': [{"filters": filters_shairus, "skip_rows": 5},
                    {"filters": filters_shairu, "skip_rows": 5}],
        'SANGHVI': [{"filters": filters_sanghvi_new, "skip_rows": 0}],
                    # {"filters": filters_sanghvi_sons, "skip_rows": 0}],
        'SHEETAL': {"filters": filters_sheetal, "skip_rows": 8},
        'SHYAM': {"filters": filters_shyam, "skip_rows": 0},
        'SKE': {"filters": filters_ske, "skip_rows": 0},
        'SNJ LIST': {"filters": filters_snj_list, "skip_rows": 5},
        'SRD': {"filters": filters_srd, "skip_rows": 4},
        'SUNRAYS': {"filters": filters_sunrays, "skip_rows": 0},    
        'STAR RAYS': {"filters": filters_starrays, "skip_rows": 3},    
        "VD GLOBAL": {"filters": filters_vd_global, "skip_rows": 4},
        "VINAY":{"filters":filter_vinay_diamond,"skip_rows":1},
        "VAIBHAV":{"filters":filter_vaibhav,"skip_rows":1},
        'VEER GEMS': {"filters": filters_veer_gems, "skip_rows": 3},
        'UNIQUE': [{"filters": filters_unique, "skip_rows": 0},
                   {"filters": filters_unique, "skip_rows": 4}], 
    }

    with st.spinner("Processing files..."):
        for file in files:
            file_name = file.name
            prefix, configured = supplier_for_file(file_name, prefix_info)
            if configured:
                configurations = prefix_info[prefix]
            else:
                # A supplier not in prefix_info has no dedicated filters.
                # Run strictly validated, generic mapping and quality rules.
                configurations = {"filters": {}, "skip_rows": 0}
                st.info(
                    f"{file_name}: '{prefix}' has no supplier-specific rules. "
                    "Trying generic DEAL filters; confirm its output before use."
                )

            try:
                candidates = find_supplier_header_candidates(
                    file, configurations, column_mapping, generic=not configured
                )
            except Exception as exc:
                st.error(f"Could not inspect headers in {file_name}: {exc}")
                continue

            attempts = []
            successful = False
            for detected_skip_rows, config in candidates:
                try:
                    csv_data, row_count, output_data = process_data(
                        file_path=file,
                        filters=config["filters"],
                        column_mapping=column_mapping,
                        prefix=prefix,
                        skip_rows=detected_skip_rows,
                        generic=not configured,
                    )
                    if row_count <= 0:
                        attempts.append(
                            f"header row {detected_skip_rows + 1}: no records passed filters"
                        )
                        continue

                    download_name = f"{prefix.lower()} {current_datetime}.csv"
                    existing_names = {
                        info['download_name'] for info in st.session_state.deal_results.values()
                    }
                    if download_name in existing_names:
                        base_name = os.path.splitext(file_name)[0]
                        download_name = f"{base_name} {current_datetime}.csv"
                    st.session_state.deal_results[(file_name, prefix)] = {
                        "csv": csv_data,
                        "download_name": download_name,
                        "row_count": row_count,
                        "df": output_data,
                        "header_row": detected_skip_rows + 1,
                        "generic": not configured,
                    }
                    successful = True
                    break
                except Exception as exc:
                    attempts.append(f"header row {detected_skip_rows + 1}: {exc}")

            if not successful:
                st.warning(
                    f"No valid output for {file_name} ({prefix}). "
                    "Tried automatic header detection and skip_rows=0. "
                    + " | ".join(attempts[:5])
                    + (" | Generic supplier requires recognizable report, carat, color, "
                       "clarity, and discount columns."
                       if not configured else "")
                )

        # ---- Show every processed file inside an expander ----
        st.markdown("---")

        if not st.session_state.deal_results:
            st.warning("No matching DEAL output was created. Please check file names, supplier prefix, and file format.")
            return

        st.subheader("Processed Files")
        
        workbook_bytes = build_deal_workbook(st.session_state.deal_results)
        workbook_name = f"DEAL Files {current_datetime}.xlsx"
        st.download_button(
            label=f"Download {workbook_name}",
            data=workbook_bytes,
            file_name=workbook_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="download_deal_workbook",
        )

        total_files = len(st.session_state.deal_results)
        total_rows = sum(info.get("row_count", 0) for info in st.session_state.deal_results.values())

        with st.expander(f"All Files ({total_files} files | {total_rows} rows)", expanded=True):
            for index, ((file_name, prefix), info) in enumerate(st.session_state.deal_results.items(), start=1):
                csv_data = info["csv"]
                download_name = info["download_name"]
                row_count = info.get("row_count", 0)
                output_df = info.get("df")

                if output_df is None:
                    output_df = pd.read_csv(io.StringIO(csv_data))

                st.markdown(f"#### {index}. Sheet: {info.get('sheet_name', prefix)}")
                st.caption(
                    f"Source: {file_name} | Prefix: {prefix} | Rows: {row_count} "
                    f"| Header row detected: {info.get('header_row', '?')}"
                    + (" | GENERIC FILTERS (unconfigured supplier)" if info.get("generic") else "")
                )
                st.dataframe(output_df)

                if index < total_files:
                    st.markdown("---")

if __name__ == "__main__":
    all_deal()