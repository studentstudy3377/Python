from __future__ import annotations

import argparse
import html
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import parse_qs, quote, quote_plus, urlparse

import pandas as pd
import undetected_chromedriver as uc
from selenium.common.exceptions import SessionNotCreatedException, TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# -----------------------------------------------------------------------------
# Windows/undetected_chromedriver cleanup guard
# -----------------------------------------------------------------------------
# On Windows, especially with newer Python versions, undetected_chromedriver may
# print a harmless cleanup warning after all data is saved:
#   OSError: [WinError 6] The handle is invalid
# This happens inside uc.Chrome.__del__ after driver.quit() has already run.
# The guard below keeps the terminal clean without changing extraction behavior.
_ORIGINAL_UC_CHROME_DEL = getattr(uc.Chrome, "__del__", None)


def _quiet_uc_chrome_del(self) -> None:
    if _ORIGINAL_UC_CHROME_DEL is None:
        return
    try:
        _ORIGINAL_UC_CHROME_DEL(self)
    except OSError:
        # Destructor cleanup errors are non-fatal; the output file is already saved.
        pass
    except Exception:
        pass


uc.Chrome.__del__ = _quiet_uc_chrome_del


try:
    import tkinter as tk
    from tkinter import messagebox
    from tkinter.filedialog import askopenfilename
except Exception:
    tk = None
    messagebox = None
    askopenfilename = None

GIA_REPORT_CHECK_URL_TEMPLATE = "https://www.gia.edu/report-check?reportno={report_no}"
GIA_REPORT_CHECK_LANDING = "https://www.gia.edu/report-check-landing"

# Fast-processing defaults
FAST_POLL_INTERVAL = 0.20
FAST_PAGE_TIMEOUT = 20
FAST_INTER_REPORT_DELAY = 0.10
FAST_SAVE_EVERY = 10

# Default output Excel path. Change this one line if you want a different default.

DEFAULT_OUTPUT = fr"D:\Gaurav\Download\Certificate_{pd.Timestamp.now():%S}.xlsx"
# DEFAULT_OUTPUT = f"akarsh .csv"

DEFAULT_INPUT_COLUMNS = [
    "Report Number",
    "Input Report Number",
    "GIA Report Number",
    "Certificate Number",
    "Certificate No",
    "Report No",
    "Report No.",
]

OUTPUT_COLUMNS = [
    "Input Report Number",
    "GIA Report Number",
    "Report Date",
    "Report Date ISO",
    "Report Type",
    "Report Type Code",
    "Description",
    
    "Disclosed Source Name",
    "Disclosed Source Code",
    "Shape and Cut",
    "Measurements",
    "Carat Weight",
    "Color Grade",
    "Color Origin",
    "Color Distribution",
    "Clarity Grade",
    "Cut Grade",
    "Polish",
    "Symmetry",
    "Fluorescence",
    "Clarity Characteristics",
    "Inscription(s)",
    "Comments",
    "Total Depth",
    "Table Size",
    "Crn Ag",
    "Crn Ht",
    "Pav Ag",
    "Pav Dp",
    "Star Length",
    "Lower Half",
    "Girdle Thickness",
    "Culet",
]

DEBUG_COLUMNS = [
    "Report URL",
    "PDF URL",
    "Status",
    "Error",
    "Raw Page Text",
]

LABEL_VARIANTS: Dict[str, List[str]] = {
    "GIA Report Number": [
        "GIA Report Number",
        "Report Number",
        "Report No",
        "Report No.",
        "Certificate Number",
    ],
    "Report Date": ["Date of Issue", "Issue Date", "Report Date"],
    "Report Date ISO": ["Report Date ISO", "Issue Date ISO", "Date of Issue ISO"],
    "Report Type": [
        "Report Type",
        "Diamond Grading Report",
        "Diamond Dossier",
        "Colored Diamond Grading Report",
        "Laboratory-Grown Diamond Report",
    ],
    "Report Type Code": ["Report Type Code", "report_type_code"],
    "Description": [
        "Description",
        "Report Type",
        "Diamond Grading Report",
        "Diamond Dossier",
        "Colored Diamond Grading Report",
        "Laboratory-Grown Diamond Report",
    ],
    "Disclosed Source Name": ["Disclosed Source", "Disclosed Source Name", "Source Name"],
    "Disclosed Source Code": ["Disclosed Source Code", "Source Code"],
    "Shape and Cut": [
        "Shape and Cutting Style",
        "Shape and Cut",
        "Shape",
        "Cutting Style",
    ],
    "Measurements": ["Measurements", "Measurement"],
    "Carat Weight": ["Carat Weight", "Weight"],
    "Color Grade": ["Color Grade", "Color"],
    "Color Origin": ["Origin", "Color Origin", "color_origin"],
    "Color Distribution": ["Distribution", "Color Distribution", "color_distribution"],
    "Clarity Grade": ["Clarity Grade", "Clarity"],
    "Cut Grade": ["Cut Grade", "Cut"],
    "Polish": ["Polish"],
    "Symmetry": ["Symmetry"],
    "Fluorescence": ["Fluorescence"],
    "Clarity Characteristics": ["Clarity Characteristics", "clarity_characteristics"],
    "Inscription(s)": ["Inscription(s)", "Inscriptions", "Inscription", "Laser Inscription"],
    "Comments": ["Report Comments", "Comments", "Comment"],
    "Total Depth": ["Depth", "Depth %", "Total Depth", "depth_pct"],
    "Table Size": ["Table", "Table %", "Table Size", "Table Pct", "table_pct"],
    "Crn Ag": ["Crown Angle", "crown_angle"],
    "Crn Ht": ["Crown Height", "crown_height"],
    "Pav Ag": ["Pavilion Angle", "pavilion_angle"],
    "Pav Dp": ["Pavilion Depth", "pavilion_depth"],
    "Star Length": ["Star Length", "star_length"],
    "Lower Half": ["Lower Half", "lower_half"],
    "Girdle Thickness": ["Girdle", "Girdle Thickness"],
    "Culet": ["Culet"],
}

JSON_KEY_VARIANTS: Dict[str, List[str]] = {
    "GIA Report Number": ["report_number", "reportNumber", "reportno", "reportNo"],
    "Report Date": ["report_date", "reportDate", "dateOfIssue", "issueDate"],
    "Report Date ISO": ["report_date_iso", "reportDateIso", "reportDateISO"],
    "Report Type": ["report_type", "reportType"],
    "Report Type Code": ["report_type_code", "reportTypeCode"],
    "Description": ["report_type", "reportType", "description"],
    "Disclosed Source Name": ["disclosed_source_name", "disclosedSourceName"],
    "Disclosed Source Code": ["disclosed_source_code", "disclosedSourceCode"],
    "Shape and Cut": ["shape_and_cutting_style", "shapeAndCuttingStyle", "shape", "cuttingStyle"],
    "Measurements": ["measurements", "measurement"],
    "Carat Weight": ["carat_weight", "caratWeight", "weight"],
    "Color Grade": ["color_grade", "colorGrade", "color"],
    "Color Origin": ["color_origin", "colorOrigin", "origin"],
    "Color Distribution": ["color_distribution", "colorDistribution", "distribution"],
    "Clarity Grade": ["clarity_grade", "clarityGrade", "clarity"],
    "Cut Grade": ["cut_grade", "cutGrade", "cut"],
    "Polish": ["polish"],
    "Symmetry": ["symmetry"],
    "Fluorescence": ["fluorescence"],
    "Clarity Characteristics": ["clarity_characteristics", "clarityCharacteristics"],
    "Inscription(s)": ["inscriptions", "inscription", "laserInscription"],
    "Comments": ["report_comments", "reportComments", "comments"],
    "Total Depth": ["depth_pct", "depthPct", "depth"],
    "Table Size": ["table_pct", "tablePct", "table"],
    "Crn Ag": ["crown_angle", "crownAngle"],
    "Crn Ht": ["crown_height", "crownHeight"],
    "Pav Ag": ["pavilion_angle", "pavilionAngle"],
    "Pav Dp": ["pavilion_depth", "pavilionDepth"],
    "Star Length": ["star_length", "starLength"],
    "Lower Half": ["lower_half", "lowerHalf"],
    "Girdle Thickness": ["girdle", "girdleThickness"],
    "Culet": ["culet"],
}

JSON_PATH_VARIANTS: Dict[str, List[Tuple[str, ...]]] = {
    "GIA Report Number": [("report_number",), ("reportNumber",), ("reportno",), ("reportNo",)],
    "Report Date": [("report_date",), ("reportDate",), ("dateOfIssue",), ("issueDate",)],
    "Report Date ISO": [("report_date_iso",), ("reportDateIso",), ("reportDateISO",)],
    "Report Type": [("report_type",), ("reportType",)],
    "Report Type Code": [("report_type_code",), ("reportTypeCode",)],
    "Description": [("report_type",), ("reportType",), ("description",)],
    "Disclosed Source Name": [
        ("industry_disclosures", "disclosed_source", "name"),
        ("industryDisclosures", "disclosedSource", "name"),
        ("disclosed_source", "name"),
        ("disclosedSource", "name"),
    ],
    "Disclosed Source Code": [
        ("industry_disclosures", "disclosed_source", "code"),
        ("industryDisclosures", "disclosedSource", "code"),
        ("disclosed_source", "code"),
        ("disclosedSource", "code"),
    ],
    "Shape and Cut": [
        ("results", "shape_and_cutting_style"),
        ("results", "shapeAndCuttingStyle"),
        ("shape_and_cutting_style",),
        ("shapeAndCuttingStyle",),
    ],
    "Measurements": [("results", "measurements"), ("measurements",)],
    "Carat Weight": [
        ("results", "carat_weight"),
        ("results", "caratWeight"),
        ("carat_weight",),
        ("caratWeight",),
    ],
    "Color Grade": [
        ("results", "color_grade"),
        ("results", "colorGrade"),
        ("color_grade",),
        ("colorGrade",),
    ],
    "Color Origin": [
        ("results", "color_origin"),
        ("results", "colorOrigin"),
        ("color_origin",),
        ("colorOrigin",),
    ],
    "Color Distribution": [
        ("results", "color_distribution"),
        ("results", "colorDistribution"),
        ("color_distribution",),
        ("colorDistribution",),
    ],
    "Clarity Grade": [
        ("results", "clarity_grade"),
        ("results", "clarityGrade"),
        ("clarity_grade",),
        ("clarityGrade",),
    ],
    "Cut Grade": [
        ("results", "cut_grade"),
        ("results", "cutGrade"),
        ("cut_grade",),
        ("cutGrade",),
    ],
    "Polish": [("results", "polish"), ("polish",)],
    "Symmetry": [("results", "symmetry"), ("symmetry",)],
    "Fluorescence": [("results", "fluorescence"), ("fluorescence",)],
    "Clarity Characteristics": [
        ("results", "clarity_characteristics"),
        ("results", "clarityCharacteristics"),
        ("clarity_characteristics",),
        ("clarityCharacteristics",),
    ],
    "Inscription(s)": [("results", "inscriptions"), ("inscriptions",), ("laserInscription",)],
    "Comments": [
        ("results", "report_comments"),
        ("results", "reportComments"),
        ("report_comments",),
        ("reportComments",),
        ("comments",),
    ],
    "Total Depth": [
        ("results", "proportions", "depth_pct"),
        ("results", "proportions", "depthPct"),
        ("proportions", "depth_pct"),
        ("proportions", "depthPct"),
        ("depth_pct",),
        ("depthPct",),
    ],
    "Table Size": [
        ("results", "proportions", "table_pct"),
        ("results", "proportions", "tablePct"),
        ("proportions", "table_pct"),
        ("proportions", "tablePct"),
        ("table_pct",),
        ("tablePct",),
    ],
    "Crn Ag": [
        ("results", "proportions", "crown_angle"),
        ("results", "proportions", "crownAngle"),
        ("proportions", "crown_angle"),
        ("proportions", "crownAngle"),
        ("crown_angle",),
        ("crownAngle",),
    ],
    "Crn Ht": [
        ("results", "proportions", "crown_height"),
        ("results", "proportions", "crownHeight"),
        ("proportions", "crown_height"),
        ("proportions", "crownHeight"),
        ("crown_height",),
        ("crownHeight",),
    ],
    "Pav Ag": [
        ("results", "proportions", "pavilion_angle"),
        ("results", "proportions", "pavilionAngle"),
        ("proportions", "pavilion_angle"),
        ("proportions", "pavilionAngle"),
        ("pavilion_angle",),
        ("pavilionAngle",),
    ],
    "Pav Dp": [
        ("results", "proportions", "pavilion_depth"),
        ("results", "proportions", "pavilionDepth"),
        ("proportions", "pavilion_depth"),
        ("proportions", "pavilionDepth"),
        ("pavilion_depth",),
        ("pavilionDepth",),
    ],
    "Star Length": [
        ("results", "proportions", "star_length"),
        ("results", "proportions", "starLength"),
        ("proportions", "star_length"),
        ("proportions", "starLength"),
        ("star_length",),
        ("starLength",),
    ],
    "Lower Half": [
        ("results", "proportions", "lower_half"),
        ("results", "proportions", "lowerHalf"),
        ("proportions", "lower_half"),
        ("proportions", "lowerHalf"),
        ("lower_half",),
        ("lowerHalf",),
    ],
    "Girdle Thickness": [
        ("results", "proportions", "girdle"),
        ("proportions", "girdle"),
        ("girdle",),
        ("girdleThickness",),
    ],
    "Culet": [("results", "proportions", "culet"), ("proportions", "culet"), ("culet",)],
}

ALL_KNOWN_LABELS = {
    re.sub(r"[^a-z0-9]+", " ", variant.lower()).strip()
    for variants in LABEL_VARIANTS.values()
    for variant in variants
}


def clean_report_number(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    text = re.sub(r"\s+", "", text)
    return text


def clean_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if isinstance(value, (list, tuple, set)):
        return " | ".join(part for part in (clean_value(v) for v in value) if part)
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    text = str(value)
    text = html.unescape(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
    text = re.sub(r"[ \t\f\v]+", " ", text)
    text = re.sub(r"\n+", " | ", text)
    return text.strip(" :|/")


def canonical_label(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).strip()


def normalize_spaces(text: str) -> str:
    text = html.unescape(text or "").replace("\xa0", " ")
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def is_known_label(line: str) -> bool:
    return canonical_label(line) in ALL_KNOWN_LABELS


def is_boilerplate(line: str) -> bool:
    normalized = canonical_label(line)
    if not normalized:
        return True
    bad_fragments = [
        "report check",
        "look up",
        "verify another report",
        "enter report",
        "learn more",
        "gem encyclopedia",
        "cookies and related technologies",
        "privacy notice",
        "terms of use",
        "copyright",
        "download report pdf",
        "email report",
        "important limitations",
        "all report check information",
        "submit a gem",
        "list your store",
        "gia store",
        "my gia",
        "download the new gia app",
    ]
    return any(fragment in normalized for fragment in bad_fragments)


def split_clean_lines(text: str) -> List[str]:
    lines: List[str] = []
    for line in normalize_spaces(text).splitlines():
        cleaned = clean_value(line)
        if cleaned:
            lines.append(cleaned)
    return lines


def same_line_label_value(lines: List[str], variants: List[str]) -> str:
    for line in lines:
        for variant in variants:
            pattern = rf"^\s*{re.escape(variant)}\s*[:\-.]?\s+(.+?)\s*$"
            match = re.match(pattern, line, flags=re.I)
            if not match:
                continue
            value = clean_value(match.group(1))
            if value and not is_known_label(value):
                return value
    return ""


def next_value_after_label(lines: List[str], variants: List[str]) -> str:
    normalized_variants = {canonical_label(v) for v in variants}
    for idx, line in enumerate(lines):
        if canonical_label(line) not in normalized_variants:
            continue
        for candidate in lines[idx + 1 : idx + 8]:
            if is_known_label(candidate):
                break
            if is_boilerplate(candidate):
                continue
            return clean_value(candidate)
    return ""


def parse_text_fields(page_text: str) -> Dict[str, str]:
    lines = [line for line in split_clean_lines(page_text) if not is_boilerplate(line)]
    data: Dict[str, str] = {}

    for output_field, variants in LABEL_VARIANTS.items():
        value = same_line_label_value(lines, variants)
        if not value:
            value = next_value_after_label(lines, variants)
        if value:
            data[output_field] = value

    if not data.get("Description"):
        for line in lines:
            if re.search(r"\b(Diamond|Colored Diamond|Laboratory-Grown).+\bReport\b", line, re.I):
                data["Description"] = clean_value(line)
                break

    return data


def stringify_json_value(value: Any) -> str:
    if isinstance(value, dict):
        for key in ("full", "text", "value", "name", "code", "display", "label"):
            if key in value:
                cleaned = stringify_json_value(value[key])
                if cleaned:
                    return cleaned
        return clean_value(value)
    return clean_value(value)


def walk_json(obj: Any) -> Iterable[Tuple[str, Any]]:
    if isinstance(obj, dict):
        for key, value in obj.items():
            yield str(key), value
            yield from walk_json(value)
    elif isinstance(obj, list):
        for item in obj:
            yield from walk_json(item)


def walk_dicts(obj: Any) -> Iterable[Dict[str, Any]]:
    if isinstance(obj, dict):
        yield obj
        for value in obj.values():
            yield from walk_dicts(value)
    elif isinstance(obj, list):
        for item in obj:
            yield from walk_dicts(item)


def dict_get_case_insensitive(obj: Dict[str, Any], wanted_key: str) -> Tuple[bool, Any]:
    if wanted_key in obj:
        return True, obj[wanted_key]
    wanted_norm = wanted_key.lower()
    for key, value in obj.items():
        if str(key).lower() == wanted_norm:
            return True, value
    return False, None


def get_json_path_value(obj: Dict[str, Any], path: Tuple[str, ...]) -> Any:
    current: Any = obj
    for key in path:
        if not isinstance(current, dict):
            return None
        found, current = dict_get_case_insensitive(current, key)
        if not found:
            return None
    return current


def parse_json_path_fields(obj: Any) -> Dict[str, str]:
    """Extract exact nested GraphQL-style fields without confusing generic keys like name/code."""
    data: Dict[str, str] = {}
    for candidate in walk_dicts(obj):
        for output_field, paths in JSON_PATH_VARIANTS.items():
            if data.get(output_field):
                continue
            for path in paths:
                value = get_json_path_value(candidate, path)
                cleaned = stringify_json_value(value)
                if cleaned and cleaned.lower() not in {"null", "none"}:
                    data[output_field] = cleaned
                    break
    return data


def extract_json_objects_from_scripts(page_source: str) -> Iterable[Any]:
    source = html.unescape(page_source or "")
    for script_text in re.findall(r"<script[^>]*>(.*?)</script>", source, flags=re.I | re.S):
        script_text = script_text.strip()
        if not script_text:
            continue
        if script_text.startswith("{") or script_text.startswith("["):
            try:
                yield json.loads(script_text)
                continue
            except Exception:
                pass
        for match in re.finditer(r"=\s*({.*?})\s*;", script_text, flags=re.S):
            candidate = match.group(1)
            if len(candidate) < 30:
                continue
            try:
                yield json.loads(candidate)
            except Exception:
                continue


def parse_embedded_json_fields(page_source: str) -> Dict[str, str]:
    data: Dict[str, str] = {}

    for obj in extract_json_objects_from_scripts(page_source):
        # Prefer exact nested GraphQL paths first, so fields like
        # industry_disclosures.disclosed_source.name/code are not confused with
        # unrelated generic keys elsewhere on the page.
        for key, value in parse_json_path_fields(obj).items():
            data.setdefault(key, value)

        for output_field, keys in JSON_KEY_VARIANTS.items():
            if data.get(output_field):
                continue
            key_set = {k.lower() for k in keys}
            for key, value in walk_json(obj):
                if key.lower() in key_set:
                    cleaned = stringify_json_value(value)
                    if cleaned and cleaned.lower() not in {"null", "none"}:
                        data[output_field] = cleaned
                        break

    unescaped = html.unescape(page_source or "").replace("\\u0022", '"').replace("\\/", "/")
    for output_field, keys in JSON_KEY_VARIANTS.items():
        if data.get(output_field):
            continue
        for key in keys:
            for pattern in [
                rf'"{re.escape(key)}"\s*:\s*"((?:\\.|[^"\\])*)"',
                rf"'{re.escape(key)}'\s*:\s*'((?:\\.|[^'\\])*)'",
            ]:
                match = re.search(pattern, unescaped, flags=re.I)
                if not match:
                    continue
                raw = match.group(1)
                try:
                    raw = bytes(raw, "utf-8").decode("unicode_escape")
                except Exception:
                    pass
                cleaned = clean_value(raw)
                if cleaned:
                    data[output_field] = cleaned
                    break
            if data.get(output_field):
                break

    return data


def page_indicates_block_or_captcha(text: str) -> bool:
    normalized = canonical_label(text)
    signals = [
        "captcha",
        "verify you are human",
        "access denied",
        "blocked",
        "unusual traffic",
        "cloudflare",
    ]
    return any(signal in normalized for signal in signals)


def get_body_text(driver: uc.Chrome) -> str:
    try:
        text = driver.execute_script("return document.body ? document.body.innerText : '';")
        if text:
            return str(text)
    except Exception:
        pass
    try:
        return driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""


def wait_for_page_body(driver: uc.Chrome, timeout: int = 45) -> None:
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.TAG_NAME, "body")))


def js_click(driver: uc.Chrome, element: WebElement) -> bool:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.1)
        driver.execute_script("arguments[0].click();", element)
        return True
    except Exception:
        try:
            element.click()
            return True
        except Exception:
            return False


def find_clickable_by_text(driver: uc.Chrome, labels: List[str]) -> Optional[WebElement]:
    wanted = [canonical_label(label) for label in labels]
    xpath = "//*[(self::button or self::a or @role='button' or self::input or self::div or self::span)]"
    try:
        elements = driver.find_elements(By.XPATH, xpath)
    except Exception:
        return None
    for element in elements:
        try:
            if not element.is_displayed():
                continue
            text = element.text or element.get_attribute("value") or element.get_attribute("aria-label") or ""
            text_norm = canonical_label(text)
            if not text_norm:
                continue
            if any(text_norm == label or label in text_norm for label in wanted):
                return element
        except Exception:
            continue
    return None


def settings_submitted_visible(driver: uc.Chrome) -> bool:
    body = canonical_label(get_body_text(driver))
    return "settings submitted" in body and "cookie settings" in body


def cookie_consent_modal_visible(driver: uc.Chrome) -> bool:
    """Return True only for the blocking TrustArc/GIA consent modal, not footer links."""
    body = canonical_label(get_body_text(driver))
    has_modal_words = (
        "cookies and related technologies" in body
        or ("required cookies" in body and "functional cookies" in body and "advertising cookies" in body)
    )
    has_actions = any(
        find_clickable_by_text(driver, [label]) is not None
        for label in ["Accept All", "Decline All", "Submit Preferences", "Close"]
    )
    return bool(has_modal_words and has_actions) or settings_submitted_visible(driver)


def close_settings_submitted_dialog(driver: uc.Chrome) -> bool:
    if not settings_submitted_visible(driver):
        return False
    close_btn = find_clickable_by_text(driver, ["Close"])
    if close_btn and js_click(driver, close_btn):
        time.sleep(0.8)
        return True
    return False


def click_cookie_action(driver: uc.Chrome, labels: List[str]) -> bool:
    for label in labels:
        btn = find_clickable_by_text(driver, [label])
        if btn and js_click(driver, btn):
            time.sleep(1.0)
            close_settings_submitted_dialog(driver)
            return True
    return False


def wait_for_user_cookie_acceptance(driver: uc.Chrome, context: str = "") -> None:
    """Soft/manual cookie handling: user accepts cookies, then extraction continues."""
    print("\n" + "=" * 72)
    print("GIA cookie step")
    print("Chrome is open on the GIA page.")
    print("Please click ACCEPT ALL in the GIA cookie popup.")
    print("If a 'Settings Submitted' popup appears, you may leave it open; the script will close it.")
    if context:
        print(f"Context: {context}")
    input("After you click ACCEPT ALL, press ENTER here to start extraction...")

    # Smoothly close the follow-up TrustArc dialog if it appears after Accept All.
    for _ in range(8):
        if close_settings_submitted_dialog(driver):
            break
        if not settings_submitted_visible(driver):
            break
        time.sleep(0.5)

    if cookie_consent_modal_visible(driver):
        print("Cookie popup still looks visible. Please make sure ACCEPT ALL was clicked. Continuing anyway...")
    print("Cookie step finished. Starting extraction...")
    print("=" * 72 + "\n")


def handle_cookie_popups(
    driver: uc.Chrome,
    cookie_mode: str = "manual",
    context: str = "",
) -> None:
    """Handle GIA/TrustArc cookies without creating a rough stop-start flow.

    cookie_mode:
        manual       - ask the user to click ACCEPT ALL, then continue.
        auto-accept  - click ACCEPT ALL automatically.
        auto-decline - click DECLINE ALL automatically.
        skip         - do nothing.
    """
    close_settings_submitted_dialog(driver)

    if cookie_mode == "skip":
        return

    if not cookie_consent_modal_visible(driver):
        return

    if cookie_mode == "auto-accept":
        click_cookie_action(driver, ["Accept All"])
        return

    if cookie_mode == "auto-decline":
        click_cookie_action(driver, ["Decline All", "Submit Preferences"])
        return

    # Default: soft/manual handling requested by the user.
    wait_for_user_cookie_acceptance(driver, context=context)


def prepare_gia_cookie_session(
    driver: uc.Chrome,
    cookie_mode: str,
    timeout: int,
) -> None:
    """Open GIA once, let cookie consent settle, then start batch extraction."""
    if cookie_mode == "skip":
        return

    print("Opening GIA once for cookie/session setup...")
    driver.get(GIA_REPORT_CHECK_LANDING)
    wait_for_page_body(driver, timeout=timeout)
    handle_cookie_popups(driver, cookie_mode=cookie_mode, context="initial setup before extracting reports")
    close_settings_submitted_dialog(driver)


def wait_for_report_loaded(driver: uc.Chrome, report_no: str, timeout: int, cookie_mode: str) -> str:
    deadline = time.time() + timeout
    best_text = ""
    cookie_prompted_this_page = False

    while time.time() < deadline:
        if not cookie_prompted_this_page and cookie_consent_modal_visible(driver):
            handle_cookie_popups(driver, cookie_mode=cookie_mode, context=f"report {report_no}")
            cookie_prompted_this_page = True

        close_settings_submitted_dialog(driver)
        text = get_body_text(driver)
        if text:
            best_text = text
        text_norm = canonical_label(text)
        report_norm = canonical_label(report_no)

        if page_indicates_block_or_captcha(text):
            return text

        valid_report_tokens = ["gia report number", "date of issue", "grading results", "diamond grading report"]
        if report_norm and report_norm in text_norm and any(token in text_norm for token in valid_report_tokens):
            return text

        invalid_tokens = ["does not exist", "not available", "invalid report", "no report", "unable to process"]
        if any(token in text_norm for token in invalid_tokens):
            return text

        time.sleep(FAST_POLL_INTERVAL)
    return best_text


def find_report_input(driver: uc.Chrome) -> Optional[WebElement]:
    xpaths = [
        "//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'report')]",
        "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'report')]",
        "//input[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'report')]",
        "//input[@type='text' or @type='search' or not(@type)]",
    ]
    for xpath in xpaths:
        try:
            for element in driver.find_elements(By.XPATH, xpath):
                if element.is_displayed() and element.is_enabled():
                    return element
        except Exception:
            continue
    return None


def submit_report_number(driver: uc.Chrome, report_no: str) -> bool:
    input_element = find_report_input(driver)
    if input_element is None:
        return False
    try:
        input_element.clear()
        input_element.send_keys(report_no)
        time.sleep(0.2)
        button = find_clickable_by_text(driver, ["Look Up", "Lookup", "Search", "Submit", "Go"])
        if button and js_click(driver, button):
            return True
        input_element.send_keys(Keys.ENTER)
        return True
    except Exception:
        return False



def build_report_entry_page_url() -> str:
    # Build the first browser tab for entering multiple GIA report numbers.
    page = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>GIA Multiple Report Entry</title>
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      font-family: Arial, Helvetica, sans-serif;
      background: #f3f5f8;
      color: #182230;
    }
    .card {
      width: min(680px, calc(100vw - 32px));
      padding: 34px;
      background: #ffffff;
      border: 1px solid #dfe5ec;
      border-radius: 16px;
      box-shadow: 0 16px 50px rgba(20, 31, 48, 0.12);
    }
    h1 { margin: 0 0 10px; font-size: 28px; }
    p { margin: 0 0 22px; color: #5b6574; line-height: 1.5; }
    label { display: block; margin-bottom: 8px; font-weight: 700; }
    textarea {
      width: 100%;
      min-height: 230px;
      resize: vertical;
      padding: 14px;
      border: 1px solid #b8c2cf;
      border-radius: 9px;
      font-family: Consolas, "Courier New", monospace;
      font-size: 17px;
      line-height: 1.5;
      outline: none;
    }
    textarea:focus { border-color: #246bfd; box-shadow: 0 0 0 3px rgba(36,107,253,.15); }
    button {
      width: 100%;
      height: 50px;
      margin-top: 16px;
      border: 0;
      border-radius: 9px;
      background: #246bfd;
      color: white;
      font-size: 17px;
      font-weight: 700;
      cursor: pointer;
    }
    button:hover { background: #1556d8; }
    button:disabled { background: #91a9d7; cursor: default; }
    #message { min-height: 24px; margin-top: 15px; font-weight: 700; }
    #resultsLog {
      max-height: 190px;
      overflow-y: auto;
      margin-top: 12px;
      padding: 0;
      font-family: Consolas, "Courier New", monospace;
      font-size: 14px;
      line-height: 1.55;
      white-space: pre-wrap;
    }
    .error { color: #b42318; }
    .success { color: #067647; }
    .working { color: #175cd3; }
    .help { margin-top: 8px; font-size: 13px; color: #667085; }
  </style>
</head>
<body>
  <main class="card">
    <h1>GIA Multiple Report Check</h1>
    <p>Paste all GIA report numbers below, then click Submit. They will be processed one by one and saved into one Excel file.</p>
    <form id="reportForm" autocomplete="off">
      <label for="reportNumbers">Report Numbers</label>
      <textarea id="reportNumbers" name="reportNumbers"
                placeholder="Paste one report number per line&#10;Example:&#10;1234567890&#10;2345678901&#10;3456789012"
                autofocus></textarea>
      <div class="help">Separate numbers using new lines, commas, semicolons, tabs, or spaces. Duplicate numbers are removed.</div>
      <button id="submitButton" type="submit">Submit All Reports</button>
      <div id="message" aria-live="polite"></div>
      <div id="resultsLog" aria-live="polite"></div>
      <div id="submittedReportNumbers" hidden></div>
    </form>
  </main>
  <script>
    const form = document.getElementById('reportForm');
    const input = document.getElementById('reportNumbers');
    const button = document.getElementById('submitButton');
    const message = document.getElementById('message');
    const submitted = document.getElementById('submittedReportNumbers');

    form.addEventListener('submit', function (event) {
      event.preventDefault();
      const parts = input.value
        .split(/[\s,;]+/)
        .map(value => value.trim().replace(/\s+/g, ''))
        .filter(Boolean);
      const values = [...new Set(parts)];

      if (!values.length) {
        message.className = 'error';
        message.textContent = 'Please paste at least one report number.';
        input.focus();
        return;
      }

      const invalid = values.filter(value => !/^[A-Za-z0-9-]{6,30}$/.test(value));
      if (invalid.length) {
        message.className = 'error';
        message.textContent = 'Invalid report number: ' + invalid[0];
        input.focus();
        return;
      }

      submitted.textContent = JSON.stringify(values);
      input.readOnly = true;
      button.disabled = true;
      button.textContent = 'Processing...';
      message.className = 'working';
      message.textContent = 'Submitted ' + values.length + ' report(s). Extraction is starting in a new tab.';
      document.title = 'GIA Reports Processing';
    });
  </script>
</body>
</html>"""
    return "data:text/html;charset=utf-8," + quote(page)


def wait_for_report_numbers_from_first_tab(driver: uc.Chrome, timeout: int = 0) -> Tuple[List[str], str]:
    # Show the custom first tab and wait for multiple report numbers.
    print("\n" + "=" * 72)
    print("GIA multiple report entry")
    print("1. Paste one or more report numbers into the first browser tab.")
    print("2. Click Submit All Reports.")
    print("3. GIA will open in a second tab and process every report one by one.")
    print("=" * 72 + "\n")

    driver.get(build_report_entry_page_url())
    first_tab = driver.current_window_handle
    started = time.time()

    while True:
        if timeout > 0 and (time.time() - started) > timeout:
            raise TimeoutException("Timed out waiting for report numbers on the first tab.")

        try:
            submitted = driver.find_element(By.ID, "submittedReportNumbers")
            raw = submitted.get_attribute("textContent") or ""
            if raw:
                try:
                    parsed = json.loads(raw)
                except Exception:
                    parsed = []
                report_numbers = parse_pasted_report_numbers("\n".join(str(v) for v in parsed))
                if report_numbers:
                    return report_numbers, first_tab
        except Exception:
            pass

        time.sleep(0.10)

def open_extraction_tab(driver: uc.Chrome, first_tab: str) -> str:
    """Open and switch to the second tab where GIA extraction will run."""
    existing = set(driver.window_handles)
    driver.execute_script("window.open('about:blank', '_blank');")
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(existing))
    new_tabs = [handle for handle in driver.window_handles if handle not in existing]
    extraction_tab = new_tabs[-1]
    driver.switch_to.window(extraction_tab)
    return extraction_tab


def update_first_tab_progress(
    driver: uc.Chrome,
    first_tab: str,
    report_no: str,
    position: int,
    total: int,
    status: str,
    error: str = "",
) -> None:
    # Update progress and append each report result in the first tab.
    try:
        current_tab = driver.current_window_handle
        driver.switch_to.window(first_tab)

        succeeded = status.lower() == "success"
        line = f"[{position}/{total}] {report_no}: " + (
            "Success" if succeeded else f"Failed - {error or 'No report data found'}"
        )
        completed = position >= total

        if completed:
            headline = f"Completed {total} report(s). The Excel file has been saved."
            button_text = "Completed"
            title = "GIA Reports Completed"
            css_class = "success"
        else:
            headline = f"Processed {position} of {total}. Continuing with the next report..."
            button_text = f"Processing {position}/{total}..."
            title = f"GIA Processing {position}/{total}"
            css_class = "working"

        driver.execute_script(
            """
            const message = document.getElementById('message');
            const button = document.getElementById('submitButton');
            const log = document.getElementById('resultsLog');
            if (message) {
              message.className = arguments[1];
              message.textContent = arguments[0];
            }
            if (button) button.textContent = arguments[2];
            if (log) {
              log.textContent += (log.textContent ? '\\n' : '') + arguments[3];
              log.scrollTop = log.scrollHeight;
            }
            document.title = arguments[4];
            """,
            headline,
            css_class,
            button_text,
            line,
            title,
        )
        driver.switch_to.window(current_tab)
    except Exception:
        pass

def report_number_from_url(url: str) -> str:
    """Read reportno/report-number values from the current GIA URL."""
    try:
        query = parse_qs(urlparse(url or "").query)
    except Exception:
        return ""

    for key in ("reportno", "reportNo", "report_number", "reportNumber"):
        values = query.get(key) or []
        if values:
            value = clean_report_number(values[0])
            if value:
                return value
    return ""


def report_number_from_page(driver: uc.Chrome, page_text: str, last_entered: str = "") -> str:
    """Determine the report number the user submitted on the visible GIA page."""
    from_url = report_number_from_url(driver.current_url)
    if from_url:
        return from_url

    try:
        parsed = parse_embedded_json_fields(driver.page_source or "")
        from_json = clean_report_number(parsed.get("GIA Report Number", ""))
        if from_json:
            return from_json
    except Exception:
        pass

    patterns = [
        r"GIA\s+Report\s+Number\s*[:#-]?\s*([A-Za-z0-9-]{6,24})",
        r"Report\s+(?:Number|No\.?)[\s:#-]*([A-Za-z0-9-]{6,24})",
    ]
    for pattern in patterns:
        match = re.search(pattern, page_text or "", flags=re.I)
        if match:
            value = clean_report_number(match.group(1))
            if value:
                return value

    return clean_report_number(last_entered)


def wait_for_user_to_submit_report(
    driver: uc.Chrome,
    timeout: int,
    cookie_mode: str,
) -> Tuple[str, str]:
    """Open GIA and wait until the user enters a report number and clicks Look Up/Send."""
    print("\n" + "=" * 72)
    print("GIA manual report entry")
    print("1. Chrome is now opening the GIA Report Check page.")
    print("2. Paste the GIA report number into the report-number text box.")
    print("3. Click LOOK UP / SEARCH (the send button).")
    print("4. The script will detect the result automatically and save it to Excel.")
    print("=" * 72 + "\n")

    driver.get(GIA_REPORT_CHECK_LANDING)
    wait_for_page_body(driver, timeout=timeout)

    # Automatic cookie modes can be handled immediately. In manual mode the user
    # can simply click the cookie button in Chrome; no terminal ENTER is required.
    if cookie_mode == "auto-accept":
        click_cookie_action(driver, ["Accept All"])
    elif cookie_mode == "auto-decline":
        click_cookie_action(driver, ["Decline All", "Submit Preferences"])

    started = time.time()
    last_entered = ""
    last_result_signature = ""
    stable_result_checks = 0
    cookie_notice_shown = False
    captcha_notice_shown = False

    while True:
        if timeout > 0 and (time.time() - started) > timeout:
            raise TimeoutException(
                "Timed out waiting for a report submission. Run again and submit the number in Chrome."
            )

        close_settings_submitted_dialog(driver)

        if cookie_mode == "manual" and cookie_consent_modal_visible(driver):
            if not cookie_notice_shown:
                print("Please accept/close the cookie popup in Chrome, then enter the report number.")
                cookie_notice_shown = True
        else:
            cookie_notice_shown = False

        # Keep the value while the user is typing. Some GIA pages clear the input
        # immediately after submission, so this stored value is used as fallback.
        input_element = find_report_input(driver)
        if input_element is not None:
            try:
                current_value = clean_report_number(input_element.get_attribute("value") or "")
                if current_value:
                    last_entered = current_value
            except Exception:
                pass

        page_text = get_body_text(driver)
        text_norm = canonical_label(page_text)
        report_no = report_number_from_page(driver, page_text, last_entered=last_entered)

        invalid_tokens = [
            "does not exist",
            "not available through report check",
            "invalid report",
            "no report found",
            "unable to process",
        ]
        result_tokens = [
            "date of issue",
            "grading results",
            "shape and cutting style",
            "carat weight",
            "color grade",
            "clarity grade",
        ]

        matched_result_tokens = sum(token in text_norm for token in result_tokens)
        invalid_result_visible = any(token in text_norm for token in invalid_tokens)
        report_visible_in_page = canonical_label(report_no) in text_norm if report_no else False
        report_visible_in_url = bool(report_number_from_url(driver.current_url))

        submitted_result_visible = bool(report_no) and (
            invalid_result_visible
            or (matched_result_tokens >= 2 and (report_visible_in_page or report_visible_in_url))
        )

        if page_indicates_block_or_captcha(page_text):
            if not captcha_notice_shown:
                print("GIA is showing a CAPTCHA/access check. Complete it in Chrome; extraction will continue automatically.")
                captcha_notice_shown = True
            time.sleep(FAST_POLL_INTERVAL)
            continue
        captcha_notice_shown = False

        if submitted_result_visible:
            # Require two stable checks so extraction does not begin halfway through
            # the page's JavaScript rendering.
            signature = (
                f"{report_no}|{driver.current_url}|{matched_result_tokens}|"
                f"{int(invalid_result_visible)}"
            )
            if signature == last_result_signature:
                stable_result_checks += 1
            else:
                last_result_signature = signature
                stable_result_checks = 1

            if stable_result_checks >= 2:
                print(f"Detected submitted GIA report: {report_no}")
                return report_no, page_text
        else:
            stable_result_checks = 0
            last_result_signature = ""

        time.sleep(FAST_POLL_INTERVAL)


def extract_current_report_page(
    driver: uc.Chrome,
    report_no: str,
    raw_text_limit: int,
) -> Dict[str, str]:
    """Extract the report currently visible in Chrome without navigating away."""
    report_url = driver.current_url or GIA_REPORT_CHECK_URL_TEMPLATE.format(report_no=quote_plus(report_no))
    empty = {column: "" for column in OUTPUT_COLUMNS + DEBUG_COLUMNS}
    empty["Input Report Number"] = report_no
    empty["Report URL"] = report_url

    try:
        page_source = driver.page_source or ""
        page_text = get_body_text(driver)
        raw_text = normalize_spaces(page_text)[:raw_text_limit]

        if page_indicates_block_or_captcha(raw_text):
            empty["Status"] = "Failed"
            empty["Error"] = (
                "GIA showed a CAPTCHA/access-check page. Complete it manually in Chrome, "
                "then submit the report again."
            )
            empty["Raw Page Text"] = raw_text
            return empty

        invalid_signals = [
            "does not exist",
            "not available through report check",
            "invalid report",
            "no report found",
            "unable to process",
        ]
        if any(signal in raw_text.lower() for signal in invalid_signals):
            empty["Status"] = "Failed"
            empty["Error"] = "GIA Report Check did not return a valid report for this number."
            empty["Raw Page Text"] = raw_text
            return empty

        scraped: Dict[str, str] = {}
        scraped.update(parse_embedded_json_fields(page_source))
        for key, value in parse_text_fields(raw_text).items():
            scraped.setdefault(key, value)

        pdf_url = extract_pdf_url(driver, raw_text)
        result = normalize_scraped_result(report_no, scraped, report_url, pdf_url, raw_text)

        found_any = any(result.get(column) for column in OUTPUT_COLUMNS[1:])
        result["Status"] = "Success" if found_any else "Failed"
        if not found_any:
            result["Error"] = (
                "The report page loaded, but its fields could not be parsed. "
                "Run with --include-debug-columns to inspect Raw Page Text."
            )
        return result

    except WebDriverException as exc:
        empty["Status"] = "Failed"
        empty["Error"] = f"Browser/WebDriver error: {exc}"
        return empty
    except Exception as exc:
        empty["Status"] = "Failed"
        empty["Error"] = str(exc)
        return empty


def extract_pdf_url(driver: uc.Chrome, page_text: str) -> str:
    try:
        for anchor in driver.find_elements(By.XPATH, "//a[@href]"):
            href = (anchor.get_attribute("href") or "").strip()
            if not href:
                continue
            text = (anchor.text or "").strip().lower()
            href_lower = href.lower()
            if href_lower.endswith(".pdf") or "download-report" in href_lower or "pdf" in href_lower or "download report pdf" in text:
                return href
    except Exception:
        pass
    match = re.search(r"https?://[^\s'\"]+\.pdf(?:\?[^\s'\"]*)?", page_text or "", flags=re.I)
    return match.group(0).strip() if match else ""


def build_chrome_options(headless: bool, user_data_dir: str = "") -> uc.ChromeOptions:
    options = uc.ChromeOptions()

    if headless:
        options.add_argument("--headless=new")

    # Faster navigation: Selenium continues once the DOM is interactive instead
    # of waiting for every image/font/analytics request to finish.
    options.page_load_strategy = "eager"

    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-features=IsolateOrigins,site-per-process")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-component-update")
    options.add_argument("--disable-default-apps")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-extensions")
    options.add_argument("--lang=en-US,en")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")

    # Report extraction is text/JSON based, so images are unnecessary.
    options.add_experimental_option(
        "prefs",
        {
            "profile.managed_default_content_settings.images": 2,
            "profile.default_content_setting_values.notifications": 2,
        },
    )

    if user_data_dir:
        options.add_argument(f"--user-data-dir={user_data_dir}")

    return options


def create_driver(headless: bool, user_data_dir: str = "") -> uc.Chrome:
    try:
        return uc.Chrome(
            options=build_chrome_options(headless=headless, user_data_dir=user_data_dir)
        )

    except SessionNotCreatedException as exc:
        message = str(exc)

        match = re.search(r"Current browser version is (\d+)\.", message)
        if match:
            chrome_major = int(match.group(1))
            print(f"ChromeDriver mismatch detected. Retrying with Chrome major version {chrome_major}...")

            return uc.Chrome(
                options=build_chrome_options(headless=headless, user_data_dir=user_data_dir),
                version_main=chrome_major,
            )

        raise


def normalize_scraped_result(report_no: str, scraped: Dict[str, str], report_url: str, pdf_url: str, raw_text: str) -> Dict[str, str]:
    result = {column: "" for column in OUTPUT_COLUMNS + DEBUG_COLUMNS}
    result["Input Report Number"] = report_no
    result["Report URL"] = report_url
    result["PDF URL"] = pdf_url
    result["Raw Page Text"] = raw_text

    for key, value in scraped.items():
        if key in result and clean_value(value):
            result[key] = clean_value(value)

    # Backward compatibility: older output used Description for report_type.
    if result.get("Report Type") and not result.get("Description"):
        result["Description"] = result["Report Type"]
    if result.get("Description") and not result.get("Report Type"):
        result["Report Type"] = result["Description"]

    if not result.get("GIA Report Number"):
        if any(result.get(col) for col in OUTPUT_COLUMNS[2:]):
            result["GIA Report Number"] = report_no

    return result


def fetch_report(driver: uc.Chrome, report_no: str, timeout: int, raw_text_limit: int, cookie_mode: str) -> Dict[str, str]:
    report_url = GIA_REPORT_CHECK_URL_TEMPLATE.format(report_no=quote_plus(report_no))
    empty = {column: "" for column in OUTPUT_COLUMNS + DEBUG_COLUMNS}
    empty["Input Report Number"] = report_no
    empty["Report URL"] = report_url

    try:
        # Try direct URL first. If GIA does not auto-load from querystring, fall back to the form.
        driver.get(report_url)
        wait_for_page_body(driver, timeout=timeout)
        text = wait_for_report_loaded(driver, report_no, timeout=timeout, cookie_mode=cookie_mode)

        text_norm = canonical_label(text)
        if report_no not in text or not any(token in text_norm for token in ["gia report number", "date of issue", "grading results"]):
            driver.get(GIA_REPORT_CHECK_LANDING)
            wait_for_page_body(driver, timeout=timeout)
            handle_cookie_popups(driver, cookie_mode=cookie_mode, context=f"landing page for report {report_no}")
            if submit_report_number(driver, report_no):
                text = wait_for_report_loaded(driver, report_no, timeout=timeout, cookie_mode=cookie_mode)
            else:
                driver.get("https://www.gia.edu/report-check")
                wait_for_page_body(driver, timeout=timeout)
                handle_cookie_popups(driver, cookie_mode=cookie_mode, context=f"lookup page for report {report_no}")
                if submit_report_number(driver, report_no):
                    text = wait_for_report_loaded(driver, report_no, timeout=timeout, cookie_mode=cookie_mode)

        page_source = driver.page_source or ""
        text = get_body_text(driver) or text
        raw_text = normalize_spaces(text)[:raw_text_limit]

        if page_indicates_block_or_captcha(raw_text):
            empty["Status"] = "Failed"
            empty["Error"] = "GIA showed a CAPTCHA/access-check page. Complete it manually in Chrome, then run again without --headless."
            empty["Raw Page Text"] = raw_text
            return empty

        invalid_signals = [
            "does not exist",
            "not available through report check",
            "invalid report",
            "no report found",
            "unable to process",
        ]
        if any(signal in raw_text.lower() for signal in invalid_signals):
            empty["Status"] = "Failed"
            empty["Error"] = "GIA Report Check did not return a valid report for this number."
            empty["Raw Page Text"] = raw_text
            return empty

        scraped: Dict[str, str] = {}
        scraped.update(parse_embedded_json_fields(page_source))
        for key, value in parse_text_fields(raw_text).items():
            scraped.setdefault(key, value)

        pdf_url = extract_pdf_url(driver, raw_text)
        result = normalize_scraped_result(report_no, scraped, report_url, pdf_url, raw_text)

        found_any = any(result.get(column) for column in OUTPUT_COLUMNS[1:])
        result["Status"] = "Success" if found_any else "Failed"
        if not found_any:
            result["Error"] = "Page loaded, but the script could not parse report fields. Run again with --include-debug-columns and inspect Raw Page Text."
        return result

    except TimeoutException as exc:
        empty["Status"] = "Failed"
        empty["Error"] = f"Timeout waiting for GIA page: {exc}"
        return empty
    except WebDriverException as exc:
        empty["Status"] = "Failed"
        empty["Error"] = f"Browser/WebDriver error: {exc}"
        return empty
    except Exception as exc:
        empty["Status"] = "Failed"
        empty["Error"] = str(exc)
        return empty


def safe_quit_driver(driver: Optional[uc.Chrome]) -> None:
    if driver is None:
        return
    try:
        driver.quit()
    except OSError:
        pass
    except Exception:
        pass


def parse_pasted_report_numbers(value: str) -> List[str]:
    """Parse report numbers pasted on separate lines or separated by commas/spaces."""
    report_numbers: List[str] = []
    seen = set()

    for part in re.split(r"[\s,;|]+", value or ""):
        report_no = clean_report_number(part)
        if not report_no or report_no in seen:
            continue
        seen.add(report_no)
        report_numbers.append(report_no)

    return report_numbers


def prompt_report_numbers_gui() -> List[str]:
    """Open a small Windows-friendly box where the user can paste report numbers."""
    if tk is None:
        return []

    result: List[str] = []
    try:
        root = tk.Tk()
    except Exception:
        return prompt_report_numbers_console()

    root.title("Paste GIA Report Numbers")
    root.geometry("620x430")
    root.minsize(520, 340)

    frame = tk.Frame(root, padx=18, pady=18)
    frame.pack(fill="both", expand=True)

    title = tk.Label(
        frame,
        text="Paste GIA report numbers",
        font=("Segoe UI", 14, "bold"),
        anchor="w",
    )
    title.pack(fill="x")

    instructions = tk.Label(
        frame,
        text=(
            "Enter one report number per line. You can also paste numbers separated "
            "by commas or spaces. Duplicate numbers will be removed."
        ),
        font=("Segoe UI", 10),
        justify="left",
        anchor="w",
        wraplength=560,
        pady=8,
    )
    instructions.pack(fill="x")

    text_box = tk.Text(frame, font=("Consolas", 11), wrap="word", undo=True)
    text_box.pack(fill="both", expand=True, pady=(4, 14))
    text_box.focus_set()

    button_frame = tk.Frame(frame)
    button_frame.pack(fill="x")

    def start_extraction() -> None:
        numbers = parse_pasted_report_numbers(text_box.get("1.0", "end"))
        if not numbers:
            if messagebox is not None:
                messagebox.showwarning("No report numbers", "Please paste at least one GIA report number.")
            return
        result.extend(numbers)
        root.destroy()

    def cancel() -> None:
        root.destroy()

    start_button = tk.Button(
        button_frame,
        text="Start Extraction",
        command=start_extraction,
        font=("Segoe UI", 10, "bold"),
        padx=18,
        pady=7,
    )
    start_button.pack(side="right")

    cancel_button = tk.Button(
        button_frame,
        text="Cancel",
        command=cancel,
        font=("Segoe UI", 10),
        padx=18,
        pady=7,
    )
    cancel_button.pack(side="right", padx=(0, 10))

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.bind("<Control-Return>", lambda _event: start_extraction())
    root.mainloop()
    return result


def prompt_report_numbers_console() -> List[str]:
    """Console fallback when Tkinter is unavailable. Finish with a blank line."""
    print("Paste GIA report numbers below. Use one per line, or commas/spaces.")
    print("Press ENTER on an empty line when finished.")
    lines: List[str] = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        if not line.strip():
            break
        lines.append(line)
    return parse_pasted_report_numbers("\n".join(lines))


def get_pasted_report_numbers() -> List[str]:
    numbers = prompt_report_numbers_gui()
    if numbers or tk is not None:
        return numbers
    return prompt_report_numbers_console()


def choose_input_file() -> str:
    if askopenfilename is None:
        return ""
    return askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])


def detect_input_column(df: pd.DataFrame, requested_column: str = "") -> str:
    if requested_column:
        if requested_column not in df.columns:
            raise ValueError(f"Column '{requested_column}' not found. Available columns: {list(df.columns)}")
        return requested_column
    for column in DEFAULT_INPUT_COLUMNS:
        if column in df.columns:
            return column
    raise ValueError(
        "Could not detect the report-number column. "
        f"Use --column. Available columns: {list(df.columns)}"
    )


def read_input_rows(input_path: str, requested_column: str = "") -> Tuple[pd.DataFrame, str]:
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Excel file not found: {input_path}")
    df = pd.read_excel(input_path, dtype=str)
    df.fillna("", inplace=True)
    column = detect_input_column(df, requested_column)
    df["__report_no__"] = df[column].apply(clean_report_number)
    df = df[df["__report_no__"] != ""].copy()
    if df.empty:
        raise ValueError(f"No report numbers found in column '{column}'.")
    return df, column


def order_columns(df: pd.DataFrame, include_debug_columns: bool) -> pd.DataFrame:
    wanted = OUTPUT_COLUMNS + (DEBUG_COLUMNS if include_debug_columns else [])
    for col in wanted:
        if col not in df.columns:
            df[col] = ""
    return df[wanted]


def save_excel(df: pd.DataFrame, output_path: str, include_debug_columns: bool) -> None:
    final_df = df.copy()
    final_df.fillna("", inplace=True)
    final_df = order_columns(final_df, include_debug_columns=include_debug_columns)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="GIA Results")
        ws = writer.book["GIA Results"]
        ws.freeze_panes = "A2"
        for column_cells in ws.columns:
            header = str(column_cells[0].value or "")
            max_len = max([len(header)] + [len(str(cell.value or "")) for cell in column_cells[1:]])
            ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 60)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Open a custom first tab for multiple GIA report numbers, process them in a second tab, and save all results to Excel.")
    parser.add_argument(
        "--input",
        "-i",
        default="",
        help="Optional Excel file for automatic batch mode. If omitted, use the custom report-entry page in the first Chrome tab.",
    )
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"Output Excel path. Default: {DEFAULT_OUTPUT}")
    parser.add_argument("--column", "-c", default="", help="Report-number column name. Auto-detected when omitted.")
    parser.add_argument("--report-number", "-r", default="", help="Fetch just one report number instead of reading Excel.")
    parser.add_argument("--headless", action="store_true", help="Run Chrome headless. Do not use this with manual cookie mode.")
    parser.add_argument(
        "--cookie-mode",
        choices=["manual", "auto-accept", "auto-decline", "skip"],
        default="manual",
        help="Cookie handling. Default manual: you click ACCEPT ALL, then extraction starts.",
    )
    parser.add_argument("--include-debug-columns", action="store_true", help="Include URL/status/error/raw text columns in output.")
    parser.add_argument("--delay", type=float, default=FAST_INTER_REPORT_DELAY, help="Seconds to wait between lookups. Fast default: 0.10.")
    parser.add_argument("--timeout", type=int, default=FAST_PAGE_TIMEOUT, help="Seconds to wait for each page/report. Fast default: 20.")
    parser.add_argument("--raw-text-limit", type=int, default=15000, help="Max Raw Page Text characters saved when debug columns are enabled.")
    parser.add_argument("--user-data-dir", default="", help="Optional Chrome profile folder to reuse cookies/session.")
    parser.add_argument("--limit", type=int, default=0, help="Optional max number of reports to process in batch mode.")
    parser.add_argument(
        "--save-every",
        type=int,
        default=FAST_SAVE_EVERY,
        help="Write Excel every N processed reports. Fast default: 10. Use 1 to save after every report.",
    )
    parser.add_argument(
        "--manual-entry-timeout",
        type=int,
        default=0,
        help="Seconds to wait for the first-tab Submit button. Default 0 waits until you submit the reports.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.headless and args.cookie_mode == "manual":
        print("Manual browser entry needs a visible Chrome window. Remove --headless.")
        return

    manual_browser_mode = not args.report_number and not args.input
    results: List[Dict[str, str]] = []
    driver = create_driver(headless=args.headless, user_data_dir=args.user_data_dir)

    try:
        if manual_browser_mode:
            report_numbers, first_tab = wait_for_report_numbers_from_first_tab(
                driver=driver,
                timeout=args.manual_entry_timeout,
            )

            if args.limit and args.limit > 0:
                report_numbers = report_numbers[:args.limit]

            total = len(report_numbers)
            print(f"Submitted {total} report number(s).")
            open_extraction_tab(driver, first_tab)

            extraction_cookie_mode = (
                "auto-accept" if args.cookie_mode == "manual" else args.cookie_mode
            )

            for position, report_no in enumerate(report_numbers, start=1):
                print(f"[{position}/{total}] Opening GIA and extracting report {report_no}...")
                result = fetch_report(
                    driver=driver,
                    report_no=report_no,
                    timeout=args.timeout,
                    raw_text_limit=args.raw_text_limit,
                    cookie_mode=extraction_cookie_mode,
                )
                results.append(result)

                if (
                    args.save_every <= 1
                    or position % args.save_every == 0
                    or position == total
                ):
                    save_excel(
                        pd.DataFrame(results),
                        args.output,
                        include_debug_columns=args.include_debug_columns,
                    )

                status = result.get("Status", "")
                error = result.get("Error", "")
                update_first_tab_progress(
                    driver=driver,
                    first_tab=first_tab,
                    report_no=report_no,
                    position=position,
                    total=total,
                    status=status,
                    error=error,
                )
                print(f"    {status}: {report_no}" + (f" - {error}" if error else ""))

                if position < total and args.delay > 0:
                    time.sleep(args.delay)

        else:
            if args.report_number:
                report_numbers = parse_pasted_report_numbers(args.report_number)
                if not report_numbers:
                    print("No valid report number was supplied.")
                    return
                input_rows = pd.DataFrame({"__report_no__": report_numbers})
                print(f"Loaded {len(input_rows)} report number(s) from --report-number.")
            else:
                input_rows, source_column = read_input_rows(args.input, args.column)
                print(
                    f"Loaded {len(input_rows)} report number(s) from {args.input}, "
                    f"column '{source_column}'."
                )

            if args.limit and args.limit > 0:
                input_rows = input_rows.head(args.limit).copy()

            prepare_gia_cookie_session(
                driver,
                cookie_mode=args.cookie_mode,
                timeout=args.timeout,
            )

            total = len(input_rows)
            print(f"Starting extraction for {total} report(s)...")
            for position, (_, row) in enumerate(input_rows.iterrows(), start=1):
                report_no = clean_report_number(row["__report_no__"])
                if not report_no:
                    continue

                print(f"[{position}/{total}] Fetching GIA report {report_no}...")
                result = fetch_report(
                    driver=driver,
                    report_no=report_no,
                    timeout=args.timeout,
                    raw_text_limit=args.raw_text_limit,
                    cookie_mode=args.cookie_mode,
                )
                results.append(result)

                status = result.get("Status", "")
                error = result.get("Error", "")
                print(f"    {status}: {report_no}" + (f" - {error}" if error else ""))

                if (
                    args.save_every <= 1
                    or position % args.save_every == 0
                    or position == total
                ):
                    save_excel(
                        pd.DataFrame(results),
                        args.output,
                        include_debug_columns=args.include_debug_columns,
                    )
                if position < total and args.delay > 0:
                    time.sleep(args.delay)

    finally:
        safe_quit_driver(driver)
        driver = None

    if not results:
        print("No report result was captured.")
        return

    save_excel(
        pd.DataFrame(results),
        args.output,
        include_debug_columns=args.include_debug_columns,
    )
    print(f"Results saved to: {args.output}")

    if sys.platform.startswith("win"):
        try:
            os.startfile(args.output)  # type: ignore[attr-defined]
        except Exception:
            pass

if __name__ == "__main__":
    main()