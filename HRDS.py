from __future__ import annotations

import argparse
import os
import re
import time
from typing import Dict, List
from dataclasses import dataclass
import pandas as pd
from tifffile import askopenfilename
import undetected_chromedriver as uc
from selenium.common.exceptions import SessionNotCreatedException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# Default file names used if no command-line arguments are supplied.
DEFAULT_INPUT = askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
DEFAULT_OUTPUT = fr"D:\Gaurav\Download\Certificate_{pd.Timestamp.now():%S}.xlsx"

# Column header in the input spreadsheet that contains report numbers.
REPORT_COLUMN_NAME = "Report Number"

# HRD Antwerp report-check page.
HRD_URL = "https://my.hrdantwerp.com/"


# undetected_chromedriver can print harmless Windows cleanup messages after the
# Excel file is already saved. This wrapper suppresses only that cleanup noise.
try:
    _ORIGINAL_UC_CHROME_DEL = uc.Chrome.__del__

    def _quiet_uc_chrome_del(self) -> None:
        try:
            _ORIGINAL_UC_CHROME_DEL(self)
        except OSError as exc:
            if getattr(exc, "winerror", None) != 6:
                raise
        except Exception:
            pass

    uc.Chrome.__del__ = _quiet_uc_chrome_del
except Exception:
    pass


OUTPUT_COLUMNS = [
    "Input Report Number",
    "Report Number",
    "Lab",
    "Date of Issue",
    "Shape",
    "Measurements",
    "Carat Weight",
    "Color Grade",
    "Clarity Grade",
    "Cut Grade",
    "Depth",
    "Table",
    "Crown Angle",
    "Crown Height",
    "Pavilion Angle",
    "Pavilion Depth",
    "Star Length",
    "Lower Half",
    "Girdle",
    "Culet",
    "Polish",
    "Symmetry",
    "Fluorescence",
    "Clarity Characteristics",
    "INSCRIPTION(S)"
]


LABEL_VARIANTS: Dict[str, List[str]] = {
    "Report Number": ["Report Number", "#"],
    "Date of Issue": ["Date of Issue", "Issue Date"],
    "Shape": ["Shape"],
    "Weight (carat)": ["Weight (carat)", "Weight", "Carat Weight"],
    "Fluorescence": ["Fluorescence"],
    "Colour": ["Colour", "Color"],
    "Clarity": ["Clarity"],
    "Cut": ["Cut"],
    "Proportions": ["Proportions"],
    "Polish": ["Polish"],
    "Symmetry": ["Symmetry"],
    "Measurements": ["Measurements", "Measurement"],
    "Girdle": ["Girdle"],
    "Culet": ["Culet"],
    "Total Depth": ["Total Depth", "Depth"],
    "Table Width": ["Table Width", "Table"],
    "Crown Height (β)": ["Crown Height (β)", "Crown Height", "Crown Height (B)"],
    "Pavilion Depth (α)": ["Pavilion Depth (α)", "Pavilion Depth", "Pavilion Depth (a)"],
    "Length Halves Crown": ["Length Halves Crown"],
    "Length Halves Pavilion": ["Length Halves Pavilion"],
}

SECTION_HEADINGS = {
    "CUT",
    "INSCRIPTIONS",
    "INSCRIPTION",
    "BASIC",
    "EXTRA INFO",
}

DATE_RE = re.compile(r"\b\d{4}-\d{2}-\d{2}\b")
REPORT_NO_RE = re.compile(r"\b\d{6,}\b")


def safe_quit_driver(driver) -> None:
    if driver is None:
        return
    try:
        driver.quit()
    except OSError as exc:
        if getattr(exc, "winerror", None) != 6:
            raise
    except Exception:
        pass


def clean_text(value: str) -> str:
    value = value or ""
    value = value.replace("\xa0", " ")
    value = value.replace("\u200b", "")
    value = re.sub(r"[ \t\r\f\v]+", " ", value)
    value = re.sub(r"\n{2,}", "\n", value)
    return value.strip()


def clean_value(value: str) -> str:
    value = clean_text(value)
    value = value.strip(" :#•")
    value = re.sub(r"\s+", " ", value)
    return value.strip(" :#•")



def clean_inscription_value(value: str, report_no: str = "") -> str:
    """
    Keep only the actual inscription value for Excel.

    Examples:
        "Inscription on girdle: HRD Antwerp 250000208338"
        "HRD 250000219505 | Shape brilliant | Weight (carat) ..."
        "HRD DIR 250000082800 | Shape brilliant | Weight (carat) ..."

    become:
        "HRD 250000208338"
        "HRD 250000219505"
        "HRD 250000082800"
    """
    value = clean_value(value)
    value = value.replace("•", " ")

    # If the current input report number appears anywhere in the captured text,
    # this is the safest and cleanest output.
    report_no = clean_report_number(report_no) if report_no else ""
    if report_no and report_no in value:
        return f"HRD {report_no}"

    # Remove labels/prefixes.
    value = re.sub(r"^inscriptions?\s*:?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(
        r"^inscription\s+on\s+girdle\s*:\s*",
        "",
        value,
        flags=re.IGNORECASE,
    )

    # When hidden/print DOM text is collapsed, it can become:
    #   HRD 250... | Shape ... | Weight ...
    # Keep only the first inscription chunk.
    value = re.split(
        r"\s*\|\s*(?:Shape|Weight|Fluorescence|Colour|Color|Clarity|Cut|"
        r"Proportions|Polish|Symmetry|Measurements|Girdle|Culet|Total Depth|"
        r"Table Width|Crown Height|Pavilion Depth|Length Halves|Sum)",
        value,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0]

    # Normalize HRD wording and remove extra words such as Antwerp / DIR.
    value = re.sub(r"\bHRD\s+Antwerp\b", "HRD", value, flags=re.IGNORECASE)

    # Prefer exactly HRD + the first report-like number.
    match = re.search(r"\bHRD(?:\s+Antwerp|\s+DIR)?\s+(\d{6,})\b", value, flags=re.IGNORECASE)
    if match:
        return f"HRD {match.group(1)}"

    match = re.search(r"\b(\d{6,})\b", value)
    if match:
        return f"HRD {match.group(1)}"

    value = re.sub(r"\s+", " ", value).strip(" :|")
    return value


def clean_result_inscriptions(data: Dict[str, str], report_no: str) -> None:
    """
    Force inscription columns to contain exactly: HRD <Report Number>

    Example:
        Inscriptions: HRD 250000208338

    This ignores any captured text such as:
        Inscription on girdle: HRD Antwerp 250000208338
        HRD DIR 250000082800 | Shape brilliant | Weight ...
    """
    # Prefer the input report number because it is the report currently being processed.
    clean_no = clean_report_number(report_no) if report_no else ""

    # Fallback to the extracted report number if needed.
    if not clean_no:
        clean_no = clean_report_number(data.get("Report Number", ""))

    # Last fallback: search all captured inscription text for a report-like number.
    if not clean_no:
        joined = " ".join([
            str(data.get("Inscriptions", "")),
            str(data.get("Inscription on girdle", "")),
        ])
        match = re.search(r"\b(\d{6,})\b", joined)
        if match:
            clean_no = match.group(1)

    if clean_no:
        final_value = f"HRD {clean_no}"
        data["Inscriptions"] = final_value
        data["Inscription on girdle"] = final_value


def clean_report_number(value) -> str:
    if pd.isna(value):
        return ""
    report_no = str(value).strip()
    if report_no.endswith(".0"):
        report_no = report_no[:-2]
    return re.sub(r"\s+", "", report_no)


def read_report_numbers(path: str, column: str) -> List[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")

    df = pd.read_excel(path, dtype=str)

    if column not in df.columns:
        raise ValueError(
            f"Column '{column}' not found in {path}. Available columns: {list(df.columns)}"
        )

    numbers = df[column].apply(clean_report_number).tolist()
    numbers = [n for n in numbers if n]

    if not numbers:
        raise ValueError(f"No report numbers found in column '{column}'.")

    # Deduplicate while preserving order.
    return list(dict.fromkeys(numbers))


def create_driver(headless: bool = False) -> uc.Chrome:
    options = uc.ChromeOptions()

    if headless:
        options.add_argument("--headless=new")

    options.add_argument("--start-maximized")
    options.add_argument("--disable-features=IsolateOrigins,site-per-process")
    options.add_argument("--disable-blink-features=AutomationControlled")

    # These preferences do not save anything automatically. They just make
    # Chrome less likely to interrupt Selenium with prompts.
    prefs = {
        "printing.print_preview_sticky_settings.appState": "{}",
        "download.prompt_for_download": False,
    }
    options.add_experimental_option("prefs", prefs)

    try:
        return uc.Chrome(options=options)
    except SessionNotCreatedException as exc:
        message = str(exc)
        match = re.search(r"Current browser version is (\d+)\.", message)
        if match:
            version_main = int(match.group(1))
            retry_options = uc.ChromeOptions()

            if headless:
                retry_options.add_argument("--headless=new")

            retry_options.add_argument("--start-maximized")
            retry_options.add_argument("--disable-features=IsolateOrigins,site-per-process")
            retry_options.add_argument("--disable-blink-features=AutomationControlled")
            retry_options.add_experimental_option("prefs", prefs)

            return uc.Chrome(options=retry_options, version_main=version_main)

        raise


def wait_for_body(driver, timeout: int = 30) -> bool:
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        return True
    except TimeoutException:
        return False


def body_text(driver) -> str:
    try:
        return driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        return ""


def all_dom_text(driver) -> str:
    """
    Return all text in the DOM, including text that may be hidden until print.
    This helps when HRD keeps the printable report in hidden markup.
    """
    try:
        return driver.execute_script("return document.body ? document.body.textContent : '';") or ""
    except Exception:
        return body_text(driver)


def dismiss_cookie_prompt(driver) -> None:
    try:
        buttons = driver.find_elements(
            By.XPATH,
            "//button["
            "contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept') "
            "or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'agree') "
            "or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'ok') "
            "or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'allow')"
            "]",
        )
        for button in buttons:
            if button.is_displayed() and button.is_enabled():
                button.click()
                time.sleep(1)
                return
    except Exception:
        pass


def visible_text_inputs(driver):
    return driver.find_elements(
        By.XPATH,
        "//input[not(@type='hidden') and not(@type='password') and "
        "(@type='text' or @type='search' or @type='number' or @type='tel' or not(@type))]",
    )


def find_diamond_input(driver):
    candidates = []

    for inp in visible_text_inputs(driver):
        try:
            if not inp.is_displayed() or not inp.is_enabled():
                continue

            attrs = " ".join(
                [
                    inp.get_attribute("id") or "",
                    inp.get_attribute("name") or "",
                    inp.get_attribute("placeholder") or "",
                    inp.get_attribute("aria-label") or "",
                    inp.get_attribute("class") or "",
                ]
            ).lower()

            ancestor_text = ""
            try:
                ancestor = inp.find_element(
                    By.XPATH,
                    "./ancestor::*[self::div or self::section or self::form][1]",
                )
                ancestor_text = ancestor.text.lower()
            except Exception:
                pass

            score = 0
            for token in ["diamond", "report", "find"]:
                if token in attrs:
                    score += 4
                if token in ancestor_text:
                    score += 4

            if "find diamond" in ancestor_text:
                score += 25

            # Avoid login/register/jewel inputs.
            for bad in ["client", "password", "email", "username", "name", "jewel"]:
                if bad in attrs:
                    score -= 10
                if bad in ancestor_text and bad != "jewel":
                    score -= 2

            candidates.append((score, inp))

        except Exception:
            continue

    if not candidates:
        return None

    candidates.sort(key=lambda item: item[0], reverse=True)
    return candidates[0][1] if candidates[0][0] > 0 else None


def click_nearby_search_button_or_enter(driver, report_input) -> None:
    search_xpaths = [
        ".//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'search')]",
        ".//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'search')]",
        ".//input[@type='submit']",
        ".//*[self::button or self::a][.//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'search')]]",
    ]

    ancestors = []
    for depth in range(1, 8):
        try:
            ancestor = report_input.find_element(By.XPATH, f"./ancestor::*[{depth}]")
            ancestors.append(ancestor)
        except Exception:
            pass

    for ancestor in ancestors:
        for xpath in search_xpaths:
            try:
                buttons = ancestor.find_elements(By.XPATH, xpath)
                for button in buttons:
                    if button.is_displayed() and button.is_enabled():
                        driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center'});", button
                        )
                        time.sleep(0.3)
                        button.click()
                        time.sleep(2)
                        return
            except Exception:
                continue

    try:
        buttons = driver.find_elements(
            By.XPATH,
            "//*[self::button or self::a or self::input]"
            "[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'search') "
            "or @type='submit']",
        )
        for button in buttons:
            if button.is_displayed() and button.is_enabled():
                button.click()
                time.sleep(2)
                return
    except Exception:
        pass

    report_input.send_keys(Keys.ENTER)
    time.sleep(2)


def submit_report_number(driver, report_no: str) -> bool:
    driver.get(HRD_URL)

    if not wait_for_body(driver):
        return False

    time.sleep(2)
    dismiss_cookie_prompt(driver)

    report_input = find_diamond_input(driver)
    if report_input is None:
        return False

    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", report_input)
        time.sleep(0.3)
        report_input.clear()
        report_input.send_keys(report_no)
        click_nearby_search_button_or_enter(driver, report_input)
        return True
    except Exception:
        return False


def page_has_report(driver, report_no: str) -> bool:
    text = clean_text(body_text(driver)).lower()

    if not text:
        return False

    return (
        report_no.lower() in text
        and (
            "natural diamond identification report" in text
            or "date of issue" in text
            or "measurements" in text
            or "proportions" in text
            or "print these grading report results" in text
        )
    )


def wait_for_report(driver, report_no: str, timeout: int = 40) -> bool:
    try:
        WebDriverWait(driver, timeout).until(lambda d: page_has_report(d, report_no))
        return True
    except TimeoutException:
        return False


def canonical_label(label: str) -> str:
    label_clean = clean_value(label).lower()

    for canonical, variants in LABEL_VARIANTS.items():
        for variant in variants:
            if label_clean == clean_value(variant).lower():
                return canonical

    return clean_value(label)


def is_known_label(text: str) -> bool:
    text_clean = clean_value(text).lower()
    for variants in LABEL_VARIANTS.values():
        for variant in variants:
            if text_clean == clean_value(variant).lower():
                return True
    return False


def is_section_heading(text: str) -> bool:
    return clean_value(text).upper() in SECTION_HEADINGS


def add_field(data: Dict[str, str], label: str, value: str) -> None:
    label = canonical_label(label)
    value = clean_value(value)

    if not label or not value:
        return

    if label in data and data[label]:
        return

    data[label] = value


def extract_rows_from_tables(driver) -> Dict[str, str]:
    data: Dict[str, str] = {}

    try:
        rows = driver.find_elements(By.XPATH, "//table//tr")
    except Exception:
        return data

    for row in rows:
        try:
            if not row.is_displayed():
                continue

            cells = row.find_elements(By.XPATH, "./th|./td")
            values = [clean_value(cell.text) for cell in cells if clean_value(cell.text)]

            if len(values) >= 2:
                add_field(data, values[0], values[1])
        except Exception:
            continue

    return data


def extract_rows_from_divs(driver) -> Dict[str, str]:
    """
    HRD pages often use div/span rows instead of HTML tables.
    This reads two-column visible rows from the modal/printable content.
    """
    data: Dict[str, str] = {}

    row_xpaths = [
        "//*[contains(@class, 'row')]",
        "//*[contains(@class, 'table')]//*[contains(@class, 'row')]",
        "//*[contains(@class, 'report')]//*[self::div or self::li]",
    ]

    for row_xpath in row_xpaths:
        try:
            rows = driver.find_elements(By.XPATH, row_xpath)
        except Exception:
            continue

        for row in rows:
            try:
                if not row.is_displayed():
                    continue

                parts = row.find_elements(By.XPATH, "./*[normalize-space()]")
                texts = [clean_value(p.text) for p in parts if clean_value(p.text)]

                # Avoid parent containers with too much text.
                texts = [t for t in texts if len(t) <= 150]

                if len(texts) == 2:
                    add_field(data, texts[0], texts[1])
            except Exception:
                continue

    return data


def known_labels_sorted() -> List[str]:
    labels = []
    for variants in LABEL_VARIANTS.values():
        labels.extend(variants)
    return sorted(set(labels), key=len, reverse=True)


def line_is_noise(line: str) -> bool:
    lower = clean_value(line).lower()
    noise = [
        "download (pdf)",
        "download pdf",
        "print these grading report results",
        "natural diamond identification report",
        "basic",
        "extra info",
        "1/1",
        "print",
        "cancel",
    ]
    return lower in noise


def extract_from_lines(text: str) -> Dict[str, str]:
    data: Dict[str, str] = {}

    lines = [clean_value(line) for line in clean_text(text).splitlines()]
    lines = [line for line in lines if line and not line_is_noise(line)]

    labels_lower = {
        clean_value(variant).lower(): canonical
        for canonical, variants in LABEL_VARIANTS.items()
        for variant in variants
    }

    for i, line in enumerate(lines):
        line_clean = clean_value(line)
        line_lower = line_clean.lower()

        # Top card format: "# 250000208338"
        if re.fullmatch(r"#?\s*\d{6,}", line_clean):
            add_field(data, "Report Number", line_clean.replace("#", "").strip())
            continue

        # Date card format: "2025-08-13"
        if DATE_RE.fullmatch(line_clean):
            add_field(data, "Date of Issue", line_clean)
            continue

        # Label on one line, value on next line.
        if line_lower in labels_lower and i + 1 < len(lines):
            next_value = clean_value(lines[i + 1])

            # "Cut" is a section header on HRD print pages, not a value.
            if line_lower == "cut":
                continue

            if (
                next_value
                and not is_known_label(next_value)
                and not is_section_heading(next_value)
                and not line_is_noise(next_value)
            ):
                add_field(data, labels_lower[line_lower], next_value)
                continue

        # Compact format: "Total Depth 62.7 %"
        for label in known_labels_sorted():
            pattern = rf"^{re.escape(label)}\s+(.+)$"
            match = re.match(pattern, line_clean, flags=re.IGNORECASE)
            if match:
                value = clean_value(match.group(1))
                if value and not is_known_label(value) and not is_section_heading(value):
                    add_field(data, label, value)
                break

    # Description block. Sometimes the print page has only the sentence.
    for idx, line in enumerate(lines):
        if line.lower() == "description" and idx + 1 < len(lines):
            add_field(data, "Description", lines[idx + 1])
            break

    if not data.get("Description"):
        for line in lines:
            if "identified as a natural gem diamond" in line.lower():
                add_field(data, "Description", line)
                break

    # Inscriptions block.
    inscription_lines = []
    capture = False

    for line in lines:
        lower = line.lower()

        if lower.startswith("inscriptions"):
            capture = True
            continue

        if capture:
            if is_known_label(line) or is_section_heading(line) or line_is_noise(line):
                break
            inscription_lines.append(line)

    if inscription_lines:
        inscription_text = " | ".join(inscription_lines)

        match = re.search(
            r"Inscription on girdle\s*:\s*(.+)",
            inscription_text,
            flags=re.IGNORECASE,
        )

        if match:
            inscription_value = clean_inscription_value(match.group(1))
        else:
            inscription_value = clean_inscription_value(inscription_text)

        if inscription_value:
            data["Inscriptions"] = inscription_value
            data["Inscription on girdle"] = inscription_value

    return data


def extract_pdf_url(driver) -> str:
    try:
        anchors = driver.find_elements(
            By.XPATH,
            "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pdf') "
            "or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download') "
            "or contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '.pdf')]",
        )

        for anchor in anchors:
            try:
                if not anchor.is_displayed():
                    continue
            except Exception:
                pass

            href = (anchor.get_attribute("href") or "").strip()
            if href and not href.lower().startswith("javascript"):
                return href
    except Exception:
        pass

    return ""


def click_tab(driver, tab_text: str) -> bool:
    try:
        xpath = (
            "//*[self::a or self::button or self::li or self::div or self::span]"
            f"[contains(translate(normalize-space(.), "
            f"'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{tab_text.lower()}')]"
        )

        elements = driver.find_elements(By.XPATH, xpath)

        for element in elements:
            try:
                if element.is_displayed() and element.is_enabled():
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center'});", element
                    )
                    time.sleep(0.2)
                    element.click()
                    time.sleep(1.5)
                    return True
            except Exception:
                continue

    except Exception:
        pass

    return False


def block_browser_print_dialog(driver) -> None:
    """
    Prevent Chrome's native print dialog/print preview from opening.
    This version injects the override into the main page, all iframes, and any
    new document that HRD opens. This prevents the script from freezing on the
    second/third report when Chrome print preview appears.
    """
    source = r"""
        window.__hrd_print_called = false;
        window.__hrd_print_blocked = true;
        try {
            Object.defineProperty(window, 'print', {
                value: function() { window.__hrd_print_called = true; return true; },
                writable: true,
                configurable: true
            });
        } catch (e) {
            window.print = function() { window.__hrd_print_called = true; return true; };
        }
    """

    # Inject into future pages/windows.
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": source})
    except Exception:
        pass

    # Inject into the current top document.
    try:
        driver.execute_script(source)
    except Exception:
        pass

    # Inject into same-origin iframes, if HRD uses any for print/report content.
    try:
        driver.switch_to.default_content()
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for frame in frames:
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(frame)
                driver.execute_script(source)
            except Exception:
                continue
        driver.switch_to.default_content()
    except Exception:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass


def hrd_print_called(driver) -> bool:
    try:
        return bool(driver.execute_script("return window.__hrd_print_called === true;"))
    except Exception:
        return False


def scroll_all_containers_to_bottom(driver) -> None:
    """
    The print link sits at the bottom of the HRD modal. Scrolling the page alone
    is sometimes not enough, because the modal body has its own scroll area.
    """
    try:
        driver.execute_script(
            """
            window.scrollTo(0, document.body.scrollHeight);
            const nodes = Array.from(document.querySelectorAll('*'));
            for (const el of nodes) {
                try {
                    const style = window.getComputedStyle(el);
                    const canScroll = (style.overflowY === 'auto' || style.overflowY === 'scroll')
                        && el.scrollHeight > el.clientHeight;
                    if (canScroll) { el.scrollTop = el.scrollHeight; }
                } catch (e) {}
            }
            """
        )
        time.sleep(0.25)
    except Exception:
        pass


def actionable_print_element(driver, element):
    """
    Convert a text/span/div match into the real clickable element.
    """
    try:
        tag = (element.tag_name or "").lower()
        role = (element.get_attribute("role") or "").lower()
        onclick = element.get_attribute("onclick") or ""
        ng_click = element.get_attribute("ng-click") or element.get_attribute("data-ng-click") or ""
        classes = (element.get_attribute("class") or "").lower()

        if tag in {"a", "button"} or role == "button" or onclick or ng_click or "btn" in classes or "button" in classes:
            return element

        descendants = element.find_elements(
            By.XPATH,
            ".//*[self::a or self::button or @role='button' or @onclick or @ng-click or @data-ng-click]",
        )
        for child in descendants:
            try:
                if child.is_displayed() and "print" in clean_value(child.text).lower():
                    return child
            except Exception:
                continue

        parent = element
        for _ in range(5):
            try:
                parent = parent.find_element(By.XPATH, "./parent::*")
                tag = (parent.tag_name or "").lower()
                role = (parent.get_attribute("role") or "").lower()
                onclick = parent.get_attribute("onclick") or ""
                ng_click = parent.get_attribute("ng-click") or parent.get_attribute("data-ng-click") or ""
                classes = (parent.get_attribute("class") or "").lower()
                if tag in {"a", "button"} or role == "button" or onclick or ng_click or "btn" in classes or "button" in classes:
                    return parent
            except Exception:
                break
    except Exception:
        pass

    return element


def click_element_hard(driver, element) -> bool:
    """
    Safe non-blocking click.

    Do NOT use Selenium element.click() here. On HRD, the real browser print
    preview can open and make element.click() hang forever. JavaScript dispatch
    plus window.print override is much safer for batch processing.
    """
    block_browser_print_dialog(driver)

    old_script_timeout = None
    try:
        # Keep any accidental JavaScript hang short.
        driver.set_script_timeout(6)
    except Exception:
        pass

    try:
        driver.execute_script(
            """
            const el = arguments[0];
            if (!el) return false;
            el.scrollIntoView({block: 'center', inline: 'center'});

            // Ensure no native print dialog can block the automation.
            try {
                window.__hrd_print_called = false;
                window.print = function() { window.__hrd_print_called = true; return true; };
            } catch(e) {}

            // Some Angular/Material links respond better to a mouse-event sequence
            // than to only el.click().
            const opts = {bubbles: true, cancelable: true, view: window};
            for (const type of ['mouseover', 'mousedown', 'mouseup', 'click']) {
                el.dispatchEvent(new MouseEvent(type, opts));
            }
            return true;
            """,
            element,
        )
        time.sleep(0.8)
        return True
    except Exception:
        # Final fallback: JavaScript native click only. This is still safer than
        # Selenium's blocking element.click().
        try:
            driver.execute_script(
                """
                try { window.print = function(){ window.__hrd_print_called = true; return true; }; } catch(e) {}
                arguments[0].scrollIntoView({block: 'center', inline: 'center'});
                arguments[0].click();
                return true;
                """,
                element,
            )
            time.sleep(0.8)
            return True
        except Exception:
            return False


def element_priority(element) -> tuple:
    """Prefer the actual short link/button over a parent container."""
    try:
        tag = (element.tag_name or "").lower()
        txt = clean_value(element.text)
        role = (element.get_attribute("role") or "").lower()
        onclick = element.get_attribute("onclick") or ""
        ng_click = element.get_attribute("ng-click") or element.get_attribute("data-ng-click") or ""
        classes = (element.get_attribute("class") or "").lower()

        clickable_score = 0
        if tag in {"a", "button"}:
            clickable_score += 100
        if role == "button" or onclick or ng_click:
            clickable_score += 60
        if "btn" in classes or "button" in classes:
            clickable_score += 30
        if txt.lower() == "print these grading report results":
            clickable_score += 40

        return (-clickable_score, len(txt))
    except Exception:
        return (9999, 9999)


def click_print_grading_report_results(driver) -> bool:
    """
    Click the exact 'PRINT THESE GRADING REPORT RESULTS' link/button.
    Returns True only after clicking a likely actionable element.
    """
    block_browser_print_dialog(driver)
    scroll_all_containers_to_bottom(driver)

    old_handles = driver.window_handles[:]
    old_url = driver.current_url

    phrase = "print these grading report results"
    xpaths = [
        # Exact/near-exact visible text first.
        "//*[normalize-space() and contains(translate(normalize-space(.), "
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
        "'print these grading report results')]",

        # Common HRD/angular click handlers and regular print links.
        "//*[self::a or self::button or @role='button' or @onclick or @ng-click or @data-ng-click]"
        "[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'print')]",
    ]

    candidates = []
    seen = set()
    for xpath in xpaths:
        try:
            for element in driver.find_elements(By.XPATH, xpath):
                try:
                    if not element.is_displayed():
                        continue
                    txt = clean_value(element.text).lower()
                    if "print" not in txt:
                        continue
                    if phrase not in txt and len(txt) > 80:
                        # Avoid huge modal/body containers that only contain the link somewhere inside.
                        continue
                    actionable = actionable_print_element(driver, element)
                    element_id = actionable.id
                    if element_id not in seen:
                        seen.add(element_id)
                        candidates.append(actionable)
                except Exception:
                    continue
        except Exception:
            continue

    candidates.sort(key=element_priority)

    for candidate in candidates:
        try:
            label = clean_value(candidate.text)
            print(f"    Trying print button: {label[:80]}")
        except Exception:
            pass

        if click_element_hard(driver, candidate):
            # Success cases: window.print was called, a print tab/window opened, URL changed,
            # or the click simply executed without exception on the actual print link.
            try:
                if hrd_print_called(driver):
                    print("    Print action detected and browser print dialog was blocked.")
                    return True
            except Exception:
                pass

            try:
                if len(driver.window_handles) > len(old_handles):
                    return True
            except Exception:
                pass

            try:
                if driver.current_url != old_url:
                    return True
            except Exception:
                pass

            # Since we filtered down to the actual print link/button, treat this as clicked.
            return True

    # Last-resort JavaScript: click the smallest visible element containing the exact phrase.
    try:
        clicked = driver.execute_script(
            """
            const phrase = 'print these grading report results';
            const els = Array.from(document.querySelectorAll('a,button,[role=button],[onclick],[ng-click],[data-ng-click],span,div'))
                .filter(el => {
                    const txt = (el.innerText || el.textContent || '').trim().toLowerCase();
                    const r = el.getBoundingClientRect();
                    return txt.includes(phrase) && r.width > 0 && r.height > 0;
                })
                .sort((a,b) => {
                    const ac = ['A','BUTTON'].includes(a.tagName) ? 0 : 1;
                    const bc = ['A','BUTTON'].includes(b.tagName) ? 0 : 1;
                    return ac - bc || (a.innerText || '').length - (b.innerText || '').length;
                });
            if (!els.length) return false;
            let el = els[0];
            let clickable = el.closest('a,button,[role=button],[onclick],[ng-click],[data-ng-click]') || el;
            clickable.scrollIntoView({block:'center', inline:'center'});
            clickable.click();
            return true;
            """
        )
        time.sleep(2)
        return bool(clicked)
    except Exception:
        return False


def switch_to_new_print_window_if_opened(driver, old_handles) -> None:
    """
    If HRD opens the printable report in a new tab/window, switch to it.
    """
    try:
        WebDriverWait(driver, 5).until(
            lambda d: len(d.window_handles) >= len(old_handles)
        )

        new_handles = [h for h in driver.window_handles if h not in old_handles]
        if new_handles:
            driver.switch_to.window(new_handles[-1])
            wait_for_body(driver, timeout=15)
            time.sleep(2)
    except Exception:
        pass


def merge_data(target: Dict[str, str], source: Dict[str, str]) -> None:
    for key, value in source.items():
        if key not in target or not target[key]:
            target[key] = value


def extract_visible_data(driver) -> Dict[str, str]:
    data: Dict[str, str] = {}
    merge_data(data, extract_rows_from_tables(driver))
    merge_data(data, extract_rows_from_divs(driver))
    merge_data(data, extract_from_lines(body_text(driver)))
    # Hidden print templates often appear only in textContent, not visible .text.
    merge_data(data, extract_from_lines(all_dom_text(driver)))
    return data


def extract_current_report_data(driver, report_no: str) -> Dict[str, str]:
    data: Dict[str, str] = {}

    # 1) Collect data visible immediately after search.
    merge_data(data, extract_visible_data(driver))

    # 2) Click Extra info first, because HRD loads measurements/proportions there.
    click_tab(driver, "Extra info")
    time.sleep(1)
    merge_data(data, extract_visible_data(driver))

    # 3) Save PDF link before clicking print, if available.
    pdf_url = extract_pdf_url(driver)
    if pdf_url:
        data["PDF URL"] = pdf_url

    # 4) Click PRINT THESE GRADING REPORT RESULTS.
    old_handles = driver.window_handles[:]
    clicked_print = click_print_grading_report_results(driver)

    if clicked_print:
        print("    Print results button clicked.")
        switch_to_new_print_window_if_opened(driver, old_handles)
        time.sleep(2)
        merge_data(data, extract_visible_data(driver))
    else:
        print("    WARNING: Could not click PRINT THESE GRADING REPORT RESULTS; using modal/hidden DOM data.")

    # Safety: if page contains the report number, keep the input report number.
    current_text = clean_text(body_text(driver) + "\n" + all_dom_text(driver))
    if report_no in current_text:
        data["Report Number"] = report_no

    # Safety: try to recover report number/date directly from all DOM text.
    if not data.get("Report Number"):
        match = REPORT_NO_RE.search(current_text)
        if match:
            data["Report Number"] = match.group(0)

    if not data.get("Date of Issue"):
        match = DATE_RE.search(current_text)
        if match:
            data["Date of Issue"] = match.group(0)

    # Final cleanup: hidden print text can accidentally append fields after the inscription,
    # e.g. "HRD 250000219505 | Shape brilliant | Weight ...".
    clean_result_inscriptions(data, report_no)

    return data



def close_extra_windows_and_return_to_main(driver, main_handle: str) -> None:
    """Close any print/new tabs left by HRD and return to the main window."""
    try:
        for handle in list(driver.window_handles):
            if handle != main_handle:
                try:
                    driver.switch_to.window(handle)
                    driver.close()
                except Exception:
                    pass
        driver.switch_to.window(main_handle)
    except Exception:
        try:
            if driver.window_handles:
                driver.switch_to.window(driver.window_handles[0])
        except Exception:
            pass


def split_percent_and_angle(value: str) -> tuple[str, str]:
    """
    Split HRD proportion values such as "16.0 % (35.6 deg)" into:
        percent = "16.0 %"
        angle   = "35.6 deg"
    """
    value = clean_value(value)
    if not value:
        return "", ""

    angle = ""
    match = re.search(r"\(([^)]*deg[^)]*)\)", value, flags=re.IGNORECASE)
    if match:
        angle = clean_value(match.group(1))

    percent = re.sub(r"\s*\([^)]*deg[^)]*\)", "", value, flags=re.IGNORECASE)
    percent = clean_value(percent)
    return percent, angle


def first_non_empty(*values: str) -> str:
    for value in values:
        value = clean_value(str(value or ""))
        if value:
            return value
    return ""


def clean_color_grade(value: str) -> str:
    """
    Convert HRD colour text to only the letter grade for Excel.

    Examples:
        "exceptional white + (D)" -> "D"
        "rare white + (F)"        -> "F"
        "D"                       -> "D"
    """
    value = clean_value(value)
    if not value:
        return ""

    # HRD usually shows the GIA-style letter inside parentheses.
    match = re.search(r"\(([A-Z]{1,2})\)", value, flags=re.IGNORECASE)
    if match:
        return match.group(1).upper()

    # Fallback: if the field already contains a plain diamond color grade.
    match = re.search(r"\b(D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z)\b", value, flags=re.IGNORECASE)
    if match:
        return match.group(1).upper()

    return value


def normalize_to_requested_output_columns(data: Dict[str, str], report_no: str) -> Dict[str, str]:
    """
    Convert HRD field names into the exact output columns requested.

    HRD page fields are named like "Weight (carat)", "Colour", "Clarity",
    "Total Depth", and "Table Width". The output Excel now uses GIA-style
    column names such as "Carat Weight", "Color Grade", "Clarity Grade",
    "Depth", and "Table".
    """
    crown_height, crown_angle = split_percent_and_angle(data.get("Crown Height (β)", ""))
    pavilion_depth, pavilion_angle = split_percent_and_angle(data.get("Pavilion Depth (α)", ""))

    clean_no = clean_report_number(report_no) or clean_report_number(data.get("Report Number", ""))
    inscription = f"HRD {clean_no}" if clean_no else clean_inscription_value(
        first_non_empty(data.get("INSCRIPTION(S)", ""), data.get("Inscriptions", ""), data.get("Inscription on girdle", "")),
        report_no=report_no,
    )

    mapped = {
        "Input Report Number": clean_no or clean_report_number(report_no),
        "Report Number": first_non_empty(data.get("Report Number", ""), clean_no),
        "Lab": "HRD",
        "Date of Issue": data.get("Date of Issue", ""),
        "Shape": data.get("Shape", ""),
        "Measurements": data.get("Measurements", ""),
        "Carat Weight": first_non_empty(data.get("Carat Weight", ""), data.get("Weight (carat)", "")),
        "Color Grade": clean_color_grade(first_non_empty(data.get("Color Grade", ""), data.get("Colour", ""), data.get("Color", ""))),
        "Clarity Grade": first_non_empty(data.get("Clarity Grade", ""), data.get("Clarity", "")),
        "Cut Grade": first_non_empty(data.get("Cut Grade", ""), data.get("Cut", "")),
        "Depth": first_non_empty(data.get("Depth", ""), data.get("Total Depth", "")),
        "Table": first_non_empty(data.get("Table", ""), data.get("Table Width", "")),
        "Crown Angle": first_non_empty(data.get("Crown Angle", ""), crown_angle),
        "Crown Height": first_non_empty(data.get("Crown Height", ""), crown_height),
        "Pavilion Angle": first_non_empty(data.get("Pavilion Angle", ""), pavilion_angle),
        "Pavilion Depth": first_non_empty(data.get("Pavilion Depth", ""), pavilion_depth),
        # HRD printable field: Length Halves Crown. Requested GIA-style column: Star Length.
        "Star Length": first_non_empty(data.get("Star Length", ""), data.get("Length Halves Crown", "")),
        # HRD printable field: Length Halves Pavilion. Requested GIA-style column: Lower Half.
        "Lower Half": first_non_empty(data.get("Lower Half", ""), data.get("Length Halves Pavilion", "")),
        "Girdle": data.get("Girdle", ""),
        "Culet": data.get("Culet", ""),
        "Polish": data.get("Polish", ""),
        "Symmetry": data.get("Symmetry", ""),
        "Fluorescence": data.get("Fluorescence", ""),
        "Clarity Characteristics": first_non_empty(
            data.get("Clarity Characteristics", ""),
            data.get("Clarity Characteristic", ""),
        ),
        "INSCRIPTION(S)": inscription,
    }

    # Keep exact requested columns only. Status and Error are added by fetch_report.
    return {column: mapped.get(column, "") for column in OUTPUT_COLUMNS if column not in {"Status", "Error"}}


def fetch_report(report_no: str, driver, headless: bool = False) -> Dict[str, str]:
    result = {column: "" for column in OUTPUT_COLUMNS}
    result["Input Report Number"] = report_no

    try:
        submitted = submit_report_number(driver, report_no)

        if not submitted:
            result["Status"] = "Failed"
            result["Error"] = "Could not find or submit the HRD Find Diamond input."
            return result

        if not wait_for_report(driver, report_no):
            if headless:
                result["Status"] = "Failed"
                result["Error"] = (
                    "Timed out waiting for HRD report. Run without --headless "
                    "if manual verification appears."
                )
                return result

            print(
                f"Report {report_no}: result did not load. "
                "If any manual check appears, solve it in Chrome, then press Enter..."
            )
            input()

            if not wait_for_report(driver, report_no):
                result["Status"] = "Failed"
                result["Error"] = "Timed out waiting for HRD report after manual check."
                return result

        data = extract_current_report_data(driver, report_no)
        data = normalize_to_requested_output_columns(data, report_no)

        if data:
            result.update(data)
            result["Status"] = "Success"
        else:
            result["Status"] = "Failed"
            result["Error"] = "No HRD report data found."

    except Exception as exc:
        result["Status"] = "Failed"
        result["Error"] = str(exc)

    return result


def order_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Force the Excel output to contain exactly the requested columns, in this order.
    return df.reindex(columns=OUTPUT_COLUMNS, fill_value="")


def save_results(results: List[Dict[str, str]], output_path: str) -> None:
    df_new = pd.DataFrame(results)
    df_new.fillna("", inplace=True)

    if os.path.exists(output_path):
        df_old = pd.read_excel(output_path, dtype=str)
        df_old.fillna("", inplace=True)

        df_combined = pd.concat([df_old, df_new], ignore_index=True)

        if "Input Report Number" in df_combined.columns:
            df_combined.drop_duplicates(
                subset=["Input Report Number"],
                keep="last",
                inplace=True,
            )
    else:
        df_combined = df_new

    df_combined.fillna("", inplace=True)
    df_combined = order_columns(df_combined)
    df_combined.to_excel(output_path, index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Scrape HRD Antwerp diamond grading report data from MyHRD, "
            "click the print-results link, and save data to Excel."
        )
    )

    parser.add_argument(
        "--input",
        "-i",
        default=DEFAULT_INPUT,
        help=f"Input Excel file. Default: {DEFAULT_INPUT}",
    )

    parser.add_argument(
        "--output",
        "-o",
        default=DEFAULT_OUTPUT,
        help=f"Output Excel file. Default: {DEFAULT_OUTPUT}",
    )

    parser.add_argument(
        "--column",
        "-c",
        default=REPORT_COLUMN_NAME,
        help=f"Column containing report numbers. Default: {REPORT_COLUMN_NAME}",
    )

    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run Chrome in headless mode. Avoid this if HRD shows manual verification.",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        report_numbers = read_report_numbers(args.input, args.column)
    except Exception as exc:
        print(f"Error reading report numbers: {exc}")
        return

    print(f"Loaded {len(report_numbers)} report numbers from {args.input}")

    driver = create_driver(headless=args.headless)
    main_handle = driver.current_window_handle
    results: List[Dict[str, str]] = []

    try:
        for idx, report_no in enumerate(report_numbers, start=1):
            print(f"[{idx}/{len(report_numbers)}] Fetching HRD report {report_no}...")

            result = fetch_report(report_no, driver, headless=args.headless)
            results.append(result)

            print(f"    {result.get('Status', '')}: {report_no}")
            if result.get("Error"):
                print(f"    Error: {result.get('Error')}")

            # Save progress after each report to avoid data loss.
            save_results(results, args.output)
            close_extra_windows_and_return_to_main(driver, main_handle)
            time.sleep(0.75)

    finally:
        safe_quit_driver(driver)

    save_results(results, args.output)
    print(f"Results saved to {args.output}")
    os.startfile(args.output)

if __name__ == "__main__":
    main()