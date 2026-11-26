from __future__ import annotations

import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, quote_plus

import requests
import undetected_chromedriver as uc
from openpyxl import Workbook, load_workbook
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait


# =============================================================================
# SETTINGS
# =============================================================================

GIA_DOWNLOAD_FOLDER = Path(
    r"D:\Gaurav\Download\New folder\Downloade Certificate\GIA"
)

IGI_DOWNLOAD_FOLDER = Path(r"D:\Gaurav\Download\New folder\Downloade Certificate\IGI")
HRD_DOWNLOAD_FOLDER = Path(r"D:\Gaurav\Download\New folder\Downloade Certificate\HRD")
LOG_FILE = Path(r"D:\Gaurav\Download\New folder\Downloade Certificate\Certificate Download Log.xlsx")

IGI_PDF_URL_TEMPLATE = "https://pdf.igi.org/ID{report_no}.pdf"
HRD_PDF_URL_TEMPLATE = (
    "https://my.hrdantwerp.com/Download/GetGradingReportPdf/"
    "?reportNumber={report_no}&printDocumentType=DiamondIdentificationReportPlusMini"
)

GIA_REPORT_URL = "https://www.gia.edu/report-check?reportno={report_no}"
GIA_LANDING_URL = "https://www.gia.edu/report-check-landing"

PAGE_TIMEOUT = 18
BUTTON_TIMEOUT = 10
DOWNLOAD_TIMEOUT = 40
INTER_REPORT_DELAY = 2.5
GIA_TYPE_DELAY = 0.16
GIA_AFTER_LOOKUP_DELAY = 2.0
MAX_ATTEMPTS = 3


# =============================================================================
# UNDTECTED-CHROMEDRIVER CLEANUP
# =============================================================================

_ORIGINAL_UC_DEL = getattr(uc.Chrome, "__del__", None)


def _quiet_uc_del(self):
    if _ORIGINAL_UC_DEL is None:
        return
    try:
        _ORIGINAL_UC_DEL(self)
    except Exception:
        pass


uc.Chrome.__del__ = _quiet_uc_del


# =============================================================================
# INPUT TAB
# =============================================================================

def clean_report_number(value: str) -> str:
    return re.sub(r"\s+", "", str(value or "").strip())


def parse_report_numbers(text: str) -> list[str]:
    values = []
    for part in re.split(r"[\s,;]+", text or ""):
        number = clean_report_number(part)
        if not number:
            continue
        if not re.fullmatch(r"[A-Za-z0-9-]{6,30}", number):
            continue
        if number not in values:
            values.append(number)
    return values


def build_entry_page() -> str:
    page = r"""<!doctype html><html><head><meta charset="utf-8">
<title>Certificate Downloader</title>
<style>
*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;background:#f3f5f8;font-family:Arial;color:#182230}
.card{width:min(720px,calc(100vw - 32px));background:#fff;padding:34px;border-radius:16px;box-shadow:0 16px 50px rgba(20,31,48,.12)}
h1{margin:0 0 10px}p{color:#5b6574;line-height:1.5}label{display:block;font-weight:700;margin:18px 0 8px}
select{width:100%;height:48px;padding:0 12px;border:1px solid #b8c2cf;border-radius:9px;background:#fff;font-size:16px}
textarea{width:100%;min-height:250px;padding:14px;border:1px solid #b8c2cf;border-radius:9px;font:17px/1.5 Consolas,monospace;resize:vertical}
button{width:100%;height:50px;margin-top:16px;border:0;border-radius:9px;background:#e30613;color:#fff;font-size:17px;font-weight:700;cursor:pointer}
button:disabled{opacity:.55}#message{margin-top:15px;font-weight:700}#log{margin-top:12px;max-height:220px;overflow:auto;white-space:pre-wrap;font:14px/1.5 Consolas,monospace}
.folder{margin-top:18px;padding:12px;background:#f8fafc;border-radius:8px;font-size:13px;line-height:1.55}
</style></head><body><div class="card">
<h1>GIA / IGI / HRD Certificate Downloader</h1>
<p>Select a certificate, paste report numbers and download. After completion, select another certificate and continue. No report data is extracted.</p>
<form id="f"><label>Certificate</label><select id="lab"><option value="GIA">GIA</option><option value="IGI">IGI</option><option value="HRD">HRD Antwerp</option></select>
<label>Report Numbers</label><textarea id="numbers" placeholder="Paste one report number per line" autofocus></textarea>
<button id="btn">DOWNLOAD CERTIFICATES</button><div id="message">Ready</div><div id="log"></div><div id="submitted" hidden></div></form>
<div class="folder"><b>GIA:</b> D:\Gaurav\Download\New folder\Downloade Certificate\GIA<br>
<b>IGI:</b> D:\Gaurav\Download\New folder\Downloade Certificate\IGI<br>
<b>HRD:</b> D:\Gaurav\Download\New folder\Downloade Certificate\HRD<br>
<b>Log:</b> D:\Gaurav\Download\New folder\Downloade Certificate\Certificate Download Log.xlsx</div>
</div><script>
const f=document.getElementById('f'),n=document.getElementById('numbers'),b=document.getElementById('btn'),
m=document.getElementById('message'),s=document.getElementById('submitted'),lab=document.getElementById('lab');
f.addEventListener('submit',e=>{e.preventDefault();const vals=[...new Set(n.value.split(/[\s,;]+/).map(x=>x.trim()).filter(Boolean))];
if(!vals.length){m.textContent='Paste at least one report number.';return;}
s.textContent=JSON.stringify({lab:lab.value,numbers:vals,token:Date.now()});
n.readOnly=true;lab.disabled=true;b.disabled=true;b.textContent='DOWNLOADING...';m.textContent='Starting '+lab.value+' downloads...';});
</script></body></html>"""
    return "data:text/html;charset=utf-8," + quote(page)


def wait_for_numbers(driver, first_tab=None, last_token=None):
    if first_tab is None:
        driver.get(build_entry_page())
        first_tab = driver.current_window_handle
    else:
        driver.switch_to.window(first_tab)

    while True:
        try:
            raw = driver.find_element(By.ID, "submitted").get_attribute("textContent") or ""
            if raw:
                payload = json.loads(raw)
                token = payload.get("token")
                if token != last_token:
                    lab = str(payload.get("lab", "GIA")).upper()
                    numbers = parse_report_numbers("\n".join(map(str, payload.get("numbers", []))))
                    if lab in {"GIA", "IGI", "HRD"} and numbers:
                        return lab, numbers, first_tab, token
        except Exception:
            pass
        time.sleep(0.1)


def reset_entry_page(driver, first_tab, lab, downloaded, failed):
    driver.switch_to.window(first_tab)
    driver.execute_script("""
    const n=document.getElementById('numbers'),lab=document.getElementById('lab'),
    b=document.getElementById('btn'),m=document.getElementById('message'),s=document.getElementById('submitted');
    n.readOnly=false;n.value='';n.focus();lab.disabled=false;b.disabled=false;b.textContent='DOWNLOAD CERTIFICATES';
    m.textContent=arguments[0]+' completed: '+arguments[1]+' downloaded, '+arguments[2]+' failed. Select another certificate and continue.';
    s.textContent='';document.title='Certificate Downloader';
    """, lab, downloaded, failed)


# =============================================================================
# CHROME
# =============================================================================

def build_options() -> uc.ChromeOptions:
    GIA_DOWNLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    IGI_DOWNLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    HRD_DOWNLOAD_FOLDER.mkdir(parents=True, exist_ok=True)

    options = uc.ChromeOptions()
    options.page_load_strategy = "eager"
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-component-update")
    options.add_argument("--disable-sync")
    options.add_argument("--lang=en-US,en")

    # Keep images enabled because GIA's UI/download control may depend on normal page rendering.
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(GIA_DOWNLOAD_FOLDER),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.notifications": 2,
        },
    )
    return options


def create_driver():
    try:
        return uc.Chrome(options=build_options())
    except SessionNotCreatedException as exc:
        match = re.search(r"Current browser version is (\d+)\.", str(exc))
        if not match:
            raise
        return uc.Chrome(
            options=build_options(),
            version_main=int(match.group(1)),
        )


def open_worker_tab(driver, first_tab):
    existing = set(driver.window_handles)
    driver.execute_script("window.open('about:blank','_blank');")
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > len(existing))
    worker = [h for h in driver.window_handles if h not in existing][-1]
    driver.switch_to.window(worker)
    return worker


# =============================================================================
# GIA DOWNLOAD
# =============================================================================

def canonical(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).strip()


def dismiss_cookie_popup(driver):
    """Fast best-effort cookie acceptance; does not stop the batch."""
    labels = ["Accept All", "Accept all", "ACCEPT ALL"]
    for label in labels:
        try:
            elements = driver.find_elements(
                By.XPATH,
                f"//*[self::button or self::a or @role='button'][contains(normalize-space(.), {json.dumps(label)})]"
            )
            for el in elements:
                if el.is_displayed() and el.is_enabled():
                    driver.execute_script("arguments[0].click();", el)
                    time.sleep(0.3)
                    return
        except Exception:
            pass


def wait_for_manual_gia_cookie_accept(driver, first_tab=None, timeout=300):
    """
    Open GIA first and let the USER handle the cookie banner.

    Nothing is clicked automatically.  If a visible cookie consent control is
    present, the program waits until the user accepts/closes it.  If no banner
    is present (for example, consent was saved earlier), processing continues
    after a short settle delay.
    """
    driver.get(GIA_LANDING_URL)
    time.sleep(1.5)

    print("GIA opened. If a cookie banner is shown, please accept it manually.")
    print("Downloads will start automatically after the cookie banner is gone.")

    deadline = time.time() + timeout
    saw_banner = False

    while time.time() < deadline:
        visible_cookie_control = False
        try:
            candidates = driver.find_elements(
                By.XPATH,
                "//*[self::button or self::a or @role='button']["
                "contains(translate(normalize-space(.),"
                "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'ACCEPT') "
                "or contains(translate(normalize-space(.),"
                "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'AGREE')]"
            )
            visible_cookie_control = any(
                el.is_displayed() for el in candidates
            )
        except Exception:
            visible_cookie_control = False

        if visible_cookie_control:
            saw_banner = True
            time.sleep(0.5)
            continue

        # If the banner was seen and is now gone, the user handled it.
        if saw_banner:
            time.sleep(1.5)
            return True

        # No banner appeared. Consent may already be stored. Give the page a
        # few seconds to settle before continuing.
        if time.time() + 0.1 >= deadline - (timeout - 4):
            time.sleep(1.0)
            return True

        time.sleep(0.5)

    return False


def find_download_button(driver):
    """
    Locate the actual red GIA 'DOWNLOAD REPORT PDF' button visible on the
    report-check page.
    """
    driver.switch_to.default_content()

    selectors = [
        # Exact visible text shown on GIA report-check.
        (By.XPATH, "//a[normalize-space()='DOWNLOAD REPORT PDF']"),
        (By.XPATH, "//button[normalize-space()='DOWNLOAD REPORT PDF']"),
        (By.XPATH, "//*[@role='button' and normalize-space()='DOWNLOAD REPORT PDF']"),

        # Case-insensitive / nested-text fallbacks.
        (
            By.XPATH,
            "//a[contains(translate(normalize-space(.),"
            "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
            "'DOWNLOAD REPORT PDF')]"
        ),
        (
            By.XPATH,
            "//button[contains(translate(normalize-space(.),"
            "'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
            "'DOWNLOAD REPORT PDF')]"
        ),

        # GIA PDF/download href fallback.
        (
            By.XPATH,
            "//a[contains(translate(@href,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'.pdf') "
            "or contains(translate(@href,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'download')]"
        ),
    ]

    for by, selector in selectors:
        try:
            elements = driver.find_elements(by, selector)
            for element in elements:
                if element.is_displayed() and element.is_enabled():
                    return element
        except Exception:
            pass

    return None


def wait_for_download_button(driver, timeout=BUTTON_TIMEOUT):
    deadline = time.time() + timeout
    while time.time() < deadline:
        button = find_download_button(driver)
        if button is not None:
            return button

        body = ""
        try:
            body = canonical(driver.find_element(By.TAG_NAME, "body").text)
        except Exception:
            pass

        if any(x in body for x in [
            "does not exist",
            "invalid report",
            "no report found",
            "not available through report check",
        ]):
            return None

        time.sleep(0.40)

    return None


def snapshot_files():
    GIA_DOWNLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
    return {
        p.name: (p.stat().st_mtime_ns, p.stat().st_size)
        for p in GIA_DOWNLOAD_FOLDER.iterdir()
        if p.is_file()
    }


def wait_for_new_download(before, report_no, driver, report_tab, tabs_before_click):
    """
    Wait for either:
      1) Chrome to download the PDF directly, or
      2) GIA to open pdf.gia.edu in a new tab.

    If a PDF tab opens, close it after triggering/saving the PDF and return
    to the report tab before continuing to the next report.
    """
    deadline = time.time() + DOWNLOAD_TIMEOUT
    target = GIA_DOWNLOAD_FOLDER / f"{report_no}.pdf"
    handled_pdf_tabs = set()

    while time.time() < deadline:

        # -------------------------------------------------------------
        # GIA often opens the PDF in a NEW TAB instead of downloading it.
        # Detect that tab, obtain its PDF URL, download via Chrome DevTools,
        # then close the PDF tab and return to the report tab.
        # -------------------------------------------------------------
        try:
            current_tabs = set(driver.window_handles)
            new_tabs = [
                handle
                for handle in current_tabs
                if handle not in tabs_before_click
                and handle not in handled_pdf_tabs
            ]

            for handle in new_tabs:
                handled_pdf_tabs.add(handle)

                try:
                    driver.switch_to.window(handle)

                    # Give Chrome a moment to resolve the new URL.
                    end_wait = time.time() + 15
                    pdf_url = ""

                    while time.time() < end_wait:
                        pdf_url = driver.current_url or ""
                        if (
                            "pdf.gia.edu" in pdf_url.lower()
                            or ".pdf" in pdf_url.lower()
                            or "reportnumber=" in pdf_url.lower()
                        ):
                            break
                        time.sleep(0.25)

                    if pdf_url and (
                        "pdf.gia.edu" in pdf_url.lower()
                        or ".pdf" in pdf_url.lower()
                        or "reportnumber=" in pdf_url.lower()
                    ):
                        # Use Chrome itself to fetch the PDF bytes so the same
                        # browser session/cookies are used.
                        result = driver.execute_async_script(
                            """
                            const url = arguments[0];
                            const done = arguments[arguments.length - 1];

                            fetch(url, {credentials: 'include'})
                              .then(r => {
                                if (!r.ok) {
                                  throw new Error('HTTP ' + r.status);
                                }
                                return r.arrayBuffer();
                              })
                              .then(buf => {
                                const bytes = new Uint8Array(buf);
                                let binary = '';
                                const chunk = 0x8000;

                                for (let i = 0; i < bytes.length; i += chunk) {
                                  binary += String.fromCharCode.apply(
                                    null,
                                    bytes.subarray(i, i + chunk)
                                  );
                                }

                                done({
                                  ok: true,
                                  data: btoa(binary)
                                });
                              })
                              .catch(err => {
                                done({
                                  ok: false,
                                  error: String(err)
                                });
                              });
                            """,
                            pdf_url,
                        )

                        if result and result.get("ok"):
                            import base64

                            pdf_bytes = base64.b64decode(result["data"])

                            if pdf_bytes.startswith(b"%PDF"):
                                target.write_bytes(pdf_bytes)

                                try:
                                    driver.close()
                                except Exception:
                                    pass

                                driver.switch_to.window(report_tab)

                                return True, str(target)

                except Exception:
                    pass
                finally:
                    # Close any PDF/new tab that remains open.
                    try:
                        if handle in driver.window_handles:
                            driver.switch_to.window(handle)
                            driver.close()
                    except Exception:
                        pass

                    try:
                        driver.switch_to.window(report_tab)
                    except Exception:
                        pass

        except Exception:
            try:
                driver.switch_to.window(report_tab)
            except Exception:
                pass

        # -------------------------------------------------------------
        # Normal Chrome-download case.
        # -------------------------------------------------------------
        try:
            partial = list(GIA_DOWNLOAD_FOLDER.glob("*.crdownload"))

            files = [
                p
                for p in GIA_DOWNLOAD_FOLDER.iterdir()
                if p.is_file()
                and p.suffix.lower() != ".crdownload"
            ]

            candidates = []

            for p in files:
                stat = p.stat()
                old = before.get(p.name)

                if old is None or old != (stat.st_mtime_ns, stat.st_size):
                    candidates.append(p)

            if candidates and not partial:
                newest = max(
                    candidates,
                    key=lambda p: p.stat().st_mtime_ns,
                )

                if newest.resolve() != target.resolve():
                    if target.exists():
                        target.unlink()

                    newest.replace(target)

                if target.exists() and target.stat().st_size > 500:
                    try:
                        driver.switch_to.window(report_tab)
                    except Exception:
                        pass

                    return True, str(target)

        except Exception:
            pass

        time.sleep(0.10)

    # Cleanup any PDF tabs that appeared during the attempt.
    try:
        for handle in list(driver.window_handles):
            if handle not in tabs_before_click:
                try:
                    driver.switch_to.window(handle)
                    driver.close()
                except Exception:
                    pass

        driver.switch_to.window(report_tab)
    except Exception:
        pass

    return False, "PDF did not download before timeout."


def find_gia_report_input(driver):
    """Find the visible 'Look Up a Report' input on GIA's landing page."""
    selectors = [
        (By.CSS_SELECTOR, "input[type='text']"),
        (By.XPATH, "//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'report')]"),
        (By.XPATH, "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'report')]"),
        (By.XPATH, "//input[contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'report')]"),
    ]
    for by, selector in selectors:
        try:
            for el in driver.find_elements(by, selector):
                if el.is_displayed() and el.is_enabled():
                    return el
        except Exception:
            pass
    return None


def find_gia_lookup_button(driver):
    """Find the visible LOOK UP button without moving the mouse pointer."""
    selectors = [
        (By.XPATH, "//button[normalize-space()='LOOK UP']"),
        (By.XPATH, "//input[@type='submit' and translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='LOOK UP']"),
        (By.XPATH, "//*[self::button or self::a or @role='button'][contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'LOOK UP')]"),
    ]
    for by, selector in selectors:
        try:
            for el in driver.find_elements(by, selector):
                if el.is_displayed() and el.is_enabled():
                    return el
        except Exception:
            pass
    return None


def softly_lookup_gia_report(driver, report_no):
    """
    Use GIA's visible Report Check form like a normal user:
      landing page -> slowly type report number -> click LOOK UP -> wait.
    Selenium clicks do not move the user's physical mouse cursor.
    """
    if "report-check-landing" not in (driver.current_url or "").lower():
        driver.get(GIA_LANDING_URL)
        time.sleep(2.0)

    deadline = time.time() + 15
    field = None
    while time.time() < deadline:
        field = find_gia_report_input(driver)
        if field is not None:
            break
        time.sleep(0.4)
    if field is None:
        return False, "GIA report-number input was not found."

    # Focus with DOM/Selenium only; this leaves the physical mouse cursor free.
    try:
        field.click()
    except Exception:
        driver.execute_script("arguments[0].focus();", field)
    time.sleep(0.5)

    # Clear gently, then type one character at a time.
    try:
        field.send_keys(Keys.CONTROL, "a")
        time.sleep(0.2)
        field.send_keys(Keys.BACKSPACE)
    except Exception:
        try:
            field.clear()
        except Exception:
            pass

    for ch in str(report_no):
        field.send_keys(ch)
        time.sleep(GIA_TYPE_DELAY)

    time.sleep(0.8)
    lookup = find_gia_lookup_button(driver)
    if lookup is None:
        return False, "GIA LOOK UP button was not found."

    try:
        lookup.click()
    except Exception:
        driver.execute_script("arguments[0].click();", lookup)

    # Let GIA render the report naturally before looking for the PDF control.
    time.sleep(GIA_AFTER_LOOKUP_DELAY)
    return True, ""


def download_one(driver, report_no):
    target = GIA_DOWNLOAD_FOLDER / f"{report_no}.pdf"

    if target.exists() and target.stat().st_size > 500:
        return True, f"Already exists: {target}"

    last_error = ""

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            # Always use the visible GIA form instead of jumping directly to a
            # report-number URL. This is slower and gentler on the website.
            driver.get(GIA_LANDING_URL)
            time.sleep(2.0)

            ok_lookup, lookup_error = softly_lookup_gia_report(driver, report_no)
            if not ok_lookup:
                last_error = lookup_error
                if attempt < MAX_ATTEMPTS:
                    time.sleep(3 * attempt)
                continue

            button = wait_for_download_button(driver, timeout=15)
            if button is None:
                last_error = "DOWNLOAD REPORT PDF button not found after LOOK UP."
                if attempt < MAX_ATTEMPTS:
                    time.sleep(3 * attempt)
                continue

            before = snapshot_files()
            report_tab = driver.current_window_handle
            tabs_before_click = set(driver.window_handles)

            # Scroll without moving the physical mouse pointer.
            try:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center',inline:'center'});",
                    button,
                )
                time.sleep(0.8)
            except Exception:
                pass

            clicked = False
            try:
                button.click()
                clicked = True
            except Exception:
                pass

            if not clicked:
                try:
                    driver.execute_script("arguments[0].click();", button)
                    clicked = True
                except Exception:
                    pass

            if not clicked:
                last_error = "Found DOWNLOAD REPORT PDF but could not activate it."
                if attempt < MAX_ATTEMPTS:
                    time.sleep(3 * attempt)
                continue

            ok, result = wait_for_new_download(
                before=before,
                report_no=report_no,
                driver=driver,
                report_tab=report_tab,
                tabs_before_click=tabs_before_click,
            )
            if ok:
                return True, result

            last_error = result

        except Exception as exc:
            last_error = str(exc)

        if attempt < MAX_ATTEMPTS:
            # Increasing pause prevents rapid repeated requests when GIA is slow.
            time.sleep(3 * attempt)

    return False, last_error or "Download failed."


# =============================================================================
# IGI / HRD + EXCEL LOG
# =============================================================================

def direct_pdf_download(lab, report_no):
    folder = IGI_DOWNLOAD_FOLDER if lab == "IGI" else HRD_DOWNLOAD_FOLDER
    url = (IGI_PDF_URL_TEMPLATE if lab == "IGI" else HRD_PDF_URL_TEMPLATE).format(report_no=quote_plus(report_no))
    referer = "https://www.igi.org/" if lab == "IGI" else "https://my.hrdantwerp.com/"
    target = folder / f"{report_no}.pdf"
    folder.mkdir(parents=True, exist_ok=True)

    if target.exists() and target.stat().st_size > 500:
        return True, str(target)

    session = requests.Session()
    session.headers.update({"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0 Safari/537.36",
                            "Accept":"application/pdf,application/octet-stream;q=0.9,*/*;q=0.8","Referer":referer})
    if lab == "HRD":
        try: session.get("https://my.hrdantwerp.com/", timeout=10)
        except Exception: pass

    last_error = ""
    for attempt in range(1, 4):
        try:
            r = session.get(url, timeout=30, allow_redirects=True)
            content = r.content or b""
            ctype = (r.headers.get("Content-Type") or "").lower()
            if r.status_code != 200: last_error = f"HTTP {r.status_code}"
            elif not content.startswith(b"%PDF") and "application/pdf" not in ctype: last_error = "Website did not return a PDF"
            elif len(content) < 500: last_error = "Returned PDF was too small"
            else:
                target.write_bytes(content)
                return True, str(target)
        except Exception as exc:
            last_error = str(exc)
        if attempt < 3: time.sleep(attempt)
    return False, last_error or "Download failed."


def append_download_log(lab, report_no, ok, detail):
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    try:
        if LOG_FILE.exists():
            wb = load_workbook(LOG_FILE)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Download Log"
            ws.append(["Date/Time","Certificate","Report Number","Status","Saved File","Error"])

        ws.append([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            lab, report_no,
            "Downloaded" if ok else "Failed",
            detail if ok else "",
            "" if ok else detail,
        ])
        for col, width in {"A":21,"B":14,"C":22,"D":14,"E":80,"F":60}.items():
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "A2"
        wb.save(LOG_FILE)
    except Exception as exc:
        print(f"    WARNING: Excel log could not be updated: {exc}")


# =============================================================================
# PROGRESS
# =============================================================================

def update_progress(driver, first_tab, worker_tab, report_no, pos, total, ok, detail):
    try:
        driver.switch_to.window(first_tab)

        line = f"[{pos}/{total}] {report_no}: " + ("Downloaded" if ok else f"Failed - {detail}")
        headline = f"Processed {pos} of {total}."
        if pos == total:
            headline = f"Completed. {total} report(s) processed."

        driver.execute_script(
            """
            const m=document.getElementById('message');
            const l=document.getElementById('log');
            const b=document.getElementById('btn');
            if(m) m.textContent=arguments[0];
            if(l){l.textContent+=(l.textContent?'\\n':'')+arguments[1];l.scrollTop=l.scrollHeight;}
            if(b) b.textContent=arguments[2];
            """,
            headline,
            line,
            "COMPLETED" if pos == total else f"DOWNLOADING {pos}/{total}...",
        )

        driver.switch_to.window(worker_tab)
    except Exception:
        try:
            driver.switch_to.window(worker_tab)
        except Exception:
            pass


# =============================================================================
# MAIN
# =============================================================================

def main():
    for folder in (GIA_DOWNLOAD_FOLDER, IGI_DOWNLOAD_FOLDER, HRD_DOWNLOAD_FOLDER):
        folder.mkdir(parents=True, exist_ok=True)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

    driver = None
    try:
        driver = create_driver()
        first_tab = None
        last_token = None
        gia_cookie_ready = False
        folders = {"GIA":GIA_DOWNLOAD_FOLDER,"IGI":IGI_DOWNLOAD_FOLDER,"HRD":HRD_DOWNLOAD_FOLDER}

        print(f"Excel log: {LOG_FILE}")
        print("After each batch, select another certificate and continue.\n")

        while True:
            lab, report_numbers, first_tab, last_token = wait_for_numbers(driver, first_tab, last_token)
            total = len(report_numbers)
            downloaded = failed = 0
            worker_tab = open_worker_tab(driver, first_tab) if lab == "GIA" else None

            if lab == "GIA" and worker_tab:
                driver.switch_to.window(worker_tab)

                # Ask the user to handle GIA cookie consent only once per
                # program/browser session. The same Chrome profile/session is
                # then reused for all later GIA batches.
                if not gia_cookie_ready:
                    cookie_ready = wait_for_manual_gia_cookie_accept(driver, first_tab)
                    if not cookie_ready:
                        print("Cookie confirmation timed out. Submit the GIA batch again when ready.")
                        try:
                            if worker_tab in driver.window_handles:
                                driver.close()
                        except Exception:
                            pass
                        reset_entry_page(driver, first_tab, lab, 0, total)
                        continue
                    gia_cookie_ready = True

                # Cookie consent is complete. Minimize the automation browser
                # so the user can continue working in other applications.
                # Selenium still sends keys/clicks directly to DOM elements;
                # it does not use or move the user's physical mouse pointer.
                try:
                    driver.minimize_window()
                    time.sleep(1.0)
                except Exception:
                    pass

            for pos, report_no in enumerate(report_numbers, 1):
                print(f"[{pos}/{total}] {lab} {report_no}: downloading...")
                if lab == "GIA":
                    driver.switch_to.window(worker_tab)
                    ok, detail = download_one(driver, report_no)
                else:
                    ok, detail = direct_pdf_download(lab, report_no)

                if ok: downloaded += 1
                else: failed += 1

                append_download_log(lab, report_no, ok, detail)

                try:
                    driver.switch_to.window(first_tab)
                    line = f"[{pos}/{total}] {report_no}: " + ("Downloaded" if ok else f"Failed - {detail}")
                    driver.execute_script("""
                    const m=document.getElementById('message'),l=document.getElementById('log'),b=document.getElementById('btn');
                    m.textContent=arguments[0];l.textContent+=(l.textContent?'\\n':'')+arguments[1];l.scrollTop=l.scrollHeight;
                    b.textContent=arguments[2];
                    """, f"Processed {pos} of {total}.", line, "COMPLETED" if pos == total else f"DOWNLOADING {pos}/{total}...")
                except Exception:
                    pass

                if lab == "GIA" and worker_tab:
                    try: driver.switch_to.window(worker_tab)
                    except Exception: pass

                # Be gentle with certificate websites between reports.
                if pos < total:
                    time.sleep(INTER_REPORT_DELAY)

            if worker_tab:
                try:
                    if worker_tab in driver.window_handles:
                        driver.switch_to.window(worker_tab); driver.close()
                except Exception: pass

            print(f"{lab} complete: {downloaded} downloaded, {failed} failed.")
            reset_entry_page(driver, first_tab, lab, downloaded, failed)

            # Bring the entry page back only after the batch is finished so
            # another batch can be submitted. During GIA processing Chrome
            # stays minimized in the background.
            if lab == "GIA":
                try:
                    driver.switch_to.window(first_tab)
                    driver.maximize_window()
                except Exception:
                    pass

    except KeyboardInterrupt:
        print("\nStopped by user.")
    except Exception as exc:
        print(f"\nERROR: {exc}")
        input("\nPress ENTER to close...")
    finally:
        if driver:
            try: driver.quit()
            except Exception: pass


if __name__ == "__main__":
    main()
