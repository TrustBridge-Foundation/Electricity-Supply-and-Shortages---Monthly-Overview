"""
Grid-India daily VRE (NLDC REMC) report: download the PDFs and extract
    Table 8  Total Curtailment  -> Wind MU, Solar MU
    Table 9  SUM (TRAS Down)    -> Wind MU, Solar MU
into one monthly Excel workbook.

Observation date = the internal "Report for" date. A PDF whose file-name date
disagrees with its internal date is rejected (and listed in the audit sheet).

Two steps, as in as_scuc.py:
    fetch_vre_month(y, m)         opens Chrome (Selenium), lists the month's reports,
                                  downloads the PDFs into GRID_VRE_REPORTS/PDF/<FY>/<month>/
    build_vre_month_excel(y, m)   parses the local PDFs (no browser needed) and writes
                                  GRID_VRE_REPORTS/Panels/VRE_TRAS_Down_<month><year>.xlsx
                                    sheet 'Daily Data'   : one row per calendar day
                                    sheet 'Report Audit' : one row per PDF, with status
    download_and_build_vre_month(y, m)   both, in order

Requires: python -m pip install selenium webdriver-manager requests pdfplumber pandas openpyxl
"""

import re
import time
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse, unquote, urljoin

import requests
import pandas as pd
import pdfplumber

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
)

from webdriver_manager.chrome import ChromeDriverManager

from config import (
    VRE_ROOT, VRE_METADATA_FILE,
    financial_year, expected_days,
    vre_month_folder, vre_panel_path,
)

BASE_URL = "https://grid-india.in/en/reports/daily-vre-report"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/154.0.0.0 Safari/537.36"
    ),
    "Accept": "application/pdf,*/*",
    "Referer": BASE_URL,
})

DEBUG_DIR = VRE_ROOT / "_debug"          # page dumps written when the site layout changes

VALUE_KEYS = [
    ("Wind Curtailment (MU)",  "wind_curtailment_mu"),
    ("Solar Curtailment (MU)", "solar_curtailment_mu"),
    ("TRAS Down Wind (MU)",    "tras_down_wind_mu"),
    ("TRAS Down Solar (MU)",   "tras_down_solar_mu"),
]
PDF_NAME = re.compile(r"^(\d{4})-(\d{2})-(\d{2})_NLDC_REMC_REPORT\.pdf$", re.I)


# --- GENERAL HELPERS ---

def clean_text(value):
    if value is None:
        return ""
    return re.sub(
        r"\s+",
        " ",
        str(value).replace("\n", " ").replace("\r", " "),
    ).strip()


def norm(value):
    return clean_text(value).lower()


def numeric_value(value):
    if value is None:
        return None

    text = clean_text(value).replace(",", "").replace("\u2212", "-")

    if text.lower() in {"", "-", "--", "\u2014", "\u2013", "nan", "none"}:
        return None

    match = re.search(r"-?\d+(?:\.\d+)?", text)

    if not match:
        return None

    try:
        return float(match.group())
    except ValueError:
        return None


def parse_date_text(text):
    if not text:
        return None

    text = str(text)

    patterns = [
        r"(?<!\d)(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})(?!\d)",
        r"(?<!\d)(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2})(?!\d)",
        r"(?<!\d)(\d{1,2})[- ]([A-Za-z]{3,9})[- ](\d{2,4})(?!\d)",
        r"(?<!\d)(\d{4})-(\d{2})-(\d{2})(?!\d)",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, text):
            try:
                a, b, c = match.groups()

                if len(a) == 4:
                    year, month, day = int(a), int(b), int(c)
                elif b.isalpha():
                    day = int(a)
                    month = datetime.strptime(b[:3], "%b").month
                    year = int(c)
                else:
                    day, month, year = int(a), int(b), int(c)

                if year < 100:
                    year += 2000

                return date(year, month, day)

            except (ValueError, IndexError):
                continue

    return None


def filename_date(url):
    filename = unquote(Path(urlparse(url).path).name)

    match = re.search(
        r"(\d{2})\.(\d{2})\.(\d{4})_NLDC_REMC_REPORT_\d+\.pdf$",
        filename,
        re.IGNORECASE,
    )

    if match:
        try:
            day, month, year = map(int, match.groups())
            return date(year, month, day)
        except ValueError:
            return None

    return parse_date_text(filename)


def save_diagnostic(driver, filename):
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    path = DEBUG_DIR / filename
    path.write_text(driver.page_source, encoding="utf-8")
    print(f"Diagnostic HTML saved: {path}")
    return path


# --- SELENIUM SETUP ---

def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )

    driver.set_page_load_timeout(60)
    return driver




# --- REACT DROPDOWN SELECTION ---
# The inputs are hidden focus proxies, not visible controls.
# Financial year: input#options   |   Month: input#react-select-3-input
# Always re-find elements after a React re-render.

def dropdown_input_locator(dropdown_index):
    if dropdown_index == 0:
        return (By.CSS_SELECTOR, "input#options")

    if dropdown_index == 1:
        return (By.CSS_SELECTOR, "input#react-select-3-input")

    raise ValueError("Dropdown index must be 0 or 1.")


def get_dropdown_control(driver, dropdown_index):
    """
    Return the visible React Select control.

    The month dropdown is located by its stable #month-options ID.
    The financial-year dropdown is located through input#options.
    """

    if dropdown_index == 1:
        # Use the explicit month container first.
        containers = driver.find_elements(
            By.CSS_SELECTOR,
            "#month-options",
        )

        for container in containers:
            try:
                controls = container.find_elements(
                    By.CSS_SELECTOR,
                    ".my-select__control",
                )

                for control in controls:
                    if control.is_displayed() and control.is_enabled():
                        return control

            except StaleElementReferenceException:
                continue

        # Fallback: use the month input's ancestors.
        month_input = driver.find_elements(
            By.CSS_SELECTOR,
            "input#react-select-3-input",
        )

        for input_element in month_input:
            try:
                control = driver.execute_script(
                    """
                    let el = arguments[0];

                    while (el && el !== document.body) {
                        if (
                            el.classList &&
                            el.classList.contains("my-select__control")
                        ) {
                            return el;
                        }
                        el = el.parentElement;
                    }

                    return null;
                    """,
                    input_element,
                )

                if control and control.is_displayed():
                    return control

            except StaleElementReferenceException:
                continue

    else:
        # Financial-year dropdown.
        fy_inputs = driver.find_elements(
            By.CSS_SELECTOR,
            "input#options",
        )

        for input_element in fy_inputs:
            try:
                control = driver.execute_script(
                    """
                    let el = arguments[0];

                    while (el && el !== document.body) {
                        if (
                            el.classList &&
                            el.classList.contains("my-select__control")
                        ) {
                            return el;
                        }
                        el = el.parentElement;
                    }

                    return null;
                    """,
                    input_element,
                )

                if control and control.is_displayed():
                    return control

            except StaleElementReferenceException:
                continue

    # Last fallback: identify controls from the dropdown inputs'
    # nearest parent containers, without relying on a count of
    # visible .period_drp_select elements.
    input_locator = dropdown_input_locator(dropdown_index)

    inputs = driver.find_elements(*input_locator)

    for input_element in inputs:
        try:
            control = driver.execute_script(
                """
                const input = arguments[0];

                let el = input.parentElement;

                while (el && el !== document.body) {
                    const controls = el.querySelectorAll(
                        ".my-select__control"
                    );

                    if (controls.length > 0) {
                        for (const c of controls) {
                            const rect = c.getBoundingClientRect();
                            const style = window.getComputedStyle(c);

                            if (
                                rect.width > 0 &&
                                rect.height > 0 &&
                                style.display !== "none" &&
                                style.visibility !== "hidden"
                            ) {
                                return c;
                            }
                        }
                    }

                    el = el.parentElement;
                }

                return null;
                """,
                input_element,
            )

            if control and control.is_displayed():
                return control

        except StaleElementReferenceException:
            continue

    return None


def get_dropdown_selected_value(driver, dropdown_index):
    """Read the visible selected value from the correct control."""

    if dropdown_index == 1:
        # The month control has a stable container ID.
        selectors = [
            "#month-options .my-select__single-value",
            "#month-options [class*='single-value']",
        ]
    else:
        selectors = [
            "#options",
            ".period_drp_select .my-select__single-value",
        ]

    if dropdown_index == 0:
        # First try the financial-year input's ancestor.
        inputs = driver.find_elements(By.CSS_SELECTOR, "input#options")

        for input_element in inputs:
            try:
                value = driver.execute_script(
                    """
                    let el = arguments[0];

                    while (el && el !== document.body) {
                        const selected = el.querySelector(
                            ".my-select__single-value"
                        );

                        if (selected) {
                            return selected.textContent.trim();
                        }

                        el = el.parentElement;
                    }

                    return "";
                    """,
                    input_element,
                )

                if value:
                    return value.strip()

            except StaleElementReferenceException:
                continue

    for selector in selectors:
        try:
            elements = driver.find_elements(
                By.CSS_SELECTOR,
                selector,
            )

            for element in elements:
                if dropdown_index == 0 and selector == "#options":
                    continue

                if element.is_displayed():
                    text = element.text.strip()
                    if text:
                        return text

        except StaleElementReferenceException:
            continue

    return ""


def select_react_dropdown(
    driver,
    dropdown_index,
    target_value,
    description,
    timeout=25,
):
    """Select a value using the visible React Select control."""

    target_upper = target_value.upper()

    print(f"Selecting {description}: {target_value}")

    wait = WebDriverWait(driver, timeout)

    def control_is_available(d):
        control = get_dropdown_control(d, dropdown_index)
        return control if control is not None else False

    control = wait.until(control_is_available)

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        control,
    )

    time.sleep(0.4)

    # Re-find the control immediately before clicking.
    control = get_dropdown_control(driver, dropdown_index)

    if control is None:
        raise RuntimeError(
            f"Could not locate the visible {description} control."
        )

    try:
        control.click()
    except Exception:
        driver.execute_script(
            "arguments[0].click();",
            control,
        )

    time.sleep(0.5)

    # Try clicking the exact option.
    option_xpath = (
        "//*[@role='option' and "
        "translate(normalize-space(.), "
        "'abcdefghijklmnopqrstuvwxyz', "
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ')="
        f"'{target_upper}']"
    )

    try:
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, option_xpath)
            )
        )

        try:
            option.click()
        except Exception:
            driver.execute_script(
                "arguments[0].click();",
                option,
            )

    except TimeoutException:
        # Keyboard fallback: re-find the hidden input.
        input_locator = dropdown_input_locator(dropdown_index)

        input_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(input_locator)
        )

        driver.execute_script(
            "arguments[0].focus();",
            input_element,
        )

        input_element.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.3)
        input_element.send_keys(Keys.ENTER)

    # React may briefly unmount and rebuild controls.
    # Verification therefore repeatedly re-finds the control.
    def value_is_selected(d):
        try:
            selected = get_dropdown_selected_value(
                d,
                dropdown_index,
            )

            return selected.upper() == target_upper

        except StaleElementReferenceException:
            return False

    try:
        wait.until(value_is_selected)

    except TimeoutException:
        diagnostic_path = save_diagnostic(
            driver,
            f"debug_dropdown_{dropdown_index + 1}.html",
        )

        current_value = get_dropdown_selected_value(
            driver,
            dropdown_index,
        )

        raise RuntimeError(
            f"Could not verify {description}. "
            f"Expected {target_value!r}; visible value is "
            f"{current_value!r}. Diagnostic: {diagnostic_path}"
        )

    print(f"Selected successfully: {description} = {target_value}")


def select_target_period(driver, fy, month_name):
    """Select FY first, then re-find and select the month."""

    select_react_dropdown(
        driver=driver,
        dropdown_index=0,
        target_value=fy,
        description="financial year",
    )

    # The website may rebuild the filter controls after FY changes.
    time.sleep(2)

    WebDriverWait(driver, 20).until(
        lambda d: get_dropdown_control(d, 1) is not None
    )

    select_react_dropdown(
        driver=driver,
        dropdown_index=1,
        target_value=month_name.upper(),
        description="month",
    )

    time.sleep(3)


def set_largest_page_size(driver):
    """Set the table page size to the largest available option."""

    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    'select[aria-label="Choose a page size"]',
                )
            )
        )

        page_select = Select(element)

        numeric_options = [
            int(option.text.strip())
            for option in page_select.options
            if option.text.strip().isdigit()
        ]

        if numeric_options:
            largest = max(numeric_options)

            page_select.select_by_visible_text(
                str(largest)
            )

            time.sleep(2)
            print(f"Page size set to {largest}.")

    except TimeoutException:
        print(
            "Page-size dropdown not found; "
            "using the site's default page size."
        )


# --- COLLECT REPORT LINKS ---

def collect_links_from_current_page(driver, cal_year, cal_month):
    """Collect PDF links for the requested calendar month."""

    links = []
    seen = set()

    for anchor in driver.find_elements(By.TAG_NAME, "a"):
        try:
            href = anchor.get_attribute("href")
            link_text = anchor.text.strip()

        except StaleElementReferenceException:
            continue

        if not href:
            continue

        combined = f"{link_text} {href}".lower()

        if ".pdf" not in combined and "nldc_remc_report" not in combined:
            continue

        full_url = urljoin(driver.current_url, href)

        report_date = filename_date(full_url)

        if report_date is None:
            report_date = parse_date_text(
                f"{link_text} {href}"
            )

        if report_date is None:
            continue

        if report_date.year != cal_year or report_date.month != cal_month:
            continue

        if full_url in seen:
            continue

        seen.add(full_url)

        links.append({
            "report_date": report_date,
            "url": full_url,
            "link_text": link_text,
        })

    return links


def click_next_page(driver):
    """Click Next Page if another results page exists."""

    try:
        button = driver.find_element(
            By.CSS_SELECTOR,
            'button[aria-label="Next Page"]',
        )

        if (
            not button.is_displayed()
            or not button.is_enabled()
            or button.get_attribute("disabled") is not None
            or button.get_attribute("aria-disabled") == "true"
        ):
            return False

        try:
            old_page = driver.find_element(
                By.CSS_SELECTOR,
                'input[aria-label="page counter"]',
            ).get_attribute("value")
        except Exception:
            old_page = None

        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            button,
        )

        try:
            button.click()
        except Exception:
            driver.execute_script(
                "arguments[0].click();",
                button,
            )

        if old_page is not None:
            try:
                WebDriverWait(driver, 8).until(
                    lambda d: d.find_element(
                        By.CSS_SELECTOR,
                        'input[aria-label="page counter"]',
                    ).get_attribute("value") != old_page
                )
            except TimeoutException:
                time.sleep(2)
        else:
            time.sleep(2)

        return True

    except Exception:
        return False


def collect_report_links(driver, cal_year, cal_month):
    month_name = date(cal_year, cal_month, 1).strftime("%B")
    fy = financial_year(cal_year, cal_month)

    driver.get(BASE_URL)

    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    time.sleep(5)

    select_target_period(driver, fy, month_name)
    set_largest_page_size(driver)

    links_by_date = {}

    for page_number in range(1, 101):
        page_links = collect_links_from_current_page(driver, cal_year, cal_month)

        for item in page_links:
            links_by_date.setdefault(item["report_date"], item)

        print(f"  Page {page_number}: {len(page_links)} links; "
              f"{len(links_by_date)} unique dates collected.")

        if not click_next_page(driver):
            break

    links = sorted(links_by_date.values(), key=lambda item: item["report_date"])

    if not links:
        diagnostic_path = save_diagnostic(driver, "debug_report_page.html")
        raise RuntimeError(
            f"No report links found for {month_name} {cal_year}. "
            f"Inspect {diagnostic_path}"
        )
    return links


# --- DOWNLOADER ---

def _is_valid_pdf(path: Path) -> bool:
    if not (path.exists() and path.stat().st_size > 1000):
        return False
    with path.open("rb") as f:
        return f.read(5) == b"%PDF-"


def download_pdf(url, pdf_path: Path):
    """Download a PDF unless a valid cached copy exists. Returns 'cached' or 'downloaded'."""
    if _is_valid_pdf(pdf_path):
        return "cached"

    response = SESSION.get(url, timeout=60, verify=False)
    response.raise_for_status()

    if not response.content.startswith(b"%PDF-"):
        raise ValueError(
            "Response is not a PDF. "
            f"HTTP status={response.status_code}; "
            f"Content-Type={response.headers.get('Content-Type')}; "
            f"URL={response.url}"
        )
    pdf_path.write_bytes(response.content)
    return "downloaded"


def fetch_vre_month(cal_year: int, cal_month: int):
    """Open the Grid-India VRE page in Chrome, list the month's reports and
    download any PDFs not already on disk."""
    month_name = date(cal_year, cal_month, 1).strftime("%B")
    days_in_month = expected_days(cal_year, cal_month)
    folder = vre_month_folder(cal_year, cal_month)
    folder.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  Fetching Daily VRE (NLDC REMC) Report: {month_name} {cal_year}  |  "
          f"FY {financial_year(cal_year, cal_month)}")
    print(f"  Expected days in month: {days_in_month}")
    print(f"  Save folder: {folder}")
    print(f"{'='*60}")

    driver = setup_driver()
    try:
        links = collect_report_links(driver, cal_year, cal_month)
    finally:
        driver.quit()

    print(f"  {len(links)} report(s) listed for {month_name} {cal_year}")

    downloaded = skipped = failed = 0
    saved_records = []

    for item in links:
        d = item["report_date"]
        filename = f"{d:%Y-%m-%d}_NLDC_REMC_REPORT.pdf"
        save_path = folder / filename

        try:
            outcome = download_pdf(item["url"], save_path)
        except Exception as e:
            print(f"  [ERR]  {filename}: {e}")
            failed += 1
            continue

        if outcome == "cached":
            print(f"  [SKIP] {filename}")
            skipped += 1
        else:
            print(f"  [DOWN] {filename}")
            downloaded += 1
            time.sleep(1)

        saved_records.append({
            "report_date": d.isoformat(),
            "url": item["url"],
            "filename": unquote(Path(urlparse(item["url"]).path).name),
            "local_file": str(save_path),
        })

    total = downloaded + skipped
    print(f"\n  Downloaded : {downloaded}")
    print(f"  Skipped    : {skipped}  (already existed)")
    print(f"  Failed     : {failed}")
    print(f"  Total PDFs : {total}  |  Expected: {days_in_month}")
    if total < days_in_month:
        print(f"  WARNING: {days_in_month - total} day(s) missing - may not yet be published.")
    else:
        print(f"  OK: All {days_in_month} daily reports accounted for.")

    if saved_records:
        VRE_METADATA_FILE.parent.mkdir(parents=True, exist_ok=True)
        df_new = pd.DataFrame(saved_records)
        if VRE_METADATA_FILE.exists():
            df_new = pd.concat([pd.read_csv(VRE_METADATA_FILE), df_new],
                               ignore_index=True).drop_duplicates(subset=["local_file"], keep="last")
        df_new.to_csv(VRE_METADATA_FILE, index=False)

    return saved_records


# --- PDF PARSING ---

def extract_report_date(pdf_bytes):
    """Read the first page and extract the observation date."""

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        text = (
            pdf.pages[0].extract_text()
            if pdf.pages
            else ""
        ) or ""

    patterns = [
        r"Report\s+for\s*:?\s*"
        r"(\d{1,2}[-./ ]+[A-Za-z]{3,9}[-./ ]+\d{2,4})",

        r"Report\s+for\s*:?\s*"
        r"(\d{1,2}[-./]\d{1,2}[-./]\d{2,4})",

        r"Report\s+for\s*:?\s*"
        r"(\d{4}-\d{2}-\d{2})",
    ]

    formats = [
        "%d-%b-%y",
        "%d-%B-%Y",
        "%d-%b-%Y",
        "%d-%B-%y",
        "%d/%b/%y",
        "%d/%B/%Y",
        "%d.%b.%y",
        "%d.%B.%Y",
        "%d-%m-%Y",
        "%d-%m-%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d.%m.%Y",
        "%d.%m.%y",
        "%d %b %Y",
        "%d %B %Y",
        "%Y-%m-%d",
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            re.IGNORECASE,
        )

        if not match:
            continue

        value = clean_text(match.group(1))

        for fmt in formats:
            try:
                return datetime.strptime(value, fmt).date(), text
            except ValueError:
                continue

    return None, text


def extract_table_8_and_9(pdf_bytes):
    """
    Table 8 Total Curtailment:
        Wind MU  = row[4]
        Solar MU = row[6]

    Table 9 SUM:
        Wind MU  = row[3]
        Solar MU = row[5]
    """

    result = {
        "wind_curtailment_mu": None,
        "solar_curtailment_mu": None,
        "tras_down_wind_mu": None,
        "tras_down_solar_mu": None,
    }

    table8_candidates = []
    table9_candidates = []

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        if not pdf.pages:
            raise ValueError("PDF contains no pages.")

        last_page = pdf.pages[-1]
        tables = last_page.extract_tables() or []

        if not tables:
            raise ValueError("No tables found on the final page.")

        for table in tables:
            for row in table or []:
                if not row:
                    continue

                row_text = " ".join(
                    clean_text(cell)
                    for cell in row
                    if cell is not None
                )

                row_lower = row_text.lower()

                if "total curtailment" in row_lower:
                    table8_candidates.append(row)

                    if len(row) >= 7:
                        wind_value = numeric_value(row[4])
                        solar_value = numeric_value(row[6])

                        if wind_value is not None:
                            result["wind_curtailment_mu"] = wind_value

                        if solar_value is not None:
                            result["solar_curtailment_mu"] = solar_value

                sum_index = next(
                    (
                        index
                        for index, cell in enumerate(row)
                        if clean_text(cell).upper() == "SUM"
                    ),
                    None,
                )

                if sum_index is not None:
                    table9_candidates.append(row)

                    if sum_index == 1 and len(row) >= 6:
                        wind_value = numeric_value(row[3])
                        solar_value = numeric_value(row[5])

                        if wind_value is not None:
                            result["tras_down_wind_mu"] = wind_value

                        if solar_value is not None:
                            result["tras_down_solar_mu"] = solar_value

    missing = []

    if result["wind_curtailment_mu"] is None:
        missing.append("Table 8 Wind")

    if result["solar_curtailment_mu"] is None:
        missing.append("Table 8 Solar")

    if result["tras_down_wind_mu"] is None:
        missing.append("Table 9 Wind")

    if result["tras_down_solar_mu"] is None:
        missing.append("Table 9 Solar")

    if missing:
        raise ValueError(
            f"Could not extract {', '.join(missing)}. "
            f"Table 8 candidate rows: {table8_candidates[:5]!r}. "
            f"Table 9 candidate rows: {table9_candidates[:5]!r}"
        )

    return result


# --- MONTHLY EXCEL BUILDER ---

def _url_lookup():
    """local PDF file name -> (original file name, URL), from the fetch metadata."""
    if not VRE_METADATA_FILE.exists():
        return {}
    meta = pd.read_csv(VRE_METADATA_FILE)
    return {Path(r.local_file).name: (r.filename, r.url) for r in meta.itertuples()}


def process_reports(pdf_files, month_start, month_end):
    """Parse each local PDF. Returns (valid_records, audit_rows)."""
    lookup = _url_lookup()
    valid_records, audit_rows = [], []

    for i, pdf_path in enumerate(pdf_files, start=1):
        m = PDF_NAME.match(pdf_path.name)
        fname_date = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        original_name, url = lookup.get(pdf_path.name, (pdf_path.name, None))

        audit = {
            "Filename Date": fname_date,
            "Report for Date": None,
            "Filename": original_name,
            "PDF URL": url,
            "Wind Curtailment (MU)": None,
            "Solar Curtailment (MU)": None,
            "TRAS Down Wind (MU)": None,
            "TRAS Down Solar (MU)": None,
            "Status": None,
            "Details": None,
        }

        print(f"[{i}/{len(pdf_files)}] {pdf_path.name}")

        try:
            pdf_bytes = pdf_path.read_bytes()
            report_date, first_page_text = extract_report_date(pdf_bytes)
            audit["Report for Date"] = report_date

            if report_date is None:
                audit["Status"] = "INVALID - REPORT DATE NOT FOUND"
                audit["Details"] = first_page_text[:500]
                audit_rows.append(audit)
                print("  ERROR: Report date not found.")
                continue

            if report_date != fname_date:
                audit["Status"] = "REJECTED - DATE MISMATCH"
                audit["Details"] = (f"Filename date {fname_date} differs from "
                                    f"internal Report for date {report_date}.")
                audit_rows.append(audit)
                print(f"  Rejected: filename date {fname_date} != report date {report_date}")
                continue

            if not month_start <= report_date <= month_end:
                audit["Status"] = "SKIPPED - OUTSIDE MONTH"
                audit["Details"] = "Outside requested month."
                audit_rows.append(audit)
                continue

            values = extract_table_8_and_9(pdf_bytes)
            for col, key in VALUE_KEYS:
                audit[col] = values[key]

            audit["Status"] = "VALID"
            audit["Details"] = "Dates match; all four values extracted."
            valid_records.append({
                "Report for Date": report_date,
                **{col: values[key] for col, key in VALUE_KEYS},
                "PDF URL": url,
                "PDF File": pdf_path.name,
            })
            audit_rows.append(audit)
            print(f"  OK: {report_date}  {values}")

        except Exception as exc:
            audit["Status"] = "ERROR"
            audit["Details"] = str(exc)
            audit_rows.append(audit)
            print(f"  ERROR: {exc}")

    return valid_records, audit_rows


def build_daily_dataset(valid_records, month_start, month_end):
    """One row per calendar day, with an explicit Status for days without a clean report."""
    by_date = {}
    for record in valid_records:
        by_date.setdefault(record["Report for Date"], []).append(record)

    empty = {col: None for col, _ in VALUE_KEYS}
    rows = []
    for timestamp in pd.date_range(month_start, month_end, freq="D"):
        current_date = timestamp.date()
        matches = by_date.get(current_date, [])

        if len(matches) == 1:
            r = matches[0]
            rows.append({"Report for Date": current_date,
                         **{col: r[col] for col, _ in VALUE_KEYS},
                         "PDF URL": r["PDF URL"], "PDF File": r["PDF File"], "Status": "OK"})
        elif len(matches) > 1:
            rows.append({"Report for Date": current_date, **empty, "PDF URL": None,
                         "PDF File": None, "Status": "DUPLICATE VALID REPORTS - REVIEW AUDIT"})
        else:
            rows.append({"Report for Date": current_date, **empty, "PDF URL": None,
                         "PDF File": None, "Status": "NO VALID REPORT FOUND"})
    return pd.DataFrame(rows)


def export_results(daily_df, audit_df, out_path: Path):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        daily_df.to_excel(writer, sheet_name="Daily Data", index=False)
        audit_df.to_excel(writer, sheet_name="Report Audit", index=False)

        for sheet_name in ["Daily Data", "Report Audit"]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column_cells in ws.columns:
                max_length = max(len(str(c.value or "")) for c in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 65)

        for row in writer.sheets["Daily Data"].iter_rows(min_row=2, max_col=1):
            row[0].number_format = "dd-mm-yyyy"


def build_vre_month_excel(cal_year: int, cal_month: int, out_path=None):
    month_name = date(cal_year, cal_month, 1).strftime("%B")
    folder = vre_month_folder(cal_year, cal_month)

    if not folder.exists():
        print(f"No folder found at {folder} - run fetch_vre_month({cal_year}, {cal_month}) first.")
        return None

    pdf_files = sorted(f for f in folder.glob("*.pdf") if PDF_NAME.match(f.name))
    if not pdf_files:
        print(f"No VRE PDF files found in {folder}.")
        return None

    month_start = date(cal_year, cal_month, 1)
    month_end = date(cal_year, cal_month, expected_days(cal_year, cal_month))

    print(f"\nFound {len(pdf_files)} file(s) in {folder}. Processing...")
    valid_records, audit_rows = process_reports(pdf_files, month_start, month_end)

    daily_df = build_daily_dataset(valid_records, month_start, month_end)
    audit_df = pd.DataFrame(audit_rows).sort_values(
        ["Report for Date", "Filename Date"], na_position="last")

    out_path = Path(out_path) if out_path else vre_panel_path(cal_year, cal_month)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    export_results(daily_df, audit_df, out_path)

    n_ok = int((daily_df["Status"] == "OK").sum())
    n_missing = int((daily_df["Status"] == "NO VALID REPORT FOUND").sum())
    n_dup = int(daily_df["Status"].str.contains("DUPLICATE", na=False).sum())
    n_err = int(audit_df["Status"].eq("ERROR").sum())
    n_mis = int(audit_df["Status"].eq("REJECTED - DATE MISMATCH").sum())

    print(f"\n{'='*60}")
    print(f"  Month     : {month_name} {cal_year}  ({len(daily_df)} days)")
    print(f"  Extracted : {n_ok} days")
    if n_missing:
        missing_days = daily_df.loc[daily_df["Status"] == "NO VALID REPORT FOUND",
                                    "Report for Date"].map(lambda d: d.day).tolist()
        print(f"  Missing   : {n_missing} day(s) -> {missing_days}")
    else:
        print(f"  Complete  : all {len(daily_df)} days present")
    if n_dup:
        print(f"  Duplicate valid reports: {n_dup} day(s) - see 'Report Audit'")
    if n_mis:
        print(f"  Date mismatches rejected: {n_mis} - see 'Report Audit'")
    if n_err:
        print(f"  Extraction errors: {n_err} - see 'Report Audit'")
    print(f"  Saved     : {out_path}")
    print(f"{'='*60}")

    return daily_df


# --- ORCHESTRATOR ---

def download_and_build_vre_month(cal_year: int, cal_month: int):
    fetch_vre_month(cal_year, cal_month)
    return build_vre_month_excel(cal_year, cal_month)
