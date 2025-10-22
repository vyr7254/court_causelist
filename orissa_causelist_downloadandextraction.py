import os
import re
import time
import logging
import tempfile
import pandas as pd
import PyPDF2
from datetime import datetime, timedelta
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# === CONFIGURATION ===
OUTPUT_FOLDER = r"C:\Users\Dell\OneDrive\Desktop\orissa_causelists"
LOG_FILE = os.path.join(OUTPUT_FOLDER, "orissa_download_log.txt")
EXCEL_FILE = os.path.join(OUTPUT_FOLDER, "orissa_causelists_data.xlsx")
CAUSELIST_URL = "https://hcservices.ecourts.gov.in/ecourtindiaHC/cases/highcourt_causelist.php?state_cd=11&dist_cd=1&court_code=1&stateNm=Odisha"

# Date range configuration
START_DATE = datetime(2025, 9, 1)
END_DATE = datetime(2025, 10, 30)

# === LOGGING SETUP ===
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# === CHROME DRIVER SETUP ===
def setup_driver():
    """Configure Chrome driver with download preferences."""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    temp_dir = tempfile.mkdtemp()
    chrome_options.add_argument(f"--user-data-dir={temp_dir}")
    
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    prefs = {
        "download.default_directory": OUTPUT_FOLDER,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "plugins.plugins_disabled": ["Chrome PDF Viewer"],
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    return driver


def wait_for_download(download_folder, timeout=60):
    """Wait for download to complete."""
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        files = os.listdir(download_folder)
        if not any(f.endswith('.crdownload') or f.endswith('.tmp') for f in files):
            time.sleep(2)
            return True
        seconds += 1
    logging.warning(f"Download timeout after {timeout} seconds")
    return False


def get_latest_pdf(folder):
    """Get the most recently downloaded PDF."""
    pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
    if not pdf_files:
        return None
    pdf_files.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)), reverse=True)
    return pdf_files[0]


def select_date_in_picker(driver, target_date):
    """Select a specific date in the date picker."""
    try:
        wait = WebDriverWait(driver, 15)
        date_str = target_date.strftime("%d-%m-%Y")
        
        selectors = [
            (By.ID, "date_in_01"),
            (By.NAME, "date_in_01"),
            (By.XPATH, "//input[@type='text' and contains(@placeholder, 'date')]"),
            (By.XPATH, "//input[@type='text' and @id='date_in_01']"),
            (By.XPATH, "//input[contains(@class, 'date')]"),
            (By.CSS_SELECTOR, "input[type='text'][id='date_in_01']")
        ]
        
        date_input = None
        for by_type, selector in selectors:
            try:
                date_input = wait.until(EC.element_to_be_clickable((by_type, selector)))
                logging.info(f"Found date input using: {by_type} = {selector}")
                break
            except:
                continue
        
        if not date_input:
            logging.error("Could not find date input field")
            return False
        
        date_input.click()
        time.sleep(0.5)
        date_input.clear()
        time.sleep(0.5)
        driver.execute_script("arguments[0].value = arguments[1];", date_input, date_str)
        time.sleep(0.5)
        date_input.send_keys(date_str)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", date_input)
        
        logging.info(f"✅ Selected date: {date_str}")
        time.sleep(1)
        return True
        
    except Exception as e:
        logging.error(f"Error selecting date: {e}")
        return False


def click_go_button(driver):
    """Click the GO button."""
    try:
        wait = WebDriverWait(driver, 10)
        selectors = [
            (By.XPATH, "//input[@value='GO']"),
            (By.XPATH, "//input[@value='Go']"),
            (By.XPATH, "//button[contains(text(), 'GO')]"),
            (By.XPATH, "//input[@type='submit' and contains(@value, 'GO')]"),
            (By.CSS_SELECTOR, "input[value='GO']")
        ]
        
        go_button = None
        for by_type, selector in selectors:
            try:
                go_button = wait.until(EC.element_to_be_clickable((by_type, selector)))
                break
            except:
                continue
        
        if not go_button:
            logging.error("Could not find GO button")
            return False
        
        go_button.click()
        logging.info("✅ Clicked GO button")
        time.sleep(3)
        return True
        
    except Exception as e:
        logging.error(f"Error clicking GO button: {e}")
        return False


def get_causelist_table_rows(driver):
    """Extract all rows from the causelist table."""
    try:
        wait = WebDriverWait(driver, 10)
        
        # Wait for table to load
        table = wait.until(
            EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'table') or .//th[contains(text(), 'Bench')] or .//th[contains(text(), 'Sr No')]]"))
        )
        
        # Get all rows - try different approaches
        try:
            # Method 1: Get tbody rows
            tbody = table.find_element(By.TAG_NAME, "tbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")
        except:
            # Method 2: Get all tr except first (header)
            all_rows = table.find_elements(By.TAG_NAME, "tr")
            rows = all_rows[1:] if len(all_rows) > 1 else all_rows
        
        logging.info(f"Found {len(rows)} causelist entries in table")
        
        # Log first few rows for debugging
        for i, row in enumerate(rows[:3], start=1):
            cells = row.find_elements(By.TAG_NAME, "td")
            if cells:
                logging.info(f"  Row {i}: {len(cells)} columns - First cell: '{cells[0].text.strip()}'")
        
        return rows
        
    except TimeoutException:
        logging.warning("No causelist table found for this date")
        return []
    except Exception as e:
        logging.error(f"Error getting table rows: {e}")
        return []


def download_causelist_pdf(driver, row, sr_no, current_date):
    """Download PDF for a specific causelist row.

    **CHANGES MADE HERE**
    - Always returns two values: (pdf_filename, bench_name) or (None, bench_name)/(None, None)
    - Sanitizes new filename to remove illegal Windows characters/newlines to avoid WinError 123.
    """
    try:
        # Get all cells in the row
        cells = row.find_elements(By.TAG_NAME, "td")
        
        if len(cells) < 3:
            logging.warning(f"  Sr No {sr_no}: Row has insufficient columns ({len(cells)})")
            # return two values so unpacking doesn't fail
            return None, None
        
        # Extract information from cells
        # Typically: [Sr No, Bench, Causelist Type, View Causelist]
        sr_no_text = cells[0].text.strip()
        bench_name = cells[1].text.strip() if len(cells) > 1 else "Unknown"
        causelist_type = cells[2].text.strip() if len(cells) > 2 else "Unknown"
        
        logging.info(f"  Sr No {sr_no_text}: Bench - {bench_name}, Type - {causelist_type}")
        
        # Find the View link - try multiple approaches
        view_link = None
        try:
            # Method 1: Find by link text in the last cell
            view_link = cells[-1].find_element(By.LINK_TEXT, "View")
        except:
            try:
                # Method 2: Find by partial link text
                view_link = cells[-1].find_element(By.PARTIAL_LINK_TEXT, "View")
            except:
                try:
                    # Method 3: Find any anchor tag in last cell
                    view_link = cells[-1].find_element(By.TAG_NAME, "a")
                except:
                    logging.warning(f"    ⚠️ Could not find View link for Sr No {sr_no_text}")
                    return None, bench_name
        
        if not view_link:
            logging.warning(f"    ⚠️ No View link found for Sr No {sr_no_text}")
            return None, bench_name
        
        # Store current window handle
        main_window = driver.current_window_handle
        
        # Click View link
        view_link.click()
        time.sleep(3)
        
        # Check if new window/tab opened
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(2)
            
            # Try to find and click download button
            try:
                # Try multiple selectors for download button
                download_selectors = [
                    (By.XPATH, "//button[contains(@title, 'Download')]"),
                    (By.XPATH, "//button[contains(@class, 'download')]"),
                    (By.XPATH, "//a[contains(@title, 'Download')]"),
                    (By.XPATH, "//button[contains(text(), 'Download')]"),
                    (By.ID, "download"),
                    (By.CSS_SELECTOR, "button[title*='Download']")
                ]
                
                download_btn = None
                for by_type, selector in download_selectors:
                    try:
                        download_btn = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((by_type, selector))
                        )
                        break
                    except:
                        continue
                
                if download_btn:
                    download_btn.click()
                    logging.info(f"    ✅ Clicked download button for Sr No {sr_no_text}")
                else:
                    logging.info(f"    📄 PDF opened (auto-download expected) for Sr No {sr_no_text}")
                
            except TimeoutException:
                logging.info(f"    📄 PDF auto-downloading for Sr No {sr_no_text}")
            
            # Wait for download to complete
            if wait_for_download(OUTPUT_FOLDER, timeout=40):
                latest_pdf = get_latest_pdf(OUTPUT_FOLDER)
                if latest_pdf:
                    # Create new filename with proper format
                    date_str = current_date.strftime("%d-%m-%Y")  # DD-MM-YYYY format
                    new_name = f"causelist_{date_str}_{sr_no_text}.pdf"
                    
                    # sanitize filename to avoid invalid characters/newlines
                    # remove characters: \ / : * ? " < > | and newline/carriage returns
                    safe_name = re.sub(r'[\\/:*?"<>|\r\n]+', '_', new_name).strip()
                    
                    old_path = os.path.join(OUTPUT_FOLDER, latest_pdf)
                    new_path = os.path.join(OUTPUT_FOLDER, safe_name)
                    
                    # Check if file already exists
                    if os.path.exists(new_path):
                        logging.info(f"    ⚠️ PDF already exists: {safe_name}")
                        # Delete the duplicate download
                        try:
                            os.remove(old_path)
                        except:
                            pass
                    else:
                        # Rename the file
                        try:
                            os.rename(old_path, new_path)
                            logging.info(f"    ✅ Downloaded: {safe_name}")
                            new_name = safe_name
                        except Exception as e:
                            logging.error(f"    ❌ Error renaming file: {e}")
                            # fallback to existing filename
                            new_name = latest_pdf
                    
                    # Close the PDF tab and switch back
                    driver.close()
                    driver.switch_to.window(main_window)
                    time.sleep(1)
                    # RETURN TWO VALUES (fixed)
                    return new_name, bench_name
            
            # Close the PDF tab and switch back
            driver.close()
            driver.switch_to.window(main_window)
            time.sleep(1)
        else:
            # PDF might have downloaded directly without opening new window
            if wait_for_download(OUTPUT_FOLDER, timeout=30):
                latest_pdf = get_latest_pdf(OUTPUT_FOLDER)
                if latest_pdf:
                    date_str = current_date.strftime("%d-%m-%Y")
                    new_name = f"causelist_{date_str}_{sr_no_text}.pdf"
                    
                    # sanitize filename
                    safe_name = re.sub(r'[\\/:*?"<>|\r\n]+', '_', new_name).strip()
                    
                    old_path = os.path.join(OUTPUT_FOLDER, latest_pdf)
                    new_path = os.path.join(OUTPUT_FOLDER, safe_name)
                    
                    if not os.path.exists(new_path):
                        try:
                            os.rename(old_path, new_path)
                            logging.info(f"    ✅ Downloaded: {safe_name}")
                            # RETURN TWO VALUES (fixed)
                            return safe_name, bench_name
                        except Exception as e:
                            logging.error(f"    ❌ Error renaming file: {e}")
                            # fallback return existing name
                            return latest_pdf, bench_name
        
        # If nothing downloaded
        return None, bench_name
        
    except Exception as e:
        logging.error(f"    ❌ Error downloading Sr No {sr_no}: {e}")
        # Ensure we're back on main window
        try:
            if len(driver.window_handles) > 1:
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
        except:
            pass
        # RETURN TWO VALUES (fixed)
        return None, None

# === PDF TEXT EXTRACTION WITH LAYOUT PRESERVATION ===
def extract_text_from_pdf_with_layout(pdf_path):
    """Extract text from PDF maintaining layout structure."""
    try:
        pages_text = []
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page_num, page in enumerate(reader.pages):
                page_text = page.extract_text()
                if page_text:
                    pages_text.append(page_text)
        return pages_text
    except Exception as e:
        logging.error(f"Error extracting text from {pdf_path}: {e}")
        return []


# === ENHANCED PDF PARSING WITH COLUMN-BASED EXTRACTION ===
def extract_court_hall_and_justice_and_time(first_page_text):
    """Extract court hall, chief justice, and time from PDF header."""
    lines = first_page_text.split('\n')
    court_hall = "N/A"
    chief_justice_lines = []
    hearing_time = "N/A"
    
    for i, line in enumerate(lines[:40]):
        line_upper = line.upper().strip()
        
        # Extract court hall number
        if "COURT NO" in line_upper and "FLOOR" in line_upper:
            court_hall = line.strip()
        elif "CHIEF JUSTICE'S COURT" in line_upper and "FLOOR" in line_upper:
            court_hall = line.strip()
        
        # Extract time
        if re.search(r'\d{1,2}:\d{2}\s*(?:AM|PM)', line, re.IGNORECASE):
            hearing_time = line.strip()
        
        # Extract chief justice
        if "THE HON'BLE" in line_upper:
            chief_justice_lines.append(line.strip())
            # Check next few lines for additional justice names
            for j in range(i+1, min(i+3, len(lines))):
                next_line = lines[j].strip()
                next_upper = next_line.upper()
                if "JUSTICE" in next_upper and "HON'BLE" not in next_upper:
                    chief_justice_lines.append(next_line)
                elif next_line and not any(kw in next_upper for kw in ["HYBRID", "ARRANGEMENT", "FLOOR", "FUNCTION", "THROUGH"]):
                    break
    
    chief_justice = " ".join(chief_justice_lines) if chief_justice_lines else "N/A"
    
    return court_hall, chief_justice, hearing_time


def parse_case_identifier(case_id_text):
    """Parse case type, number, year, and IA number."""
    case_type = "N/A"
    case_number = "N/A"
    case_year = "N/A"
    ia_no = "N/A"
    
    # Pattern: WP(C)/31043/2024 or RVWPET/193/2025 or CONTC/7667/2024
    match = re.search(r'([A-Z]+(?:\([A-Z]+\))?)/(\d+)/(\d{4})', case_id_text)
    if match:
        case_type = match.group(1)
        case_number = match.group(2)
        case_year = match.group(3)
    
    # Extract IA number: IA No.328/2025
    ia_match = re.search(r'IA\s*No\.?(\d+/\d{4})', case_id_text, re.IGNORECASE)
    if ia_match:
        ia_no = ia_match.group(1)
    
    return case_type, case_number, case_year, ia_no


def parse_orissa_causelist_structured(pdf_path, pdf_filename, cause_date, bench_name_from_table):
    """Parse Orissa High Court causelist with precise column-based extraction."""
    cases = []
    
    try:
        pages_text = extract_text_from_pdf_with_layout(pdf_path)
        if not pages_text:
            logging.warning(f"No text extracted from {pdf_filename}")
            return cases
        
        # Extract header information from first page
        first_page = pages_text[0]
        court_hall, chief_justice, hearing_time = extract_court_hall_and_justice_and_time(first_page)
        
        # Process all pages
        full_text = "\n".join(pages_text)
        lines = [line for line in full_text.split('\n') if line.strip()]
        
        # Find start of case listing (skip header)
        case_start_idx = 0
        for i, line in enumerate(lines):
            # Look for first case number pattern: "1)"
            if re.match(r'^\s*1\)', line.strip()):
                case_start_idx = i
                break
        
        if case_start_idx == 0:
            logging.warning(f"Could not find case listing start in {pdf_filename}")
            return cases
        
        # Parse cases
        i = case_start_idx
        while i < len(lines):
            line = lines[i].strip()
            
            # Check if line starts with case number: "1)" "2)" etc.
            case_num_match = re.match(r'^(\d+)\)', line)
            if case_num_match:
                causelist_slno = case_num_match.group(1)
                
                # Collect lines for this case until next case number
                case_lines = []
                j = i
                while j < len(lines):
                    current_line = lines[j].strip()
                    # Stop at next case number (but not the current one)
                    if j > i and re.match(r'^\d+\)', current_line):
                        break
                    case_lines.append(current_line)
                    j += 1
                
                # Parse the case data
                case_data = parse_single_case(
                    case_lines, 
                    causelist_slno,
                    court_hall,
                    chief_justice,
                    hearing_time,
                    bench_name_from_table,
                    cause_date,
                    pdf_filename
                )
                
                if case_data:
                    cases.append(case_data)
                
                i = j  # Move to next case
            else:
                i += 1
        
        logging.info(f"✅ Extracted {len(cases)} cases from {pdf_filename}")
        
    except Exception as e:
        logging.error(f"Error parsing {pdf_filename}: {e}", exc_info=True)
    
    return cases


def parse_single_case(case_lines, causelist_slno, court_hall, chief_justice, hearing_time, 
                      bench_name, cause_date, pdf_filename):
    """Parse a single case from its text lines."""
    
    # Initialize case data
    case_data = {
        "id": None,
        "causelist_slno": causelist_slno,
        "court_hall_number": court_hall,
        "Case_number": "N/A",
        "Case_type": "N/A",
        "case_year": "N/A",
        "bench_name": bench_name if bench_name else "Orissa High Court",
        "cause_date": cause_date.strftime("%d-%m-%Y"),
        "time": hearing_time,
        "chief_justice": chief_justice,
        "petitioner": "N/A",
        "respondent": "N/A",
        "petitioner_advocate": "N/A",
        "respondent_advocate": "N/A",
        "particulars": "",
        "Pdf_name": pdf_filename,
        "case_status": "N/A",
        "IA_no": "N/A"
    }
    
    if not case_lines:
        return case_data
    
    # Combine all lines for processing
    full_case_text = "\n".join(case_lines)
    case_data["particulars"] = full_case_text[:1000]  # Store first 1000 chars
    
    # Extract case identifier from second column (usually in first or second line)
    case_id_found = False
    for line in case_lines[:3]:
        if not case_id_found:
            case_type, case_number, case_year, ia_no = parse_case_identifier(line)
            if case_type != "N/A":
                case_data["Case_type"] = case_type
                case_data["Case_number"] = case_number
                case_data["case_year"] = case_year
                case_data["IA_no"] = ia_no
                case_id_found = True
    
    # Extract petitioner and respondent (split on "Vs")
    # The text usually contains: PETITIONER_NAME Vs RESPONDENT_NAME
    petitioner_parts = []
    respondent_parts = []
    found_vs = False
    
    for line in case_lines:
        if re.search(r'\bVs\b|\bVS\b|\bvs\b', line):
            # Split on Vs
            parts = re.split(r'\s+(?:Vs|VS|vs)\s+', line, maxsplit=1)
            if len(parts) == 2:
                petitioner_parts.append(parts[0])
                respondent_parts.append(parts[1])
                found_vs = True
            else:
                if not found_vs:
                    petitioner_parts.append(line)
                else:
                    respondent_parts.append(line)
        else:
            # Lines without Vs
            if not found_vs:
                # Before Vs - likely petitioner or case details
                if not any(kw in line.upper() for kw in ["ORDER", "MENTION", "SPEC.ADJ", "DTRD"]):
                    # Skip administrative text
                    if case_data["Case_type"] in line:
                        continue  # Skip line with case identifier
                    petitioner_parts.append(line)
            else:
                # After Vs - likely respondent
                respondent_parts.append(line)
    
    # Clean and set petitioner
    petitioner_text = " ".join(petitioner_parts).strip()
    # Remove case number pattern from petitioner
    petitioner_text = re.sub(r'\d+\)\s*[A-Z]+(?:\([A-Z]+\))?/\d+/\d{4}.*?(?=\s+[A-Z/])', '', petitioner_text)
    petitioner_text = re.sub(r'^\d+\)', '', petitioner_text).strip()
    
    if petitioner_text:
        case_data["petitioner"] = petitioner_text[:200]  # Limit length
    
    # Clean and set respondent
    respondent_text = " ".join(respondent_parts).strip()
    if respondent_text:
        case_data["respondent"] = respondent_text[:200]
    
    # Extract advocates - typically names with specific patterns
    # Petitioner advocate usually appears on same line as petitioner (4th column)
    # Respondent advocate appears on same line as respondent (4th column)
    
    advocate_patterns = [
        r'(?:M/S\.?|MR\.|MS\.|SR\.\s*ADV\.?)\s*[A-Z][A-Z\s.,]+',
        r'[A-Z][A-Z\s.]+\([^)]*ADV[^)]*\)',
        r'[A-Z][A-Z\s.,]+(?:,\s*[A-Z]\.[A-Z][A-Z.]*)+',
    ]
    
    # Try to extract petitioner advocate
    for line in case_lines:
        if not found_vs or "Vs" in line:
            for pattern in advocate_patterns:
                matches = re.findall(pattern, line)
                if matches and case_data["petitioner_advocate"] == "N/A":
                    case_data["petitioner_advocate"] = matches[0].strip()
                    break
    
    # Try to extract respondent advocate
    found_vs_line = False
    for line in case_lines:
        if "Vs" in line or "VS" in line or "vs" in line:
            found_vs_line = True
        if found_vs_line:
            for pattern in advocate_patterns:
                matches = re.findall(pattern, line)
                if matches:
                    # Check if this advocate is different from petitioner advocate
                    adv_text = matches[0].strip()
                    if adv_text != case_data["petitioner_advocate"] and case_data["respondent_advocate"] == "N/A":
                        case_data["respondent_advocate"] = adv_text
                        break
    
    return case_data


# === EXCEL OPERATIONS ===
def save_to_excel(cases_data, excel_path):
    """Save or append case data to Excel file."""
    try:
        if not cases_data:
            logging.warning("No case data to save")
            return False
        
        # Define columns WITHOUT subject and act
        columns = [
            "id", "causelist_slno", "court_hall_number", "Case_number", "Case_type",
            "case_year", "bench_name", "cause_date", "time", "chief_justice",
            "petitioner", "respondent", "petitioner_advocate", "respondent_advocate",
            "particulars", "Pdf_name", "case_status", "IA_no"
        ]
        
        df_new = pd.DataFrame(cases_data)
        
        for col in columns:
            if col not in df_new.columns:
                df_new[col] = "N/A"
        
        df_new = df_new[columns]
        
        if os.path.exists(excel_path):
            df_existing = pd.read_excel(excel_path)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined["id"] = range(1, len(df_combined) + 1)
            df_combined.to_excel(excel_path, index=False)
            logging.info(f"✅ Appended {len(df_new)} cases → Total: {len(df_combined)}")
        else:
            df_new["id"] = range(1, len(df_new) + 1)
            df_new.to_excel(excel_path, index=False)
            logging.info(f"✅ Created Excel with {len(df_new)} cases")
        
        return True
        
    except Exception as e:
        logging.error(f"Error saving to Excel: {e}")
        return False


# === MAIN EXECUTION ===
def main():
    logging.info("=" * 80)
    logging.info("ORISSA HIGH COURT CAUSELIST DOWNLOADER & EXTRACTOR")
    logging.info("=" * 80)
    
    driver = setup_driver()
    total_pdfs_downloaded = 0
    total_cases_extracted = 0
    failed_downloads = []
    
    try:
        driver.get(CAUSELIST_URL)
        time.sleep(3)
        logging.info(f"Opened URL: {CAUSELIST_URL}")
        
        current_date = START_DATE
        
        while current_date <= END_DATE:
            logging.info("\n" + "=" * 80)
            logging.info(f"PROCESSING DATE: {current_date.strftime('%d-%m-%Y')}")
            logging.info("=" * 80)
            
            if not select_date_in_picker(driver, current_date):
                logging.error(f"Failed to select date: {current_date}")
                failed_downloads.append(f"{current_date.strftime('%d-%m-%Y')} - Date selection failed")
                current_date += timedelta(days=1)
                continue
            
            if not click_go_button(driver):
                logging.error("Failed to click GO button")
                failed_downloads.append(f"{current_date.strftime('%d-%m-%Y')} - GO button failed")
                current_date += timedelta(days=1)
                continue
            
            rows = get_causelist_table_rows(driver)
            
            if not rows:
                logging.warning(f"No cause lists for {current_date.strftime('%d-%m-%Y')}")
                current_date += timedelta(days=1)
                continue
            
            date_pdfs = 0
            date_cases = []
            
            for idx, row in enumerate(rows, start=1):
                pdf_filename, bench_name = download_causelist_pdf(driver, row, idx, current_date)
                
                if pdf_filename:
                    total_pdfs_downloaded += 1
                    date_pdfs += 1
                    
                    # Extract data from downloaded PDF
                    pdf_path = os.path.join(OUTPUT_FOLDER, pdf_filename)
                    if os.path.exists(pdf_path):
                        cases = parse_orissa_causelist_structured(
                            pdf_path, pdf_filename, current_date, bench_name
                        )
                        date_cases.extend(cases)
                        total_cases_extracted += len(cases)
                    else:
                        logging.warning(f"PDF file not found: {pdf_path}")
                else:
                    failed_downloads.append(f"{current_date.strftime('%d-%m-%Y')} - Sr No {idx}")
                
                time.sleep(2)
            
            # Save extracted cases to Excel
            if date_cases:
                save_to_excel(date_cases, EXCEL_FILE)
            
            logging.info(f"Downloaded {date_pdfs} PDFs, Extracted {len(date_cases)} cases")
            
            current_date += timedelta(days=1)
            time.sleep(3)
        
        logging.info("\n" + "=" * 80)
        logging.info("PROCESSING COMPLETED")
        logging.info("=" * 80)
        logging.info(f"Total PDFs Downloaded: {total_pdfs_downloaded}")
        logging.info(f"Total Cases Extracted: {total_cases_extracted}")
        logging.info(f"Failed Downloads: {len(failed_downloads)}")
        
        if failed_downloads:
            logging.info("\nFailed Downloads:")
            for fail in failed_downloads:
                logging.info(f"  ❌ {fail}")
        
        logging.info(f"\nPDFs: {OUTPUT_FOLDER}")
        logging.info(f"Excel: {EXCEL_FILE}")
        logging.info(f"Log: {LOG_FILE}")
        
    except Exception as e:
        logging.error(f"Critical error: {e}", exc_info=True)
        
    finally:
        driver.quit()
        logging.info("\nBrowser closed. Processing finished.")


if __name__ == "__main__":
    main()
