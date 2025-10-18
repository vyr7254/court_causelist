import os
import time
import requests
import re
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import PyPDF2

# === CONFIGURATION ===
OUTPUT_FOLDER = r"C:\Users\Dell\OneDrive\Desktop\jshc_code\jhc_causelists"
EXCEL_OUTPUT = os.path.join(OUTPUT_FOLDER, "JHC_Complete_Cases.xlsx")
LOG_FILE = os.path.join(OUTPUT_FOLDER, "scraper_log.txt")
CAUSELIST_URL = "https://jharkhandhighcourt.nic.in/entire-cause-list.php"

# Setup logging
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# === CHROME SETUP ===
def setup_driver():
    """Setup Chrome driver with download preferences"""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    # Create unique user data directory to avoid conflicts
    import tempfile
    temp_dir = tempfile.mkdtemp()
    chrome_options.add_argument(f"--user-data-dir={temp_dir}")
    
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    prefs = {
        "download.default_directory": OUTPUT_FOLDER,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    return driver

def wait_for_download(download_folder, timeout=30):
    """Wait for download to complete"""
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        downloading = False
        for filename in os.listdir(download_folder):
            if filename.endswith('.crdownload'):
                downloading = True
                break
        if not downloading:
            return True
        seconds += 1
    return False

def get_daily_causelist_links(driver):
    """Get all DAILY CAUSELIST links from the ARCHIVES section"""
    try:
        driver.get(CAUSELIST_URL)
        time.sleep(3)
        
        logging.info("Looking for ARCHIVES section...")
        archives_section = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'ARCHIVES')]"))
        )
        logging.info("Found ARCHIVES section")
        
        all_archive_links = driver.find_elements(
            By.XPATH, 
            "//h3[contains(text(), 'ARCHIVES')]/following-sibling::*//a"
        )
        
        if not all_archive_links:
            all_archive_links = driver.find_elements(By.TAG_NAME, "a")
        
        links_info = []
        for link in all_archive_links:
            link_text = link.text.strip()
            link_href = link.get_attribute("href")
            
            if "ARCHIVES BEFORE" in link_text.upper() or "01.01.2025" in link_text:
                logging.info(f"Reached stop condition: {link_text}")
                break
            
            if link_text.startswith("DAILY CAUSELIST") and link_href:
                links_info.append({
                    'text': link_text,
                    'href': link_href
                })
                
                if "2nd JANUARY 2025" in link_text or "2ND JANUARY 2025" in link_text:
                    logging.info(f"Reached final target: {link_text}")
                    break
        
        logging.info(f"Found {len(links_info)} DAILY CAUSELIST links")
        return links_info
    
    except Exception as e:
        logging.error(f"Error getting causelist links: {e}")
        return []

def download_pdf_from_viewer(driver, link_info):
    """Navigate to PDF viewer URL and download the PDF"""
    try:
        link_text = link_info['text']
        link_href = link_info['href']
        logging.info(f"Processing: {link_text}")
        
        driver.get(link_href)
        time.sleep(5)
        
        # Check for iframes
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        
        for iframe in iframes:
            try:
                driver.switch_to.frame(iframe)
                time.sleep(2)
                
                open_elements = driver.find_elements(By.XPATH, 
                    "//*[self::button or self::a][contains(translate(text(), 'OPEN', 'open'), 'open')]")
                
                if open_elements:
                    for elem in open_elements:
                        if elem.is_displayed():
                            driver.execute_script("arguments[0].click();", elem)
                            driver.switch_to.default_content()
                            
                            if wait_for_download(OUTPUT_FOLDER, timeout=20):
                                logging.info(f"SUCCESS: Downloaded {link_text}")
                                return True
                
                driver.switch_to.default_content()
            except:
                driver.switch_to.default_content()
                continue
        
        return False
    except Exception as e:
        logging.error(f"Error processing '{link_info['text']}': {e}")
        return False

# === PDF EXTRACTION FUNCTIONS ===
def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            return text
    except Exception as e:
        logging.error(f"Error extracting PDF: {e}")
        return None

# ==========================
# UPDATED FUNCTIONS START
# ==========================
def parse_causelist_data(pdf_text, pdf_filename):
    """Parse causelist and extract case information accurately with time and court mapping"""
    cases = []
    lines = [ln.strip() for ln in pdf_text.splitlines() if ln.strip()]

    court_no = "N/A"
    cause_date = "N/A"
    chief_justice = "N/A"
    time_val = "N/A"
    bench_name = "Jharkhand"

    current_sno = None
    current_case_lines = []
    case_counter = 0

    for line in lines:
        # Detect court header section
        if re.search(r"DAILY\s+CAUSELIST\s+COURT\s+NO", line, re.IGNORECASE):
            m = re.search(r"COURT\s*NO[.\s]*(\d+)", line, re.IGNORECASE)
            if m:
                court_no = m.group(1).strip()

            m = re.search(r"FOR\s+\w+\s+THE\s+(\d{1,2})(?:ST|ND|RD|TH)?\s+(\w+)\s+(\d{4})", line, re.IGNORECASE)
            if m:
                cause_date = f"{m.group(1)} {m.group(2).title()} {m.group(3)}"

            logging.info(f"\n=== COURT NO {court_no} | DATE: {cause_date} ===")
            continue

        # Detect time (AT 10:30 AM)
        m = re.search(r"\bAT\s+(\d{1,2}:\d{2})\s*(?:AM|PM|A\.M\.|P\.M\.)?", line, re.IGNORECASE)
        if m:
            time_val = m.group(1)
            logging.info(f"⏰ Time detected: {time_val}")
            continue

        # Detect chief justice
        if re.search(r"HON'?BLE.*JUSTICE", line, re.IGNORECASE):
            chief_justice = ' '.join(line.split())
            logging.info(f"👨‍⚖️ Chief Justice: {chief_justice}")
            continue

        # Detect new case by serial number
        sno_match = re.match(r"^(\d+)\s+", line)
        if sno_match:
            # Save previous case
            if current_sno and current_case_lines:
                case_counter += 1
                process_case_block_fixed(
                    case_counter, current_sno, ' '.join(current_case_lines),
                    court_no, bench_name, chief_justice,
                    cause_date, time_val, pdf_filename, cases
                )

            # Start new case
            current_sno = sno_match.group(1)
            current_case_lines = [line[len(sno_match.group(0)):]]
            continue

        # Continue current case
        if current_sno:
            current_case_lines.append(line)

    # Process last case
    if current_sno and current_case_lines:
        case_counter += 1
        process_case_block_fixed(
            case_counter, current_sno, ' '.join(current_case_lines),
            court_no, bench_name, chief_justice,
            cause_date, time_val, pdf_filename, cases
        )

    logging.info(f"Extracted {len(cases)} cases from {pdf_filename}")
    return cases


def process_case_block_fixed(id_no, sno, case_text, court_number, bench_name,
                             chief_justice, cause_date, time_val, pdf_filename, cases):
    """Extract details from one case entry cleanly"""
    raw = ' '.join(case_text.split())

    # Extract case type, number, year
    case_type = case_number = case_year = "N/A"
    m = re.search(r'([A-Za-z\.\(\)/]+)\s*/\s*(\d+)\s*/\s*(\d{4})', raw)
    if m:
        case_type, case_number, case_year = m.group(1).strip(), m.group(2).strip(), m.group(3).strip()

    # Extract Petitioner vs Respondent
    petitioner = respondent = "N/A"
    m = re.search(r'(.+?)\s+(?:VS|V/S|Vs|vs)\s+(.+)', raw, re.IGNORECASE)
    if m:
        petitioner = m.group(1).strip()
        respondent = re.split(r'\s{2,}|IA\s+NO|SUBJECT|ACT', m.group(2), 1, re.IGNORECASE)[0].strip()

    # Extract Advocates
    petitioner_adv = respondent_adv = "N/A"
    adv_text = raw[m.end():] if m else ""
    adv_names = re.findall(
        r'\b([A-Z][A-Z\s\.]+(?:KUMAR|SINGH|PRASAD|VERMA|MISHRA|SHARMA|YADAV|MEHTA|ROY|ALAM|KHAN|PATI|DUBEY|TIWARI|HASSAN|NARAYAN))\b',
        adv_text, re.IGNORECASE)
    adv_names = [a.strip() for a in adv_names]
    if adv_names:
        petitioner_adv = adv_names[0]
        if len(adv_names) > 1:
            respondent_adv = adv_names[1]

    # Extract IA NO
    ia_no = "N/A"
    m = re.search(r'IA\s+NO\.?\s*([0-9/]+)', raw, re.IGNORECASE)
    if m:
        ia_no = m.group(1).strip()

    # Extract SUBJECT
    subject = "N/A"
    m = re.search(r'SUBJECT\s*[:-]?\s*([A-Z][A-Z\s,\.]+)', raw, re.IGNORECASE)
    if m:
        subject = m.group(1).strip()

    # Extract ACT
    act = "N/A"
    m = re.search(r'ACT\s*[:-]?\s*([A-Za-z0-9\s,\.()/-]+)', raw, re.IGNORECASE)
    if m:
        act = m.group(1).strip()

    # Append clean structured data
    case_data = {
        'id': id_no,
        'causelist_slno': sno,
        'court_hall_number': f"Court {court_number}",
        'Case_number': case_number,
        'Case_type': case_type,
        'case_year': case_year,
        'bench_name': bench_name,
        'cause_date': cause_date,
        'time': time_val,
        'chief_justice': chief_justice,
        'petitioner': petitioner,
        'respondent': respondent,
        'petitioner_advocate': petitioner_adv,
        'respondent_advocate': respondent_adv,
        'particulars': 'list downloaded',
        'Pdf_name': pdf_filename,
        'case_status': 'N/A',
        'IA_no': ia_no,
        'subject': subject,
        'act': act
    }

    cases.append(case_data)
    logging.info(f"[Court {court_number} | {time_val}] Case {sno} → {case_type}/{case_number}/{case_year}")
# ==========================
# UPDATED FUNCTIONS END
# ==========================


def save_to_excel(cases_data, excel_path):
    """Save data to Excel"""
    try:
        columns = [
            'id', 'causelist_slno', 'court_hall_number', 'Case_number', 'Case_type', 
            'case_year', 'bench_name', 'cause_date', 'time', 'chief_justice', 'petitioner', 'respondent', 
            'petitioner_advocate', 'respondent_advocate', 'particulars', 'Pdf_name',
            'case_status', 'IA_no', 'subject', 'act'
        ]
        
        if os.path.exists(excel_path):
            existing_df = pd.read_excel(excel_path)
            new_df = pd.DataFrame(cases_data)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            combined_df['id'] = range(1, len(combined_df) + 1)
            combined_df = combined_df[columns]
            combined_df.to_excel(excel_path, index=False)
            logging.info(f"Appended {len(cases_data)} cases to Excel")
        else:
            df = pd.DataFrame(cases_data)
            df = df[columns]
            df.to_excel(excel_path, index=False)
            logging.info(f"Created Excel with {len(cases_data)} cases")
        
        return True
    except Exception as e:
        logging.error(f"Error saving to Excel: {e}")
        return False

# === MAIN EXECUTION ===
def main():
    logging.info("=== JHARKHAND HIGH COURT CAUSELIST SCRAPER ===")
    
    driver = setup_driver()
    
    try:
        daily_links = get_daily_causelist_links(driver)
        
        if not daily_links:
            logging.error("No links found. Exiting.")
            return
        
        logging.info(f"Found {len(daily_links)} causelists to process")
        
        success_count = 0
        total_cases_extracted = 0
        
        for idx, link_info in enumerate(daily_links, 1):
            logging.info(f"\n{'='*60}")
            logging.info(f"Processing {idx}/{len(daily_links)}: {link_info['text']}")
            logging.info(f"{'='*60}")
            
            if download_pdf_from_viewer(driver, link_info):
                success_count += 1
                time.sleep(2)
                
                pdf_files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.pdf')]
                if pdf_files:
                    pdf_files.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_FOLDER, x)), reverse=True)
                    latest_pdf = pdf_files[0]
                    pdf_path = os.path.join(OUTPUT_FOLDER, latest_pdf)
                    
                    logging.info(f"Extracting data from: {latest_pdf}")
                    pdf_text = extract_text_from_pdf(pdf_path)
                    if pdf_text:
                        cases = parse_causelist_data(pdf_text, latest_pdf)
                        if cases:
                            logging.info(f"Extracted {len(cases)} cases from {latest_pdf}")
                            save_to_excel(cases, EXCEL_OUTPUT)
                            total_cases_extracted += len(cases)
                            logging.info(f"✓ Total cases in Excel so far: {total_cases_extracted}")
                        else:
                            logging.warning(f"No cases extracted from {latest_pdf}")
                    else:
                        logging.error(f"Could not extract text from {latest_pdf}")
                else:
                    logging.warning("No PDF file found after download")
            else:
                logging.error(f"Failed to download: {link_info['text']}")
            
            time.sleep(2)
        
        logging.info(f"\n{'='*60}")
        logging.info("=== COMPLETED ===")
        logging.info(f"{'='*60}")
        logging.info(f"PDFs downloaded: {success_count}/{len(daily_links)}")
        logging.info(f"Total cases extracted: {total_cases_extracted}")
        logging.info(f"Excel file: {EXCEL_OUTPUT}")
        
    finally:
        driver.quit()
        logging.info("Browser closed.")

if __name__ == "__main__":
    main()

