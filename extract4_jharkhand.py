import os
import time
import re
import logging
import tempfile
import pandas as pd
import PyPDF2
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# === CONFIGURATION ===
OUTPUT_FOLDER = r"C:\Users\Dell\OneDrive\Desktop\jshc_code\jhc_causelists"
EXCEL_OUTPUT = os.path.join(OUTPUT_FOLDER, "JHC_Complete_Cases.xlsx")
LOG_FILE = os.path.join(OUTPUT_FOLDER, "scraper_log.txt")
CAUSELIST_URL = "https://jharkhandhighcourt.nic.in/entire-cause-list.php"

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

# === CHROME SETUP ===
def setup_driver():
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
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )
    return driver


def wait_for_download(download_folder, timeout=40):
    """Wait until file download completes."""
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        if not any(f.endswith('.crdownload') for f in os.listdir(download_folder)):
            return True
        seconds += 1
    return False


# === FETCH LINKS ===
def get_daily_causelist_links(driver):
    """Fetch all DAILY CAUSELIST links from archive section."""
    try:
        driver.get(CAUSELIST_URL)
        time.sleep(3)
        logging.info("Fetching DAILY CAUSELIST links...")

        all_links = driver.find_elements(By.TAG_NAME, "a")
        links_info = []
        for link in all_links:
            text = link.text.strip()
            href = link.get_attribute("href")
            if text.startswith("DAILY CAUSELIST") and href and href.endswith(".pdf"):
                links_info.append({"text": text, "href": href})

        logging.info(f"Found {len(links_info)} cause list links.")
        return links_info
    except Exception as e:
        logging.error(f"Error getting cause list links: {e}")
        return []


# === DOWNLOAD PDF ===
def download_pdf(driver, link_info):
    """Download cause list PDF."""
    try:
        link_text = link_info["text"]
        pdf_url = link_info["href"]
        logging.info(f"Downloading: {link_text}")

        driver.get(pdf_url)
        if wait_for_download(OUTPUT_FOLDER, timeout=30):
            logging.info(f"✅ Download complete: {link_text}")
            return True
        else:
            logging.warning(f"⚠️ Timeout waiting for: {link_text}")
            return False
    except Exception as e:
        logging.error(f"Error downloading {link_info['text']}: {e}")
        return False


# === PDF TEXT EXTRACTION ===
def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file."""
    try:
        text = ""
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                if page.extract_text():
                    text += page.extract_text() + "\n"
        return text
    except Exception as e:
        logging.error(f"Error reading {pdf_path}: {e}")
        return ""


# === IMPROVED PARSING LOGIC ===
def parse_causelist_data(pdf_text, pdf_filename):
    """Extract clean structured case data with improved column-based extraction."""
    cases = []
    if not pdf_text.strip():
        return cases

    lines = [ln.strip() for ln in pdf_text.splitlines() if ln.strip()]
    
    # Extract header information
    court_hall = "N/A"
    bench_name = "Jharkhand"
    cause_date = "N/A"
    time_val = "N/A"
    chief_justice = "N/A"
    
    for line in lines[:20]:  # Check first 20 lines for header info
        if "COURT NO" in line.upper() or "COURT HALL" in line.upper():
            court_match = re.search(r"COURT\s*(?:NO\.?|HALL)?\s*[:.]?\s*(\d+)", line, re.IGNORECASE)
            if court_match:
                court_hall = f"Court {court_match.group(1)}"
        
        if "AT" in line and re.search(r"\d{1,2}:\d{2}", line):
            time_match = re.search(r"(\d{1,2}:\d{2}\s*(?:AM|PM)?)", line, re.IGNORECASE)
            if time_match:
                time_val = time_match.group(1)
        
        if "CHIEF JUSTICE" in line.upper() or "HON'BLE" in line.upper():
            chief_justice = line.replace("HON'BLE", "").replace("THE CHIEF JUSTICE", "").strip()
    
    # If court hall is still N/A, look for "Court 1", "Court 2" pattern in data
    if court_hall == "N/A":
        for line in lines:
            court_match = re.search(r"(\d+)\s+Court\s+(\d+)", line, re.IGNORECASE)
            if court_match:
                court_hall = f"Court {court_match.group(2)}"
                break
    
    # Parse case rows - looking for tabular data with columns
    case_pattern = re.compile(r"([A-Z]\.[A-Z]\([A-Z]+\)|[A-Z]+\.[A-Z]+\([A-Z]+\))\s*[/\-]?\s*(\d+)\s*[/\-]?\s*(\d{4})")
    
    current_row = []
    all_rows = []
    
    for i, line in enumerate(lines):
        # Check if line starts with a serial number (case row)
        if re.match(r"^\d+\s+", line):
            if current_row:
                all_rows.append(current_row)
            current_row = [line]
        elif current_row:
            current_row.append(line)
    
    if current_row:
        all_rows.append(current_row)
    
    id_counter = 0
    last_valid_court = court_hall
    
    for row in all_rows:
        full_text = " ".join(row)
        
        # Extract serial number and check for court number
        sno_match = re.match(r"^(\d+)\s+(.+)", full_text)
        if not sno_match:
            continue
        
        sno = sno_match.group(1)
        rest_text = sno_match.group(2)
        
        # Check for court number in the row
        court_in_row_match = re.search(r"Court\s+(\d+)", rest_text, re.IGNORECASE)
        if court_in_row_match:
            last_valid_court = f"Court {court_in_row_match.group(1)}"
            # Remove court info from text
            rest_text = re.sub(r"Court\s+\d+", "", rest_text, flags=re.IGNORECASE).strip()
        
        # Try to match case number pattern
        case_match = case_pattern.search(rest_text)
        if not case_match:
            logging.debug(f"Skipping row {sno} - no valid case number found")
            continue  # Skip if no case number found
        
        case_type = case_match.group(1).strip()
        case_number = case_match.group(2).strip()
        case_year = case_match.group(3).strip()
        
        # Split text into columns by looking for multiple spaces or specific patterns
        # The structure is: [Serial] [Court] [Case] [Petitioner vs Respondent] [Pet Advocate] [Resp Advocate]
        
        # Remove case number pattern from text to parse parties
        text_after_case = rest_text[case_match.end():].strip()
        
        # Extract petitioner and respondent by splitting on VS
        petitioner = "N/A"
        respondent = "N/A"
        petitioner_advocate = "N/A"
        respondent_advocate = "N/A"
        
        vs_match = re.search(r"\b(VS|V/S|V\.S\.)\b", text_after_case, re.IGNORECASE)
        if vs_match:
            # Text before VS is petitioner
            before_vs = text_after_case[:vs_match.start()].strip()
            after_vs = text_after_case[vs_match.end():].strip()
            
            # Petitioner is the first part before VS
            petitioner = before_vs
            
            # Split remaining text into columns
            # Look for multiple consecutive spaces as column delimiters
            parts = re.split(r'\s{2,}', after_vs)
            
            if len(parts) >= 1:
                respondent = parts[0].strip()
            if len(parts) >= 2:
                petitioner_advocate = parts[1].strip()
            if len(parts) >= 3:
                respondent_advocate = parts[2].strip()
        else:
            # If no VS found, try to parse columns differently
            parts = re.split(r'\s{2,}', text_after_case)
            if len(parts) >= 1:
                petitioner = parts[0].strip()
            if len(parts) >= 2:
                respondent = parts[1].strip()
            if len(parts) >= 3:
                petitioner_advocate = parts[2].strip()
            if len(parts) >= 4:
                respondent_advocate = parts[3].strip()
        
        # Extract IA, Subject, and Act information
        ia_match = re.search(r"IA\s*NO\.?\s*([0-9/]+)", full_text, re.IGNORECASE)
        subject_match = re.search(r"SUBJECT\s*[:-]?\s*([A-Z\s,]+)", full_text, re.IGNORECASE)
        act_match = re.search(r"ACT\s*[:-]?\s*(.+)", full_text, re.IGNORECASE)
        
        ia_no = ia_match.group(1).strip() if ia_match else "N/A"
        subject = subject_match.group(1).strip() if subject_match else "N/A"
        act = act_match.group(1).strip() if act_match else "N/A"
        
        id_counter += 1
        case_data = {
            "id": id_counter,
            "causelist_slno": sno,
            "court_hall_number": last_valid_court,
            "Case_number": case_number,
            "Case_type": case_type,
            "case_year": case_year,
            "bench_name": bench_name,
            "cause_date": cause_date,
            "time": time_val,
            "chief_justice": chief_justice,
            "petitioner": petitioner,
            "respondent": respondent,
            "petitioner_advocate": petitioner_advocate,
            "respondent_advocate": respondent_advocate,
            "particulars": "list downloaded",
            "Pdf_name": pdf_filename,
            "case_status": "N/A",
            "IA_no": ia_no,
            "subject": subject,
            "act": act,
        }
        cases.append(case_data)
    
    logging.info(f"Extracted {len(cases)} valid cases from {pdf_filename}")
    return cases


# === SAVE TO EXCEL ===
def save_to_excel(cases_data, excel_path):
    """Save extracted cases to Excel."""
    try:
        if not cases_data:
            logging.warning("No case data to write.")
            return False

        columns = [
            "id", "causelist_slno", "court_hall_number", "Case_number", "Case_type",
            "case_year", "bench_name", "cause_date", "time", "chief_justice",
            "petitioner", "respondent", "petitioner_advocate", "respondent_advocate",
            "particulars", "Pdf_name", "case_status", "IA_no", "subject", "act"
        ]

        df_new = pd.DataFrame(cases_data)
        for c in columns:
            if c not in df_new.columns:
                df_new[c] = "N/A"
        df_new = df_new[columns]

        if os.path.exists(excel_path):
            df_old = pd.read_excel(excel_path)
            combined = pd.concat([df_old, df_new], ignore_index=True)
            combined["id"] = range(1, len(combined) + 1)
            combined.to_excel(excel_path, index=False)
            logging.info(f"Appended {len(df_new)} cases → total {len(combined)}.")
        else:
            df_new.to_excel(excel_path, index=False)
            logging.info(f"Created Excel file with {len(df_new)} cases.")

        return True
    except Exception as e:
        logging.error(f"Error saving Excel: {e}")
        return False


# === MAIN EXECUTION ===
def main():
    logging.info("=== JHARKHAND HIGH COURT CAUSELIST SCRAPER STARTED ===")

    driver = setup_driver()
    try:
        links = get_daily_causelist_links(driver)
        if not links:
            logging.error("No cause list links found.")
            return

        total_cases = 0
        for i, link in enumerate(links, start=1):
            logging.info(f"\n{'='*70}")
            logging.info(f"Processing {i}/{len(links)} → {link['text']}")
            logging.info(f"{'='*70}")

            if download_pdf(driver, link):
                pdfs = [f for f in os.listdir(OUTPUT_FOLDER) if f.lower().endswith(".pdf")]
                pdfs.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_FOLDER, x)), reverse=True)
                latest_pdf = pdfs[0] if pdfs else None

                if latest_pdf:
                    pdf_path = os.path.join(OUTPUT_FOLDER, latest_pdf)
                    text = extract_text_from_pdf(pdf_path)
                    cases = parse_causelist_data(text, latest_pdf)
                    if cases:
                        save_to_excel(cases, EXCEL_OUTPUT)
                        total_cases += len(cases)
            else:
                logging.warning(f"Skipping {link['text']} due to download failure.")

        logging.info(f"\n=== ALL DONE ===\nTotal cases extracted: {total_cases}")
        logging.info(f"Excel file saved to: {EXCEL_OUTPUT}")

    finally:
        driver.quit()
        logging.info("Browser closed. Scraper finished.")


if __name__ == "__main__":
    main()