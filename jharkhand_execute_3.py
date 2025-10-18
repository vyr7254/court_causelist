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

    # ✅ ensure PDFs are downloaded instead of opened in Chrome Viewer
    prefs = {
        "download.default_directory": OUTPUT_FOLDER,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "plugins.plugins_disabled": ["Chrome PDF Viewer"],  # critical fix
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


# === PARSING LOGIC ===
def parse_causelist_data(pdf_text, pdf_filename):
    """Extract clean structured case data."""
    cases = []
    if not pdf_text.strip():
        return cases

    lines = [ln.strip() for ln in pdf_text.splitlines() if ln.strip()]
    blocks = []
    block = None

    for line in lines:
        if re.match(r"^\d+\s+[A-Za-z]", line):
            if block:
                blocks.append(block)
            sno = line.split()[0]
            content = " ".join(line.split()[1:])
            block = {"sno": sno, "lines": [content]}
        elif block:
            block["lines"].append(line)

    if block:
        blocks.append(block)

    case_pattern = re.compile(r"([A-Za-z\.\(\)/\s-]+)\s*[/\-]?\s*(\d+)\s*[/\-]?\s*(\d{4})")

    id_counter = 0
    for blk in blocks:
        text = " ".join(blk["lines"])
        m_case = case_pattern.search(text)
        if not m_case:
            continue

        case_type = m_case.group(1).strip()
        case_number = m_case.group(2).strip()
        case_year = m_case.group(3).strip()

        m_vs = re.search(r"\b(VS|V/S|V\.S\.)\b", text, re.IGNORECASE)
        if m_vs:
            petitioner = text[:m_vs.start()].strip()
            respondent = text[m_vs.end():].split("IA", 1)[0].strip()
        else:
            parts = text.split("  ", 1)
            petitioner = parts[0].strip()
            respondent = parts[1].strip() if len(parts) > 1 else "N/A"

        ia_match = re.search(r"IA\s*NO\.?\s*([0-9/]+)", text, re.IGNORECASE)
        subject_match = re.search(r"SUBJECT\s*[:-]?\s*([A-Z\s,]+)", text, re.IGNORECASE)
        act_match = re.search(r"ACT\s*[:-]?\s*(.+)", text, re.IGNORECASE)

        ia_no = ia_match.group(1).strip() if ia_match else "N/A"
        subject = subject_match.group(1).strip() if subject_match else "N/A"
        act = act_match.group(1).strip() if act_match else "N/A"

        id_counter += 1
        case_data = {
            "id": id_counter,
            "causelist_slno": blk["sno"],
            "court_hall_number": "N/A",
            "Case_number": case_number,
            "Case_type": case_type,
            "case_year": case_year,
            "bench_name": "Jharkhand",
            "cause_date": "N/A",
            "time": "N/A",
            "chief_justice": "N/A",
            "petitioner": petitioner or "N/A",
            "respondent": respondent or "N/A",
            "petitioner_advocate": "N/A",
            "respondent_advocate": "N/A",
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
