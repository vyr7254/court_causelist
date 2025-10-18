import os
import time
import requests
import re
import logging
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import PyPDF2

# === CONFIGURATION ===
START_DATE = datetime(2025, 1, 1)
END_DATE = datetime(2025, 12, 31)
OUTPUT_FOLDER = r"C:\Users\Dell\OneDrive\Desktop\ghc_script\GHC_CauseLists"
EXCEL_OUTPUT = os.path.join(OUTPUT_FOLDER, "GHC_Complete_Cases.xlsx")
LOG_FILE = os.path.join(OUTPUT_FOLDER, "scraper_log.txt")

URL = "https://gujarathc-casestatus.nic.in/gujarathc/#"

# Setup logging
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)

# === CHROME SETUP ===
def setup_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
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

def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page_num in range(len(pdf_reader.pages)):
                if page_num == 0:
                    continue
                page = pdf_reader.pages[page_num]
                text += page.extract_text() + "\n"
            return text
    except Exception as e:
        logging.error(f"Error extracting PDF text: {e}")
        return None

# === DATA PARSING FUNCTIONS ===
def parse_causelist_data(pdf_text, cause_date, pdf_filename):
    """Parse the causelist text and extract case information"""
    cases = []
    
    court_sections = re.split(r'COURT\s+ROOM\s+NO[:\s]*(\d+)', pdf_text, flags=re.IGNORECASE)
    
    current_court_room = "N/A"
    current_chief_justice = "N/A"
    
    for i in range(1, len(court_sections), 2):
        if i >= len(court_sections):
            break
            
        current_court_room = court_sections[i].strip()
        section_text = court_sections[i + 1] if i + 1 < len(court_sections) else ""
        
        cj_patterns = [
            r'(CHIEF JUSTICE[^:\n]+)',
            r'(MRS?\.\s+JUSTICE[^:\n]+)',
            r"(HON['']BLE[^:\n]+JUSTICE[^:\n]+)",
        ]
        
        for pattern in cj_patterns:
            cj_match = re.search(pattern, section_text, re.IGNORECASE)
            if cj_match:
                current_chief_justice = cj_match.group(1).strip()
                break
        
        lines = section_text.split('\n')
        
        current_sno = None
        current_case_block = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            sno_match = re.match(r'^(\d+)\s+', line)
            
            if sno_match:
                if current_sno and current_case_block:
                    process_case_block(current_sno, current_case_block, current_court_room, 
                                     current_chief_justice, cause_date, pdf_filename, cases)
                
                current_sno = sno_match.group(1)
                current_case_block = [line]
            elif current_sno:
                current_case_block.append(line)
        
        if current_sno and current_case_block:
            process_case_block(current_sno, current_case_block, current_court_room, 
                             current_chief_justice, cause_date, pdf_filename, cases)
    
    return cases

def process_case_block(sno, case_lines, court_room, chief_justice, cause_date, pdf_filename, cases):
    """Process a single case block and extract all details"""
    
    full_text = ' '.join(case_lines)
    
    listed_match = re.search(r'LISTED\s+(\d+)\s+TIME[S]?', full_text, re.IGNORECASE)
    num_listings = int(listed_match.group(1)) if listed_match else 1
    
    case_numbers = re.findall(r'([A-Z]+)/(\d+)/(\d+)', full_text)
    
    vs_matches = re.finditer(r'([A-Z][A-Z\s\.\,\&\(\)]+?)\s+V[/\\]?S\s+([A-Z][A-Z\s\.\,\&\(\)]+?)(?=\s+MR|MS|ADVOCATE|\n|$)', full_text, re.IGNORECASE)
    
    parties_list = []
    for match in vs_matches:
        petitioner = match.group(1).strip()
        respondent = match.group(2).strip()
        parties_list.append((petitioner, respondent))
    
    advocate_pattern = r'(MR\.?|MS\.?|MRS\.?|ADVOCATE)\s+([A-Z][A-Z\s\.]+?)(?=\s+\d+|\s+MR|MS|FOR|LISTED|$)'
    advocates = re.findall(advocate_pattern, full_text, re.IGNORECASE)
    
    remarks_patterns = [
        r'(FOR\s+[A-Z\s]+)',
        r'(UNDER\s+[A-Z\s]+)',
        r'(FOR\s+CONDONATION\s+OF\s+DELAY)',
        r'(FOR\s+STAY)',
        r'(FOR\s+AMENDMENT)',
    ]
    
    remarks = "N/A"
    for pattern in remarks_patterns:
        remarks_match = re.search(pattern, full_text, re.IGNORECASE)
        if remarks_match:
            remarks = remarks_match.group(1).strip()
            break
    
    for listing_idx in range(num_listings):
        case_sl_no = f"{sno}.{listing_idx + 1}" if num_listings > 1 else sno
        
        case_number = "N/A"
        case_type = "N/A"
        
        if listing_idx < len(case_numbers):
            case_type = case_numbers[listing_idx][0]
            case_number = case_numbers[listing_idx][1]
        elif case_numbers:
            case_type = case_numbers[0][0]
            case_number = case_numbers[0][1]
        
        petitioner = "N/A"
        respondent = "N/A"
        
        if parties_list:
            petitioner = parties_list[0][0]
            respondent = parties_list[0][1]
        
        petitioner_advocate = "N/A"
        respondent_advocate = "N/A"
        
        if len(advocates) >= 1:
            petitioner_advocate = f"{advocates[0][0]} {advocates[0][1]}".strip()
        if len(advocates) >= 2:
            respondent_advocate = f"{advocates[1][0]} {advocates[1][1]}".strip()
        
        if num_listings > 1:
            listed_times = f"LISTED {num_listings} TIMES"
        else:
            listed_times = "LISTED 1 TIME"
        
        case_data = {
            'id': len(cases) + 1,
            'causelist_slno': case_sl_no,
            'court_hall_number': court_room,
            'Case_number': case_number,
            'Case_type': case_type,
            'bench_name': chief_justice if chief_justice != "N/A" else "GUJARAT",
            'cause_date': cause_date,
            'chief_justice': chief_justice,
            'petitioner': petitioner,
            'respondent': respondent,
            'petitioner_advocate': petitioner_advocate,
            'respondent_advocate': respondent_advocate,
            'particulars': 'list downloaded',
            'Pdf_name': pdf_filename,
            'listed_times': listed_times,
            'remarks': remarks
        }
        
        cases.append(case_data)

# === EXCEL MANAGEMENT ===
def save_to_excel(cases_data, excel_path):
    """Save or append data to Excel file"""
    try:
        columns = [
            'id', 'causelist_slno', 'court_hall_number', 'Case_number', 'Case_type', 
            'bench_name', 'cause_date', 'chief_justice', 'petitioner', 'respondent', 
            'petitioner_advocate', 'respondent_advocate', 'particulars', 'Pdf_name', 
            'listed_times', 'remarks'
        ]
        
        if os.path.exists(excel_path):
            existing_df = pd.read_excel(excel_path)
            new_df = pd.DataFrame(cases_data)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            combined_df['id'] = range(1, len(combined_df) + 1)
            combined_df = combined_df[columns]
            combined_df.to_excel(excel_path, index=False)
            logging.info(f"Appended {len(cases_data)} cases to existing Excel file")
        else:
            df = pd.DataFrame(cases_data)
            df = df[columns]
            df.to_excel(excel_path, index=False)
            logging.info(f"Created new Excel file with {len(cases_data)} cases")
        
        return True
    except Exception as e:
        logging.error(f"Error saving to Excel: {e}")
        return False

# === MAIN SCRAPING FUNCTION ===
def navigate_to_causelist_page(driver):
    """Navigate to the causelist page from main page"""
    try:
        driver.get(URL)
        time.sleep(5)
        
        causelist_found = False
        
        # Method 1: Look for link with text containing "CAUSELIST"
        try:
            causelist_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[contains(translate(text(), 'CAUSELIST', 'causelist'), 'causelist')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", causelist_link)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", causelist_link)
            causelist_found = True
            logging.info("Clicked CAUSELIST using Method 1")
        except:
            pass
        
        if not causelist_found:
            try:
                causelist_link = driver.find_element(By.XPATH, "//a[contains(@href, 'causelist') or contains(@onclick, 'causelist')]")
                driver.execute_script("arguments[0].scrollIntoView(true);", causelist_link)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", causelist_link)
                causelist_found = True
                logging.info("Clicked CAUSELIST using Method 2")
            except:
                pass
        
        if not causelist_found:
            try:
                all_links = driver.find_elements(By.TAG_NAME, "a")
                for link in all_links:
                    if "cause" in link.text.lower() and "list" in link.text.lower():
                        driver.execute_script("arguments[0].scrollIntoView(true);", link)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", link)
                        causelist_found = True
                        logging.info(f"Clicked CAUSELIST using Method 3: {link.text}")
                        break
            except:
                pass
        
        if not causelist_found:
            logging.error("Could not find CAUSELIST link on main page")
            return False
        
        time.sleep(5)
        return True
        
    except Exception as e:
        logging.error(f"Error navigating to causelist page: {e}")
        return False

def download_and_process_causelist(driver, date):
    """Download causelist for a specific date and extract data"""
    date_str = date.strftime("%d/%m/%Y")
    date_filename = date.strftime("%d_%m_%Y")
    
    logging.info(f"Processing date: {date_str}")
    
    # Get list of existing PDFs before download
    existing_pdfs = set()
    try:
        existing_pdfs = set([f for f in os.listdir(OUTPUT_FOLDER) if f.lower().endswith('.pdf')])
    except:
        pass
    
    try:
        # Wait for page to fully load
        time.sleep(3)
        
        # Step 1: Find and interact with the date input field
        try:
            date_input = None
            
            # Method 1: Try to find visible input with value
            try:
                all_inputs = driver.find_elements(By.XPATH, "//input[@type='text']")
                for inp in all_inputs:
                    if inp.is_displayed() and inp.is_enabled():
                        # Check if it has a date-like value or is empty
                        val = inp.get_attribute('value')
                        if val and ('/' in val or len(val) == 10):
                            date_input = inp
                            logging.info(f"Found date input with existing value: {val}")
                            break
                        elif not val:
                            date_input = inp
                            logging.info("Found empty visible input")
                            break
            except Exception as e:
                logging.warning(f"Method 1 failed: {e}")
            
            # Method 2: Find input near "DATE" text
            if not date_input:
                try:
                    # Look for the DATE label/heading and find input near it
                    date_input = driver.find_element(By.XPATH, 
                        "//input[@type='text' and (preceding-sibling::*[contains(text(), 'DATE')] or following-sibling::*[contains(text(), 'DATE')])]")
                    logging.info("Found date input near DATE heading")
                except Exception as e:
                    logging.warning(f"Method 2 failed: {e}")
            
            # Method 3: Try using JavaScript to find the right input
            if not date_input:
                try:
                    date_input = driver.execute_script("""
                        var inputs = document.querySelectorAll('input[type="text"]');
                        for(var i=0; i<inputs.length; i++) {
                            if(inputs[i].offsetWidth > 0 && inputs[i].offsetHeight > 0) {
                                return inputs[i];
                            }
                        }
                        return null;
                    """)
                    if date_input:
                        logging.info("Found date input using JavaScript")
                except Exception as e:
                    logging.warning(f"Method 3 failed: {e}")
            
            if not date_input:
                logging.error("Could not find date input field")
                # Save screenshot for debugging
                driver.save_screenshot(os.path.join(OUTPUT_FOLDER, f"error_{date_filename}.png"))
                return False
            
            # Make sure element is visible and enabled
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", date_input)
            time.sleep(1)
            
            # Use JavaScript to set the value directly
            try:
                # First try clicking with JavaScript
                driver.execute_script("arguments[0].click();", date_input)
                time.sleep(0.5)
                
                # Clear the field using JavaScript
                driver.execute_script("arguments[0].value = '';", date_input)
                time.sleep(0.5)
                
                # Set the value using JavaScript
                driver.execute_script(f"arguments[0].value = '{date_str}';", date_input)
                
                # Trigger change event to make sure the form registers the change
                driver.execute_script("""
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """, date_input)
                
                logging.info(f"Set date using JavaScript: {date_str}")
                time.sleep(1)
                
            except Exception as e:
                logging.error(f"JavaScript method failed, trying keyboard input: {e}")
                
                # Fallback: Try keyboard input
                try:
                    date_input.click()
                    time.sleep(0.5)
                    date_input.clear()
                    time.sleep(0.5)
                    date_input.send_keys(date_str)
                    logging.info(f"Entered date via keyboard: {date_str}")
                    time.sleep(1)
                except Exception as e2:
                    logging.error(f"Keyboard input also failed: {e2}")
                    return False
            
        except Exception as e:
            logging.error(f"Could not enter date: {e}")
            driver.save_screenshot(os.path.join(OUTPUT_FOLDER, f"error_{date_filename}.png"))
            return False

        # Step 2: Click GO button
        try:
            # Wait a moment for any JavaScript to process
            time.sleep(1)
            
            go_button = None
            
            # Method 1: Find by text content
            try:
                go_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'GO') or contains(., 'GO')]"))
                )
                logging.info("Found GO button")
            except:
                pass
            
            # Method 2: Try to find any button near the date field
            if not go_button:
                try:
                    go_button = driver.find_element(By.XPATH, "//button[@type='submit' or @type='button']")
                    logging.info("Found submit button")
                except:
                    pass
            
            if not go_button:
                logging.error("Could not find GO button")
                driver.save_screenshot(os.path.join(OUTPUT_FOLDER, f"error_go_button_{date_filename}.png"))
                return False
            
            # Click GO button
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", go_button)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", go_button)
            logging.info("Clicked GO button")
            time.sleep(5)
            
        except Exception as e:
            logging.error(f"Could not click GO button: {e}")
            driver.save_screenshot(os.path.join(OUTPUT_FOLDER, f"error_go_{date_filename}.png"))
            return False

        # Step 3: Click on "COMPLETE" button/tab
        try:
            time.sleep(2)
            
            complete_button = None
            
            # Method 1: Find by text "COMPLETE"
            try:
                complete_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'COMPLETE')] | //a[contains(text(), 'COMPLETE')]"))
                )
                logging.info("Found COMPLETE button by text")
            except:
                pass
            
            # Method 2: Try finding by class or any clickable element with COMPLETE
            if not complete_button:
                try:
                    complete_button = driver.find_element(By.XPATH, "//*[contains(., 'COMPLETE') and (self::button or self::a)]")
                    logging.info("Found COMPLETE element")
                except:
                    pass
            
            if complete_button:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", complete_button)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", complete_button)
                logging.info("Clicked COMPLETE button")
                time.sleep(3)
            else:
                logging.warning("Could not find COMPLETE button, continuing anyway")
                
        except Exception as e:
            logging.warning(f"Error clicking COMPLETE button: {e}")
            # Continue anyway as the page might already show complete causelist
        
        # Step 4: Click "GET CAUSELIST" button (with download icon)
        try:
            time.sleep(2)
            
            get_causelist_button = None
            
            # Method 1: Find by text containing "GET CAUSELIST"
            try:
                get_causelist_button = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'GET CAUSELIST')] | //a[contains(text(), 'GET CAUSELIST')]"))
                )
                logging.info("Found GET CAUSELIST button by text")
            except:
                pass
            
            # Method 2: Find button with download icon (class might contain 'download' or icon)
            if not get_causelist_button:
                try:
                    get_causelist_button = driver.find_element(By.XPATH, 
                        "//button[contains(., 'GET CAUSELIST') or .//i[contains(@class, 'download')]] | " +
                        "//a[contains(., 'GET CAUSELIST') or .//i[contains(@class, 'download')]]")
                    logging.info("Found GET CAUSELIST button with icon")
                except:
                    pass
            
            # Method 3: Find by looking for download icon symbol
            if not get_causelist_button:
                try:
                    # Look for button with download symbol or near "GET CAUSELIST" text
                    buttons = driver.find_elements(By.TAG_NAME, "button")
                    for btn in buttons:
                        if "GET CAUSELIST" in btn.text.upper() or "CAUSELIST" in btn.text.upper():
                            get_causelist_button = btn
                            logging.info(f"Found button with text: {btn.text}")
                            break
                except:
                    pass
            
            # Method 4: Try finding by class attributes common in download buttons
            if not get_causelist_button:
                try:
                    get_causelist_button = driver.find_element(By.XPATH, 
                        "//button[contains(@class, 'btn') and contains(., 'CAUSELIST')]")
                    logging.info("Found GET CAUSELIST by class")
                except:
                    pass
            
            if not get_causelist_button:
                logging.error("Could not find GET CAUSELIST button")
                driver.save_screenshot(os.path.join(OUTPUT_FOLDER, f"error_get_causelist_{date_filename}.png"))
                return False
            
            # Scroll to button and click
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", get_causelist_button)
            time.sleep(1)
            
            # Try multiple click methods
            try:
                # First try JavaScript click
                driver.execute_script("arguments[0].click();", get_causelist_button)
                logging.info("Clicked GET CAUSELIST button using JavaScript")
            except:
                try:
                    # Fallback to regular click
                    get_causelist_button.click()
                    logging.info("Clicked GET CAUSELIST button using regular click")
                except Exception as click_error:
                    logging.error(f"Failed to click GET CAUSELIST button: {click_error}")
                    driver.save_screenshot(os.path.join(OUTPUT_FOLDER, f"error_click_causelist_{date_filename}.png"))
                    return False
            
            time.sleep(8)  # Wait for PDF to load/download
            
        except Exception as e:
            logging.error(f"Error with GET CAUSELIST button: {e}")
            driver.save_screenshot(os.path.join(OUTPUT_FOLDER, f"error_causelist_button_{date_filename}.png"))
            return False
        
        # Wait for PDF to download to folder
        # The PDF is downloaded directly, not opened in a new tab
        logging.info(f"Waiting for PDF to download for {date_str}")
        
        # Generate possible PDF filenames
        # Try different formats the website might use
        possible_filenames = [
            f"Complete_Causelist_1st_January_2025.pdf",  # Example from user
            f"Complete_Causelist_gujarat_{date_filename}.pdf",
            f"causelist_{date_filename}.pdf",
            f"Causelist_{date_filename}.pdf",
            f"Complete_Causelist_{date_filename}.pdf"
        ]
        
        # Also try to construct the date in word format
        date_obj = datetime.strptime(date_str, "%d/%m/%Y")
        day_suffix = "th"
        day_num = date_obj.day
        if day_num in [1, 21, 31]:
            day_suffix = "st"
        elif day_num in [2, 22]:
            day_suffix = "nd"
        elif day_num in [3, 23]:
            day_suffix = "rd"
        
        month_name = date_obj.strftime("%B")
        year_name = date_obj.strftime("%Y")
        formatted_date = f"{day_num}{day_suffix}_{month_name}_{year_name}"
        
        possible_filenames.insert(0, f"Complete_Causelist_{formatted_date}.pdf")
        
        # Wait for file to appear and download to complete
        max_wait_time = 30  # seconds
        wait_interval = 1  # second
        downloaded_file = None
        
        for attempt in range(max_wait_time):
            time.sleep(wait_interval)
            
            # Check for any of the possible filenames
            for filename in possible_filenames:
                pdf_path = os.path.join(OUTPUT_FOLDER, filename)
                if os.path.exists(pdf_path):
                    # Check if file size is stable (download complete)
                    try:
                        size1 = os.path.getsize(pdf_path)
                        time.sleep(1)
                        size2 = os.path.getsize(pdf_path)
                        if size1 == size2 and size1 > 1000:  # File size stable and > 1KB
                            downloaded_file = pdf_path
                            logging.info(f"Found downloaded PDF: {filename}")
                            break
                    except:
                        continue
            
            if downloaded_file:
                break
            
            # Also check for any newly created PDF files in the folder
            if attempt % 5 == 0:  # Check every 5 seconds
                try:
                    files = os.listdir(OUTPUT_FOLDER)
                    pdf_files = [f for f in files if f.lower().endswith('.pdf')]
                    if pdf_files:
                        # Get the most recently modified PDF
                        latest_pdf = max(pdf_files, key=lambda f: os.path.getmtime(os.path.join(OUTPUT_FOLDER, f)))
                        pdf_path = os.path.join(OUTPUT_FOLDER, latest_pdf)
                        
                        # Check if it was modified in the last 30 seconds
                        if time.time() - os.path.getmtime(pdf_path) < 30:
                            # Check if file size is stable
                            size1 = os.path.getsize(pdf_path)
                            time.sleep(1)
                            size2 = os.path.getsize(pdf_path)
                            if size1 == size2 and size1 > 1000:
                                downloaded_file = pdf_path
                                logging.info(f"Found recently downloaded PDF: {latest_pdf}")
                                break
                except Exception as e:
                    logging.warning(f"Error checking for downloaded files: {e}")
        
        if not downloaded_file:
            logging.warning(f"No PDF downloaded for {date_str} after {max_wait_time} seconds")
            # Close any extra windows/tabs
            if len(driver.window_handles) > 1:
                for handle in driver.window_handles[1:]:
                    driver.switch_to.window(handle)
                    driver.close()
                driver.switch_to.window(driver.window_handles[0])
            return False
        
        # Close any extra windows/tabs
        if len(driver.window_handles) > 1:
            for handle in driver.window_handles[1:]:
                driver.switch_to.window(handle)
                driver.close()
            driver.switch_to.window(driver.window_handles[0])
        
        # Extract and parse PDF
        try:
            pdf_text = extract_text_from_pdf(downloaded_file)
            
            if pdf_text:
                # Use the actual filename without extension
                pdf_filename = os.path.splitext(os.path.basename(downloaded_file))[0]
                cases = parse_causelist_data(pdf_text, date_str, pdf_filename)
                
                if cases:
                    save_to_excel(cases, EXCEL_OUTPUT)
                    logging.info(f"Extracted {len(cases)} cases for {date_str}")
                    return True
                else:
                    logging.warning(f"No cases extracted from PDF for {date_str}")
                    return False
            else:
                logging.error(f"Could not extract text from PDF for {date_str}")
                return False
                
        except Exception as e:
            logging.error(f"Error processing PDF for {date_str}: {e}")
            return False
            
    except Exception as e:
        logging.error(f"Error processing {date_str}: {e}")
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
        return False

# === MAIN EXECUTION ===
def main():
    logging.info("Starting Gujarat High Court Cause List Scraper")
    logging.info(f"Date range: {START_DATE.strftime('%d/%m/%Y')} to {END_DATE.strftime('%d/%m/%Y')}")
    
    driver = setup_driver()
    
    try:
        # Navigate to causelist page first
        if not navigate_to_causelist_page(driver):
            logging.error("Failed to navigate to causelist page. Exiting.")
            driver.quit()
            return
        
        current_date = START_DATE
        success_count = 0
        failure_count = 0
        
        while current_date <= END_DATE:
            result = download_and_process_causelist(driver, current_date)
            
            if result:
                success_count += 1
            else:
                failure_count += 1
            
            # Move to next day
            current_date += timedelta(days=1)
            
            # Small delay between requests
            time.sleep(2)
        
        logging.info(f"Scraping completed. Success: {success_count}, Failed: {failure_count}")
        
    except Exception as e:
        logging.error(f"Critical error in main loop: {e}")
    finally:
        driver.quit()
        logging.info("Browser closed. Script finished.")

if __name__ == "__main__":
    main()
