import os
import re
import time
import json
import ctypes
import threading
import traceback
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# ------------------------------
# CONFIG
# ------------------------------
YEAR = 2025
BASE_FOLDER = r"C:\Users\Dell\OneDrive\Desktop\cause list\Karnataka_CauseLists"
OUTPUT_EXCEL = os.path.join(BASE_FOLDER, f"Karnataka_AllBenches_{YEAR}.xlsx")
PROGRESS_FILE = os.path.join(BASE_FOLDER, f"progress_{YEAR}.json")
CAUSELIST_URL = "https://judiciary.karnataka.gov.in/causelistSearch.php"

BENCHES = {
    "K": "Kalaburagi Bench"
}

MIN_COURTS_TO_CHECK = 7
MAX_COURTS = 40
SPRINT_DAYS = 0  # 0 means daily sprint (same date for from and to)

# ------------------------------
# UTILS AND NO-SLEEP HELPER
# ------------------------------
def debug_print(msg):
    print(msg, flush=True)

def ensure_folder():
    os.makedirs(BASE_FOLDER, exist_ok=True)

# Windows: prevent sleep using SetThreadExecutionState in a background thread
def start_prevent_sleep_thread():
    try:
        ES_CONTINUOUS = 0x80000000
        ES_SYSTEM_REQUIRED = 0x00000001
        # optional: ES_AWAYMODE_REQUIRED = 0x00000040
        flags = ES_CONTINUOUS | ES_SYSTEM_REQUIRED

        def prevent_sleep_loop():
            kernel32 = ctypes.windll.kernel32
            while not prevent_sleep_thread_stop.is_set():
                kernel32.SetThreadExecutionState(flags)
                # call every 30 seconds
                time.sleep(30)

        prevent_sleep_thread_stop.clear()
        t = threading.Thread(target=prevent_sleep_loop, daemon=True)
        t.start()
        debug_print("🔒 Prevent-sleep thread started (Windows).")
    except Exception as e:
        debug_print(f"⚠️ Prevent-sleep not available: {repr(e)}")

def stop_prevent_sleep_thread():
    try:
        prevent_sleep_thread_stop.set()
        # call once to clear flags (ES_CONTINUOUS alone won't disable, so call without SYSTEM_REQUIRED)
        try:
            kernel32 = ctypes.windll.kernel32
            kernel32.SetThreadExecutionState(0x80000000)
        except Exception:
            pass
        debug_print("🔓 Prevent-sleep thread stopped.")
    except Exception:
        pass

prevent_sleep_thread_stop = threading.Event()

# ------------------------------
# Selenium setup
# ------------------------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-notifications")
# headless disabled to reduce detection issues; enable if you want: chrome_options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.set_page_load_timeout(60)
wait = WebDriverWait(driver, 25)

def dispatch_events_on(el):
    try:
        driver.execute_script("""
            var el = arguments[0];
            ['input','change','blur','keyup'].forEach(function(t){
                try { el.dispatchEvent(new Event(t, {bubbles:true})); } catch(e) {}
            });
        """, el)
    except Exception:
        pass

# ------------------------------
# Date setting helpers (kept from your original)
# ------------------------------
def try_jquery_datepicker_set(id_or_selector, date_str):
    try:
        sel = id_or_selector
        script = f"""
            try {{
                if (window.jQuery && jQuery('{sel}').datepicker) {{
                    jQuery('{sel}').datepicker('setDate', '{date_str}');
                    jQuery('{sel}').trigger('change');
                    return true;
                }}
            }} catch(e) {{ return false; }}
            return false;
        """
        return driver.execute_script(script)
    except Exception:
        return False

def set_date_on_elements(from_el, to_el, from_val, to_val, debug_label=""):
    formats = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"]
    def read_back_valid(el, expected_day, expected_month):
        try:
            val = (el.get_attribute("value") or "").strip()
            if not val:
                return False, val
            if str(expected_day).zfill(2) in val and str(expected_month).zfill(2) in val:
                return True, val
            return False, val
        except Exception:
            return False, ""
    try:
        dt_test = datetime.strptime(from_val, formats[0])
        exp_day = dt_test.day
        exp_month = dt_test.month
    except Exception:
        try:
            parts = re.split(r'[-/]', from_val)
            exp_day = int(parts[0])
            exp_month = int(parts[1])
        except Exception:
            exp_day = None
            exp_month = None
    # try jQuery by id
    try:
        fid = from_el.get_attribute("id") or ""
        tid = to_el.get_attribute("id") or ""
        if fid:
            for fmt in formats:
                s_from = datetime.strptime(from_val, formats[0]).strftime(fmt) if fmt != formats[0] else from_val
                s_to = datetime.strptime(to_val, formats[0]).strftime(fmt) if fmt != formats[0] else to_val
                ok1 = try_jquery_datepicker_set(f"#{fid}", s_from)
                ok2 = try_jquery_datepicker_set(f"#{tid}", s_to) if tid else ok1
                time.sleep(0.2)
                valid_from, read_from = read_back_valid(from_el, exp_day, exp_month)
                if ok1 and valid_from:
                    debug_print(f"      ℹ jQuery datepicker set by id succeeded ({fid}) -> {read_from}")
                    return True
    except Exception:
        pass
    # try jQuery by name
    try:
        fname = from_el.get_attribute("name") or ""
        tname = to_el.get_attribute("name") or ""
        if fname:
            for fmt in formats:
                s_from = datetime.strptime(from_val, formats[0]).strftime(fmt) if fmt != formats[0] else from_val
                s_to = datetime.strptime(to_val, formats[0]).strftime(fmt) if fmt != formats[0] else to_val
                sel_from = f"input[name='{fname}']"
                sel_to = f"input[name='{tname}']" if tname else sel_from
                ok1 = try_jquery_datepicker_set(sel_from, s_from)
                ok2 = try_jquery_datepicker_set(sel_to, s_to)
                time.sleep(0.2)
                valid_from, read_from = read_back_valid(from_el, exp_day, exp_month)
                if ok1 and valid_from:
                    debug_print(f"      ℹ jQuery datepicker set by name succeeded ({fname}) -> {read_from}")
                    return True
    except Exception:
        pass
    # try JS direct set and dispatch events
    for fmt in formats:
        try:
            val_from = datetime.strptime(from_val, formats[0]).strftime(fmt) if fmt != formats[0] else from_val
            val_to = datetime.strptime(to_val, formats[0]).strftime(fmt) if fmt != formats[0] else to_val
            driver.execute_script("""
                var f = arguments[0], v1 = arguments[1], t = arguments[2], v2 = arguments[3];
                try { f.value = v1; } catch(e) {}
                try { t.value = v2; } catch(e) {}
                ['input','change','blur','keyup'].forEach(function(ev){
                    try { f.dispatchEvent(new Event(ev, {bubbles:true})); } catch(e) {}
                    try { t.dispatchEvent(new Event(ev, {bubbles:true})); } catch(e) {}
                });
            """, from_el, val_from, to_el, val_to)
            dispatch_events_on(from_el)
            dispatch_events_on(to_el)
            time.sleep(0.4)
            valid_from, read_from = read_back_valid(from_el, exp_day, exp_month)
            if valid_from:
                debug_print(f"      ℹ JS set value succeeded -> {read_from} (fmt={fmt})")
                return True
        except Exception:
            time.sleep(0.1)
    # fallback to send_keys
    try:
        for fmt in formats:
            try_val = datetime.strptime(from_val, formats[0]).strftime(fmt) if fmt != formats[0] else from_val
            try:
                from_el.clear()
                from_el.click()
                from_el.send_keys(try_val)
                from_el.send_keys(Keys.TAB)
                time.sleep(0.2)
                to_el.clear()
                to_el.send_keys(datetime.strptime(to_val, formats[0]).strftime(fmt) if fmt != formats[0] else to_val)
                to_el.send_keys(Keys.TAB)
                time.sleep(0.4)
                dispatch_events_on(from_el)
                valid_from, read_from = read_back_valid(from_el, exp_day, exp_month)
                if valid_from:
                    debug_print(f"      ℹ send_keys succeeded -> {read_from} (fmt={fmt})")
                    return True
            except Exception:
                time.sleep(0.1)
    except Exception:
        pass
    return False

# ------------------------------
# DOM helpers from original code
# ------------------------------
def find_form_and_bench_select(driver, bench_code):
    forms = driver.find_elements(By.TAG_NAME, "form")
    for form in forms:
        try:
            selects = form.find_elements(By.TAG_NAME, "select")
            for sel in selects:
                try:
                    opts = sel.find_elements(By.TAG_NAME, "option")
                    for opt in opts:
                        if (opt.get_attribute("value") or "").strip() == bench_code:
                            return form, sel
                except Exception:
                    continue
        except Exception:
            continue
    selects = driver.find_elements(By.XPATH, "//select[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'bench') or contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'bench')]")
    for sel in selects:
        try:
            opts = sel.find_elements(By.TAG_NAME, "option")
            for opt in opts:
                if (opt.get_attribute("value") or "").strip() == bench_code:
                    try:
                        parent_form = sel.find_element(By.XPATH, "./ancestor::form[1]")
                    except Exception:
                        parent_form = None
                    return parent_form, sel
        except Exception:
            continue
    return None, None

def find_court_select(form=None):
    try:
        if form:
            selects = form.find_elements(By.TAG_NAME, "select")
        else:
            selects = driver.find_elements(By.TAG_NAME, "select")
        for sel in selects:
            try:
                options = sel.find_elements(By.TAG_NAME, "option")
                option_texts = [opt.text.strip().upper() for opt in options if opt.text.strip()]
                court_hall_pattern_count = sum(1 for txt in option_texts if "COURT HALL -" in txt or "COURT HALL-" in txt)
                if court_hall_pattern_count >= 2:
                    return sel
            except Exception:
                continue
        for sel in selects:
            try:
                name = (sel.get_attribute("name") or "").lower()
                id_attr = (sel.get_attribute("id") or "").lower()
                if "search" in name and "by" in name:
                    continue
                if "searchby" in name or "searchby" in id_attr:
                    continue
                if any(term in name or term in id_attr for term in ["courthall", "court_hall", "hallno", "hall_no"]):
                    return sel
            except Exception:
                continue
    except Exception:
        pass
    return None

def click_get_button_in_form(form):
    try:
        candidates = form.find_elements(By.XPATH, ".//input[@type='button'] | .//button | .//input[@type='submit']")
        for c in candidates:
            try:
                txt = (c.get_attribute("value") or c.text or "").strip().lower()
                if txt and ("get" in txt or "details" in txt or "search" in txt):
                    driver.execute_script("arguments[0].click();", c)
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False

# ------------------------------
# Improved party & advocate parsing
# ------------------------------
def split_party_and_advocate(text, prefix=None):
    if not text:
        return "", ""
    # Normalize whitespace and line breaks
    text = re.sub(r'\r', '\n', text)
    text = re.sub(r'\n+', '\n', text).strip()
    # Remove prefix tokens like "PET:" "RES:" (case-insensitive)
    if prefix:
        text = re.sub(r'(?i)^\s*' + re.escape(prefix) + r'\s*:\s*', '', text).strip()
    # Split by newlines
    parts = [p.strip() for p in text.split("\n") if p.strip()]
    if len(parts) == 0:
        return "", ""
    if len(parts) == 1:
        # Try to split inline by known tokens (RES:, PET:, GA, SD, ADV) if appear
        inline = parts[0]
        # split by " GA " or " SD " or " ADV " if present (common advocate abbreviations)
        m = re.split(r'\b(GA|SD|ADV|ADVOCATE)\b', inline, flags=re.IGNORECASE)
        if len(m) > 1:
            # m[0] is party, rest join as advocate text
            party = m[0].strip()
            adv = " ".join([p.strip() for p in m[1:] if p.strip()])
            return party, adv
        return parts[0], ""
    # multiple lines -> first line = party, rest = advocates combined
    party = parts[0]
    adv = " ".join(parts[1:])
    return party, adv

# ------------------------------
# Page parsing of table rows (ORIGINAL VERSION - No changes)
# ------------------------------
def extract_case_data_from_page(bench_name, date_str, court_no):
    records = []
    try:
        time.sleep(1.5)
        try:
            page_text = driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            page_text = ""
        judge_name = ""
        try:
            judge_elem = driver.find_element(By.XPATH, "//*[contains(text(), 'HON') and contains(text(), 'JUSTICE')]")
            judge_name = judge_elem.text.strip()
        except Exception:
            pass
        court_hall_display = str(court_no)
        try:
            if "COURT HALL NO" in page_text.upper():
                match = re.search(r'COURT HALL NO\s*:?\s*(\d+)', page_text, re.IGNORECASE)
                if match:
                    court_hall_display = match.group(1)
        except Exception:
            pass
        if "no record" in page_text.lower() or "no data" in page_text.lower():
            if "sl.no" not in page_text.lower() and "case no" not in page_text.lower():
                debug_print(f"      ⚠️ No records found for Court {court_no}")
                return records
        tables = driver.find_elements(By.TAG_NAME, "table")
        debug_print(f"      📍 Total tables found on page: {len(tables)}")
        if not tables:
            debug_print(f"      ⚠️ No table found for Court {court_no}")
            return records
        
        # Find the first table with actual case data (has both "sl" and "case" in text)
        data_table = None
        for table in tables:
            table_text = table.text.lower()
            if "sl" in table_text and "case" in table_text:
                rows = table.find_elements(By.TAG_NAME, "tr")
                if len(rows) >= 2:
                    data_table = table
                    debug_print(f"      📊 Found data table with {len(rows)} rows")
                    break
        
        if not data_table:
            debug_print(f"      ⚠️ No valid data table found for Court {court_no}")
            return records
        
        # Process only the found data table
        tables = [data_table]
        
        for table in tables:
            try:
                table_text = table.text.lower()
                rows = table.find_elements(By.TAG_NAME, "tr")
                if len(rows) < 2:
                    continue
                header_row = None
                header_indices = {}
                for idx, row in enumerate(rows):
                    try:
                        cells = row.find_elements(By.TAG_NAME, "th")
                        if not cells:
                            cells = row.find_elements(By.TAG_NAME, "td")
                        if cells:
                            header_text = [c.text.strip().lower() for c in cells]
                            joined = " ".join(header_text)
                            if "sl.no" in joined or "case no" in joined:
                                header_row = idx
                                for i, txt in enumerate(header_text):
                                    if "sl" in txt or "sr" in txt:
                                        header_indices["sl_no"] = i
                                    elif "case" in txt and "no" in txt:
                                        header_indices["case_no"] = i
                                    elif "pet" in txt or "appl" in txt or "comp" in txt:
                                        header_indices["petitioner"] = i
                                    elif "resp" in txt:
                                        header_indices["respondent"] = i
                                break
                    except Exception:
                        continue
                # now parse rows after header_row
                for row in rows[header_row+1:] if header_row is not None else rows[1:]:
                    try:
                        cols = row.find_elements(By.TAG_NAME, "td")
                        if len(cols) < 2:
                            continue
                        cell_texts = [c.text.strip() for c in cols]
                        
                        # CHANGE 2: Extract serial number from causelist_slno
                        causelist_slno = ""
                        if "sl_no" in header_indices and header_indices["sl_no"] < len(cell_texts):
                            causelist_slno = cell_texts[header_indices["sl_no"]].strip()
                        
                        record = {
                            "Bench": bench_name,
                            "Cause_Date": date_str,
                            "Court_Hall": court_hall_display,
                            "Causelist_Slno": causelist_slno[:50],
                            "Judge": judge_name[:200] if judge_name else "",
                            "Mode": "",
                            "Case_Type": "",
                            "Case_No": "",
                            "Year": "",
                            "Petitioner": "",
                            "Petitioner_Advocate": "",
                            "Respondent": "",
                            "Respondent_Advocate": ""
                        }
                        # extract case number: prefer header index if present
                        case_text = ""
                        if "case_no" in header_indices and header_indices["case_no"] < len(cell_texts):
                            case_text = cell_texts[header_indices["case_no"]]
                        else:
                            # try scanning columns for a pattern like "CCC 1176/2024"
                            for txt in cell_texts:
                                if re.search(r'[A-Z]+\s+\d+/\d{4}', txt):
                                    case_text = txt
                                    break
                        case_match = re.search(r'([A-Z]+)\s+(\d+)/(\d{4})', case_text)
                        if case_match:
                            record["Case_Type"] = case_match.group(1)
                            record["Case_No"] = case_match.group(2)
                            record["Year"] = case_match.group(3)
                        # Petitioner / Respondent extraction using header_indices, fallback to scanning cells
                        pet_cell_text = ""
                        res_cell_text = ""
                        if "petitioner" in header_indices and header_indices["petitioner"] < len(cell_texts):
                            pet_cell_text = cell_texts[header_indices["petitioner"]]
                        if "respondent" in header_indices and header_indices["respondent"] < len(cell_texts):
                            res_cell_text = cell_texts[header_indices["respondent"]]
                        # fallback: detect "PET:" or "RES:" tokens inside full row text
                        if not pet_cell_text or not res_cell_text:
                            joined_row = "\n".join(cell_texts)
                            # try extract blocks like "PET: ... \n RES: ..."
                            pet_match = re.search(r'(PET[:\s].*?)(?=(RES[:\s]|$))', joined_row, re.IGNORECASE | re.DOTALL)
                            res_match = re.search(r'(RES[:\s].*)', joined_row, re.IGNORECASE | re.DOTALL)
                            if pet_match and not pet_cell_text:
                                pet_cell_text = pet_match.group(1).strip()
                            if res_match and not res_cell_text:
                                res_cell_text = res_match.group(1).strip()
                        # parse petitioner
                        try:
                            if pet_cell_text:
                                pet_name, pet_adv = split_party_and_advocate(pet_cell_text, "PET")
                                record["Petitioner"] = pet_name[:300]
                                record["Petitioner_Advocate"] = pet_adv[:300]
                        except Exception:
                            pass
                        # parse respondent
                        try:
                            if res_cell_text:
                                res_name, res_adv = split_party_and_advocate(res_cell_text, "RES")
                                record["Respondent"] = res_name[:300]
                                record["Respondent_Advocate"] = res_adv[:300]
                        except Exception:
                            pass
                        # if case_no exists, append
                        if record["Case_No"]:
                            records.append(record)
                            debug_print(f"      ✓ Extracted: {record['Case_Type']} {record['Case_No']}/{record['Year']} (S/N: {causelist_slno}) | Pet: {record['Petitioner'][:30]} | Resp: {record['Respondent'][:30]}")
                    except Exception:
                        continue
            except Exception:
                continue
        if records:
            debug_print(f"      ✅ Total: {len(records)} case(s) from Court {court_no}")
        else:
            debug_print(f"      ⚠️ No valid records for Court {court_no}")
    except Exception as e:
        debug_print(f"      ⚠️ Error in extract_case_data_from_page: {repr(e)}")
    return records

# ------------------------------
# Save / Load progress and auto-save records
# ------------------------------
def load_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_progress(progress_obj):
    try:
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(progress_obj, f, indent=2)
    except Exception as e:
        debug_print(f"⚠️ Failed to save progress: {repr(e)}")

def autosave_records(records):
    if not records:
        return
    try:
        df = pd.DataFrame(records)

        if os.path.exists(OUTPUT_EXCEL):
            existing = pd.read_excel(OUTPUT_EXCEL)
            combined = pd.concat([existing, df], ignore_index=True).drop_duplicates()
        else:
            combined = df

        # Reset ID column safely
        if "ID" in combined.columns:
            combined = combined.drop(columns=["ID"])
        combined.insert(0, "ID", range(1, len(combined) + 1))

        combined.to_excel(OUTPUT_EXCEL, index=False)
        debug_print(f"💾 Auto-saved {len(records)} new record(s) to {OUTPUT_EXCEL}")
    except Exception as e:
        debug_print(f"⚠️ Auto-save failed: {repr(e)}")

# ------------------------------
# Main scraping loop (with resume)
# ------------------------------
def main():
    ensure_folder()
    start_prevent_sleep_thread()

    all_records = []
    existing_progress = load_progress()
    # determine start date based on progress file (resume by sprint start)
    start_date = datetime(YEAR, 1, 1)
    end_date = datetime(YEAR, 12, 31)
    today = datetime.now()
    current = start_date

    if existing_progress and "current_date" in existing_progress:
        try:
            saved = datetime.strptime(existing_progress["current_date"], "%Y-%m-%d")
            if saved > current:
                current = saved
                debug_print(f"🔁 Resuming from saved date: {current.strftime('%d/%m/%Y')}")
        except Exception:
            pass

    try:
        while current <= end_date:
            sprint_end = min(current + timedelta(days=SPRINT_DAYS), end_date)
            from_str = current.strftime("%d/%m/%Y")
            to_str = sprint_end.strftime("%d/%m/%Y")
            debug_print(f"\n📅 Sprint: {from_str} → {to_str}")

            # mode: previous if sprint_end < today else daily
            mode_value = "P" if sprint_end.date() < today.date() else "D"
            debug_print(f"🔁 Mode: {'Previous' if mode_value=='P' else 'Daily & Advance'}")

            for bench_code, bench_name in BENCHES.items():
                debug_print(f"  ➤ {bench_name} | {from_str} → {to_str}")
                found_data_after_min = False

                for court_no in range(1, MAX_COURTS + 1):
                    debug_print(f"    🏛️ Court Hall {court_no}")

                    try:
                        driver.get(CAUSELIST_URL)
                        wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "form")))

                        form, bench_select = find_form_and_bench_select(driver, bench_code)
                        if not bench_select:
                            try:
                                bench_select = driver.find_element(By.NAME, "bench")
                                form = bench_select.find_element(By.XPATH, "./ancestor::form[1]")
                            except Exception:
                                raise RuntimeError("Bench select not found.")

                        sel = Select(bench_select)
                        sel.select_by_value(bench_code)
                        dispatch_events_on(bench_select)
                        time.sleep(0.6)

                        # set Search By to Court Hall if available
                        try:
                            search_by_select = None
                            if form:
                                selects = form.find_elements(By.TAG_NAME, "select")
                            else:
                                selects = driver.find_elements(By.TAG_NAME, "select")
                            for s in selects:
                                try:
                                    options = s.find_elements(By.TAG_NAME, "option")
                                    option_texts = [opt.text.strip() for opt in options]
                                    if "Court Hall" in option_texts and "Judge" in option_texts:
                                        search_by_select = s
                                        break
                                except Exception:
                                    continue
                            if search_by_select:
                                search_by_sel = Select(search_by_select)
                                # try to select matching visible text
                                try:
                                    search_by_sel.select_by_visible_text("Court Hall")
                                except Exception:
                                    # fallback: choose first option that contains Court Hall
                                    for opt in search_by_select.find_elements(By.TAG_NAME, "option"):
                                        if "Court Hall" in (opt.text or ""):
                                            search_by_sel.select_by_visible_text(opt.text)
                                            break
                                dispatch_events_on(search_by_select)
                                debug_print(f"      ✓ Fixed 'Search By' to 'Court Hall'")
                                time.sleep(0.5)
                        except Exception as e:
                            debug_print(f"      ⚠️ Error setting 'Search By': {repr(e)}")

                        # select P/D radio if present
                        try:
                            radios = (form.find_elements(By.XPATH, ".//input[@type='radio']") if form else driver.find_elements(By.XPATH, "//input[@type='radio']"))
                            for r in radios:
                                if (r.get_attribute("value") or "").strip().upper() == mode_value:
                                    driver.execute_script("arguments[0].click();", r)
                                    break
                        except Exception:
                            pass
                        time.sleep(0.4)

                        # find court select
                        court_select = find_court_select(form)
                        if court_select:
                            try:
                                court_sel = Select(court_select)
                                all_options = court_sel.options
                                court_found = False
                                for opt in all_options:
                                    opt_text = opt.text.strip()
                                    opt_value = (opt.get_attribute("value") or "").strip()
                                    if (f"COURT HALL - {court_no}" in opt_text.upper() or 
                                        f"COURT HALL {court_no}" in opt_text.upper() or
                                        f"HALL - {court_no}" in opt_text.upper() or
                                        opt_value == str(court_no)):
                                        try:
                                            court_sel.select_by_visible_text(opt_text)
                                            court_found = True
                                            debug_print(f"      ✓ Selected: {opt_text}")
                                            break
                                        except Exception:
                                            try:
                                                court_sel.select_by_value(opt_value)
                                                court_found = True
                                                debug_print(f"      ✓ Selected Court by value: {opt_value}")
                                                break
                                            except Exception:
                                                continue
                                if not court_found:
                                    debug_print(f"      ⚠️ Court Hall {court_no} not found")
                                    if court_no <= MIN_COURTS_TO_CHECK:
                                        debug_print(f"      ℹ️ Continuing (court {court_no}/{MIN_COURTS_TO_CHECK} mandatory)")
                                        continue
                                    else:
                                        debug_print(f"      ⏩ Stopping - court not available")
                                        break
                                dispatch_events_on(court_select)
                                time.sleep(0.5)
                            except Exception as e:
                                debug_print(f"      ⚠️ Error selecting Court {court_no}: {repr(e)}")
                                if court_no <= MIN_COURTS_TO_CHECK:
                                    debug_print(f"      ℹ️ Continuing despite error")
                                    continue
                                else:
                                    break
                        else:
                            debug_print(f"      ⚠️ Court select dropdown not found")
                            if court_no > 1:
                                break

                        # find date inputs
                        if form:
                            date_inputs = form.find_elements(By.XPATH, ".//input[@type='text' or @type='date']")
                        else:
                            date_inputs = driver.find_elements(By.XPATH, "//input[@type='text' or @type='date']")
                        visible_date_inputs = [i for i in date_inputs if i.is_displayed()]
                        if len(visible_date_inputs) >= 2:
                            from_el = visible_date_inputs[0]
                            to_el = visible_date_inputs[1]
                        else:
                            raise RuntimeError("Date inputs not found.")
                        ok = set_date_on_elements(from_el, to_el, from_str, to_str, debug_label=f"{bench_name} Court {court_no}")
                        if not ok:
                            raise RuntimeError(f"Failed to set dates")
                        time.sleep(0.8)
                        if form:
                            click_get_button_in_form(form)
                        time.sleep(1.5)

                        # extract cases
                        court_records = extract_case_data_from_page(bench_name, from_str, court_no)
                        if court_records:
                            all_records.extend(court_records)
                            # write autosave immediately for the batch found
                            autosave_records(court_records)
                            if court_no > MIN_COURTS_TO_CHECK:
                                found_data_after_min = True
                                debug_print(f"      ✅ Found data beyond court {MIN_COURTS_TO_CHECK}")
                        else:
                            debug_print(f"      ⚠️ No records for Court {court_no}")
                            if court_no > MIN_COURTS_TO_CHECK and not found_data_after_min:
                                debug_print(f"    ⏩ Stopping - no data after {MIN_COURTS_TO_CHECK} courts")
                                break
                    except Exception as e:
                        debug_print(f"    ❌ Error for Court {court_no}: {repr(e)}")
                        if court_no <= MIN_COURTS_TO_CHECK:
                            debug_print(f"      ℹ️ Continuing despite error")
                            continue
                        else:
                            debug_print(f"    ⏩ Stopping due to error")
                            break

            # finished sprint across benches -> save progress and move on
            # progress saved as next start date
            next_start = (sprint_end + timedelta(days=1)).strftime("%Y-%m-%d")
            save_progress({"current_date": next_start})
            debug_print(f"🔖 Progress saved. Next start will be {next_start}")
            # small pause to reduce server load
            time.sleep(1.0)
            current = sprint_end + timedelta(days=1)

    finally:
        debug_print("\n✅ Data collection finished or interrupted. Closing browser.")
        try:
            driver.quit()
        except Exception:
            pass
        stop_prevent_sleep_thread()

        # final save of accumulated records (if any)
        if all_records:
            try:
                df = pd.DataFrame(all_records)
                if os.path.exists(OUTPUT_EXCEL):
                    existing = pd.read_excel(OUTPUT_EXCEL)
                    combined = pd.concat([existing, df], ignore_index=True).drop_duplicates()
                else:
                    combined = df
                combined.insert(0, "ID", range(1, len(combined) + 1))
                combined.to_excel(OUTPUT_EXCEL, index=False)
                debug_print(f"✅ Excel saved with {len(combined)} records: {OUTPUT_EXCEL}")
            except Exception as e:
                debug_print(f"⚠️ Final save failed: {repr(e)}")
        else:
            debug_print("⚠️ No records extracted in this run. Excel not created/updated.")

if __name__ == "__main__":
    main()