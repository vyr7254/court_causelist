import requests
from datetime import datetime, timedelta
import os
from io import BytesIO
import PyPDF2
import time
import pandas as pd
import re
from pathlib import Path

# === CONFIG ===
SAVE_DIR = r"C:\Users\Dell\OneDrive\Desktop\tshc_script_for_download_and_extraction\tshc_pdfs"
OUTPUT_EXCEL = r"C:\Users\Dell\OneDrive\Desktop\tshc_script_for_download_and_extraction\TSHC_CaseList_Extracted.xlsx"
BASE_URL = "https://tshc.gov.in/getPdfForDate"
START_DATE = datetime(2025, 1, 1)
END_DATE = datetime(2025, 10, 21)
DELAY_BETWEEN = 1

# === SETUP ===
os.makedirs(SAVE_DIR, exist_ok=True)

def download_pdf(date_obj):
    """Download PDF for a specific date"""
    date_str = date_obj.strftime("%d-%m-%Y")
    file_date = date_obj.strftime("%Y_%m_%d")
    filename = os.path.join(SAVE_DIR, f"TSHC-CauseList_{file_date}.pdf")
    
    if os.path.exists(filename):
        if os.path.getsize(filename) > 1500:
            return (date_str, filename, "⚙️ Already exists, will extract")
        else:
            os.remove(filename)
            print(f"  → Deleted corrupt file for {date_str}, re-downloading...")

    params = {"id": "0", "arc-date": date_str}

    try:
        response = requests.get(BASE_URL, params=params, timeout=20)
        if response.status_code == 200 and len(response.content) > 1500:
            pdf_file = BytesIO(response.content)
            try:
                reader = PyPDF2.PdfReader(pdf_file)
                text = "".join(page.extract_text() or "" for page in reader.pages)
                if text.strip():
                    with open(filename, "wb") as f:
                        f.write(response.content)
                    return (date_str, filename, "✅ Downloaded, will extract")
                else:
                    return (date_str, None, "⚠️ Empty PDF, skipped")
            except Exception as e:
                return (date_str, None, f"⚠️ PDF read error: {e}")
        else:
            return (date_str, None, "❌ No valid PDF")
    except Exception as e:
        return (date_str, None, f"❌ Error: {e}")
    finally:
        time.sleep(DELAY_BETWEEN)

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file"""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"
            return text
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return ""

def parse_date_from_header(header_text):
    """Parse date from header text like 'Thursday the 2nd day of January 2025'"""
    date_pattern = r'(\d+)(?:st|nd|rd|th)?\s+day\s+of\s+(\w+)\s+(\d{4})'
    match = re.search(date_pattern, header_text, re.IGNORECASE)
    if match:
        day = match.group(1)
        month = match.group(2)
        year = match.group(3)
        try:
            date_obj = datetime.strptime(f"{day} {month} {year}", "%d %B %Y")
            return date_obj.strftime("%d/%m/%Y")
        except:
            pass
    return ""

def clean_text(text):
    """Clean text by removing extra spaces"""
    if not text:
        return ""
    return re.sub(r'\s+', ' ', text).strip()

def extract_cases_from_pdf(pdf_path):
    """
    Extract all case information from PDF with proper column-based parsing
    Handles both table formats:
    1. WITH Party Details: SNO | CASE | PARTY DETAILS | PETITIONER ADV | RESPONDENT ADV | DISTRICT
    2. WITHOUT Party Details: SNO | CASE | PETITIONER ADV | RESPONDENT ADV | DISTRICT
    """
    text = extract_text_from_pdf(pdf_path)
    if not text:
        return []
    
    pdf_name = Path(pdf_path).name
    lines = text.split('\n')
    
    cases = []
    
    # Current context
    current_court_hall = ""
    current_cause_date = ""
    current_time = ""
    current_chief_justice = ""
    current_section = ""
    
    in_table = False
    table_has_party_details = False
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        if not line:
            i += 1
            continue
        
        # Detect COURT NO header
        court_match = re.search(r'COURT\s+NO\.?\s*(\d+)', line, re.IGNORECASE)
        if court_match:
            current_court_hall = court_match.group(1)
            current_chief_justice = ""
            current_cause_date = ""
            current_time = ""
            in_table = False
            
            # Extract Chief Justice (next few lines)
            for j in range(i+1, min(i+10, len(lines))):
                justice_line = lines[j].strip()
                if 'HONOURABLE' in justice_line.upper() and 'JUSTICE' in justice_line.upper():
                    justice_parts = [justice_line]
                    for k in range(j+1, min(j+3, len(lines))):
                        next_line = lines[k].strip()
                        if next_line and 'To be heard' not in next_line and 'COURT NO' not in next_line:
                            justice_parts.append(next_line)
                        else:
                            break
                    current_chief_justice = clean_text(' '.join(justice_parts))
                    break
            
            # Extract date and time
            for j in range(i+1, min(i+15, len(lines))):
                date_line = lines[j].strip()
                if 'day of' in date_line.lower():
                    current_cause_date = parse_date_from_header(date_line)
                    time_match = re.search(r'(\d{1,2}:\d{2}\s*(?:AM|PM))', date_line, re.IGNORECASE)
                    if time_match:
                        current_time = time_match.group(1).strip()
                    break
            
            i += 1
            continue
        
        # Detect section headers
        if line.startswith('FOR ') and line.isupper():
            current_section = line
            i += 1
            continue
        
        # Detect table header
        if re.search(r'\bSNO\b.*\bCASE\b', line, re.IGNORECASE):
            in_table = True
            table_has_party_details = 'PARTY DETAILS' in line.upper()
            i += 1
            continue
        
        # Parse case rows
        if in_table and current_court_hall:
            # Check if line starts with SNO (number)
            sno_match = re.match(r'^(\d+)\s+', line)
            
            if sno_match:
                causelist_slno = sno_match.group(1)
                
                # Extract case number pattern: CASE_TYPE/NUMBER/YEAR
                case_pattern = r'([A-Z]+)/(\d+)/(\d{4})'
                case_match = re.search(case_pattern, line)
                
                if case_match:
                    case_type = case_match.group(1)
                    case_number = case_match.group(2)
                    case_year = case_match.group(3)
                    
                    # Get the position where case number ends
                    case_end_pos = case_match.end()
                    
                    # Get everything after the case number
                    remaining_line = line[case_end_pos:].strip()
                    
                    # Collect continuation lines (party details, advocates may span multiple lines)
                    continuation_lines = [remaining_line]
                    for j in range(i+1, min(i+8, len(lines))):
                        next_line = lines[j].strip()
                        # Stop if we hit next case or section
                        if re.match(r'^\d+\s+[A-Z]+/\d+/\d{4}', next_line):
                            break
                        if next_line.startswith('FOR ') and next_line.isupper():
                            break
                        if next_line:
                            continuation_lines.append(next_line)
                    
                    # Join all continuation lines
                    full_data = ' '.join(continuation_lines)
                    full_data = clean_text(full_data)
                    
                    # Initialize case
                    case_data = {
                        'causelist_slno': causelist_slno,
                        'court_hall_number': current_court_hall,
                        'case_number': case_number,
                        'case_type': case_type,
                        'case_year': case_year,
                        'bench_name': 'HYDERABAD',
                        'cause_date': current_cause_date,
                        'time': current_time,
                        'chief_justice': current_chief_justice,
                        'section': current_section,
                        'petitioner': '',
                        'respondent': '',
                        'petitioner_advocate': '',
                        'respondent_advocate': '',
                        'particulars': 'list downloaded',
                        'pdf_name': pdf_name
                    }
                    
                    # Parse based on table format
                    if table_has_party_details:
                        # Format: PARTY DETAILS | PETITIONER ADV | RESPONDENT ADV | DISTRICT
                        
                        # Step 1: Extract Party Details (look for Vs pattern)
                        vs_match = re.search(r'^(.*?)\s+(?:Vs|V/s|V/S|VS)\s+(.+?)(?=\s{2,}|$)', full_data, re.IGNORECASE)
                        if vs_match:
                            petitioner_text = vs_match.group(1).strip()
                            respondent_full = vs_match.group(2).strip()
                            
                            # Petitioner is straightforward
                            case_data['petitioner'] = petitioner_text
                            
                            # Respondent ends where next column starts (usually before 2+ spaces or uppercase names)
                            # Look for where petitioner advocate starts (usually all caps name)
                            resp_parts = re.split(r'\s{2,}', respondent_full)
                            if resp_parts:
                                case_data['respondent'] = resp_parts[0].strip()
                            
                            # Step 2: Extract Petitioner Advocate (after party details, before respondent adv)
                            # Look for pattern after respondent, typically 2+ spaces followed by name
                            # The advocates are usually in format: NAME1  NAME2  DISTRICT
                            remaining_after_parties = full_data[vs_match.end():].strip()
                            
                            # Split by 2+ spaces to get columns
                            columns = re.split(r'\s{2,}', remaining_after_parties)
                            columns = [c.strip() for c in columns if c.strip()]
                            
                            # Typically: [PETITIONER_ADV, RESPONDENT_ADV, DISTRICT] or [RESPONDENT_ADV, DISTRICT]
                            if len(columns) >= 2:
                                # First column is petitioner advocate
                                case_data['petitioner_advocate'] = columns[0]
                                # Second column is respondent advocate
                                case_data['respondent_advocate'] = columns[1]
                            elif len(columns) == 1:
                                # Only one advocate listed
                                case_data['petitioner_advocate'] = columns[0]
                        
                        else:
                            # No Vs pattern found, try alternative parsing
                            # Split by 2+ spaces
                            parts = re.split(r'\s{2,}', full_data)
                            parts = [p.strip() for p in parts if p.strip()]
                            
                            if len(parts) >= 2:
                                # Last parts are usually advocates
                                if len(parts) >= 3:
                                    case_data['petitioner_advocate'] = parts[-3]
                                    case_data['respondent_advocate'] = parts[-2]
                                else:
                                    case_data['petitioner_advocate'] = parts[0]
                                    case_data['respondent_advocate'] = parts[1]
                    
                    else:
                        # Format: PETITIONER ADV | RESPONDENT ADV | DISTRICT (NO Party Details column)
                        
                        # Split by 2+ spaces to separate columns
                        columns = re.split(r'\s{2,}', full_data)
                        columns = [c.strip() for c in columns if c.strip()]
                        
                        # Expected format: [PETITIONER_ADV, RESPONDENT_ADV, DISTRICT]
                        if len(columns) >= 2:
                            case_data['petitioner_advocate'] = columns[0]
                            case_data['respondent_advocate'] = columns[1]
                        elif len(columns) == 1:
                            case_data['petitioner_advocate'] = columns[0]
                    
                    # Clean all extracted fields
                    case_data['petitioner'] = clean_text(case_data['petitioner'])
                    case_data['respondent'] = clean_text(case_data['respondent'])
                    case_data['petitioner_advocate'] = clean_text(case_data['petitioner_advocate'])
                    case_data['respondent_advocate'] = clean_text(case_data['respondent_advocate'])
                    
                    # Remove DISTRICT names from respondent advocate if present
                    district_pattern = r'\s+(HYDERABAD|RANGAREDDY|WARANGAL|KARIMNAGAR|NIZAMABAD|KHAMMAM|ADILABAD|MEDAK|NALGONDA|MAHBUBNAGAR).*$'
                    case_data['respondent_advocate'] = re.sub(district_pattern, '', case_data['respondent_advocate'], flags=re.IGNORECASE)
                    case_data['respondent_advocate'] = clean_text(case_data['respondent_advocate'])
                    
                    cases.append(case_data)
        
        i += 1
    
    return cases

def append_to_excel(new_cases):
    """Append new cases to existing Excel file or create new one"""
    if not new_cases:
        return
    
    # Define column order
    columns = [
        'id', 'causelist_slno', 'court_hall_number', 'case_number', 
        'case_type', 'case_year', 'bench_name', 'cause_date', 
        'time', 'chief_justice', 'section', 'petitioner', 'respondent', 
        'petitioner_advocate', 'respondent_advocate', 'particulars', 'pdf_name'
    ]
    
    # Load existing data if file exists
    if os.path.exists(OUTPUT_EXCEL):
        try:
            existing_df = pd.read_excel(OUTPUT_EXCEL)
            start_id = existing_df['id'].max() + 1 if 'id' in existing_df.columns and len(existing_df) > 0 else 1
        except Exception as e:
            print(f"  ⚠️ Error reading existing Excel: {e}")
            existing_df = pd.DataFrame(columns=columns)
            start_id = 1
    else:
        existing_df = pd.DataFrame(columns=columns)
        start_id = 1
    
    # Add ID to new cases
    for idx, case in enumerate(new_cases, start=start_id):
        case['id'] = idx
    
    # Create DataFrame for new cases
    new_df = pd.DataFrame(new_cases, columns=columns)
    
    # Combine with existing data
    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    
    # Clean up text columns only
    text_columns = ['causelist_slno', 'court_hall_number', 'case_number', 
                    'case_type', 'case_year', 'bench_name', 'cause_date', 
                    'time', 'chief_justice', 'section', 'petitioner', 'respondent', 
                    'petitioner_advocate', 'respondent_advocate', 'particulars', 'pdf_name']
    
    for col in text_columns:
        if col in combined_df.columns:
            combined_df[col] = combined_df[col].fillna('')
            combined_df[col] = combined_df[col].astype(str).str.strip()
            combined_df[col] = combined_df[col].replace('nan', '')
    
    # Save to Excel
    try:
        with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='TSHC Cases')
            
            worksheet = writer.sheets['TSHC Cases']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 60)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"  ✅ Appended {len(new_cases)} cases (Total: {len(combined_df)})")
    except Exception as e:
        print(f"  ❌ Error saving Excel: {e}")

def main():
    print("="*70)
    print("TSHC CAUSE LIST DOWNLOADER & CASE EXTRACTOR (COLUMN-BASED v3)")
    print("="*70)
    
    # Prepare date range
    all_dates = []
    current_date = START_DATE
    while current_date <= END_DATE:
        all_dates.append(current_date)
        current_date += timedelta(days=1)

    success, skipped, failed = 0, 0, 0

    print(f"\n🚀 Processing {len(all_dates)} days (Download + Extract + Append)...\n")

    for date_obj in all_dates:
        date_str = date_obj.strftime("%d-%m-%Y")
        
        # Download PDF
        date_str_result, filename, result = download_pdf(date_obj)
        print(f"{date_str}: {result}")
        
        # Extract and append cases immediately
        if filename and os.path.exists(filename):
            print(f"  → Extracting cases from: {Path(filename).name}")
            cases = extract_cases_from_pdf(filename)
            print(f"  → Extracted {len(cases)} cases")
            
            # Append to Excel immediately
            if cases:
                append_to_excel(cases)
            
            if "Downloaded" in result:
                success += 1
            elif "Already exists" in result:
                skipped += 1
        else:
            failed += 1
        
        time.sleep(0.5)

    print("\n" + "="*70)
    print(f"✅ Downloaded: {success}")
    print(f"⚙️ Skipped/Existing: {skipped}")
    print(f"❌ Failed: {failed}")
    print(f"📅 Total dates processed: {success + skipped + failed}")
    
    # Print final summary
    if os.path.exists(OUTPUT_EXCEL):
        final_df = pd.read_excel(OUTPUT_EXCEL)
        print(f"\n📊 FINAL EXCEL SUMMARY:")
        print(f"  Total cases: {len(final_df)}")
        if len(final_df) > 0:
            print(f"  Date range: {final_df['cause_date'].min()} to {final_df['cause_date'].max()}")
            print(f"  Courts covered: {final_df['court_hall_number'].nunique()}")
            print(f"  Cases with petitioner: {(final_df['petitioner'].str.strip() != '').sum()}")
            print(f"  Cases with pet. advocate: {(final_df['petitioner_advocate'].str.strip() != '').sum()}")
            print(f"  Cases with resp. advocate: {(final_df['respondent_advocate'].str.strip() != '').sum()}")
    
    print("="*70)
    print("✅ PROCESS COMPLETED!")
    print("="*70)

if __name__ == "__main__":
    main()
