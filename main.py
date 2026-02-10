import time
import pandas as pd
import gspread
import os
import json
from datetime import datetime
from dotenv import load_dotenv 
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

# Load local env vars
load_dotenv()

# --- CONFIGURATION ---
LOGIN_URL = "https://erp.xlri.ac.in/login.htm"
USERNAME = os.getenv("XLRI_USERNAME")
PASSWORD = os.getenv("XLRI_PASSWORD")

SHEET_NAME = "BM 27 Term 3"
BATCH_NAME = "PGDM (BMD) (2025-2027) III"

# --- COLORS ---
SUBJECT_COLORS = {
    "OPR":  {"red": 0.65, "green": 0.82, "blue": 1.0},
    "STM":  {"red": 1.0,  "green": 0.85, "blue": 0.6},
    "FM2":  {"red": 0.6,  "green": 1.0,  "blue": 0.6},
    "BRM":  {"red": 0.8,  "green": 0.7,  "blue": 1.0},
    "ORM2": {"red": 1.0,  "green": 0.95, "blue": 0.6},
    "HRM":  {"red": 1.0,  "green": 0.65, "blue": 0.65},
    "BLA":  {"red": 0.5,  "green": 0.9,  "blue": 0.9},
    "BOB2": {"red": 0.6,  "green": 0.6,  "blue": 1.0},
    "Act":  {"red": 0.85, "green": 0.85, "blue": 0.85}
}

COURSE_DETAILS_LIST = [
    {"code": "FM2",  "name": "Financial Management - II",            "credit": 3, "faculty": "Dr Vaibhav Lalwani, Dr Gourav Vallabh"},
    {"code": "ORM2", "name": "Operations Management - II",           "credit": 3, "faculty": "Dipankar Bose, Alok Raj"},
    {"code": "BOB2", "name": "Org. Structure, Design and Change",    "credit": 3, "faculty": "Dipak Bhattacharyya"},
    {"code": "STM",  "name": "Strategic Management",                 "credit": 3, "faculty": "Dr Faisal Mohammad Ahsan, Dr Sanchayan Nath"},
    {"code": "BLA",  "name": "Business Law",                         "credit": 2, "faculty": "Dr Paramjyot Singh"},
    {"code": "HRM",  "name": "Human Resource Management",            "credit": 2, "faculty": "Dr Abhishek Singh, Mr Harbhajan Singh"},
    {"code": "OPR",  "name": "Operations Research",                  "credit": 2, "faculty": "Dr Sayan Mukherjee, Visiting Faculty"},
    {"code": "BRM",  "name": "Business Research Methods",            "credit": 2, "faculty": "Dr Sakhhi Chhabra"}
]

def clean_slate(sheet):
    """Deep Clean."""
    try:
        sheet.clear()
        sheet.format("A:Z", {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}})
        sheet.spreadsheet.batch_update({
            "requests": [{"unmergeCells": {"range": {"sheetId": sheet.id}}}]
        })
        metadata = sheet.spreadsheet.fetch_sheet_metadata()
        sheet_meta = next((s for s in metadata['sheets'] if s['properties']['title'] == sheet.title), None)
        if sheet_meta and 'conditionalFormats' in sheet_meta:
            requests = [{"deleteConditionalFormatRule": {"index": 0, "sheetId": sheet.id}} 
                        for _ in range(len(sheet_meta['conditionalFormats']))]
            sheet.spreadsheet.batch_update({"requests": requests})
    except Exception as e:
        print(f"Warning during cleaning: {e}")

def parse_cell_data(html_content):
    if not html_content or not html_content.strip():
        return []
    soup = BeautifulSoup(html_content, 'html.parser')
    for br in soup.find_all("br"): br.replace_with("\n")
    full_text = soup.get_text().replace('\xa0', ' ')
    raw_blocks = full_text.split('----------------------------------------------')
    parsed_classes = []
    for block in raw_blocks:
        if not block.strip(): continue
        lines = [l.strip() for l in block.split('\n') if l.strip()]
        display_lines = []
        for line in lines:
            line = line.replace('[III]', '').replace('[2025-2027]', '')
            if not line.strip() or line.startswith('[-]') or line == '[0]': continue
            display_lines.append(line)
        final_text = "\n".join(display_lines)
        c_type = "CLASS"
        if any(x in final_text.lower() for x in ['quiz', 'mid term', 'end term', 'exam']): c_type = "EXAM"
        elif "maxi" in final_text.lower(): c_type = "MAXI"
        parsed_classes.append({"text": final_text, "full_content": block, "type": c_type})
    return parsed_classes

def get_shifted_slots(date_str, s0, s1, s2, s3, s4):
    """
    Inputs: 
    s0: 08:30-21:30 (Weird Slot)
    s1: 09:00
    s2: 11:00
    s3: 14:00
    s4: 16:00
    """
    try:
        parts = date_str.split()
        if len(parts) >= 2:
            clean_date_str = f"{parts[0]} {parts[1]}"
            dt = datetime.strptime(clean_date_str, "%b %d,%Y")
            
            # --- SHIFT LOGIC ---
            period_1 = (dt <= datetime(2026, 1, 26))
            period_2 = (datetime(2026, 1, 29) <= dt <= datetime(2026, 1, 31))
            
            if period_1 or period_2:
                # Cold Shift: 9->11, 11->2, 2->4. 
                # s0 (Weird) stays. s4 (Original 4pm) falls off.
                return [s0, [], s1, s2, s3]
            else:
                # Normal Schedule
                return [s0, s1, s2, s3, s4]
        else: 
            return [s0, s1, s2, s3, s4]
    except Exception as e: 
        print(f"Date Parse Warning: {e}")
        return [s0, s1, s2, s3, s4]

def update_google_sheet(sheet, data_rows):
    clean_slate(sheet)
    
    # --- UPDATED HEADERS FOR 5 SLOTS ---
    headers = [
        "Date", 
        "08:30 AM - 21:30 PM",  # Slot 0 (The Mistake/Weird Slot)
        "09:00 AM - 10:30 AM",  # Slot 1
        "11:00 AM - 12:30 PM",  # Slot 2
        "14:00 PM - 15:30 PM",  # Slot 3
        "16:00 PM - 17:30 PM"   # Slot 4
    ]
    
    final_rows = [headers]
    requests = []
    current_row_idx = 1
    
    for row in data_rows:
        date_str = row['Date']
        # Pass all 5 slots
        final_slots = get_shifted_slots(
            date_str, 
            row['Slot0'], row['Slot1'], row['Slot2'], row['Slot3'], row['Slot4']
        )
        
        max_depth = max(len(s) for s in final_slots) if final_slots else 1
        if max_depth == 0: max_depth = 1
        
        if max_depth > 1:
            requests.append({
                "mergeCells": {
                    "range": {"sheetId": sheet.id, "startRowIndex": current_row_idx, "endRowIndex": current_row_idx + max_depth, "startColumnIndex": 0, "endColumnIndex": 1},
                    "mergeType": "MERGE_ALL"
                }
            })

        for i in range(max_depth):
            current_data = ["", "", "", "", "", ""] # 6 Columns
            if i == 0: current_data[0] = date_str
            for slot_idx, slot_content in enumerate(final_slots):
                col_index = slot_idx + 1
                if i < len(slot_content):
                    cls = slot_content[i]
                    current_data[col_index] = cls['text']
                    bg_color = None
                    text_format = {"bold": True}
                    content_to_check = cls['full_content'].upper() 
                    for code in SUBJECT_COLORS:
                        if code in content_to_check:
                            bg_color = SUBJECT_COLORS[code]
                            break
                    if cls['type'] == "EXAM":
                        text_format["foregroundColor"] = {"red": 1.0, "green": 0.0, "blue": 0.0}
                        bg_color = {"red": 1.0, "green": 1.0, "blue": 1.0}
                    elif cls['type'] == "MAXI":
                        text_format["foregroundColor"] = {"red": 0.0, "green": 0.5, "blue": 0.0}

                    # Field Mask Construction
                    user_fmt = {"textFormat": text_format, "verticalAlignment": "TOP", "wrapStrategy": "WRAP"}
                    fields_list = ["textFormat", "verticalAlignment", "wrapStrategy"]
                    
                    if bg_color:
                        user_fmt["backgroundColor"] = bg_color
                        fields_list.append("backgroundColor")
                    
                    requests.append({
                        "repeatCell": {
                            "range": {"sheetId": sheet.id, "startRowIndex": current_row_idx, "endRowIndex": current_row_idx + 1, "startColumnIndex": col_index, "endColumnIndex": col_index + 1},
                            "cell": {"userEnteredFormat": user_fmt},
                            "fields": "userEnteredFormat(" + ",".join(fields_list) + ")"
                        }
                    })
            final_rows.append(current_data)
            current_row_idx += 1

    sheet.update(range_name="A1", values=final_rows)
    
    # Headers Format
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 6},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.8, "green": 0.0, "blue": 0.0}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }
    })
    # Borders
    requests.append({
        "updateBorders": {
            "range": {"sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": current_row_idx, "startColumnIndex": 0, "endColumnIndex": 6},
            "top": {"style": "SOLID", "width": 1}, "bottom": {"style": "SOLID", "width": 1}, "left": {"style": "SOLID", "width": 1}, "right": {"style": "SOLID", "width": 1}, "innerHorizontal": {"style": "SOLID", "width": 1}, "innerVertical": {"style": "SOLID", "width": 1}
        }
    })
    # Column Widths
    requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 6}, "properties": {"pixelSize": 200}, "fields": "pixelSize"}})

    if requests: sheet.spreadsheet.batch_update({"requests": requests})

def update_info_table(sheet):
    headers = ["Course Code", "Course Name", "Course Credit", "Faculty"]
    data_rows = []
    for course in COURSE_DETAILS_LIST:
        data_rows.append([course["code"], course["name"], course["credit"], course["faculty"]])
    
    sheet.update(range_name="H1", values=[headers] + data_rows)
    requests = []
    
    # Header Format (H1:K1)
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 7, "endColumnIndex": 11},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.8, "green": 0.0, "blue": 0.0}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }
    })
    # Rows
    current_row_idx = 1
    for course in COURSE_DETAILS_LIST:
        bg_color = SUBJECT_COLORS.get(course["code"], {"red": 1, "green": 1, "blue": 1})
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet.id, "startRowIndex": current_row_idx, "endRowIndex": current_row_idx + 1, "startColumnIndex": 7, "endColumnIndex": 11},
                "cell": {"userEnteredFormat": {"backgroundColor": bg_color, "textFormat": {"bold": False}, "verticalAlignment": "MIDDLE"}},
                "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)"
            }
        })
        current_row_idx += 1
    # Borders
    requests.append({
        "updateBorders": {
            "range": {"sheetId": sheet.id, "startRowIndex": 0, "endRowIndex": current_row_idx, "startColumnIndex": 7, "endColumnIndex": 11},
            "top": {"style": "SOLID", "width": 1}, "bottom": {"style": "SOLID", "width": 1}, "left": {"style": "SOLID", "width": 1}, "right": {"style": "SOLID", "width": 1}, "innerHorizontal": {"style": "SOLID", "width": 1}, "innerVertical": {"style": "SOLID", "width": 1}
        }
    })
    # Widths
    requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 7, "endIndex": 11}, "properties": {"pixelSize": 250}, "fields": "pixelSize"}})

    if requests: sheet.spreadsheet.batch_update({"requests": requests})

def fetch_and_update():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)

    try:
        print("Bot: Logging in...")
        driver.get(LOGIN_URL)
        wait.until(EC.presence_of_element_located((By.ID, "j_username"))).send_keys(USERNAME)
        driver.find_element(By.ID, "password-1").send_keys(PASSWORD)
        driver.find_element(By.XPATH, "//button[contains(text(), 'Login')]").click()
        
        print("Bot: Navigating...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Academic Schedules')]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Class Schedule System (CSS)')]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Program-wise Class Schedule')]"))).click()

        print("Bot: Selecting Batch...")
        time.sleep(2)
        driver.execute_script("document.getElementById('cmdProgrammeWiseList').style.display = 'block';")
        driver.execute_script("document.getElementById('cmdProgrammeWiseList').style.visibility = 'visible';")
        select = Select(driver.find_element(By.ID, "cmdProgrammeWiseList"))
        select.select_by_visible_text(BATCH_NAME)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", driver.find_element(By.ID, "cmdProgrammeWiseList"))
        
        print("Bot: Submitting...")
        time.sleep(5)
        submit_btn = driver.find_element(By.XPATH, "//a[contains(@onclick, 'getProgramwiseClassSchedule')]")
        submit_btn.click()
        
        print("Bot: Scraping...")
        wait.until(EC.presence_of_element_located((By.ID, "programwiseClassScheduleReport")))
        time.sleep(3) 
        
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        table = soup.find('table', {'id': 'programwiseClassScheduleReport'})
        rows = table.find("tbody").find_all("tr")
        
        scraped_data = []
        last_date = "Jan 01,2026 Thursday" 

        for row in rows:
            cells = row.find_all(['th', 'td'])
            # CHECK: We need Date + 5 Slots = 6 cells
            if len(cells) < 6: continue
            
            raw_date = cells[0].get_text(separator=" ").replace('\xa0', '').strip()
            if len(raw_date) > 2: 
                last_date = raw_date
                final_date = raw_date
            else: final_date = last_date
            
            scraped_data.append({
                "Date": final_date,
                "Slot0": parse_cell_data(str(cells[1])), # 08:30 (Weird Slot)
                "Slot1": parse_cell_data(str(cells[2])), # 09:00
                "Slot2": parse_cell_data(str(cells[3])), # 11:00
                "Slot3": parse_cell_data(str(cells[4])), # 14:00
                "Slot4": parse_cell_data(str(cells[5]))  # 16:00
            })
            
        print(f"Bot: Found {len(scraped_data)} rows. Updating Sheets...")
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        
        if os.getenv("GOOGLE_CREDENTIALS_JSON"):
            creds_dict = json.loads(os.getenv("GOOGLE_CREDENTIALS_JSON"))
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        else:
            creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
            
        client = gspread.authorize(creds)
        sheet = client.open(SHEET_NAME).sheet1
        update_google_sheet(sheet, scraped_data)
        update_info_table(sheet)
        print("Bot: SUCCESS! Schedule restored with all 5 slots.")
    except Exception as e:
        print(f"ERROR: {e}")
        driver.save_screenshot("error_final.png")
    finally:
        driver.quit()

if __name__ == "__main__":
    fetch_and_update()