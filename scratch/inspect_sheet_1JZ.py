import os
import sys
import time

BASE_DIR = r"c:\Users\lap4all\Documents\Auto report"
output_path = os.path.join(BASE_DIR, "scratch", "inspect_sheet_1JZ_output.txt")

def log(msg):
    with open(output_path, "a", encoding="utf-8") as out:
        out.write(f"[{time.strftime('%H:%M:%S')}] {msg}\n")

# Clean output first
with open(output_path, "w", encoding="utf-8") as out:
    out.write("Diagnostic Start\n")

log("Step 1: Importing packages...")
try:
    import gspread
    from google.oauth2.service_account import Credentials
    log("Step 1: Imports succeeded!")
except Exception as e:
    log(f"Step 1: Imports failed: {e}")
    sys.exit(1)

log("Step 2: Loading credentials...")
try:
    JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(credentials)
    log("Step 2: Credentials loaded and gspread client authorized!")
except Exception as e:
    log(f"Step 2: Credentials loading failed: {e}")
    sys.exit(1)

log("Step 3: Opening Main Spreadsheet...")
try:
    sheet_id = "1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ"
    sh = gc_client.open_by_key(sheet_id)
    log(f"Step 3: Opened Main Spreadsheet! Title: {sh.title}")
    
    log("Step 4: Fetching worksheets...")
    worksheets = sh.worksheets()
    log(f"Step 4: Worksheets: {[w.title for w in worksheets]}")
    
    for w in worksheets:
        log(f"Step 5: Reading sheet '{w.title}' headers...")
        try:
            row1 = w.row_values(1)
            log(f"  Headers: {row1[:25]}")
            row2 = w.row_values(2)
            log(f"  Sample row 2: {row2[:25]}")
            row3 = w.row_values(3)
            log(f"  Sample row 3: {row3[:25]}")
        except Exception as e:
            log(f"  Error reading sheet '{w.title}': {e}")
except Exception as e:
    log(f"Step 3/4: Failed to read Main Spreadsheet: {e}")

log("Step 6: Opening Rot LC Spreadsheet...")
try:
    sheet_id = "14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4"
    sh = gc_client.open_by_key(sheet_id)
    log(f"Step 6: Opened Rot LC Spreadsheet! Title: {sh.title}")
    
    log("Step 7: Fetching worksheets...")
    worksheets = sh.worksheets()
    log(f"Step 7: Worksheets: {[w.title for w in worksheets]}")
    for w in worksheets:
        log(f"Step 8: Reading sheet '{w.title}' headers...")
        try:
            row1 = w.row_values(1)
            log(f"  Headers: {row1[:25]}")
            row2 = w.row_values(2)
            log(f"  Sample row 2: {row2[:25]}")
        except Exception as e:
            log(f"  Error reading sheet '{w.title}': {e}")
except Exception as e:
    log(f"Step 6/7: Failed to read Rot LC Spreadsheet: {e}")

log("Diagnostic finished!")
