import os
import sys
import json
import gspread
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

SPREADSHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"

def get_service_account():
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json"),
        r"C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json",
        "credentials.json",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    raise FileNotFoundError("credentials.json not found")

def main():
    cred_file = get_service_account()
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(cred_file, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    print("Worksheets:")
    for ws in sh.worksheets():
        print(f"  Title: {ws.title}, ID: {ws.id}")

    target_ws = None
    for ws in sh.worksheets():
        if str(ws.id) == "1525919864" or "lời" in ws.title.lower() or "loi" in ws.title.lower():
            target_ws = ws
            break
    
    if not target_ws:
        target_ws = sh.worksheets()[0]

    print(f"\nUsing worksheet: {target_ws.title} (id={target_ws.id})")
    rows = target_ws.get_all_values()
    print(f"Total rows: {len(rows)}")
    for i, row in enumerate(rows):
        print(f"Row {i+1}: {row}")

if __name__ == "__main__":
    main()
