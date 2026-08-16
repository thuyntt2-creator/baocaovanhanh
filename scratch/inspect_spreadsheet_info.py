import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def inspect():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    
    # Check sheet last modified time using Drive API
    drive_service = build('drive', 'v3', credentials=creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    
    file_info = drive_service.files().get(fileId=key, fields='name, modifiedTime').execute()
    print("Spreadsheet Info from Drive API:")
    print("Name:", file_info.get('name'))
    print("Modified Time:", file_info.get('modifiedTime'))
    
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(key)
    print("\nWorksheets in spreadsheet:")
    for ws in sh.worksheets():
        print(f"  - '{ws.title}' (ID: {ws.id})")

if __name__ == "__main__":
    inspect()
