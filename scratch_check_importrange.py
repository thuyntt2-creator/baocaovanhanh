import json
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc_client = gspread.authorize(creds)

try:
    sh = gc_client.open_by_key('1dvOBw93Q-W5j9kuTOO_S5Bl_dE37FJXsF1dAxSGMr5Y')
    print("Title of 1dvOBw:", sh.title)
    print("Worksheets:")
    for ws in sh.worksheets():
        print(f" - {ws.title}")
except Exception as e:
    import traceback
    print("Error class:", e.__class__.__name__)
    print("Error message:", str(e))
    traceback.print_exc()
