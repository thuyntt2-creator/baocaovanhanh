import json
import sys
import io
import gspread
from google.oauth2.service_account import Credentials

try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except Exception:
    pass

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc_client = gspread.authorize(creds)

new_key = '1vCxSTNgSpO9ETvVRElGyuGc7lnx7LxLRhAB4-lJMHLU'
sh = gc_client.open_by_key(new_key)

print("Checking PUSH REGION sheet:")
try:
    ws_push = sh.worksheet("PUSH REGION")
    print("PUSH REGION rows count:", len(ws_push.get_all_values()))
    print("First 3 rows:")
    print(ws_push.get_all_values()[:3])
except Exception as e:
    print("Error in PUSH REGION:", e)

print("\nChecking Lượt gán sheet:")
try:
    ws_gan = sh.worksheet("Lượt gán")
    print("Lượt gán rows count:", len(ws_gan.get_all_values()))
    print("First 3 rows:")
    print(ws_gan.get_all_values()[:3])
except Exception as e:
    print("Error in Lượt gán:", e)
