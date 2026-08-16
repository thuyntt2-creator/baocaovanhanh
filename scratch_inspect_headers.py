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

sh = gc_client.open_by_key('1l2jZGLFoqxta2jz1RRJDTa1x15nHmXxGa2ZXhWXjM4M')
ws = sh.worksheet("Đơn giao aging trên 5 ngày")
headers = ws.get_all_values()[0]
print("Headers of Đơn giao aging trên 5 ngày on 1l2j:")
print(headers)
