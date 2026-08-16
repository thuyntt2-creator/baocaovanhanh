import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import sys
import io

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

JSON_FILE = r"c:\Users\lap4all\Documents\Auto report\credentials.json"
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
gc_client = gspread.authorize(creds)

sheet_key = '10cq3DUggZ4vXffcxweIRTRK3qiyMeWnV8gksdGwvp7s'
sh = gc_client.open_by_key(sheet_key)

ws = sh.worksheet("No attempt")
data = ws.get_all_values()
df = pd.DataFrame(data[1:], columns=data[0])

# Filter out empty ma_don
df_active = df[df['ma_don'].str.strip() != '']
df_empty_am = df_active[df_active['am_name'].str.strip() == '']

print(f"Total rows with empty AM: {len(df_empty_am)}")
print("Unique 'buu_cuc' values with empty AM:")
print(df_empty_am['buu_cuc'].value_counts())

print("\nUnique 'warehouse_id' values with empty AM:")
print(df_empty_am['warehouse_id'].value_counts())

print("\nUnique 'warehouse_name' values with empty AM:")
print(df_empty_am['warehouse_name'].value_counts())
