# -*- coding: utf-8 -*-
import sys
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

SHEET_ID = "1MtbZBgRFwCWj6uQKsSqddiJ2GsTiEvKxRIPSshDa5PM"
CREDENTIALS_FILE = "credentials.json"

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)
ss = client.open_by_key(SHEET_ID)

tn_vals = ss.worksheet("thu nhập").get_all_values()
df_tn = pd.DataFrame(tn_vals[1:], columns=tn_vals[0])
print("\n--- THU NHẬP ---")
print("Unique 'Thời gian' values:", df_tn['Thời gian'].unique())

ns_vals = ss.worksheet("năng suất").get_all_values()
df_ns = pd.DataFrame(ns_vals[1:], columns=ns_vals[0])
print("\n--- NĂNG SUẤT ---")
print("Unique 'Ngay' values:", df_ns['Ngay'].unique())
