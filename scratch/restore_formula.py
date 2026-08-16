import sys
import io
import os
import gspread
from google.oauth2.service_account import Credentials

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SHEET_KEY = '14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4'
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def restore():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    ws_rot = sh.worksheet('data rớt LC')
    ws_rot.clear()
    
    formula = '=QUERY(IMPORTRANGE("1LdnXYFnkACLUnVK2-I7MQ8YGRRNc88RcXjDu-lbQQ2Q", "NTB!I1:T"), "SELECT * WHERE Col1 IS NOT NULL", 1)'
    print("Writing formula back to A1 of sheet 'data rớt LC'...")
    ws_rot.update_acell('A1', formula)
    print("✅ Formula restored successfully in 'data rớt LC'!")

if __name__ == "__main__":
    restore()
