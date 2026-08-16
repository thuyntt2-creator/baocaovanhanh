import os
import io
import sys
import gspread
from google.oauth2.service_account import Credentials

# Fix encoding
sys.stdout.reconfigure(encoding='utf-8')

JSON_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'credentials.json')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def test():
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    key = '1DAwY-46twFrHIs77R4p4IMuIZ6JTE-e58Aj-9Kcr5Jk'
    sh = gc_client.open_by_key(key)
    ws = sh.worksheet("gtc")
    
    print("Writing numeric 1080 and float 0.6648 to cells J2 and K2...")
    ws.update_acell('J2', 1080)
    ws.update_acell('K2', 0.6648)
    
    # Read back formatted
    print("Read back formatted:")
    print("J2:", ws.acell('J2').value)
    print("K2:", ws.acell('K2').value)
    
    # Read back unformatted
    print("Read back unformatted:")
    print("J2:", ws.get('J2', value_render_option='UNFORMATTED_VALUE'))
    print("K2:", ws.get('K2', value_render_option='UNFORMATTED_VALUE'))
    
    # Clean up J2 and K2
    ws.update_acell('J2', '')
    ws.update_acell('K2', '')

if __name__ == "__main__":
    test()
