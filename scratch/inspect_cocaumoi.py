import sys, json, gspread
from google.oauth2.credentials import Credentials

sys.stdout.reconfigure(encoding='utf-8')

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_authorized_user_file(r'c:\Users\lap4all\Documents\Auto report\authorized_user.json', scopes=scopes)
gc = gspread.authorize(creds)

sh = gc.open_by_key('1DuMW_ajrtrmLlMNslJY2UIMWygVY1cFD4QhKnX9YGNQ')
ws_co = sh.worksheet('cocaumoi')

headers_co = ws_co.row_values(1)
print("cocaumoi headers:", headers_co)
print("cocaumoi sample row 2:", ws_co.row_values(2))
