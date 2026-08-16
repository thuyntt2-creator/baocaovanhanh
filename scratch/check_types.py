import sys, json, gspread
from google.oauth2.credentials import Credentials

sys.stdout.reconfigure(encoding='utf-8')

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_authorized_user_file(r'c:\Users\lap4all\Documents\Auto report\authorized_user.json', scopes=scopes)
gc = gspread.authorize(creds)

sh = gc.open_by_key('1DuMW_ajrtrmLlMNslJY2UIMWygVY1cFD4QhKnX9YGNQ')
ws_lm = sh.worksheet('LM')
ws_co = sh.worksheet('cocaumoi')

# Check first 5 rows of A in LM vs cocaumoi
lm_a = ws_lm.get('A2:A6')
co_a = ws_co.get('A2:A6')

print("LM A2:A6 values:", lm_a)
print("cocaumoi A2:A6 values:", co_a)

# Check cell H2 formula vs cell value
h2_val = ws_lm.get('H2:H6')
print("LM H2:H6 values:", h2_val)
