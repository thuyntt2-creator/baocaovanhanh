import sys, json, gspread
from google.oauth2.credentials import Credentials

sys.stdout.reconfigure(encoding='utf-8')

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_authorized_user_file(r'c:\Users\lap4all\Documents\Auto report\authorized_user.json', scopes=scopes)
gc = gspread.authorize(creds)

sh = gc.open_by_key('1DuMW_ajrtrmLlMNslJY2UIMWygVY1cFD4QhKnX9YGNQ')
ws_lm = sh.worksheet('LM')

# Check formulas across columns in row 2
row2_f = ws_lm.get('A2:Z2', value_render_option='FORMULA')[0]
headers = ws_lm.row_values(1)

print("Columns in LM sheet:")
for idx, (h, f) in enumerate(zip(headers, row2_f)):
    print(f" Col {idx+1} ({h}): Formula = '{f}'")

# Check if there are formulas in other rows
print("\nChecking row 44645 formula:")
row_last_f = ws_lm.get('A44645:Z44645', value_render_option='FORMULA')[0]
for idx, (h, f) in enumerate(zip(headers, row_last_f)):
    if f.startswith('='):
        print(f" Col {idx+1} ({h}): Formula = '{f}'")

# Check if there are formulas beyond row 44646
print("\nChecking row 45000 formula (in empty region):")
row_empty_f = ws_lm.get('A45000:Z45000', value_render_option='FORMULA')
print(row_empty_f)
