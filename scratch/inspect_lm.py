import sys, json, gspread
from google.oauth2.credentials import Credentials

sys.stdout.reconfigure(encoding='utf-8')

scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_authorized_user_file(r'c:\Users\lap4all\Documents\Auto report\authorized_user.json', scopes=scopes)
gc = gspread.authorize(creds)

sh = gc.open_by_key('1DuMW_ajrtrmLlMNslJY2UIMWygVY1cFD4QhKnX9YGNQ')
ws_lm = sh.worksheet('LM')

print(f"Sheet LM: {ws_lm.row_count} rows, {ws_lm.col_count} cols")
headers = ws_lm.row_values(1)
print(f"Headers ({len(headers)} cols): {headers[:10]}")

# Get all values to check actual filled rows
all_vals = ws_lm.get_all_values()
filled_rows = len(all_vals)
print(f"Total rows with data: {filled_rows}")

# Check sample formulas in row 2
sample_formulas = ws_lm.get('A2:Z2', value_render_option='FORMULA')
print(f"Row 2 formulas: {sample_formulas}")
