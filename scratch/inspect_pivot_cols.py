import sys, gspread, pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
gc = gspread.service_account(filename='credentials.json')
sh = gc.open_by_key('1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ')
ws = sh.worksheet('PIVOT aging ')
rows = ws.get_all_values()
print("=== PIVOT aging sheet total rows:", len(rows), "columns:", len(rows[0]) if rows else 0)
for r_idx in range(min(6, len(rows))):
    print(f"Row {r_idx}:", rows[r_idx][:25])
