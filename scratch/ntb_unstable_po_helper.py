import os
import sys
import json
import gspread
from google.oauth2.credentials import Credentials as UserCredentials
from google.oauth2.service_account import Credentials as SACredentials

sys.stdout.reconfigure(encoding='utf-8')
scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def get_gc():
    if os.path.exists('authorized_user.json'):
        try:
            return gspread.authorize(UserCredentials.from_authorized_user_file('authorized_user.json', scopes=scopes))
        except Exception:
            pass
    if os.path.exists('credentials.json'):
        try:
            return gspread.authorize(SACredentials.from_service_account_file('credentials.json', scopes=scopes))
        except Exception:
            pass
    return None

def fetch_ntb_unstable_po_data(client=None):
    gc = client or get_gc()
    if not gc:
        return {"error": "Authentication failed"}

    rows = []
    # 1. Try reading from Dashboard Sheet 1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ (tab NTB, gid 1301452336)
    try:
        sh = gc.open_by_key("1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ")
        ws = sh.get_worksheet_by_id(1301452336)
        rows = ws.get_all_values()
    except Exception as e:
        print(f"⚠️ Warning reading 1JZ NTB tab: {e}")

    # Fallback to Master Sheet 14cN4YAf01NqKtREtmhBngRP58OCdgby-PjVoixVlpZo (tab Bất ổn)
    if not rows or len(rows) < 5:
        try:
            sh_master = gc.open_by_key("14cN4YAf01NqKtREtmhBngRP58OCdgby-PjVoixVlpZo")
            ws_master = sh_master.worksheet("Bất ổn")
            rows = ws_master.get_all_values()
        except Exception as e:
            print(f"⚠️ Error reading Master Bất ổn tab: {e}")
            return {"error": str(e)}

    update_time = "Chưa rõ"
    total_warning = 0
    logic_rules = ""

    if len(rows) > 0 and len(rows[0]) > 1:
        logic_rules = rows[0][1]
    if len(rows) > 1 and len(rows[1]) > 0 and str(rows[1][0]).strip().isdigit():
        total_warning = int(rows[1][0])
    if len(rows) > 3 and len(rows[3]) > 23:
        update_time = rows[3][23]

    data_rows = rows[5:] if len(rows) > 5 else []
    records = []

    for row in data_rows:
        if not row or len(row) < 4 or not str(row[2]).strip():
            continue

        po_id = str(row[2]).strip()
        po_name = str(row[3]).strip()
        province = str(row[1]).strip() if len(row) > 1 else ""
        vung = str(row[0]).strip() if len(row) > 0 else "NTB"

        rec = {
            "vung": vung,
            "province": province,
            "po_id": po_id,
            "name": po_name,
            "gtc_tb_7d": row[4] if len(row) > 4 else "",
            "gtc_best_6m": row[5] if len(row) > 5 else "",
            "pct_gtc_vs_best": row[6] if len(row) > 6 else "",
            "gtc_n1": row[7] if len(row) > 7 else "",
            "vol_gtc_n1": row[8] if len(row) > 8 else "",
            "gtc_n2": row[9] if len(row) > 9 else "",
            "vol_gtc_n2": row[10] if len(row) > 10 else "",
            "gtc_n3": row[11] if len(row) > 11 else "",
            "vol_gtc_n3": row[12] if len(row) > 12 else "",
            "pct_vol_national_n1": row[13] if len(row) > 13 else "",
            "vol_gtc_tb_7d": row[14] if len(row) > 14 else "",
            "vol_can_giao_total": row[15] if len(row) > 15 else "",
            "vol_ton": row[16] if len(row) > 16 else "",
            "vol_moi": row[17] if len(row) > 17 else "",
            "status": row[18] if len(row) > 18 else "",
            "days_unstable": int(row[19]) if len(row) > 19 and str(row[19]).isdigit() else 0,
            "days_est_clear": int(row[20]) if len(row) > 20 and str(row[20]).isdigit() else 0,
            "staff_assigned_n1": int(row[21]) if len(row) > 21 and str(row[21]).isdigit() else 0,
            "dieu_tiet_hang": row[22] if len(row) > 22 else "",
            "backlog_ton_dong": row[23] if len(row) > 23 else "",
            "pct_backlog_assigned": row[24] if len(row) > 24 else "",
            "backlog_over_5d": row[25] if len(row) > 25 else "",
            "pct_fd_n1": row[26] if len(row) > 26 else "",
            "pct_ontime_sla_n1": row[27] if len(row) > 27 else "",
            "pct_opr_n1": row[28] if len(row) > 28 else "",
            "pct_vol_gtc_under_24h": row[29] if len(row) > 29 else "",
            "vol_created_new_n1": row[30] if len(row) > 30 else ""
        }
        records.append(rec)

    return {
        "logic_rules": logic_rules,
        "update_time": update_time,
        "total_warning": total_warning or len(records),
        "records": records
    }

if __name__ == "__main__":
    res = fetch_ntb_unstable_po_data()
    print(f"Fetched {len(res.get('records', []))} records.")
