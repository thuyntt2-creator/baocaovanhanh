# -*- coding: utf-8 -*-
"""
Script: generate_morning_questions.py
Author: Antigravity AI
Description: Analyzes operational performance indicators for N-1 date across multiple worksheets,
             identifies problematic post offices, generates natural Vietnamese business questions 
             for AM explanation, and broadcasts them to a configured GTalk channel.
"""

import os
import sys
import json
import argparse
import unicodedata
import requests
import pandas as pd
from datetime import datetime, timedelta
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

# Fix console output encoding for Vietnamese character support
try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

# ============ CONFIG & CONSTANTS ============
MAIN_SPREADSHEET_ID = "1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ"
ROT_LC_SPREADSHEET_ID = "14r8n9L2cIG1Bmz8kSH79B24QzmnOApZhniGyOU40hr4"

# Default fallback values for GTalk (overridden by env if present)
GTALK_OA_TOKEN = "2067164759710552066:RfgqBJY4QtV18udu1wMEfhmoRNI4hgBv"
GTALK_CHANNEL_ID = "2067283005274091520"  # Fallback to OPR / regional channel ID

# Load environment configuration from standard paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
env_paths = [
    os.path.join(BASE_DIR, ".env"),
    r"c:\Users\lap4all\Desktop\New folder\.env"
]
for p in env_paths:
    if os.path.exists(p):
        load_dotenv(dotenv_path=p, override=True)
        break

# Override constants with environment variables
GTALK_OA_TOKEN = os.environ.get("GTALK_OA_TOKEN") or GTALK_OA_TOKEN
# Use MORNING_QUESTIONS_GTALK_CHANNEL_ID if available, else OPR_GTALK_CHANNEL_ID, else default
GTALK_CHANNEL_ID = (
    os.environ.get("MORNING_QUESTIONS_GTALK_CHANNEL_ID") or 
    os.environ.get("OPR_GTALK_CHANNEL_ID") or 
    os.environ.get("GTALK_CHANNEL_ID") or 
    GTALK_CHANNEL_ID
)

JSON_FILE = os.path.join(BASE_DIR, 'credentials.json')
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# ============ AM PERSONAL DETAILS MAP ============
# Mapping AM names to their pronouns, call names and tag names
AM_DETAILS = {
    "Trần Văn Phước":      {"pronoun": "Anh", "call_name": "Phước", "tag": "@Trần Văn Phước"},
    "Trần Thị Nhung":      {"pronoun": "Chị", "call_name": "Nhung", "tag": "@Trần Thị Nhung"},
    "Trầm Hữu Tiến":       {"pronoun": "Anh", "call_name": "Tiến",  "tag": "@Trầm Hữu Tiến"},
    "Thái Thị Thanh Thư":  {"pronoun": "Chị", "call_name": "Thư",   "tag": "@Thái Thị Thanh Thư"},
    "Phan Đình Duy":       {"pronoun": "Anh", "call_name": "Duy",   "tag": "@Phan Đình Duy"},
    "Phạm Bá Thành Công":  {"pronoun": "Anh", "call_name": "Công",  "tag": "@Phạm Bá Thành Công"},
    "Nguyễn Tiến Lực":     {"pronoun": "Anh", "call_name": "Lực",   "tag": "@Nguyễn Tiến Lực"},
    "Nguyễn Thanh Long":   {"pronoun": "Anh", "call_name": "Long",  "tag": "@Nguyễn Thanh Long"},
    "Nguyễn Ngọc Khánh":   {"pronoun": "Anh", "call_name": "Khánh", "tag": "@Nguyễn Ngọc Khánh"},
    "Nguyễn Minh Hoàng":   {"pronoun": "Anh", "call_name": "Hoàng", "tag": "@Nguyễn Minh Hoàng"},
    "Nguyễn Lê Nguyên Vũ": {"pronoun": "Anh", "call_name": "Vũ",    "tag": "@Nguyễn Lê Nguyên Vũ"},
    "Nguyễn Hoàng Phi":    {"pronoun": "Anh", "call_name": "Phi",   "tag": "@Nguyễn Hoàng Phi"},
    "Nguyễn Duy Long":     {"pronoun": "Anh", "call_name": "Long",  "tag": "@Nguyễn Duy Long"},
    "Lê Văn Trường":       {"pronoun": "Anh", "call_name": "Trường","tag": "@Lê Văn Trường"},
    "Lê Thanh Nhựt":       {"pronoun": "Anh", "call_name": "Nhựt",  "tag": "@Lê Thanh Nhựt"},
    "Lê Minh Đại":         {"pronoun": "Anh", "call_name": "Đại",   "tag": "@Lê Minh Đại"},
    "Huỳnh Thị Kim Chi":   {"pronoun": "Chị", "call_name": "Chi",   "tag": "@Huỳnh Thị Kim Chi"},
    "Hồng Bích Nga":       {"pronoun": "Chị", "call_name": "Nga",   "tag": "@Hồng Bích Nga"}
}

def get_am_meta(am_name):
    """Fetch pronoun, call_name, and tag for AM, with defaults if not found"""
    if not am_name:
        return {"pronoun": "Anh/Chị", "call_name": "AM", "tag": "@AM"}
    
    # Try direct mapping
    norm_am = unicodedata.normalize('NFC', str(am_name).strip())
    for name, meta in AM_DETAILS.items():
        if unicodedata.normalize('NFC', name) == norm_am:
            return meta
            
    # Fallback heuristic
    pronoun = "Chị" if any(w in norm_am.lower() for w in ["thị", "bích", "nga", "chi", "thư", "nhung", "vy"]) else "Anh"
    parts = norm_am.split()
    call_name = parts[-1] if parts else "AM"
    return {"pronoun": pronoun, "call_name": call_name, "tag": f"@{norm_am}"}


# ============ STRING HELPERS ============
def normalize_str(s):
    if not s:
        return ""
    return unicodedata.normalize('NFC', str(s).strip().lower())

def clean_bc_name(name):
    name = normalize_str(name)
    # Remove region tags
    for tag in ['(dno)', '(ldo)', '(kho)', '(bth)', '(nth)']:
        name = name.replace(tag, "")
    # Remove post office prefixes
    for prefix in ['kho chuyển tiếp', 'kho trung chuyển', 'điểm xử lý hàng', 'điểm lấy hàng', 'bưu cục', 'bc', 'đl']:
        name = name.replace(prefix, "")
    # Standardize whitespace and remove symbols
    name = name.replace("-", " ").replace("_", " ")
    return " ".join(name.split())

def match_po_name(raw_name, standard_list):
    """
    Finds the best standard PO name matching the raw name.
    1. Exact Match (normalized)
    2. Clean Exact Match
    3. Clean Substring Match
    """
    raw_norm = normalize_str(raw_name)
    # Direct match
    for std in standard_list:
        if normalize_str(std) == raw_norm:
            return std
            
    # Clean matches
    raw_clean = clean_bc_name(raw_name)
    if not raw_clean:
        return None
        
    cleaned_std_list = [(std, clean_bc_name(std)) for std in standard_list]
    
    # Clean exact match
    for std, std_clean in cleaned_std_list:
        if std_clean == raw_clean:
            return std
            
    # Substring match
    matches = []
    for std, std_clean in cleaned_std_list:
        if std_clean and (raw_clean in std_clean or std_clean in raw_clean):
            matches.append(std)
            
    if matches:
        # Return the one with closest length
        matches.sort(key=lambda x: abs(len(clean_bc_name(x)) - len(raw_clean)))
        return matches[0]
        
    return None

def parse_percent_to_float(val):
    """Converts GSheet percentage string (e.g. '17.2%' or '10.20') to float (17.2 or 10.2)"""
    if val is None or val == "":
        return 0.0
    try:
        val_str = str(val).replace("%", "").replace(",", ".").strip()
        num = float(val_str)
        # GSheet raw value might be stored as e.g. 0.172 for 17.2%
        if num < 1.0 and num > 0:
            return num * 100.0
        return num
    except ValueError:
        return 0.0

def parse_int(val):
    if val is None or val == "":
        return 0
    try:
        return int(str(val).replace(",", "").replace(".", "").strip())
    except ValueError:
        return 0


# ============ GTALK MESSAGE SENDER ============
def send_gtalk_message(text):
    if not GTALK_OA_TOKEN or not GTALK_CHANNEL_ID:
        print("⚠️ Warning: Missing GTALK_OA_TOKEN or GTALK_CHANNEL_ID. Cannot send GTalk message.")
        return False
        
    print(f"📡 Broadcasting morning questions to GTalk (Channel: {GTALK_CHANNEL_ID})...")
    url = "https://mbff.ghn.vn/api/gtalk/send-message"
    client_msg_id = str(int(datetime.now().timestamp() * 1000))
    payload = {
        "channelId": GTALK_CHANNEL_ID,
        "clientMsgId": client_msg_id,
        "content": {
            "parseMode": "HTML",
            "text": text
        },
        "oaToken": GTALK_OA_TOKEN
    }
    headers = {"Content-Type": "application/json"}
    
    try:
        # Ignore SSL verification errors just like other scripts
        res = requests.post(url, json=payload, headers=headers, timeout=20, verify=False)
        if res.status_code == 200:
            res_data = res.json()
            if res_data.get("errorCode") == "success":
                print("✅ Successfully sent message to GTalk!")
                return True
            else:
                print(f"❌ Failed to send GTalk message API error: {res_data.get('error')}")
        else:
            print(f"❌ Failed to send GTalk message HTTP error {res.status_code}: {res.text}")
    except Exception as e:
         print(f"❌ Connection error sending to GTalk: {e}")
    return False


def rewrite_questions_gemini(questions_list, api_key):
    if not api_key:
        print("⚠️ GEMINI_API_KEY is not set. Skipping AI rewriting.")
        return None
        
    print(f"🤖 Sending {len(questions_list)} questions to Gemini for natural rewriting...")
    
    # Large chunk size to process all questions in 1-2 requests to save quota
    chunk_size = 50
    rewritten_results = []
    
    for i in range(0, len(questions_list), chunk_size):
        chunk = questions_list[i : i + chunk_size]
        print(f"  • Processing chunk {i//chunk_size + 1} ({len(chunk)} questions)...")
        
        # Try gemini-3.5-flash first, fallback to gemini-2.5-flash
        models_to_try = ["gemini-3.5-flash", "gemini-2.5-flash"]
        success = False
        
        for model in models_to_try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
            headers = {"Content-Type": "application/json"}
            
            prompt = f"""
Bạn là một quản lý vận hành giàu kinh nghiệm. Hãy viết lại (rewrite) các câu hỏi giải trình vận hành gửi các AM (Area Manager) dưới đây để chúng tự nhiên, đa dạng cấu trúc ngữ pháp và tránh trùng lặp máy móc.

Yêu cầu cụ thể:
1. Giữ nguyên 100% các từ khóa tên bưu cục (ví dụ: bưu cục (DNO) Quảng Tín, (LDO) Di Linh), các thẻ tag tên AM (ví dụ: @Trần Văn Phước, @Thái Thị Thanh Thư) và các số liệu thống kê trong câu.
2. Thay đổi cách đặt câu hỏi, sử dụng nhiều mẫu câu khác nhau (hỏi thăm sự cố, đề xuất rà soát, yêu cầu giải pháp, nhắc nhở deadline...).
3. Văn phong chuyên nghiệp, thân thiện nhưng quyết liệt, đúng chất chat công việc (Zalo/GTalk).
4. Không thêm lời chào chung chung ngoài nội dung câu hỏi.

Đầu vào là một chuỗi JSON như sau:
{json.dumps(chunk, ensure_ascii=False)}

Hãy trả về kết quả định dạng JSON duy nhất khớp với cấu trúc trên:
[
  {{
    "id": 1,
    "text": "Câu hỏi đã được viết lại tự nhiên và sinh động hơn"
  }}
]
            """
            
            payload = {
                "contents": [{
                    "parts": [{
                        "text": prompt
                    }]
                }],
                "generationConfig": {
                    "responseMimeType": "application/json"
                }
            }
            
            try:
                # We set timeout=60 for each chunk to allow processing larger lists
                response = requests.post(url, headers=headers, json=payload, timeout=60, verify=False)
                if response.status_code == 200:
                    res_data = response.json()
                    text_out = res_data['candidates'][0]['content']['parts'][0]['text']
                    chunk_rewritten = json.loads(text_out)
                    if isinstance(chunk_rewritten, list):
                        rewritten_results.extend(chunk_rewritten)
                        success = True
                        break
                    else:
                        print(f"⚠️ Chunk response was not a list: {text_out}")
                else:
                    print(f"⚠️ Gemini API Error on model {model}: {response.status_code} - {response.text}")
            except Exception as e:
                print(f"⚠️ Exception during Gemini API call for model {model}: {e}")
                
        if not success:
            print("⚠️ All models failed for this chunk.")
            return None
            
    return rewritten_results


# ============ MAIN PIPELINE ============
def main():
    parser = argparse.ArgumentParser(description="Generate morning explanation questions for Area Managers (AM).")
    parser.add_argument("--date", type=str, help="Target date in YYYY-MM-DD format. Defaults to yesterday.")
    parser.add_argument("--send", action="store_true", help="Send questions to GTalk channel.")
    parser.add_argument("--sheet", action="store_true", help="Generate highlights dashboard and write to Google Sheet tab.")
    args = parser.parse_args()

    if args.sheet:
        print("📊 Running Highlights Dashboard sheet writer...")
        try:
            import subprocess
            cmd = [sys.executable, os.path.join(BASE_DIR, "generate_highlights_dashboard.py")]
            if args.date:
                cmd.extend(["--date", args.date])
            subprocess.run(cmd, check=True)
            print("✅ Highlights Dashboard Sheet Writer completed successfully!")
        except Exception as e:
            print(f"⚠️ Error running highlights dashboard: {e}")

    # Determine target date
    if args.date:
        try:
            target_date = datetime.strptime(args.date.strip(), "%Y-%m-%d")
        except ValueError:
            print("❌ Invalid date format. Use YYYY-MM-DD.")
            sys.exit(1)
    else:
        # Default to N-1 (yesterday)
        target_date = datetime.now() - timedelta(days=1)
        
    date_str_iso = target_date.strftime("%Y-%m-%d")
    
    # Vietnamese weekdays mapping for 'Data' and 'TTS' sheets format
    weekday_map = {
        0: "Thứ 2", 1: "Thứ 3", 2: "Thứ 4", 3: "Thứ 5", 4: "Thứ 6", 5: "Thứ 7", 6: "Chủ Nhật"
    }
    date_str_sheet = f"{date_str_iso} - {weekday_map[target_date.weekday()]}"
    date_str_dmy = target_date.strftime("%d/%m/%Y")
    
    print(f"📅 Running analysis for date N-1: {date_str_iso} ({weekday_map[target_date.weekday()]})")
    print(f"  • Data/TTS Sheet date query string: '{date_str_sheet}'")
    print(f"  • OPR/Rot LC Sheet date query string: '{date_str_iso}'")
    
    # Authenticate with Google sheets
    if not os.path.exists(JSON_FILE):
        print(f"❌ Credentials file not found at: {JSON_FILE}")
        sys.exit(1)
        
    credentials = Credentials.from_service_account_file(JSON_FILE, scopes=SCOPES)
    gc = gspread.authorize(credentials)
    
    # ============ 1. LOAD MASTER METADATA (Cơ cấu / CoCauVung) ============
    print("📖 Loading Master CoCauVung sheet...")
    sh_main = gc.open_by_key(MAIN_SPREADSHEET_ID)
    ws_cc = None
    for sname in ["CoCauVung", "Cơ cấu", "cơ cấu"]:
        try:
            ws_cc = sh_main.worksheet(sname)
            break
        except Exception:
            pass
    if not ws_cc:
        print("❌ Không tìm thấy tab CoCauVung hoặc Cơ cấu trong spreadsheet.")
        sys.exit(1)

    cc_rows = ws_cc.get_all_values()
    df_cc = pd.DataFrame(cc_rows[1:], columns=cc_rows[0])
    
    # Remove empty lines and transit hubs (Kho trung chuyển / Kho chuyển tiếp) which don't have last-mile post offices
    df_cc = df_cc[df_cc['Bưu cục'].str.strip() != '']
    df_cc = df_cc[~df_cc['Bưu cục'].str.contains('Kho Trung Chuyển|Kho Chuyển Tiếp', case=False, na=False)]
    master_pos = df_cc['Bưu cục'].str.strip().tolist()
    po_am_map = dict(zip(df_cc['Bưu cục'].str.strip(), df_cc['AM'].str.strip()))
    
    print(f"  • Found {len(master_pos)} standard post offices in Master CoCauVung.")
    
    # Structure to hold metrics for each Post Office
    po_metrics = {po: {
        'po_name': po,
        'am_name': po_am_map[po],
        'issues': [],
        # Data (All) metrics
        'all_vol': 0, 'all_gan_vol': 0, 'all_gan_rate': 0.0,
        'all_gan_ca1': 0.0, 'all_gan_ca2': 0.0,
        'all_gtc_ca1': 0.0, 'all_gtc_ca2': 0.0,
        'all_gtc_vol': 0, 'all_gtc_rate': 0.0,
        'all_chua_gan': 0, 'all_ton': 0,
        # TTS metrics
        'tts_vol': 0, 'tts_gan_vol': 0, 'tts_gan_rate': 0.0,
        'tts_gan_ca1': 0.0, 'tts_gan_ca2': 0.0,
        'tts_gtc_ca1': 0.0, 'tts_gtc_ca2': 0.0,
        'tts_gtc_vol': 0, 'tts_gtc_rate': 0.0,
        'tts_chua_gan': 0, 'tts_ton': 0,
        # OPR TTS
        'opr_vol': 0, 'opr_ontime': 0, 'opr_rate': 0.0, 'opr_late': 0,
        # Aging > 5 days
        'aging_total': 0, 'aging_5_8d': 0, 'aging_8_15d': 0, 'aging_15d_plus': 0,
        # Hanging LC
        'treo_lc': 0,
        # FD
        'fd_rate': 0.0, 'fd_rate_n1': 0.0, 'fd_vs_n1': 0.0, 'fd_vol_giao': 0, 'fd_vol_tra': 0.0, 'fd_ty_trong_tra': 0.0,
        # Rot LC
        'rot_lc_tts': 0.0, 'rot_lc_shopee': 0.0, 'rot_lc_khac': 0.0
    } for po in master_pos}

    # ============ 2. PROCESS DATA SHEET (ALL CLIENTS) ============
    print("📖 Loading and processing 'Data' worksheet...")
    ws_data = sh_main.worksheet("Data")
    data_rows = ws_data.get_all_values()
    df_data = pd.DataFrame(data_rows[1:], columns=data_rows[0])
    
    # Process target date (N-1)
    df_data_day = df_data[df_data['Time'] == date_str_sheet]
    print(f"  • Loaded {len(df_data_day)} rows for 'Data' sheet on target date.")
    for idx, row in df_data_day.iterrows():
        raw_po = row['Chi tiết']
        std_po = match_po_name(raw_po, master_pos)
        if not std_po:
            continue
            
        m = po_metrics[std_po]
        loai_hang = row['Loại Hàng'].strip()
        vol = parse_int(row['Volume'])
        gan = parse_int(row['Sản Lượng Gán'])
        gtc = parse_int(row['Sản Lượng Giao Thành Công'])
        chua_gan = parse_int(row['Sản Lượng Chưa Gán'])
        ton = parse_int(row['Sản Lượng Tồn'])
        
        # Accumulate totals
        m['all_vol'] += vol
        m['all_gan_vol'] += gan
        m['all_gtc_vol'] += gtc
        m['all_chua_gan'] += chua_gan
        m['all_ton'] += ton
        
        # Capture shifts
        if loai_hang == "Hàng Mới Ca 1":
            m['all_gan_ca1'] = parse_percent_to_float(row['% Gán'])
            m['all_gtc_ca1'] = parse_percent_to_float(row['% GTC'])
        elif loai_hang == "Hàng Mới Ca 2":
            m['all_gan_ca2'] = parse_percent_to_float(row['% Gán'])
            m['all_gtc_ca2'] = parse_percent_to_float(row['% GTC'])

    # Recalculate GTC & Gán rates
    for po, m in po_metrics.items():
        if m['all_vol'] > 0:
            m['all_gan_rate'] = (m['all_gan_vol'] / m['all_vol']) * 100.0
            m['all_gtc_rate'] = (m['all_gtc_vol'] / m['all_vol']) * 100.0

    # Process prev date (N-2)
    prev_target_date = target_date - timedelta(days=1)
    prev_date_str_iso = prev_target_date.strftime("%Y-%m-%d")
    prev_date_str_sheet = f"{prev_date_str_iso} - {weekday_map[prev_target_date.weekday()]}"
    
    prev_po_metrics = {po: {
        'all_vol': 0, 'all_gan_vol': 0, 'all_gan_rate': 0.0,
        'all_gtc_vol': 0, 'all_gtc_rate': 0.0,
        'tts_vol': 0, 'tts_gan_vol': 0, 'tts_gan_rate': 0.0,
        'tts_gtc_vol': 0, 'tts_gtc_rate': 0.0,
        'opr_vol': 0, 'opr_ontime': 0, 'opr_rate': 0.0,
        'rot_lc_tts': 0.0, 'rot_lc_shopee': 0.0, 'rot_lc_khac': 0.0
    } for po in master_pos}

    df_data_prev = df_data[df_data['Time'] == prev_date_str_sheet]
    print(f"  • Loaded {len(df_data_prev)} rows for 'Data' sheet on yesterday.")
    for idx, row in df_data_prev.iterrows():
        raw_po = row['Chi tiết']
        std_po = match_po_name(raw_po, master_pos)
        if not std_po:
            continue
        pm = prev_po_metrics[std_po]
        vol = parse_int(row['Volume'])
        gan = parse_int(row['Sản Lượng Gán'])
        gtc = parse_int(row['Sản Lượng Giao Thành Công'])
        pm['all_vol'] += vol
        pm['all_gan_vol'] += gan
        pm['all_gtc_vol'] += gtc

    for po, pm in prev_po_metrics.items():
        if pm['all_vol'] > 0:
            pm['all_gan_rate'] = (pm['all_gan_vol'] / pm['all_vol']) * 100.0
            pm['all_gtc_rate'] = (pm['all_gtc_vol'] / pm['all_vol']) * 100.0

    # ============ 3. PROCESS TTS SHEET (TIKTOK SHOP CLIENT) ============
    print("📖 Loading and processing 'TTS' worksheet...")
    ws_tts = sh_main.worksheet("TTS")
    tts_rows = ws_tts.get_all_values()
    df_tts = pd.DataFrame(tts_rows[1:], columns=tts_rows[0])
    
    # Process target date (N-1)
    df_tts_day = df_tts[df_tts['Time'] == date_str_sheet]
    print(f"  • Loaded {len(df_tts_day)} rows for 'TTS' sheet on target date.")
    for idx, row in df_tts_day.iterrows():
        raw_po = row['Chi tiết']
        std_po = match_po_name(raw_po, master_pos)
        if not std_po:
            continue
            
        m = po_metrics[std_po]
        loai_hang = row['Loại Hàng'].strip()
        vol = parse_int(row['Volume'])
        gan = parse_int(row['Sản Lượng Gán'])
        gtc = parse_int(row['Sản Lượng Giao Thành Công'])
        chua_gan = parse_int(row['Sản Lượng Chưa Gán'])
        ton = parse_int(row['Sản Lượng Tồn'])
        
        # Accumulate totals
        m['tts_vol'] += vol
        m['tts_gan_vol'] += gan
        m['tts_gtc_vol'] += gtc
        m['tts_chua_gan'] += chua_gan
        m['tts_ton'] += ton
        
        # Capture shifts
        if loai_hang == "Hàng Mới Ca 1":
            m['tts_gan_ca1'] = parse_percent_to_float(row['% Gán'])
            m['tts_gtc_ca1'] = parse_percent_to_float(row['% GTC'])
        elif loai_hang == "Hàng Mới Ca 2":
            m['tts_gan_ca2'] = parse_percent_to_float(row['% Gán'])
            m['tts_gtc_ca2'] = parse_percent_to_float(row['% GTC'])

    # Recalculate OPR/Gán rates
    for po, m in po_metrics.items():
        if m['tts_vol'] > 0:
            m['tts_gan_rate'] = (m['tts_gan_vol'] / m['tts_vol']) * 100.0
            m['tts_gtc_rate'] = (m['tts_gtc_vol'] / m['tts_vol']) * 100.0

    # Process prev date (N-2)
    df_tts_prev = df_tts[df_tts['Time'] == prev_date_str_sheet]
    print(f"  • Loaded {len(df_tts_prev)} rows for 'TTS' sheet on yesterday.")
    for idx, row in df_tts_prev.iterrows():
        raw_po = row['Chi tiết']
        std_po = match_po_name(raw_po, master_pos)
        if not std_po:
            continue
        pm = prev_po_metrics[std_po]
        vol = parse_int(row['Volume'])
        gan = parse_int(row['Sản Lượng Gán'])
        gtc = parse_int(row['Sản Lượng Giao Thành Công'])
        pm['tts_vol'] += vol
        pm['tts_gan_vol'] += gan
        pm['tts_gtc_vol'] += gtc

    for po, pm in prev_po_metrics.items():
        if pm['tts_vol'] > 0:
            pm['tts_gan_rate'] = (pm['tts_gan_vol'] / pm['tts_vol']) * 100.0
            pm['tts_gtc_rate'] = (pm['tts_gtc_vol'] / pm['tts_vol']) * 100.0

    # ============ 4. PROCESS OPR SHEET (TTS PICKUP PERFORMANCE) ============
    print("📖 Loading and processing 'OPR' worksheet...")
    ws_opr = sh_main.worksheet("OPR")
    opr_rows = ws_opr.get_all_values()
    df_opr = pd.DataFrame(opr_rows[1:], columns=opr_rows[0])
    
    # Process target date (N-1)
    df_opr_day = df_opr[df_opr['NgayLTC'] == date_str_iso]
    print(f"  • Loaded {len(df_opr_day)} rows for 'OPR' sheet on target date.")
    for idx, row in df_opr_day.iterrows():
        raw_po = row['KhoLay']
        std_po = match_po_name(raw_po, master_pos)
        if not std_po:
            continue
            
        m = po_metrics[std_po]
        vol = parse_int(row['Don_ltc'])
        ot = parse_int(row['Don_ontime'])
        late = max(0, vol - ot)
        
        m['opr_vol'] += vol
        m['opr_ontime'] += ot
        m['opr_late'] += late
        
    for po, m in po_metrics.items():
        if m['opr_vol'] > 0:
            m['opr_rate'] = (m['opr_ontime'] / m['opr_vol']) * 100.0

    # Process prev date (N-2)
    df_opr_prev = df_opr[df_opr['NgayLTC'] == prev_date_str_iso]
    print(f"  • Loaded {len(df_opr_prev)} rows for 'OPR' sheet on yesterday.")
    for idx, row in df_opr_prev.iterrows():
        raw_po = row['KhoLay']
        std_po = match_po_name(raw_po, master_pos)
        if not std_po:
            continue
        pm = prev_po_metrics[std_po]
        vol = parse_int(row['Don_ltc'])
        ot = parse_int(row['Don_ontime'])
        pm['opr_vol'] += vol
        pm['opr_ontime'] += ot

    for po, pm in prev_po_metrics.items():
        if pm['opr_vol'] > 0:
            pm['opr_rate'] = (pm['opr_ontime'] / pm['opr_vol']) * 100.0

    # ============ 5. PROCESS AGING BACKLOG SHEET (SKIPPED) ============
    print("📖 Skipping 'Aging trên 5 ngày' worksheet (disabled as per request).")

    # ============ 6. PROCESS HANGING TRANSIT SHEET (SKIPPED) ============
    print("📖 Skipping 'Treo LC' worksheet (disabled as per request).")

    # ============ 7. PROCESS FD SHEET (PRE-CALCULATED) ============
    print("📖 Loading and processing 'FD ' worksheet...")
    ws_fd = sh_main.worksheet("FD ")
    fd_rows = ws_fd.get_all_values()
    
    # Parse date from row 1: e.g. "N = 23/06/2026"
    fd_date_label = fd_rows[0][1].replace("N =", "").strip()
    print(f"  • FD sheet date header reports: N = {fd_date_label}")
    
    # Form headers
    fd_headers = ['Bưu Cục', 'AM', '%FD (N)', '%FD (N-1)', 'vs N-1', '%FD (N-7)', 'vs N-7', 'Vol giao', 'Vol trả', 'Tỷ trọng trả']
    df_fd = pd.DataFrame(fd_rows[3:], columns=fd_headers + [f'Col_{i}' for i in range(10, len(fd_rows[3]))])
    
    for idx, row in df_fd.iterrows():
        raw_po = row['Bưu Cục']
        if not raw_po or raw_po.strip() == "":
            continue
        # Stop processing if we reach AM or Province summary sections
        if "👤" in raw_po or "🗺️" in raw_po or "THEO AM" in raw_po or "THEO TỈNH" in raw_po:
            break
            
        std_po = match_po_name(raw_po, master_pos)
        if not std_po:
            continue
            
        m = po_metrics[std_po]
        m['fd_rate'] = parse_percent_to_float(row['%FD (N)'])
        m['fd_rate_n1'] = parse_percent_to_float(row['%FD (N-1)'])
        
        # Parse vs N-1 delta string (e.g. "▲4.0%" or "▼7.3%" or "—")
        vs_str = str(row['vs N-1']).replace("▲", "+").replace("▼", "-").replace("%", "").strip()
        try:
            m['fd_vs_n1'] = float(vs_str) if vs_str not in ["", "—", "-", "N/A"] else 0.0
        except ValueError:
            m['fd_vs_n1'] = 0.0
            
        m['fd_vol_giao'] = parse_int(row['Vol giao'])
        m['fd_vol_tra'] = parse_percent_to_float(row['Vol trả'])  # It might be written as decimal or float
        m['fd_ty_trong_tra'] = parse_percent_to_float(row['Tỷ trọng trả'])

    # ============ 8. PROCESS LOST TRANSIT SPREADSHEET (RỚT LC) ============
    print("📖 Loading and processing Rot LC Spreadsheet...")
    try:
        sh_rot = gc.open_by_key(ROT_LC_SPREADSHEET_ID)
        
        # 8a. Data TTS
        ws_rot_tts = sh_rot.worksheet("Data TTS")
        rot_tts_rows = ws_rot_tts.get_all_values()
        df_rot_tts = pd.DataFrame(rot_tts_rows[1:], columns=rot_tts_rows[0])
        
        # N-1
        df_rot_tts_day = df_rot_tts[df_rot_tts['Loại ngày'] == date_str_iso]
        for idx, row in df_rot_tts_day.iterrows():
            raw_po = row['Chi tiết']
            std_po = match_po_name(raw_po, master_pos)
            if std_po:
                po_metrics[std_po]['rot_lc_tts'] = parse_percent_to_float(row['%_rot_lc'])
                
        # N-2
        df_rot_tts_prev = df_rot_tts[df_rot_tts['Loại ngày'] == prev_date_str_iso]
        for idx, row in df_rot_tts_prev.iterrows():
            raw_po = row['Chi tiết']
            std_po = match_po_name(raw_po, master_pos)
            if std_po:
                prev_po_metrics[std_po]['rot_lc_tts'] = parse_percent_to_float(row['%_rot_lc'])
                
        # 8b. Data Shopee
        ws_rot_shp = sh_rot.worksheet("Data Shopee")
        rot_shp_rows = ws_rot_shp.get_all_values()
        df_rot_shp = pd.DataFrame(rot_shp_rows[1:], columns=rot_shp_rows[0])
        
        # N-1
        df_rot_shp_day = df_rot_shp[df_rot_shp['Loại ngày'] == date_str_iso]
        for idx, row in df_rot_shp_day.iterrows():
            raw_po = row['Chi tiết']
            std_po = match_po_name(raw_po, master_pos)
            if std_po:
                po_metrics[std_po]['rot_lc_shopee'] = parse_percent_to_float(row['%_rot_lc'])
                
        # N-2
        df_rot_shp_prev = df_rot_shp[df_rot_shp['Loại ngày'] == prev_date_str_iso]
        for idx, row in df_rot_shp_prev.iterrows():
            raw_po = row['Chi tiết']
            std_po = match_po_name(raw_po, master_pos)
            if std_po:
                prev_po_metrics[std_po]['rot_lc_shopee'] = parse_percent_to_float(row['%_rot_lc'])
                
        # 8c. Data Khác
        ws_rot_khc = sh_rot.worksheet("Data Khác")
        rot_khc_rows = ws_rot_khc.get_all_values()
        df_rot_khc = pd.DataFrame(rot_khc_rows[1:], columns=rot_khc_rows[0])
        
        # N-1
        df_rot_khc_day = df_rot_khc[df_rot_khc['Loại ngày'] == date_str_iso]
        for idx, row in df_rot_khc_day.iterrows():
            raw_po = row['Chi tiết']
            std_po = match_po_name(raw_po, master_pos)
            if std_po:
                po_metrics[std_po]['rot_lc_khac'] = parse_percent_to_float(row['%_rot_lc'])
                
        # N-2
        df_rot_khc_prev = df_rot_khc[df_rot_khc['Loại ngày'] == prev_date_str_iso]
        for idx, row in df_rot_khc_prev.iterrows():
            raw_po = row['Chi tiết']
            std_po = match_po_name(raw_po, master_pos)
            if std_po:
                prev_po_metrics[std_po]['rot_lc_khac'] = parse_percent_to_float(row['%_rot_lc'])
                
        print(f"  • Successfully extracted rớt LC rates for target date and yesterday.")
    except Exception as e:
        print(f"  ⚠️ Warning processing Rot LC Spreadsheet: {e}")

    # ============ 9. ISSUE DETECTION ENGINE & QUESTION GENERATION ============
    print("\n🔍 Analyzing indicators and matching alert thresholds...")
    
    for po, m in po_metrics.items():
        meta = get_am_meta(m['am_name'])
        pronoun = meta['pronoun']
        pronoun_lower = pronoun.lower()
        tag = meta['tag']
        
        # Get N-2 metrics for comparison
        pm = prev_po_metrics.get(po)
        
        # 1. FD ALERT (🔴 = FD>4.5% & Tỷ trọng>3% or extreme FD > 10% or Day-over-day delta >= 2.0%)
        is_priority_fd = (m['fd_rate'] > 4.5 and m['fd_ty_trong_tra'] > 3.0)
        is_extreme_fd = (m['fd_rate'] > 10.0)
        is_spiked_fd = (m['fd_vs_n1'] >= 2.0)
        
        if is_priority_fd or is_extreme_fd or is_spiked_fd:
            is_drop = is_spiked_fd
            if is_extreme_fd:
                status_desc = "ở mức rất cao"
            else:
                status_desc = "ở mức cao"
                
            delta_desc = ""
            if m['fd_vs_n1'] > 0:
                delta_desc = f" (tăng mạnh {m['fd_vs_n1']:.1f}% vs hôm qua)"
            elif m['fd_vs_n1'] < 0:
                delta_desc = f" (giảm {abs(m['fd_vs_n1']):.1f}% vs hôm qua)"
                
            desc = f"tỷ lệ %FD {status_desc} đạt {m['fd_rate']:.1f}%{delta_desc}"
            m['issues'].append({
                'type': 'FD',
                'desc': desc,
                'severity_val': m['fd_rate'],
                'is_drop': is_drop,
                'base_score': 400
            })

        # 2. OPERATIONAL ALERT (%GTC / Unassigned / Backlog issues / Drops)
        # Conditions: %GTC tổng < 45%, or %GTC ca 2 < 40%, or unassigned > 100, or backlogs > 150, or Day-over-day drop >= 5.0%
        is_bad_gtc = (m['all_vol'] > 0 and m['all_gtc_rate'] < 45.0)
        is_bad_ca2_gtc = (m['all_gtc_ca2'] > 0 and m['all_gtc_ca2'] < 40.0)
        is_high_unassigned = (m['all_chua_gan'] > 100)
        is_high_backlogs = (m['all_ton'] > 150)
        
        has_gtc_drop = False
        gtc_drop_val = 0.0
        if pm and pm['all_vol'] > 0 and m['all_vol'] > 0:
            gtc_drop_val = pm['all_gtc_rate'] - m['all_gtc_rate']
            if gtc_drop_val >= 5.0:
                has_gtc_drop = True
                
        if is_bad_gtc or is_bad_ca2_gtc or is_high_unassigned or is_high_backlogs or has_gtc_drop:
            details_list = []
            is_drop = has_gtc_drop
            if has_gtc_drop:
                details_list.append(f"tỷ lệ %GTC tổng sụt giảm mạnh {gtc_drop_val:.1f}% vs hôm qua (từ {pm['all_gtc_rate']:.1f}% còn {m['all_gtc_rate']:.1f}%)")
            elif is_bad_gtc:
                details_list.append(f"tỷ lệ %GTC tổng thấp chỉ đạt {m['all_gtc_rate']:.2f}%")
            elif is_bad_ca2_gtc:
                details_list.append(f"tỷ lệ %GTC ca 2 thấp chỉ đạt {m['all_gtc_ca2']:.2f}%")
                
            if m['all_gan_ca2'] > 0 and m['all_gan_ca2'] < 50.0:
                 details_list.append(f"tỷ lệ %Gán ca 2 thấp chỉ đạt {m['all_gan_ca2']:.2f}%")
                 
            if is_high_unassigned:
                details_list.append(f"lượng hàng chưa gán vọt lên đến {m['all_chua_gan']} đơn")
            if is_high_backlogs:
                details_list.append(f"tồn kho tích lũy đến {m['all_ton']} đơn")
                
            desc = "vận hành yếu: " + ", ".join(details_list)
            m['issues'].append({
                'type': 'OPERATIONAL',
                'desc': desc,
                'severity_val': float(m['all_vol']),
                'is_drop': is_drop,
                'base_score': 100
            })

        # 2b. TTS CUSTOMER SPECIAL OPERATIONAL ALERT
        # Condition: TTS GTC < 45%, or TTS unassigned > 40, or TTS backlogs > 60, or Day-over-day drop >= 5.0%
        is_bad_tts_gtc = (m['tts_vol'] > 0 and m['tts_gtc_rate'] < 45.0)
        is_high_tts_unassigned = (m['tts_chua_gan'] > 40)
        is_high_tts_backlogs = (m['tts_ton'] > 60)
        
        has_tts_gtc_drop = False
        tts_gtc_drop_val = 0.0
        if pm and pm['tts_vol'] > 0 and m['tts_vol'] > 0:
            tts_gtc_drop_val = pm['tts_gtc_rate'] - m['tts_gtc_rate']
            if tts_gtc_drop_val >= 5.0:
                has_tts_gtc_drop = True
                
        if is_bad_tts_gtc or is_high_tts_unassigned or is_high_tts_backlogs or has_tts_gtc_drop:
            details_list = []
            is_drop = has_tts_gtc_drop
            if has_tts_gtc_drop:
                details_list.append(f"tỷ lệ %GTC TTS sụt giảm mạnh {tts_gtc_drop_val:.1f}% vs hôm qua (từ {pm['tts_gtc_rate']:.1f}% còn {m['tts_gtc_rate']:.1f}%)")
            elif is_bad_tts_gtc:
                details_list.append(f"tỷ lệ %GTC TTS thấp chỉ đạt {m['tts_gtc_rate']:.2f}%")
            if is_high_tts_unassigned:
                details_list.append(f"lượng hàng TTS chưa gán vọt lên {m['tts_chua_gan']} đơn")
            if is_high_tts_backlogs:
                details_list.append(f"tồn kho hàng TTS tích lũy đến {m['tts_ton']} đơn")
                
            desc = "vận hành TTS yếu: " + ", ".join(details_list)
            m['issues'].append({
                'type': 'TTS_OPERATIONAL',
                'desc': desc,
                'severity_val': float(m['tts_vol']),
                'is_drop': is_drop,
                'base_score': 200
            })

        # 3. OPR TTS ALERT (%OPR TTS < 80% and delayed > 5, or Day-over-day drop >= 10.0%)
        is_bad_opr = (m['opr_vol'] > 0 and m['opr_rate'] < 80.0 and m['opr_late'] > 5)
        has_opr_drop = False
        opr_drop_val = 0.0
        if pm and pm['opr_vol'] > 0 and m['opr_vol'] > 0:
            opr_drop_val = pm['opr_rate'] - m['opr_rate']
            if opr_drop_val >= 10.0:
                has_opr_drop = True
                
        if is_bad_opr or has_opr_drop:
            is_drop = has_opr_drop
            if has_opr_drop:
                desc = f"hiệu suất lấy hàng OPR TTS sụt giảm mạnh {opr_drop_val:.1f}% vs hôm qua (từ {pm['opr_rate']:.1f}% còn {m['opr_rate']:.1f}%, trễ {m['opr_late']} đơn)"
            else:
                desc = f"chỉ số lấy hàng OPR TTS không đạt KPI, chỉ đạt {m['opr_rate']:.1f}% (dưới mục tiêu 80%), để trễ {m['opr_late']} đơn lấy"
            
            m['issues'].append({
                'type': 'OPR_TTS',
                'desc': desc,
                'severity_val': float(m['opr_late']),
                'is_drop': is_drop,
                'base_score': 650
            })

        # 4 & 5. AGING & TREO LC Alerts (Skipped)

        # 6. LOST TRANSIT ALERT (RỚT LC)
        # Conditions: %_rot_lc TTS > 5.0% or Shopee > 5.0% or Day-over-day increase >= 5.0%
        is_bad_rot_tts = (m['rot_lc_tts'] > 5.0)
        is_bad_rot_shp = (m['rot_lc_shopee'] > 5.0)
        
        has_rot_tts_increase = False
        rot_tts_inc_val = 0.0
        if pm:
            rot_tts_inc_val = m['rot_lc_tts'] - pm['rot_lc_tts']
            if rot_tts_inc_val >= 5.0:
                has_rot_tts_increase = True
                
        has_rot_shp_increase = False
        rot_shp_inc_val = 0.0
        if pm:
            rot_shp_inc_val = m['rot_lc_shopee'] - pm['rot_lc_shopee']
            if rot_shp_inc_val >= 5.0:
                has_rot_shp_increase = True
                
        if is_bad_rot_tts or is_bad_rot_shp or has_rot_tts_increase or has_rot_shp_increase:
            details_list = []
            is_drop = has_rot_tts_increase or has_rot_shp_increase
            
            if has_rot_tts_increase:
                details_list.append(f"TTS rớt tăng mạnh {rot_tts_inc_val:.1f}% vs hôm qua (lên {m['rot_lc_tts']:.2f}%)")
            elif is_bad_rot_tts:
                details_list.append(f"TTS rớt {m['rot_lc_tts']:.2f}%")
                
            if has_rot_shp_increase:
                details_list.append(f"Shopee rớt tăng mạnh {rot_shp_inc_val:.1f}% vs hôm qua (lên {m['rot_lc_shopee']:.2f}%)")
            elif is_bad_rot_shp:
                details_list.append(f"Shopee rớt {m['rot_lc_shopee']:.2f}%")
                
            desc = "rớt luân chuyển (rớt LC) cao: " + " & ".join(details_list)
            m['issues'].append({
                'type': 'ROT_LC',
                'desc': desc,
                'severity_val': max(m['rot_lc_tts'], m['rot_lc_shopee']),
                'is_drop': is_drop,
                'base_score': 600
            })

    # ============ 10. GROUP QUESTIONS BY AM & FILTER BY CRITICALITY ============
    # List of all issues across the entire region
    all_issues = []
    
    for po, m in po_metrics.items():
        for issue in m['issues']:
            issue_type = issue['type']
            base = issue['base_score']
            is_drop = issue['is_drop']
            
            # Custom severity weight
            severity = 0.0
            if issue_type == 'ROT_LC':
                severity = float(max(m['rot_lc_tts'], m['rot_lc_shopee'])) * 2.0
            elif issue_type == 'OPR_TTS':
                severity = float(m['opr_late']) * 3.0
            elif issue_type == 'FD':
                severity = float(m['fd_rate'])
            elif issue_type == 'TTS_OPERATIONAL':
                severity = float(m['tts_ton']) / 5.0 + float(m['tts_chua_gan']) / 3.0
            elif issue_type == 'OPERATIONAL':
                severity = float(m['all_ton']) / 10.0 + float(m['all_chua_gan']) / 5.0
                
            crit_score = base + severity
            
            # Apply criticality bonus of +150 for sudden drops/spikes
            if is_drop:
                crit_score += 150.0
                
            all_issues.append({
                'po_name': po,
                'am_name': m['am_name'],
                'type': issue_type,
                'desc': issue['desc'],
                'severity': issue['severity_val'],
                'criticality': crit_score,
                'is_drop': is_drop
            })
            
    # Find the top 5 most critical issues in the entire region
    all_issues_sorted = sorted(all_issues, key=lambda x: -x['criticality'])
    top_5_keys = set()
    for item in all_issues_sorted[:5]:
        top_5_keys.add((item['po_name'], item['type']))
        
    # Group issues by (AM, PO)
    am_po_issues = {}
    for item in all_issues:
        am = item['am_name']
        po = item['po_name']
        if am not in am_po_issues:
            am_po_issues[am] = {}
        if po not in am_po_issues[am]:
            am_po_issues[am][po] = []
        am_po_issues[am][po].append(item)
        
    # Filter and format for each AM (Stage 1: Raw collection)
    am_questions = {}
    total_issues_count = 0
    
    for am, po_map in am_po_issues.items():
        # Sort POs by their maximum issue criticality score descending
        po_sorted = sorted(po_map.items(), key=lambda x: -max([issue['criticality'] for issue in x[1]]))
        # Keep only top 3 problematic POs per AM to keep it highly focused
        po_filtered = po_sorted[:3]
        
        am_questions[am] = []
        for po, po_issues in po_filtered:
            # Check if any issue in this PO is top 5 region-wide
            is_top_5 = any((issue['po_name'], issue['type']) in top_5_keys for issue in po_issues)
            
            # Generate the combined draft question
            meta = get_am_meta(am)
            pronoun = meta['pronoun']
            pronoun_lower = pronoun.lower()
            tag = meta['tag']
            
            descs = [x['desc'] for x in po_issues]
            descs_str = "; ".join(descs)
            
            if len(po_issues) == 1:
                text = f"{pronoun} {tag} ơi, check giúp em bưu cục {po} nha. Hiện tại bưu cục đang gặp vấn đề: {descs_str}. Nhờ {pronoun_lower} chỉ đạo rà soát và cho em giải pháp xử lý sớm nhé."
            else:
                text = f"{pronoun} {tag} ơi, rà soát giúp em tình hình bưu cục {po} hôm qua có nhiều chỉ số đuối/biến động: {descs_str}. Nhờ {pronoun_lower} kiểm tra nguyên nhân sự cố và tập trung xử lý dứt điểm các vấn đề trên nha."
                
            am_questions[am].append({
                'po_name': po,
                'text': text,
                'is_top_5': is_top_5,
                'criticality': max([x['criticality'] for x in po_issues])
            })
            
    # --- GEMINI BATCH REWRITING (Stage 2) ---
    flat_questions = []
    question_map = {}
    
    for am, po_qs in am_questions.items():
        for idx, q in enumerate(po_qs):
            q_id = len(flat_questions)
            flat_questions.append({"id": q_id, "text": q['text']})
            question_map[q_id] = q
            
    if flat_questions:
        # Load env vars again just in case
        load_dotenv(os.path.join(BASE_DIR, ".env"))
        api_key = os.environ.get("GEMINI_API_KEY")
        
        rewritten_list = rewrite_questions_gemini(flat_questions, api_key)
        if rewritten_list and isinstance(rewritten_list, list):
            print("✅ Gemini successfully rewrote all questions!")
            for item in rewritten_list:
                q_id = item.get('id')
                q_text = item.get('text')
                if q_id is not None and q_text and q_id in question_map:
                    question_map[q_id]['text'] = q_text
        else:
            print("⚠️ Falling back to original template questions.")
            
    # Prepend regional warnings prefix and update total count
    for am, po_qs in am_questions.items():
        for q in po_qs:
            prefix = "🚨 <b>[CẢNH BÁO TOÀN VÙNG]</b> " if q['is_top_5'] else ""
            q['text'] = prefix + q['text']
            total_issues_count += 1
            
    # Sort AMs by name
    sorted_ams = sorted(list(am_questions.keys()))
    
    # Generate GTalk formatted text
    gtalk_lines = []
    gtalk_lines.append(f"🔔 <b>BỘ CÂU HỎI ĐẦU NGÀY - GIẢI TRÌNH VẬN HÀNH (ĐÃ LỌC TRỌNG TÂM)</b> 📅 <b>Ngày {date_str_dmy}</b>")
    gtalk_lines.append(f"<i>Tổng hợp {total_issues_count} chỉ số nghiêm trọng nhất từ {len(am_questions)} AM khu vực NTB (tối đa 3 bưu cục/AM).</i>")
    gtalk_lines.append("")
    
    # Generate Markdown formatted text
    md_lines = []
    md_lines.append(f"# Morning Operational Questions for AM Explanation - Date: {date_str_dmy}")
    md_lines.append(f"*Filtered for key priorities, top regional drops, and instability (Max 3 bưu cục/AM)*")
    md_lines.append(f"*Total issues listed: {total_issues_count}*")
    md_lines.append("\n---\n")

    for am in sorted_ams:
        meta = get_am_meta(am)
        tag = meta['tag']
        
        # GTalk AM Block
        gtalk_lines.append(f"----------------------------------------")
        gtalk_lines.append(f"👤 <b>AM: {am}</b> ({tag})")
        gtalk_lines.append(f"----------------------------------------")
        
        # Markdown AM Block
        md_lines.append(f"## 👤 AM: {am} ({tag})")
        
        # Use directly pre-sorted questions list
        questions = am_questions[am]
        for q in questions:
            q_text = q['text']
            gtalk_lines.append(f"• {q_text}")
            md_lines.append(f"- {q_text}")
            
        gtalk_lines.append("")
        md_lines.append("")

    # If no issues found
    if total_issues_count == 0:
        msg = f"✅ Không phát hiện bưu cục nào vượt ngưỡng cảnh báo vận hành trong ngày {date_str_dmy}."
        gtalk_lines.append(msg)
        md_lines.append(f"### {msg}")
        
    gtalk_output = "\n".join(gtalk_lines)
    md_output = "\n".join(md_lines)

    # 1. Print human-readable output to console
    print("\n=================== GENERATED MORNING QUESTIONS ===================")
    print(gtalk_output)
    print("===================================================================\n")

    # 2. Write to local markdown file
    md_file_path = os.path.join(BASE_DIR, "morning_am_questions.md")
    try:
        with open(md_file_path, "w", encoding="utf-8") as f:
            f.write(md_output)
        print(f"💾 Report saved successfully to local file: {md_file_path}")
    except Exception as e:
        print(f"❌ Failed to save local markdown report: {e}")

    # 3. Broadcast to GTalk channel if requested
    if args.send:
        if total_issues_count > 0:
            print("📡 Starting sequential GTalk broadcast to prevent character limit errors...")
            # Send intro first
            intro = f"🔔 <b>BỘ CÂU HỎI ĐẦU NGÀY - GIẢI TRÌNH VẬN HÀNH (ĐÃ LỌC TRỌNG TÂM)</b> 📅 <b>Ngày {date_str_dmy}</b>\n<i>Tổng hợp {total_issues_count} chỉ số nghiêm trọng nhất từ {len(am_questions)} AM khu vực NTB (tối đa 3 bưu cục/AM).</i>"
            send_gtalk_message(intro)
            
            # Send each AM block in a separate message
            for am in sorted_ams:
                meta = get_am_meta(am)
                tag = meta['tag']
                
                lines = []
                lines.append(f"----------------------------------------")
                lines.append(f"👤 <b>AM: {am}</b> ({tag})")
                lines.append(f"----------------------------------------")
                
                questions = am_questions[am]
                for q in questions:
                    lines.append(f"• {q['text']}")
                    
                am_message = "\n".join(lines)
                send_gtalk_message(am_message)
            print("✅ Broadcast completed.")
        else:
            print("ℹ️ No issues detected. Skipping GTalk broadcast.")

if __name__ == "__main__":
    main()
