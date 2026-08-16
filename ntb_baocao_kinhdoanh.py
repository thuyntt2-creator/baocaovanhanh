import re
import sys
import os
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

sys.stdout.reconfigure(encoding='utf-8')

# ─────────────────────────────────────────────────────────────────────────────
# ⚙️  CONFIG
# ─────────────────────────────────────────────────────────────────────────────
# Dynamic check for credentials.json path
CREDENTIALS_PATH = r'C:\Users\lap4all\Desktop\Backlog_Automation\credentials.json'
if not os.path.exists(CREDENTIALS_PATH):
    CREDENTIALS_PATH = 'credentials.json'

SPREADSHEET_ID   = '1XaziS_8UB2lCwL01bxam106EdIUHtlMISDGJX1PU5A8'
OLD_SPREADSHEET_ID = '12VsqpIx1vLOqk6-JKHwgRmdrgu6fxSNPQcpyO0wpvdQ'

# Tên sheet nguồn
SHEET_NGAY   = "tong_quan"
SHEET_TUAN   = "tong_quan"
SHEET_THANG  = "tong_quan"
SHEET_KHM    = "f30"
SHEET_COCAU  = "CoCauVung"
SHEET_NHOM   = "phan_nhom"

# Tên sheet output (tạo mới hoặc ghi đè)
OUT_DAILY = "RPT_Ngày"
OUT_TUAN  = "RPT_Tuần"
OUT_MTD   = "RPT_Tháng"
OUT_KHM   = "RPT_KHM"
OUT_NHOM  = "RPT_NhomKH"

TINH_ORDER = ["Khánh Hòa", "Lâm Đồng", "Đắk Nông", "Ninh Thuận", "Bình Thuận"]

# AM đang hoạt động nhưng chưa có trong sheet Cocauvung → bổ sung thủ công
AM_EXTRA_MAP = {
    "Trần Công Hậu":          "Khánh Hòa",
    "Phạm Đức Thắng":         "Lâm Đồng",
    "Nguyễn Vĩnh Tường":      "Khánh Hòa",
    "Nguyễn Tống Hùng Phong": "Khánh Hòa",
}

# ── Màu ──────────────────────────────────────────────────────────────────────
def rgb(h):
    h = h.lstrip("#")
    return {"red":int(h[0:2],16)/255,"green":int(h[2:4],16)/255,"blue":int(h[4:6],16)/255}

NAVY   = rgb("1F3864"); BLUE   = rgb("0072BC"); ORANGE = rgb("F26522")
SUBHDR = rgb("2E75B6"); GREEN  = rgb("2E7D32"); GOLD   = rgb("FFC000")
GRAY   = rgb("F2F2F2"); WHITE  = rgb("FFFFFF"); DARK   = rgb("1F1F1F")
C_UP   = rgb("375623"); C_DOWN = rgb("C00000"); YELLOW = rgb("FFFACD")


# ─────────────────────────────────────────────────────────────────────────────
# AUTH
# ─────────────────────────────────────────────────────────────────────────────
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]

def connect():
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(SPREADSHEET_ID)

def connect_old():
    creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(OLD_SPREADSHEET_ID)


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOAD
# ─────────────────────────────────────────────────────────────────────────────
_MONTH_MAP = {f"thg {i}": i for i in range(1, 13)}

def parse_vn_date(s):
    if isinstance(s, (int, float)):
        return pd.to_datetime(s, unit='D', origin='1899-12-30').to_pydatetime()
    s = str(s).strip()
    if s.isdigit():
        return pd.to_datetime(int(s), unit='D', origin='1899-12-30').to_pydatetime()
    for m, num in _MONTH_MAP.items():
        if m in s:
            parts = re.findall(r"\d+", s)
            if len(parts) >= 2:
                return datetime(int(parts[-1]), num, int(parts[0]))
    return None

def ws_to_df(ss, name):
    print(f"  📥 {name} ...", end=" ", flush=True)
    data = ss.worksheet(name).get_all_values(value_render_option='UNFORMATTED_VALUE')
    if not data: print("trống"); return pd.DataFrame()
    df = pd.DataFrame(data[1:], columns=data[0])
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except Exception: pass
    print(f"{len(df)} dòng ✅"); return df

def load_nhom(ss):
    print(f"  📥 {SHEET_NHOM} ...", end=" ", flush=True)
    data = ss.worksheet(SHEET_NHOM).get_all_values(value_render_option='UNFORMATTED_VALUE')
    df = pd.DataFrame(data[1:], columns=data[0])

    def parse_nhom_date(val):
        p = parse_vn_date(val)
        return p if p is not None else pd.NaT

    df['Ngay'] = df['Ngay'].apply(parse_nhom_date)
    df['DT']   = pd.to_numeric(df['DT'],    errors='coerce').fillna(0)
    df['MTD']  = pd.to_numeric(df['MTD'],   errors='coerce').fillna(0)
    df = df.dropna(subset=['MaKH','Ngay'])
    df_ntb = df[df['Vung'].astype(str).str.startswith('NTB')].copy()
    df_ntb['Tinh'] = (df_ntb['Vung']
                      .str.replace('NTB-','',regex=False)
                      .str.replace(r'^BD$','Bình Thuận',regex=True)
                      .str.strip())
    print(f"{len(df_ntb)} dòng ✅")
    return df_ntb

def load_data(ss):
    print("📂 Đọc data từ Google Sheet...")
    df_tq    = ws_to_df(ss, "tong_quan")
    df_khm   = ws_to_df(ss, SHEET_KHM)
    df_coc   = ws_to_df(ss, SHEET_COCAU)
    df_nhom  = load_nhom(ss)

    try:
        ss_old = connect_old()
        df_old_day = ws_to_df(ss_old, "Theo ngày")
    except Exception as e:
        print(f"⚠️ Không load được spreadsheet cũ: {e}")
        df_old_day = pd.DataFrame()

    # Filter out empty/trailing rows from CoCauVung
    df_coc = df_coc[df_coc["AM"].astype(str).str.strip() != ""]
    df_coc = df_coc[df_coc["Tỉnh"].astype(str).str.strip() != ""]

    am_tinh = (df_coc[["AM","Tỉnh"]].dropna()
               .drop_duplicates("AM").set_index("AM")["Tỉnh"].to_dict())
    am_tinh.update(AM_EXTRA_MAP)

    def get_tinh(am):
        am = str(am).strip()
        if not am or am == "-": return None
        if am in am_tinh: return am_tinh[am]
        for k, v in am_tinh.items():
            if am in k or k in am: return v
        return None

    # Process tong_quan for Daily/Weekly/Monthly MTD
    df_tq["date"]     = df_tq["Ngay"].apply(parse_vn_date)
    df_tq["DoanhThu"] = pd.to_numeric(df_tq["DoanhThu"], errors="coerce").fillna(0)
    df_tq["Volume"]   = pd.to_numeric(df_tq["Volume"],   errors="coerce").fillna(0)
    df_tq = df_tq.dropna(subset=["date"])
    df_tq = df_tq[~df_tq["AM_format"].astype(str).str.contains(",", na=False)]
    df_tq["Tinh"] = df_tq["AM_format"].apply(get_tinh)
    df_tq = df_tq[df_tq["Tinh"].notna() & (df_tq["Tinh"] != "")]

    # Process df_old_day for YoY comparison if present
    if not df_old_day.empty:
        df_old_day["date"]     = df_old_day["Ngay"].apply(parse_vn_date)
        df_old_day["DoanhThu"] = pd.to_numeric(df_old_day["DoanhThu"], errors="coerce").fillna(0)
        df_old_day["Volume"]   = pd.to_numeric(df_old_day["Volume"],   errors="coerce").fillna(0)
        df_old_day = df_old_day.dropna(subset=["date"])
        df_old_day = df_old_day[~df_old_day["AM_format"].astype(str).str.contains(",", na=False)]
        df_old_day["Tinh"] = df_old_day["AM_format"].apply(get_tinh)
        df_old_day = df_old_day[df_old_day["Tinh"].notna() & (df_old_day["Tinh"] != "")]

    dates  = sorted(df_tq["date"].unique())
    d_cur  = dates[-1]; d_prev = dates[-2]
    d7     = min(dates, key=lambda d: abs((d-(d_cur-timedelta(days=7))).days))

    def agg_tinh(d):
        return (df_tq[df_tq["date"]==d]
                .groupby("Tinh")[["DoanhThu","Volume"]].sum()
                .rename(columns={"DoanhThu":"DT","Volume":"Vol"}))

    def agg_am(d):
        return (df_tq[df_tq["date"]==d]
                .groupby(["AM_format","Tinh"])[["DoanhThu","Volume"]].sum()
                .rename(columns={"DoanhThu":"DT","Volume":"Vol"})
                .reset_index().sort_values("DT", ascending=False))

    # Weekly (computed from df_tq dates)
    df_tq["Tuan"] = df_tq["date"].apply(lambda d: f"{d.year}/{d.isocalendar()[1]:02d}" if d else None)
    weeks   = sorted([w for w in df_tq["Tuan"].unique() if w])
    w_cur, w_prev = weeks[-1], weeks[-2]

    def agg_week(w):
        return (df_tq[df_tq["Tuan"]==w]
                .groupby("Tinh")[["DoanhThu","Volume"]].sum()
                .rename(columns={"DoanhThu":"DT","Volume":"Vol"}))

    # Monthly MTD (computed from df_tq dates using custom date bounds)
    m_start      = d_cur.replace(day=1)
    prev_m_start = (m_start-timedelta(days=1)).replace(day=1)
    prev_m_end   = prev_m_start.replace(day=d_cur.day)

    yoy_start    = m_start.replace(year=m_start.year-1)
    yoy_end      = yoy_start.replace(day=d_cur.day)

    def agg_mtd(start_dt, end_dt):
        df_source = df_tq if start_dt.year == 2026 else df_old_day
        if df_source.empty:
            return pd.DataFrame(columns=["DT", "Vol"])
        return (df_source[(df_source["date"] >= start_dt) & (df_source["date"] <= end_dt)]
                .groupby("Tinh")[["DoanhThu","Volume"]].sum()
                .rename(columns={"DoanhThu":"DT","Volume":"Vol"}))

    # KHM
    df_khm["Ngay"]           = df_khm["Ngày LTC đầu tiên"].apply(parse_vn_date)
    df_khm["DoanhThu_NoVAT"] = pd.to_numeric(df_khm["DoanhThu_NoVAT"], errors="coerce").fillna(0)
    df_khm["Volume"]         = pd.to_numeric(df_khm["Volume"],          errors="coerce").fillna(0)
    def determine_khm_tinh(row):
        t = str(row.get("Tinh", "")).strip()
        if t in TINH_ORDER:
            return t
        am = str(row.get("AM", "")).strip()
        if am:
            return get_tinh(am)
        return None

    df_khm["Tinh_mapped"] = df_khm.apply(determine_khm_tinh, axis=1)
    df_ntb = df_khm[df_khm["Tinh_mapped"].notna()].copy()
    df_ntb["Tinh"] = df_ntb["Tinh_mapped"]

    def agg_khm(s, e=None):
        mask = (df_ntb["Ngay"]==s) if e is None else (df_ntb["Ngay"]>=s)&(df_ntb["Ngay"]<=e)
        return (df_ntb[mask].groupby("Tinh")
                .agg(SLKH=("Mã KH","count"),Vol=("Volume","sum"),DT=("DoanhThu_NoVAT","sum")))

    # Calculate WTD Group Revenue
    df_nhom['AOV_numeric'] = pd.to_numeric(df_nhom['AOV'], errors='coerce').fillna(0)
    df_nhom['Revenue'] = df_nhom['DT'] * df_nhom['AOV_numeric']

    wtd_start = d_cur - timedelta(days=d_cur.weekday())
    wtd_prev_start = wtd_start - timedelta(days=7)
    d_prev_wtd_end = d_cur - timedelta(days=7)

    df_wtd_cur = df_nhom[(df_nhom['Ngay'] >= wtd_start) & (df_nhom['Ngay'] <= d_cur)]
    df_wtd_prev = df_nhom[(df_nhom['Ngay'] >= wtd_prev_start) & (df_nhom['Ngay'] <= d_prev_wtd_end)]

    wtd_cur_by_group = df_wtd_cur.groupby('Nhom')['Revenue'].sum().to_dict()
    wtd_prev_by_group = df_wtd_prev.groupby('Nhom')['Revenue'].sum().to_dict()

    print(f"\n✅ d_cur={d_cur.date()} | d_prev={d_prev.date()} | d7={d7.date()}")
    print(f"   Tuần: {w_prev} → {w_cur} | MTD: T{prev_m_start.month} vs T{m_start.month}")

    return dict(
        d_cur=d_cur, d_prev=d_prev, d7=d7,
        t_cur=agg_tinh(d_cur), t_prev=agg_tinh(d_prev), t_d7=agg_tinh(d7),
        am_cur=agg_am(d_cur),  am_prev=agg_am(d_prev),
        w_cur=w_cur, w_prev=w_prev,
        wt_cur=agg_week(w_cur), wt_prev=agg_week(w_prev),
        m_start=m_start, prev_m_start=prev_m_start, prev_m_end=prev_m_end,
        mtd_cur=agg_mtd(m_start, d_cur),
        mtd_prev=agg_mtd(prev_m_start, prev_m_end),
        mtd_yoy=agg_mtd(yoy_start, yoy_end),
        yoy_start=yoy_start, yoy_end=yoy_end,
        khm_cur=agg_khm(d_cur), khm_prev=agg_khm(d_prev), khm_d7=agg_khm(d7),
        khm_mtd_cur=agg_khm(m_start,d_cur), khm_mtd_prev=agg_khm(prev_m_start,prev_m_end),
        wtd_cur_by_group=wtd_cur_by_group,
        wtd_prev_by_group=wtd_prev_by_group,
        df_nhom=df_nhom,
    )


# ─────────────────────────────────────────────────────────────────────────────
# GSPREAD REQUEST BUILDERS
# ─────────────────────────────────────────────────────────────────────────────
def col_letter(n):
    r = ""
    while n:
        n, rem = divmod(n-1, 26)
        r = chr(65+rem) + r
    return r

def get_or_create(ss, title, rows=300, cols=20):
    try:
        ws = ss.worksheet(title)
        ss.del_worksheet(ws)
    except gspread.WorksheetNotFound:
        pass
    return ss.add_worksheet(title=title, rows=rows, cols=cols)

def _range(ws_id, r1, c1, r2, c2):
    return {"sheetId":ws_id,"startRowIndex":r1-1,"endRowIndex":r2,
            "startColumnIndex":c1-1,"endColumnIndex":c2}

def R_bg(ws_id, r1, c1, r2, c2, color):
    return {"repeatCell":{"range":_range(ws_id,r1,c1,r2,c2),
            "cell":{"userEnteredFormat":{"backgroundColor":color}},
            "fields":"userEnteredFormat.backgroundColor"}}

def R_font(ws_id, r1, c1, r2, c2, color=None, size=9, bold=True):
    color = color or WHITE
    return {"repeatCell":{"range":_range(ws_id,r1,c1,r2,c2),
            "cell":{"userEnteredFormat":{"textFormat":{
                "foregroundColor":color,"fontSize":size,"bold":bold,"fontFamily":"Arial"}}},
            "fields":"userEnteredFormat.textFormat"}}

def R_align(ws_id, r1, c1, r2, c2, h="CENTER", v="MIDDLE", wrap=True):
    return {"repeatCell":{"range":_range(ws_id,r1,c1,r2,c2),
            "cell":{"userEnteredFormat":{
                "horizontalAlignment":h,"verticalAlignment":v,
                "wrapStrategy":"WRAP" if wrap else "OVERFLOW_CELL"}},
            "fields":"userEnteredFormat(horizontalAlignment,verticalAlignment,wrapStrategy)"}}

def R_border(ws_id, r1, c1, r2, c2):
    s = {"style":"SOLID","color":rgb("BFBFBF")}
    return {"repeatCell":{"range":_range(ws_id,r1,c1,r2,c2),
            "cell":{"userEnteredFormat":{"borders":{"top":s,"bottom":s,"left":s,"right":s}}},
            "fields":"userEnteredFormat.borders"}}

def R_merge(ws_id, r1, c1, r2, c2):
    return {"mergeCells":{"range":_range(ws_id,r1,c1,r2,c2),"mergeType":"MERGE_ALL"}}

def R_row_h(ws_id, row, px):
    return {"updateDimensionProperties":{
        "range":{"sheetId":ws_id,"dimension":"ROWS","startIndex":row-1,"endIndex":row},
        "properties":{"pixelSize":px},"fields":"pixelSize"}}

def R_col_w(ws_id, col, px):
    return {"updateDimensionProperties":{
        "range":{"sheetId":ws_id,"dimension":"COLUMNS","startIndex":col-1,"endIndex":col},
        "properties":{"pixelSize":px},"fields":"pixelSize"}}

def R_freeze(ws_id, rows=1):
    return {"updateSheetProperties":{
        "properties":{"sheetId":ws_id,"gridProperties":{"frozenRowCount":rows}},
        "fields":"gridProperties.frozenRowCount"}}


# ─────────────────────────────────────────────────────────────────────────────
# WRITE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def pct_str(cur, prev):
    if not prev: return "—"
    p = (cur-prev)/prev*100
    return f"{'▲' if p>=0 else '▼'} {abs(p):.1f}%"

def fmt_dt(v):  return round(v/1e6, 1)
def fmt_vol(v): return int(round(v))
def is_up(s):   return isinstance(s,str) and "▲" in s
def is_dn(s):   return isinstance(s,str) and "▼" in s

def _safe(v):
    try:
        import numpy as np
        if isinstance(v, np.integer): return int(v)
        if isinstance(v, np.floating): return float(v)
        if isinstance(v, np.bool_): return bool(v)
    except ImportError:
        pass
    return v

def flush(ws, vals, max_row):
    if not vals: return
    max_col = max(c for _,c in vals)
    grid = [[""] * max_col for _ in range(max_row)]
    for (r,c),v in vals.items():
        if r <= max_row: grid[r-1][c-1] = _safe(v)
    ws.update(range_name="A1", values=grid)

def section_title(vals, reqs, ws_id, row, ncols, text, bg):
    vals[(row,1)] = text
    reqs += [R_merge(ws_id,row,1,row,ncols), R_bg(ws_id,row,1,row,ncols,bg),
             R_font(ws_id,row,1,row,ncols,WHITE,11,True),
             R_align(ws_id,row,1,row,ncols), R_row_h(ws_id,row,24)]

def write_hdr(vals, reqs, ws_id, row, labels, bg):
    for c,h in enumerate(labels,1): vals[(row,c)] = h
    n = len(labels)
    reqs += [R_bg(ws_id,row,1,row,n,bg), R_font(ws_id,row,1,row,n,WHITE,9,True),
             R_align(ws_id,row,1,row,n,wrap=True),
             R_border(ws_id,row,1,row,n), R_row_h(ws_id,row,34)]

def write_data_row(vals, reqs, ws_id, row, data, ncols, pct_cols=None):
    for c,v in enumerate(data,1): vals[(row,c)] = v
    reqs += [R_border(ws_id,row,1,row,ncols),
             R_font(ws_id,row,1,row,ncols,DARK,9,False),
             R_align(ws_id,row,1,row,1,"LEFT")]
    if row%2==0: reqs.append(R_bg(ws_id,row,1,row,ncols,GRAY))
    for ci in (pct_cols or []):
        v = data[ci-1]
        if is_up(v): reqs.append(R_font(ws_id,row,ci,row,ci,C_UP,9,True))
        elif is_dn(v): reqs.append(R_font(ws_id,row,ci,row,ci,C_DOWN,9,True))

def write_total(vals, reqs, ws_id, row, data, ncols):
    for c,v in enumerate(data,1): vals[(row,c)] = v
    reqs += [R_bg(ws_id,row,1,row,ncols,NAVY),
             R_font(ws_id,row,1,row,ncols,WHITE,9,True),
             R_align(ws_id,row,1,row,ncols),
             R_align(ws_id,row,1,row,1,"LEFT"),
             R_border(ws_id,row,1,row,ncols), R_row_h(ws_id,row,22)]


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 1 – DAILY
# ─────────────────────────────────────────────────────────────────────────────
def build_daily(ss, D):
    ws = get_or_create(ss, OUT_DAILY, rows=200, cols=8)
    ws_id = ws.id; reqs = []; vals = {}
    d_cur=D["d_cur"]; d_prev=D["d_prev"]; d7=D["d7"]
    NCOLS = 7

    for c,w in enumerate([200,130,130,110,130,110,150],1): reqs.append(R_col_w(ws_id,c,w))
    reqs.append(R_freeze(ws_id))

    row = 1
    vals[(row,1)] = f"BÁO CÁO KINH DOANH VÙNG NTB – NGÀY {d_cur.strftime('%d/%m/%Y')}"
    reqs += [R_merge(ws_id,row,1,row,NCOLS), R_bg(ws_id,row,1,row,NCOLS,NAVY),
             R_font(ws_id,row,1,row,NCOLS,WHITE,13,True),
             R_align(ws_id,row,1,row,NCOLS), R_row_h(ws_id,row,36)]
    row += 2

    for metric, label_unit, bg_sec in [
        ("DT","triệu đồng",BLUE), ("Vol","đơn",ORANGE)
    ]:
        fmt = fmt_dt if metric=="DT" else fmt_vol
        icon = "📊 DOANH THU" if metric=="DT" else "📦 SẢN LƯỢNG"
        section_title(vals,reqs,ws_id,row,NCOLS,f"{icon} THEO TỈNH ({label_unit})",bg_sec); row+=1
        write_hdr(vals,reqs,ws_id,row,[
            "TỈNH", f"D ({d_cur.strftime('%d/%m')})",
            f"D-1 ({d_prev.strftime('%d/%m')})", "DoD vs D-1",
            f"D-7 ({d7.strftime('%d/%m')})", "DoD vs D-7", "Ghi chú"],bg_sec); row+=1

        tc=tp=t7=0
        for tinh in TINH_ORDER:
            vc=D["t_cur"][metric].get(tinh,0)  if tinh in D["t_cur"].index  else 0
            vp=D["t_prev"][metric].get(tinh,0) if tinh in D["t_prev"].index else 0
            v7=D["t_d7"][metric].get(tinh,0)   if tinh in D["t_d7"].index   else 0
            if vc==0 and vp==0: continue
            tc+=vc; tp+=vp; t7+=v7
            write_data_row(vals,reqs,ws_id,row,
                [tinh,fmt(vc),fmt(vp),pct_str(vc,vp),fmt(v7),pct_str(vc,v7),""],
                NCOLS, pct_cols=[4,6]); row+=1
        write_total(vals,reqs,ws_id,row,
            ["TỔNG VÙNG NTB",fmt(tc),fmt(tp),pct_str(tc,tp),fmt(t7),pct_str(tc,t7),""],NCOLS)
        row+=2

    # AM table
    section_title(vals,reqs,ws_id,row,NCOLS,"👤 CHI TIẾT THEO AM",NAVY); row+=1
    write_hdr(vals,reqs,ws_id,row,[
        "AM","Tỉnh",
        f"DT {d_cur.strftime('%d/%m')} (M)", f"Vol {d_cur.strftime('%d/%m')}",
        f"DT {d_prev.strftime('%d/%m')} (M)", f"Vol {d_prev.strftime('%d/%m')}",
        "DoD DT"], SUBHDR); row+=1

    merged = (D["am_cur"].merge(D["am_prev"],on=["AM_format","Tinh"],how="outer",
              suffixes=("_c","_p")).fillna(0).sort_values("DT_c",ascending=False))
    for _,r in merged.iterrows():
        dod = pct_str(r["DT_c"],r["DT_p"])
        write_data_row(vals,reqs,ws_id,row,
            [r["AM_format"],r["Tinh"],fmt_dt(r["DT_c"]),fmt_vol(r["Vol_c"]),
             fmt_dt(r["DT_p"]),fmt_vol(r["Vol_p"]),dod],
            NCOLS, pct_cols=[7])
        reqs.append(R_align(ws_id,row,2,row,2,"LEFT"))
        row+=1

    flush(ws, vals, row); ss.batch_update({"requests":reqs})
    print(f"  ✅ {OUT_DAILY}")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 2 – WEEKLY
# ─────────────────────────────────────────────────────────────────────────────
def build_weekly(ss, D):
    ws = get_or_create(ss, OUT_TUAN, rows=100, cols=6)
    ws_id = ws.id; reqs=[]; vals={}
    NCOLS=5; w_cur=D["w_cur"]; w_prev=D["w_prev"]

    for c,w in enumerate([200,130,130,110,130],1): reqs.append(R_col_w(ws_id,c,w))
    reqs.append(R_freeze(ws_id))

    row=1
    vals[(row,1)] = f"DOANH THU & SẢN LƯỢNG THEO TUẦN | {w_prev} vs {w_cur}"
    reqs += [R_merge(ws_id,row,1,row,NCOLS), R_bg(ws_id,row,1,row,NCOLS,NAVY),
             R_font(ws_id,row,1,row,NCOLS,WHITE,13,True),
             R_align(ws_id,row,1,row,NCOLS), R_row_h(ws_id,row,36)]
    row+=2

    for metric,label,unit,bg in [("DT","📊 DOANH THU","triệu đồng",BLUE),
                                   ("Vol","📦 SẢN LƯỢNG","đơn",ORANGE)]:
        fmt = fmt_dt if metric=="DT" else fmt_vol
        section_title(vals,reqs,ws_id,row,NCOLS,f"{label} THEO TỈNH ({unit})",bg); row+=1
        write_hdr(vals,reqs,ws_id,row,["TỈNH",w_prev,w_cur,"WoW (%)","Chênh lệch"],bg); row+=1
        tc=tp=0
        for tinh in TINH_ORDER:
            vc=D["wt_cur"][metric].get(tinh,0)  if tinh in D["wt_cur"].index  else 0
            vp=D["wt_prev"][metric].get(tinh,0) if tinh in D["wt_prev"].index else 0
            if vc==0 and vp==0: continue
            tc+=vc; tp+=vp
            write_data_row(vals,reqs,ws_id,row,
                [tinh,fmt(vp),fmt(vc),pct_str(vc,vp),fmt(vc-vp)],NCOLS,pct_cols=[4]); row+=1
        write_total(vals,reqs,ws_id,row,
            ["TỔNG NTB",fmt(tp),fmt(tc),pct_str(tc,tp),fmt(tc-tp)],NCOLS); row+=2

    flush(ws,vals,row); ss.batch_update({"requests":reqs})
    print(f"  ✅ {OUT_TUAN}")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 3 – MTD
# ─────────────────────────────────────────────────────────────────────────────
def build_mtd(ss, D):
    ws = get_or_create(ss, OUT_MTD, rows=120, cols=8)
    ws_id=ws.id; reqs=[]; vals={}
    NCOLS=7
    m=D["m_start"]; pm=D["prev_m_start"]; pe=D["prev_m_end"]; d=D["d_cur"]
    yoy_start=D["yoy_start"]; yoy_end=D["yoy_end"]

    for c,w in enumerate([200,130,130,110,130,110,110],1): reqs.append(R_col_w(ws_id,c,w))
    reqs.append(R_freeze(ws_id))

    row=1
    vals[(row,1)] = (f"DOANH THU LUỸ KẾ (MTD) | T{pm.month}/{pm.year} vs T{m.month}/{m.year} "
                     f"vs T{yoy_start.month}/{yoy_start.year} (YoY)")
    reqs += [R_merge(ws_id,row,1,row,NCOLS), R_bg(ws_id,row,1,row,NCOLS,NAVY),
             R_font(ws_id,row,1,row,NCOLS,WHITE,13,True),
             R_align(ws_id,row,1,row,NCOLS), R_row_h(ws_id,row,36)]
    row+=2

    lbl_p   = f"MTD T{pm.month}/{pm.year} (1–{pe.day})"
    lbl_c   = f"MTD T{m.month}/{m.year} (1–{d.day})"
    lbl_yoy = f"MTD T{yoy_start.month}/{yoy_start.year} (1–{yoy_end.day})"

    for metric,label,unit,bg in [("DT","📊 DOANH THU MTD","triệu đồng",BLUE),
                                   ("Vol","📦 SẢN LƯỢNG MTD","đơn",ORANGE)]:
        fmt = fmt_dt if metric=="DT" else fmt_vol
        section_title(vals,reqs,ws_id,row,NCOLS,f"{label} ({unit})",bg); row+=1
        write_hdr(vals,reqs,ws_id,row,
            ["TỈNH", lbl_p, lbl_c, "MoM (%)", lbl_yoy, "YoY (%)", "Chênh lệch YoY"],
            bg); row+=1
        tc=tp=ty=0
        for tinh in TINH_ORDER:
            vc=D["mtd_cur"][metric].get(tinh,0)   if tinh in D["mtd_cur"].index  else 0
            vp=D["mtd_prev"][metric].get(tinh,0)  if tinh in D["mtd_prev"].index else 0
            vy=D["mtd_yoy"][metric].get(tinh,0)   if tinh in D["mtd_yoy"].index  else 0
            if vc==0 and vp==0 and vy==0: continue
            tc+=vc; tp+=vp; ty+=vy
            write_data_row(vals,reqs,ws_id,row,
                [tinh, fmt(vp), fmt(vc), pct_str(vc,vp),
                 fmt(vy), pct_str(vc,vy), fmt(vc-vy)],
                NCOLS, pct_cols=[4,6]); row+=1
        write_total(vals,reqs,ws_id,row,
            ["TỔNG NTB", fmt(tp), fmt(tc), pct_str(tc,tp),
             fmt(ty), pct_str(tc,ty), fmt(tc-ty)], NCOLS); row+=2

    flush(ws,vals,row); ss.batch_update({"requests":reqs})
    print(f"  ✅ {OUT_MTD}")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 4 – KHM
# ─────────────────────────────────────────────────────────────────────────────
def build_khm(ss, D):
    ws = get_or_create(ss, OUT_KHM, rows=100, cols=14)
    ws_id=ws.id; reqs=[]; vals={}
    NCOLS=13
    d_cur=D["d_cur"]; d_prev=D["d_prev"]
    m=D["m_start"]; pm=D["prev_m_start"]; pe=D["prev_m_end"]

    for c,w in enumerate([200]+[90]*12,1): reqs.append(R_col_w(ws_id,c,w))
    reqs.append(R_freeze(ws_id))

    row=1
    vals[(row,1)] = f"KHÁCH HÀNG MỚI (KHM) – VÙNG NTB | đến {d_cur.strftime('%d/%m/%Y')}"
    reqs += [R_merge(ws_id,row,1,row,NCOLS), R_bg(ws_id,row,1,row,NCOLS,NAVY),
             R_font(ws_id,row,1,row,NCOLS,WHITE,13,True),
             R_align(ws_id,row,1,row,NCOLS), R_row_h(ws_id,row,36)]
    row+=2

    section_title(vals,reqs,ws_id,row,NCOLS,"📈 KHM THEO TỈNH",GREEN); row+=1

    # Group header
    grp = [
        (1,1,   NAVY,           "TỈNH"),
        (2,4,   rgb("FFC000"),  f"Ngày {d_cur.strftime('%d/%m')}"),
        (5,7,   rgb("E6A817"),  f"Ngày {d_prev.strftime('%d/%m')}"),
        (8,10,  rgb("2E7D32"),  f"MTD T{m.month} (1–{d_cur.day})"),
        (11,13, rgb("1B5E20"),  f"MTD T{pm.month} (1–{pe.day})"),
    ]
    for cs,ce,color,label in grp:
        vals[(row,cs)] = label
        if cs!=ce: reqs.append(R_merge(ws_id,row,cs,row,ce))
        reqs += [R_bg(ws_id,row,cs,row,ce,color),
                 R_font(ws_id,row,cs,row,ce,WHITE,9,True),
                 R_align(ws_id,row,cs,row,ce),
                 R_border(ws_id,row,cs,row,ce)]
    reqs.append(R_row_h(ws_id,row,22)); row+=1

    # Sub-header
    sub    = ["TỈNH"]+["#KH","Vol","DT(M)"]*4
    colors = [NAVY]+[rgb("FFC000")]*3+[rgb("E6A817")]*3+[rgb("2E7D32")]*3+[rgb("1B5E20")]*3
    for c,(h,color) in enumerate(zip(sub,colors),1):
        vals[(row,c)] = h
        reqs += [R_bg(ws_id,row,c,row,c,color),
                 R_font(ws_id,row,c,row,c,WHITE,9,True),
                 R_align(ws_id,row,c,row,c,wrap=True),
                 R_border(ws_id,row,c,row,c)]
    reqs.append(R_row_h(ws_id,row,28)); row+=1

    def gv(df,tinh,col):
        return df.loc[tinh,col] if (tinh in df.index and col in df.columns) else 0

    for tinh in TINH_ORDER:
        kc=D["khm_cur"]; kp=D["khm_prev"]; km=D["khm_mtd_cur"]; kmp=D["khm_mtd_prev"]
        rd = [tinh,
              gv(kc,tinh,"SLKH"),gv(kc,tinh,"Vol"),round(gv(kc,tinh,"DT")/1e6,3),
              gv(kp,tinh,"SLKH"),gv(kp,tinh,"Vol"),round(gv(kp,tinh,"DT")/1e6,3),
              gv(km,tinh,"SLKH"),gv(km,tinh,"Vol"),round(gv(km,tinh,"DT")/1e6,1),
              gv(kmp,tinh,"SLKH"),gv(kmp,tinh,"Vol"),round(gv(kmp,tinh,"DT")/1e6,1)]
        if all(v==0 for v in rd[1:]): continue
        for c,v in enumerate(rd,1): vals[(row,c)] = v if v!=0 else "—"
        reqs += [R_border(ws_id,row,1,row,NCOLS),
                 R_font(ws_id,row,1,row,NCOLS,DARK,9,False),
                 R_align(ws_id,row,1,row,1,"LEFT")]
        if row%2==0: reqs.append(R_bg(ws_id,row,1,row,NCOLS,GRAY))
        row+=1

    def sums(df):
        if df.empty: return 0,0,0.0
        return int(df["SLKH"].sum()),int(df["Vol"].sum()),df["DT"].sum()/1e6
    sc,vc,dc = sums(D["khm_cur"]); sp,vp,dp = sums(D["khm_prev"])
    sm,vm,dm = sums(D["khm_mtd_cur"]); smp,vmp,dmp = sums(D["khm_mtd_prev"])

    total_row = [
        "TỔNG NTB",
        sc, vc, round(dc,3) if dc!=0 else "—",
        sp, vp, round(dp,3) if dp!=0 else "—",
        sm, vm, round(dm,1) if dm!=0 else "—",
        smp, vmp, round(dmp,1) if dmp!=0 else "—"
    ]
    write_total(vals,reqs,ws_id,row,total_row,NCOLS); row+=2

    note = (f"DoD #KH: {pct_str(sc,sp)}  |  DoD Vol: {pct_str(vc,vp)}  |  "
            f"MTD T{m.month} vs T{pm.month} #KH: {pct_str(sm,smp)}")
    vals[(row,1)] = note
    reqs += [R_merge(ws_id,row,1,row,NCOLS), R_bg(ws_id,row,1,row,NCOLS,YELLOW),
             R_font(ws_id,row,1,row,NCOLS,DARK,10,True),
             R_align(ws_id,row,1,row,NCOLS)]

    flush(ws,vals,row); ss.batch_update({"requests":reqs})
    print(f"  ✅ {OUT_KHM}")


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 5 – NHÓM KH (KH không lên đơn / KH rời bỏ / DT KHM deep dive)
# ─────────────────────────────────────────────────────────────────────────────
def build_nhom(ss, D):
    df_ntb = D["df_nhom"]
    ws     = get_or_create(ss, OUT_NHOM, rows=300, cols=10)
    ws_id  = ws.id; reqs=[]; vals={}

    d_cur  = D["d_cur"]; d_prev = D["d_prev"]; d7 = D["d7"]
    m_start = D["m_start"]; prev_m_start = D["prev_m_start"]; prev_m_end = D["prev_m_end"]
    NCOLS = 8

    for c,w in enumerate([22,12,120,120,120,90,90,130],1):
        reqs.append(R_col_w(ws_id, c, w*5))
    reqs.append(R_freeze(ws_id, rows=1))

    def on(d):
        return df_ntb[df_ntb['Ngay']==pd.Timestamp(d)]

    # helper
    def put(r,c,v): vals[(r,c)] = v

    row = 1
    put(row,1, f"PHÂN TÍCH NHÓM KH – VÙNG NTB | NGÀY {d_cur.strftime('%d/%m/%Y')}")
    reqs += [R_merge(ws_id,row,1,row,NCOLS), R_bg(ws_id,row,1,row,NCOLS,NAVY),
             R_font(ws_id,row,1,row,NCOLS,WHITE,13,True),
             R_align(ws_id,row,1,row,NCOLS), R_row_h(ws_id,row,36)]
    row += 2

    # ══ SECTION 0: DOANH THU LUỸ KẾ THEO NHÓM KH (WTD) ═══════════════════════
    section_title(vals,reqs,ws_id,row,NCOLS,"📊 DOANH THU LUỸ KẾ THEO NHÓM KH (WTD)",BLUE); row+=1
    
    wtd_cur = D.get("wtd_cur_by_group", {})
    wtd_prev = D.get("wtd_prev_by_group", {})
    total_wtd_cur = sum(wtd_cur.values())
    sorted_groups = sorted(wtd_cur.keys(), key=lambda g: wtd_cur[g], reverse=True)
    
    for group in sorted_groups:
        val_cur = wtd_cur[group]
        val_prev = wtd_prev.get(group, 0)
        pct = val_cur / total_wtd_cur * 100 if total_wtd_cur > 0 else 0
        change = val_cur - val_prev
        icon = "▲" if change >= 0 else "▼"
        sign = "+" if change >= 0 else "-"
        
        val_triieu_str = f"{val_cur / 1e6:.1f}".replace(".", ",")
        change_triieu_str = f"{abs(change) / 1e6:.1f}".replace(".", ",")
        pct_str_val = f"{pct:.1f}".replace(".", ",") + "%"
        
        text = f"• {group} : {val_triieu_str} triệu — {pct_str_val} {icon} {sign}{change_triieu_str} triệu so tuần trước"
        vals[(row, 1)] = text
        
        reqs.append(R_merge(ws_id, row, 1, row, NCOLS))
        reqs.append(R_align(ws_id, row, 1, row, NCOLS, h="LEFT"))
        reqs.append(R_font(ws_id, row, 1, row, NCOLS, DARK, 10, False))
        reqs.append(R_row_h(ws_id, row, 20))
        row += 1
    row += 1

    # ══ SECTION 1: KH KHÔNG LÊN ĐƠN ══════════════════════════════════════════
    section_title(vals,reqs,ws_id,row,NCOLS,"🔴 KH KHÔNG LÊN ĐƠN THEO TỈNH & NHÓM",ORANGE); row+=1

    # build summary: KH có trong data ngày đó nhưng DT=0
    def no_order_summary(d):
        day = on(d)
        if day.empty: return pd.DataFrame(columns=['MaKH','TenKH','Nhom','Tinh','AM','DT'])
        kh = day.groupby(['MaKH','TenKH','Nhom','Tinh','AM'])['DT'].sum().reset_index()
        return kh[kh['DT']==0]

    no_cur = no_order_summary(d_cur)

    # sub-header: tỉnh × nhóm
    NHOM_ORDER = ['A','BCD','EF','G']
    write_hdr(vals,reqs,ws_id,row,
        ['TỈNH','A','BCD','EF','G','TỔNG','DoD vs D-1','vs D-7'], ORANGE); row+=1

    def no_by_tinh_nhom(d):
        no = no_order_summary(d)
        if no.empty: return {}
        return no.groupby(['Tinh','Nhom']).size().unstack(fill_value=0).to_dict()

    nb_cur  = no_by_tinh_nhom(d_cur)
    nb_prev = no_by_tinh_nhom(d_prev)
    nb_d7   = no_by_tinh_nhom(d7)

    def gn(nb, tinh, nhom):
        return int(nb.get(nhom,{}).get(tinh,0))

    tc_total = tp_total = t7_total = 0
    for tinh in TINH_ORDER:
        row_vals = [tinh]
        tc = 0
        for nhom in NHOM_ORDER:
            v = gn(nb_cur, tinh, nhom); tc += v; row_vals.append(v if v else '—')
        tp = sum(gn(nb_prev,tinh,n) for n in NHOM_ORDER)
        t7 = sum(gn(nb_d7,  tinh,n) for n in NHOM_ORDER)
        tc_total += tc; tp_total += tp; t7_total += t7
        row_vals += [tc, pct_str(tc,tp) if tp else '—', pct_str(tc,t7) if t7 else '—']
        write_data_row(vals,reqs,ws_id,row,row_vals,NCOLS,pct_cols=[7,8]); row+=1

    write_total(vals,reqs,ws_id,row,
        ['TỔNG NTB'] + [sum(gn(nb_cur,t,n) for t in TINH_ORDER) for n in NHOM_ORDER] +
        [tc_total, pct_str(tc_total,tp_total) if tp_total else '—',
         pct_str(tc_total,t7_total) if t7_total else '—'], NCOLS); row+=2

    # ══ SECTION 2: DANH SÁCH KH RỜI BỎ ══════════════════════════════════════
    section_title(vals,reqs,ws_id,row,NCOLS,"⚠️ KH RỜI BỎ (có đơn tuần trước, không lên đơn hôm nay)",NAVY); row+=1

    kh_d7  = set(on(d7)['MaKH'].dropna().unique())
    kh_cur = set(on(d_cur)['MaKH'].dropna().unique())
    churned_ids = kh_d7 - kh_cur

    churned_df = (df_ntb[df_ntb['MaKH'].isin(churned_ids)]
                  .drop_duplicates('MaKH')[['MaKH','TenKH','Nhom','Tinh','AM','MTD']]
                  .sort_values(['Nhom','MTD'], ascending=[True,False]))

    # chỉ hiện nhóm A và BCD (priority)
    priority = churned_df[churned_df['Nhom'].isin(['A','BCD'])]
    others   = churned_df[~churned_df['Nhom'].isin(['A','BCD'])]

    write_hdr(vals,reqs,ws_id,row,
        ['Nhóm','Mã KH','Tên KH','Tỉnh','AM','MTD (đơn)','Ghi chú'], NAVY); row+=1

    for _, r in priority.iterrows():
        write_data_row(vals,reqs,ws_id,row,
            [r['Nhom'], str(int(r['MaKH'])), r['TenKH'],
             r['Tinh'], r['AM'] if pd.notna(r['AM']) else '—',
             _safe(r['MTD']), ''], 7, pct_cols=[])
        reqs.append(R_align(ws_id,row,1,row,7,"LEFT"))
        row+=1

    # Summary EF
    put(row,1, f"Nhóm EF & G: {len(others)} KH – xem danh sách đầy đủ trong sheet Nhóm gốc")
    reqs += [R_merge(ws_id,row,1,row,7), R_bg(ws_id,row,1,row,7,YELLOW),
             R_font(ws_id,row,1,row,7,DARK,9,False),
             R_align(ws_id,row,1,row,7,"LEFT")]
    row += 2

    # ══ SECTION 3: DT KHM DEEP DIVE ══════════════════════════════════════════
    section_title(vals,reqs,ws_id,row,NCOLS,"🆕 DT KHÁCH HÀNG MỚI (KHM) DEEP DIVE",GREEN); row+=1

    khm_cur  = D["khm_cur"];  khm_prev = D["khm_prev"]
    khm_mtd  = D["khm_mtd_cur"]; khm_mtd_p = D["khm_mtd_prev"]

    def sums(df2):
        if df2 is None or df2.empty: return 0,0,0
        return int(df2["SLKH"].sum()), int(df2["Vol"].sum()), round(df2["DT"].sum()/1e6,1)

    sc,vc,dc   = sums(khm_cur)
    sp,vp,dp   = sums(khm_prev)
    sm,vm,dm   = sums(khm_mtd)
    smp,vmp,dmp= sums(khm_mtd_p)

    write_hdr(vals,reqs,ws_id,row,
        ['CHỈ SỐ', f"Ngày {d_cur.strftime('%d/%m')}",
         f"Ngày {d_prev.strftime('%d/%m')}", 'DoD (%)',
         f"MTD T{m_start.month} (1–{d_cur.day})",
         f"MTD T{prev_m_start.month} cùng kỳ", 'MoM (%)',''],
         GREEN); row+=1

    for label, vc2, vp2 in [
        (f'Số KH mới (shop)',  sc, sp),
        (f'Sản lượng (đơn)',   vc, vp),
        (f'DT (triệu đồng)',   dc, dp),
    ]:
        mtd_c = sm if label.startswith('Số') else (vm if label.startswith('Sản') else dm)
        mtd_p = smp if label.startswith('Số') else (vmp if label.startswith('Sản') else dmp)
        write_data_row(vals,reqs,ws_id,row,
            [label, vc2, vp2, pct_str(vc2,vp2), mtd_c, mtd_p, pct_str(mtd_c,mtd_p),''],
            NCOLS, pct_cols=[4,7]); row+=1

    row+=1
    # KHM by tỉnh ngày cur
    write_hdr(vals,reqs,ws_id,row,
        ['TỈNH','#KH mới','Vol','DT (M)',
         f"MTD T{m_start.month} #KH",f"MTD T{m_start.month} DT(M)",
         f"MTD T{prev_m_start.month} #KH",'MoM DT (%)'],
         GREEN); row+=1

    def gv(df2, tinh, col):
        return df2.loc[tinh,col] if (df2 is not None and tinh in df2.index and col in df2.columns) else 0

    for tinh in TINH_ORDER:
        write_data_row(vals,reqs,ws_id,row,[
            tinh,
            _safe(gv(khm_cur, tinh,'SLKH')),
            _safe(gv(khm_cur, tinh,'Vol')),
            round(gv(khm_cur, tinh,'DT')/1e6,1),
            _safe(gv(khm_mtd, tinh,'SLKH')),
            round(gv(khm_mtd, tinh,'DT')/1e6,1),
            _safe(gv(khm_mtd_p,tinh,'SLKH')),
            pct_str(gv(khm_mtd,tinh,'DT'), gv(khm_mtd_p,tinh,'DT')),
        ], NCOLS, pct_cols=[8]); row+=1

    flush(ws, vals, row); ss.batch_update({"requests": reqs})
    print(f"  ✅ {OUT_NHOM}")
    return (df_ntb, no_cur, priority)


# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY PRINT
# ─────────────────────────────────────────────────────────────────────────────
def print_full_report(D, nhom_data=None):
    """In ra đoạn báo cáo text hoàn chỉnh để copy lên group."""
    d  = D["d_cur"]; dp = D["d_prev"]; d7 = D["d7"]
    m  = D["m_start"]; pm = D["prev_m_start"]; pe = D["prev_m_end"]
    ys = D["yoy_start"]; ye = D["yoy_end"]

    # ── DT tổng vùng ──
    tc_dt = D["t_cur"]["DT"].sum()  if "DT"  in D["t_cur"].columns  else 0
    tp_dt = D["t_prev"]["DT"].sum() if "DT"  in D["t_prev"].columns else 0
    t7_dt = D["t_d7"]["DT"].sum()   if "DT"  in D["t_d7"].columns   else 0
    tc_vl = D["t_cur"]["Vol"].sum() if "Vol" in D["t_cur"].columns  else 0
    tp_vl = D["t_prev"]["Vol"].sum()if "Vol" in D["t_prev"].columns else 0

    mc   = D["mtd_cur"]["DT"].sum()   if "DT" in D["mtd_cur"].columns   else 0
    mp   = D["mtd_prev"]["DT"].sum()  if "DT" in D["mtd_prev"].columns  else 0
    myoy = D["mtd_yoy"]["DT"].sum()   if "DT" in D["mtd_yoy"].columns   else 0

    # ── KHM ──
    def sums_khm(df2):
        if df2 is None or df2.empty: return 0, 0, 0
        return int(df2["SLKH"].sum()), int(df2["Vol"].sum()), df2["DT"].sum()
    sc, svc, sdt   = sums_khm(D["khm_cur"])
    sp, svp, sdp   = sums_khm(D["khm_prev"])
    s7, sv7, sd7   = sums_khm(D.get("khm_d7", D["khm_prev"]))  # fallback
    sm, svm, sdm   = sums_khm(D["khm_mtd_cur"])
    smp,svmp,sdmp  = sums_khm(D["khm_mtd_prev"])

    # ── Nhóm KH ──
    no_total = no_a = no_bcd = no_ef = no_g = 0
    churned_lines = []
    if nhom_data is not None:
        df_ntb_n, no_cur, churned_priority = nhom_data
        no_total = len(no_cur)
        by_nhom = no_cur.groupby("Nhom").size().to_dict()
        no_a   = by_nhom.get("A",   0)
        no_bcd = by_nhom.get("BCD", 0)
        no_ef  = by_nhom.get("EF",  0)
        no_g   = by_nhom.get("G",   0)
        for _, r in churned_priority.head(5).iterrows():
            mtd_val = int(r["MTD"]) if pd.notna(r["MTD"]) else 0
            churned_lines.append(
                f"     • {r['TenKH']} ({r['Nhom']}) – AM: {r['AM']} – MTD: {mtd_val:,} đơn"
            )

    def M(v):  return f"{v/1e6:.1f}".replace(".", ",") + " triệu"
    def B(v):
        ty = int(v // 1e9)
        trieu = int(round((v % 1e9) / 1e6))
        if ty == 0:
            return f"{trieu} triệu"
        if trieu == 0:
            return f"{ty} tỷ"
        return f"{ty} tỷ {trieu} triệu"
    def N(v):  return f"{int(v):,}".replace(",", ".")
    def F(v):  return f"{int(round(v)):,}".replace(",", ".") + "đ"
    def P(c,p):
        if not p: return "—"
        x = (c-p)/p*100
        return f"{'▲' if x>=0 else '▼'} {abs(x):.1f}".replace(".", ",") + "%"

    tomorrow = d + timedelta(days=1)

    lines = []
    lines.append(f"VÙNG NTB – BÁO CÁO KINH DOANH NGÀY {tomorrow.strftime('%d/%m/%Y')}")
    lines.append("")
    lines.append(f"1. Doanh thu toàn vùng ngày {d.strftime('%d/%m')}: {M(tc_dt)}")
    lines.append(f"   + {P(tc_dt,tp_dt)} so với ngày N-1 ({dp.strftime('%d/%m')}: {M(tp_dt)})")
    lines.append(f"   + {P(tc_dt,t7_dt)} so với ngày cùng kỳ tuần trước ({d7.strftime('%d/%m')}: {M(t7_dt)})")
    lines.append(f"   - Sản lượng: {N(tc_vl)} đơn ({P(tc_vl,tp_vl)} DoD)")
    lines.append(f"   - DT luỹ kế tháng {m.month}: {B(mc)} (01–{d.day}/{m.month})")
    lines.append(f"     + {P(mc,mp)} so với cùng kỳ tháng {pm.month} ({B(mp)})")
    lines.append(f"     + {P(mc,myoy)} so với cùng kỳ tháng {ys.month}/{ys.year} ({B(myoy)})")
    lines.append("")
    
    # ── KHM ──
    lines.append(f"2. Doanh thu Khách hàng mới (KHM) ngày {d.strftime('%d/%m')}: {F(sdt)} ({sc} shop)")
    lines.append(f"   + Về shop mới: {P(sc,sp)} so với ngày N-1 ({sp} shop)")
    lines.append(f"   + Về doanh thu: {P(sdt,sdp)} so với ngày N-1 ({F(sdp)})")
    lines.append(f"   + So với cùng kỳ tuần trước: Doanh thu {P(sdt,sd7)} ({F(sd7)})")
    lines.append(f"   - KHM luỹ kế tháng {m.month}: {sm} shop – {M(sdm)}")
    lines.append(f"     + Về shop mới: {P(sm,smp)} so với cùng kỳ tháng {pm.month} ({smp} shop)")
    lines.append(f"     + Về doanh thu: {P(sdm,sdmp)} so với cùng kỳ tháng {pm.month} ({M(sdmp)})")
    lines.append("")

    if nhom_data is not None:
        lines.append(f"3. Phân nhóm KH:")
        
        # WTD Revenue summary lines
        wtd_cur = D.get("wtd_cur_by_group", {})
        wtd_prev = D.get("wtd_prev_by_group", {})
        total_wtd_cur = sum(wtd_cur.values())
        
        # Sort groups descending by current WTD revenue
        sorted_groups = sorted(wtd_cur.keys(), key=lambda g: wtd_cur[g], reverse=True)
        for group in sorted_groups:
            val_cur = wtd_cur[group]
            val_prev = wtd_prev.get(group, 0)
            pct = val_cur / total_wtd_cur * 100 if total_wtd_cur > 0 else 0
            change = val_cur - val_prev
            icon = "▲" if change >= 0 else "▼"
            sign = "+" if change >= 0 else "-"
            
            val_triieu_str = f"{val_cur / 1e6:.1f}".replace(".", ",")
            change_triieu_str = f"{abs(change) / 1e6:.1f}".replace(".", ",")
            pct_str_val = f"{pct:.1f}".replace(".", ",") + "%"
            
            lines.append(f"   • {group} : {val_triieu_str} triệu — {pct_str_val} {icon} {sign}{change_triieu_str} triệu so tuần trước")
            
        lines.append(f"   - KH không lên đơn ngày {d.strftime('%d/%m')}: {N(no_total)} KH")
        lines.append(f"     + Trong đó: {no_a} KH A | {no_bcd} KH BCD | {no_ef} KH EF | {no_g} KH G")
        lines.append(f"   - KH rời bỏ (có đơn tuần trước, không lên đơn hôm nay): {len(churned_lines)} KH A/BCD ưu tiên")
        for cl in churned_lines:
            lines.append(cl)
        lines.append("   → AM đang theo sát, xác định nguyên nhân và hướng giữ chân KH")
        lines.append("")

    if d.day >= 25:
        lines.append("4. Kế hoạch hành động nước rút cuối tháng")
        lines.append("   - Vận hành tập trung nhân sự thu gom, lấy sạch đơn phát sinh trong ngày tại các bưu cục để tối đa doanh thu lấy thành công (LTC) ghi nhận trong tháng")
        lines.append("   - AM trực tiếp liên hệ chốt sản lượng cuối tháng với các shop lớn A/BCD, xử lý nhanh khiếu nại phát sinh trong 48h")
        lines.append("   - Rà soát danh sách KH rời bỏ (đặc biệt nhóm A/BCD ưu tiên trong sheet RPT_NhomKH) để liên hệ trực tiếp tháo gỡ vướng mắc")
        next_month = m.month + 1 if m.month < 12 else 1
        lines.append(f"   - Lên sẵn danh sách và chính sách giá cước đặc thù cho các KHM lớn của tháng {m.month} để chạy đà sản lượng cho tháng {next_month}")
    else:
        lines.append("4. Kế hoạch hành động")
        lines.append("   - Tập trung thương thảo chính sách cước/chiết khấu thương mại linh hoạt theo sản lượng để giữ chân nhóm KH lớn A/BCD")
        lines.append("   - Kích cầu nhóm KH vừa và nhỏ (EF/G) thông qua các chương trình đồng giá và ưu đãi cước ngắn hạn")
        lines.append("   - Kiểm soát chặt chẽ chất lượng vận hành (giao nhận đúng hạn) tại các địa bàn bưu cục trọng điểm để giảm rớt đơn")
        lines.append("   - Thúc đẩy chỉ tiêu phát triển KH mới (KHM) để bù đắp khoảng trống lũy kế (MTD)")
    lines.append("")
    lines.append(f"📊 Chi tiết: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")

    report = "\n".join(lines)
    print("\n" + "="*65)
    print("📋 BÁO CÁO GỬI GROUP (copy đoạn dưới):")
    print("="*65)
    print(report)
    print("="*65 + "\n")
    return report


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        ss = connect()
        print("Connected to Google Sheets successfully.")
        D = load_data(ss)
        
        print("Building sheets...")
        build_daily(ss, D)
        build_weekly(ss, D)
        build_mtd(ss, D)
        build_khm(ss, D)
        nhom_data = build_nhom(ss, D)
        
        print_full_report(D, nhom_data)
        print("All reports generated and sheets updated successfully!")
    except Exception as e:
        print(f"Error in main execution: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
