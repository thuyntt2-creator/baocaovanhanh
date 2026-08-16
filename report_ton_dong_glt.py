"""
report_ton_dong_glt.py
======================
- Đọc file Excel tồn đọng GLT (Tồn đọng GLT.xlsx)
- Tách 3 nhóm: Lấy / Giao (gồm Ưu tiên giao) / Trả
- Render HTML giao diện cực kỳ hiện đại, đẹp chuẩn Premium Dashboard
- Chụp ảnh PNG sắc nét bằng Playwright
- Gửi lên Telegram
"""

import os
import datetime
import pandas as pd
import requests
from playwright.sync_api import sync_playwright

# ── Cấu hình Telegram ─────────────────────────────────────────────────────────
TELEGRAM_BOT_TOKEN = "8570130113:AAGXRiUaKBknVpgtm1_i9ZA47JRjAXmB21M"
TELEGRAM_CHAT_ID   = "-5058464865"

GLT_FILE   = r"C:\Users\lap4all\Downloads\Tồn đọng GLT.xlsx"
OUTPUT_DIR = r"C:\Users\lap4all\Documents\Auto report"

# ── Mapping loại đơn ──────────────────────────────────────────────────────────
LOAI_MAP = {
    "Giao"          : "Giao",
    "Ưu tiên giao"  : "Giao",
    "Lấy"           : "Lấy",
    "Trả"           : "Trả",
}

LOAI_CONFIG = [
    {
        "key": "Lấy",
        "title": "BÁO CÁO TỒN ĐỌNG LẤY HÀNG GLT",
        "label": "📥 LẤY HÀNG",
        "header_bg": "linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%)",
        "primary_color": "#2563eb",
        "badge_bg": "#dbeafe",
        "badge_color": "#1e40af",
    },
    {
        "key": "Giao",
        "title": "BÁO CÁO TỒN ĐỌNG GIAO HÀNG GLT",
        "label": "📦 GIAO HÀNG",
        "header_bg": "linear-gradient(135deg, #064e3b 0%, #10b981 100%)",
        "primary_color": "#059669",
        "badge_bg": "#d1fae5",
        "badge_color": "#065f46",
    },
    {
        "key": "Trả",
        "title": "BÁO CÁO TỒN ĐỌNG TRẢ HÀNG GLT",
        "label": "↩️ TRẢ HÀNG",
        "header_bg": "linear-gradient(135deg, #881337 0%, #f43f5e 100%)",
        "primary_color": "#e11d48",
        "badge_bg": "#ffe4e6",
        "badge_color": "#9f1239",
    },
]

# ── Bracket ───────────────────────────────────────────────────────────────────
BRACKET_ORDER = ["0_6","6_12","12_24","24_36","36_48","48_72","72_96","96_120","120_192","192"]
BRACKET_LABEL = {
    "0_6"    : "0-6h",
    "6_12"   : "6-12h",
    "12_24"  : "12-24h",
    "24_36"  : "1-1.5N",
    "36_48"  : "1.5-2N",
    "48_72"  : "2-3N",
    "72_96"  : "3-4N",
    "96_120" : "4-5N",
    "120_192": "5-8N",
    "192"    : ">8N",
}

# Style cho từng mốc aging
BRACKET_CELL_STYLE = {
    "0_6"    : "background: #f0fdf4; color: #166534;",
    "6_12"   : "background: #ecfdf5; color: #15803d; font-weight: 600;",
    "12_24"  : "background: #fefce8; color: #854d0e; font-weight: 600;",
    "24_36"  : "background: #fef9c3; color: #a16207; font-weight: 600;",
    "36_48"  : "background: #fef3c7; color: #b45309; font-weight: 700;",
    "48_72"  : "background: #ffedd5; color: #c2410c; font-weight: 700;",
    "72_96"  : "background: #fed7aa; color: #9a3412; font-weight: 700;",
    "96_120" : "background: #fecaca; color: #991b1b; font-weight: 700;",
    "120_192": "background: #fca5a5; color: #7f1d1d; font-weight: 800;",
    "192"    : "background: #991b1b; color: #ffffff; font-weight: 800;",
}

# ── Đọc dữ liệu ───────────────────────────────────────────────────────────────
def load_glt_data(path):
    print(f"[INFO] Đọc file Excel: {path}")
    df = pd.read_excel(path, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    df["AM"]                  = df["AM"].fillna("Chưa phân AM")
    df["Thời gian tồn đọng"]  = df["Thời gian tồn đọng"].fillna("Không rõ")
    df["Loại đơn"]            = df["Loại đơn"].fillna("")
    df["Nhóm"]                = df["Loại đơn"].map(LOAI_MAP).fillna("Khác")
    return df

# ── Build Pivot ───────────────────────────────────────────────────────────────
def build_pivot(df_sub):
    if df_sub.empty:
        return None, []
    pivot = (
        df_sub.groupby(["AM", "Thời gian tồn đọng"])
              .size()
              .reset_index(name="SL")
              .pivot(index="AM", columns="Thời gian tồn đọng", values="SL")
              .fillna(0)
              .astype(int)
    )
    cols_present = [b for b in BRACKET_ORDER if b in pivot.columns]
    pivot = pivot[cols_present]
    pivot["Tổng"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("Tổng", ascending=False)
    return pivot, cols_present

# ── Render HTML & Capture với Playwright ──────────────────────────────────────
def render_html_table(pivot, cols_present, config, date_str, output_png_path):
    total_all = int(pivot["Tổng"].sum())
    
    # Header cols HTML
    header_cols_html = ""
    for c in cols_present:
        header_cols_html += f"<th>{BRACKET_LABEL.get(c, c)}</th>"
    header_cols_html += "<th class='total-col'>TỔNG</th>"

    # Data Rows HTML
    rows_html = ""
    for idx, (am, row) in enumerate(pivot.iterrows(), 1):
        row_bg = "#ffffff" if idx % 2 != 0 else "#f8fafc"
        cells_html = f"<td class='stt-col'>{idx}</td><td class='am-col'>{am}</td>"
        
        for c in cols_present:
            val = int(row[c])
            if val > 0:
                style = BRACKET_CELL_STYLE.get(c, "")
                cells_html += f"<td class='val-cell' style='{style}'>{val:,}</td>"
            else:
                cells_html += "<td class='val-cell empty-cell'>-</td>"
                
        tot_val = int(row["Tổng"])
        cells_html += f"<td class='total-cell'>{tot_val:,}</td>"
        rows_html += f"<tr style='background: {row_bg};'>{cells_html}</tr>"

    # Summary Row HTML
    sum_cells = "<td colspan='2' class='sum-label'>TỔNG TOÀN VÙNG</td>"
    for c in cols_present:
        col_sum = int(pivot[c].sum())
        sum_cells += f"<td class='sum-val'>{col_sum:,}</td>" if col_sum > 0 else "<td class='sum-val empty-cell'>-</td>"
    sum_cells += f"<td class='sum-total'>{total_all:,}</td>"
    summary_row_html = f"<tr class='summary-row'>{sum_cells}</tr>"

    html_content = f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        * {{ box-sizing: border-box; font-family: 'Plus Jakarta Sans', 'Inter', sans-serif; }}
        body {{
            margin: 0;
            padding: 24px;
            background-color: #f1f5f9;
            display: inline-block;
        }}
        .card {{
            background: #ffffff;
            border-radius: 16px;
            box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.08), 0 8px 10px -6px rgba(15, 23, 42, 0.04);
            overflow: hidden;
            border: 1px solid #e2e8f0;
        }}
        .header {{
            background: {config['header_bg']};
            color: #ffffff;
            padding: 22px 28px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .header-title {{
            font-size: 20px;
            font-weight: 800;
            letter-spacing: 0.5px;
            text-transform: uppercase;
            margin: 0;
        }}
        .header-sub {{
            font-size: 13px;
            opacity: 0.9;
            margin-top: 4px;
            font-weight: 500;
        }}
        .total-badge {{
            background: rgba(255, 255, 255, 0.2);
            backdrop-filter: blur(8px);
            border: 1px solid rgba(255, 255, 255, 0.3);
            padding: 8px 18px;
            border-radius: 99px;
            font-size: 15px;
            font-weight: 800;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            text-align: center;
        }}
        th {{
            background: #0f172a;
            color: #f8fafc;
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 12px 10px;
            border-bottom: 2px solid #334155;
        }}
        th.stt-col {{ width: 45px; }}
        th.am-header {{ text-align: left; padding-left: 16px; min-width: 190px; }}
        th.total-col {{ background: #1e293b; color: #38bdf8; min-width: 90px; }}
        
        td {{
            padding: 10px;
            font-size: 13px;
            border-bottom: 1px solid #e2e8f0;
        }}
        td.stt-col {{ color: #94a3b8; font-size: 12px; font-weight: 600; text-align: center; }}
        td.am-col {{
            text-align: left;
            padding-left: 16px;
            font-weight: 700;
            color: #1e293b;
        }}
        td.val-cell {{
            font-family: 'Inter', sans-serif;
            border-radius: 4px;
        }}
        td.empty-cell {{
            color: #cbd5e1;
        }}
        td.total-cell {{
            font-family: 'Inter', sans-serif;
            font-weight: 800;
            color: #0f172a;
            background: #f1f5f9;
        }}
        
        tr.summary-row {{
            background: #0f172a !important;
            color: #ffffff;
        }}
        tr.summary-row td {{
            border-top: 2px solid #334155;
            padding: 14px 10px;
        }}
        td.sum-label {{
            text-align: left;
            padding-left: 16px;
            font-weight: 800;
            font-size: 13px;
            letter-spacing: 0.5px;
            color: #f8fafc;
        }}
        td.sum-val {{
            font-family: 'Inter', sans-serif;
            font-weight: 700;
            color: #fbbf24;
        }}
        td.sum-total {{
            font-family: 'Inter', sans-serif;
            font-weight: 900;
            color: #38bdf8;
            font-size: 15px;
            background: #1e293b;
        }}
        .footer {{
            padding: 12px 24px;
            background: #f8fafc;
            border-top: 1px solid #e2e8f0;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 12px;
            color: #64748b;
        }}
        .legend {{
            display: flex;
            gap: 8px;
            align-items: center;
        }}
        .legend-item {{
            display: flex;
            align-items: center;
            gap: 4px;
            font-size: 11px;
        }}
        .legend-box {{
            width: 12px;
            height: 12px;
            border-radius: 3px;
        }}
    </style>
</head>
<body>
    <div class="card" id="capture-container">
        <div class="header">
            <div>
                <div class="header-title">{config['title']}</div>
                <div class="header-sub">Báo cáo tự động NTB Region • Cập nhật lúc: {date_str}</div>
            </div>
            <div class="total-badge">
                TỔNG: {total_all:,} ĐƠN
            </div>
        </div>
        <table>
            <thead>
                <tr>
                    <th class="stt-col">#</th>
                    <th class="am-header">AM PHỤ TRÁCH</th>
                    {header_cols_html}
                </tr>
            </thead>
            <tbody>
                {rows_html}
                {summary_row_html}
            </tbody>
        </table>
        <div class="footer">
            <span>🔥 Hệ thống giám sát vận hành GLT</span>
            <div class="legend">
                <div class="legend-item"><div class="legend-box" style="background:#f0fdf4; border:1px solid #bbf7d0;"></div>0-6h</div>
                <div class="legend-item"><div class="legend-box" style="background:#fef3c7; border:1px solid #fde68a;"></div>1.5-2N</div>
                <div class="legend-item"><div class="legend-box" style="background:#ffedd5; border:1px solid #fed7aa;"></div>2-3N</div>
                <div class="legend-item"><div class="legend-box" style="background:#fecaca; border:1px solid #fca5a5;"></div>4-5N</div>
                <div class="legend-item"><div class="legend-box" style="background:#991b1b;"></div>>8N</div>
            </div>
        </div>
    </div>
</body>
</html>
"""

    temp_html_path = os.path.join(OUTPUT_DIR, f"temp_{config['key'].lower()}.html")
    with open(temp_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    # Chụp ảnh bằng Playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        page.goto(f"file:///{temp_html_path.replace('\\', '/')}")
        page.wait_for_timeout(500)
        card = page.locator("#capture-container")
        card.screenshot(path=output_png_path)
        browser.close()

    try:
        os.remove(temp_html_path)
    except:
        pass

    print(f"[OK] Đã render ảnh Premium: {output_png_path}")
    return total_all

# ── Gửi Telegram ──────────────────────────────────────────────────────────────
def send_telegram_photo(token, chat_id, image_path, caption=""):
    url = f"https://api.telegram.org/bot{token}/sendPhoto"
    with open(image_path, "rb") as f:
        resp = requests.post(
            url,
            data={"chat_id": chat_id, "caption": caption, "parse_mode": "HTML"},
            files={"photo": f}
        )
    if resp.status_code == 200:
        print("[OK] Gửi Telegram thành công!")
    else:
        print(f"[ERR] Telegram {resp.status_code}: {resp.text[:200]}")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    now = datetime.datetime.now()
    date_str = now.strftime("%d/%m/%Y %H:%M")

    df = load_glt_data(GLT_FILE)
    total_all = len(df)
    print(f"[INFO] Tổng đơn tồn đọng toàn vùng: {total_all:,}")

    # Gửi tin nhắn mở đầu
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        header_msg = (
            f"📊 <b>BÁO CÁO TỒN ĐỌNG GLT - NTB REGION</b>\n"
            f"⏱️ <i>Cập nhật: {date_str}</i>\n"
            f"📦 Tổng đơn tồn đọng: <b>{total_all:,}</b> đơn\n\n"
            f" Chi tiết theo từng loại đơn hàng được cập nhật bên dưới 👇"
        )
        requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
            data={"chat_id": TELEGRAM_CHAT_ID, "text": header_msg, "parse_mode": "HTML"}
        )

    # Render & Gửi từng nhóm
    for cfg in LOAI_CONFIG:
        nhom_key = cfg["key"]
        df_sub   = df[df["Nhóm"] == nhom_key].copy()
        count    = len(df_sub)
        print(f"\n[INFO] {cfg['title']}: {count:,} đơn")

        if count == 0:
            print(f"[SKIP] Không có đơn {nhom_key}")
            continue

        pivot, cols_present = build_pivot(df_sub)
        if pivot is None:
            continue

        out_png = os.path.join(OUTPUT_DIR, f"table_glt_{nhom_key.lower()}.png")
        total   = render_html_table(pivot, cols_present, cfg, date_str, out_png)

        # Top 5 AM nhiều tồn nhất
        top5 = pivot.head(5)
        caption_lines = [
            f"<b>{cfg['label']} – {date_str}</b>",
            f"Tổng tồn đọng: <b>{total:,} đơn</b>",
            "",
            "🚨 <b>TOP AM TỒN ĐỌNG NHIỀU NHẤT:</b>",
        ]
        for am, row in top5.iterrows():
            caption_lines.append(f"  • <b>{am}</b>: {int(row['Tổng']):,} đơn")
            
        caption = "\n".join(caption_lines)

        if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
            send_telegram_photo(TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, out_png, caption)

    print("\n[DONE] Đã render & gửi báo cáo xong xuôi!")

if __name__ == "__main__":
    main()
