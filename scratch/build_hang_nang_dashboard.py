import os
import sys
import io
import json
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

os.environ['PYTHONIOENCODING'] = 'utf-8'
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(BASE_DIR)
JSON_FILE = os.path.join(PARENT_DIR, 'credentials.json')
SHEET_KEY = '1JZ1eRerRqrpwjZ4HBevQunjd8VquM_cvPFz12TaJfMQ'

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def main():
    print("🔄 Connecting to Google Sheets...")
    creds = Credentials.from_service_account_file(JSON_FILE, scopes=scopes)
    gc_client = gspread.authorize(creds)
    sh = gc_client.open_by_key(SHEET_KEY)
    
    # 1. Fetch trên10kg
    print("📥 Fetching sheet 'trên10kg'...")
    ws1 = sh.worksheet('trên10kg')
    data1 = ws1.get_all_records()
    df1 = pd.DataFrame(data1)
    df1['so_don'] = pd.to_numeric(df1['so_don'], errors='coerce').fillna(0).astype(int)
    
    # Clean records
    records_created = []
    for r in df1.to_dict(orient='records'):
        if r.get('hen_lay'):
            records_created.append({
                'date': str(r.get('hen_lay')),
                'province': str(r.get('province_name', '')),
                'bc': str(r.get('warehouse_name', '')),
                'kh': str(r.get('nhom_kh', '')),
                'kl': str(r.get('nhom_kl', '')),
                'so_don': int(r.get('so_don', 0))
            })
            
    # 2. Fetch SL > 10kg
    print("📥 Fetching sheet 'SL > 10kg'...")
    ws2 = sh.worksheet('SL > 10kg')
    data2 = ws2.get_all_records()
    df2 = pd.DataFrame(data2)
    
    records_ops = []
    for r in df2.to_dict(orient='records'):
        time_str = str(r.get('Time', ''))
        # Extract YYYY-MM-DD
        import re
        m = re.search(r'(\d{4}-\d{2}-\d{2})', time_str)
        date_val = m.group(1) if m else ''
        if date_val:
            vol = str(r.get('Volume', '')).replace(',', '')
            gtc = str(r.get('Sản Lượng Giao Thành Công', '')).replace(',', '')
            ton = str(r.get('Sản Lượng Tồn', '')).replace(',', '')
            records_ops.append({
                'date': date_val,
                'bc': str(r.get('Chi tiết', '')),
                'province': str(r.get('Tỉnh', '')),
                'am': str(r.get('AM', '')),
                'loai_hang': str(r.get('Loại Hàng', '')),
                'loai_kl': str(r.get('Loại Khối Lượng', '')),
                'volume': int(float(vol)) if vol and vol != '-' else 0,
                'gtc': int(float(gtc)) if gtc and gtc != '-' else 0,
                'ton': int(float(ton)) if ton and ton != '-' else 0
            })
            
    print(f"✅ Loaded {len(records_created)} created records and {len(records_ops)} operational records.")
    
    # Build HTML Dashboard
    html_content = generate_dashboard_html(records_created, records_ops)
    
    out_dir = os.path.join(PARENT_DIR, 'output')
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, 'ntb_hang_nang_dashboard.html')
    
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print(f"🎉 Generated Dashboard HTML: {out_file}")

def generate_dashboard_html(records_created, records_ops):
    json_created = json.dumps(records_created, ensure_ascii=False)
    json_ops = json.dumps(records_ops, ensure_ascii=False)
    
    return f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>GHN Ops - Báo Cáo Hàng Nặng > 10kg</title>
    
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@500;600;700;800&display=swap" rel="stylesheet">
    
    <!-- FontAwesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    
    <!-- ApexCharts -->
    <script src="https://cdn.jsdelivr.net/npm/apexcharts"></script>

    <style>
        :root {{
            --primary: #ff5f00;
            --primary-gradient: linear-gradient(135deg, #ff5f00, #ff8533);
            --blue: #007bc3;
            --blue-gradient: linear-gradient(135deg, #007bc3, #299ce6);
            --bg-page: #f8fafc;
            --bg-card: #ffffff;
            --border: #e2e8f0;
            --text-main: #0f172a;
            --text-sub: #475569;
            --text-muted: #94a3b8;
            --success: #10b981;
            --danger: #ef4444;
            --warning: #f59e0b;
            --shadow-card: 0 4px 20px -2px rgba(15, 23, 42, 0.06);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            font-family: 'Inter', sans-serif;
            background-color: var(--bg-page);
            color: var(--text-main);
            padding: 24px;
            min-height: 100vh;
        }}

        h1, h2, h3, h4 {{
            font-family: 'Outfit', sans-serif;
        }}

        .dashboard-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            background: var(--bg-card);
            padding: 20px 28px;
            border-radius: 16px;
            border: 1px solid var(--border);
            box-shadow: var(--shadow-card);
        }}

        .brand-section {{
            display: flex;
            align-items: center;
            gap: 16px;
        }}

        .brand-icon {{
            width: 48px;
            height: 48px;
            background: var(--primary-gradient);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 24px;
            box-shadow: 0 6px 16px rgba(255, 95, 0, 0.3);
        }}

        .brand-title h1 {{
            font-size: 22px;
            font-weight: 700;
            color: var(--text-main);
        }}

        .brand-title p {{
            font-size: 13px;
            color: var(--text-sub);
            margin-top: 2px;
        }}

        .data-source-badge {{
            display: flex;
            align-items: center;
            gap: 8px;
            background: rgba(0, 123, 195, 0.08);
            color: var(--blue);
            padding: 8px 14px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            border: 1px solid rgba(0, 123, 195, 0.2);
        }}

        /* FILTER PANEL */
        .filter-panel {{
            background: var(--bg-card);
            border-radius: 16px;
            border: 1px solid var(--border);
            padding: 20px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-card);
        }}

        .filter-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 16px;
            padding-bottom: 12px;
            border-bottom: 1px dashed var(--border);
        }}

        .filter-header h3 {{
            font-size: 15px;
            color: var(--text-main);
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .date-presets {{
            display: flex;
            gap: 8px;
        }}

        .btn-preset {{
            background: #f1f5f9;
            border: 1px solid var(--border);
            color: var(--text-sub);
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }}

        .btn-preset:hover, .btn-preset.active {{
            background: var(--primary);
            color: white;
            border-color: var(--primary);
            box-shadow: 0 4px 10px rgba(255, 95, 0, 0.2);
        }}

        .filter-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
        }}

        .filter-item {{
            display: flex;
            flex-direction: column;
            gap: 6px;
        }}

        .filter-item label {{
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-sub);
        }}

        .filter-item select, .filter-item input {{
            width: 100%;
            padding: 10px 14px;
            border-radius: 10px;
            border: 1px solid var(--border);
            background-color: #ffffff;
            font-size: 13px;
            color: var(--text-main);
            outline: none;
            transition: border-color 0.2s;
        }}

        .filter-item select:focus, .filter-item input:focus {{
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(255, 95, 0, 0.15);
        }}

        /* KPI GRID */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
        }}

        .kpi-card {{
            background: var(--bg-card);
            border-radius: 16px;
            border: 1px solid var(--border);
            padding: 20px;
            position: relative;
            overflow: hidden;
            box-shadow: var(--shadow-card);
            transition: transform 0.2s;
        }}

        .kpi-card:hover {{
            transform: translateY(-3px);
        }}

        .kpi-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 4px;
            height: 100%;
            background: var(--primary);
        }}

        .kpi-card.blue::before {{ background: var(--blue); }}
        .kpi-card.green::before {{ background: var(--success); }}
        .kpi-card.purple::before {{ background: #8b5cf6; }}

        .kpi-top {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
        }}

        .kpi-title {{
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: var(--text-sub);
        }}

        .kpi-icon {{
            width: 36px;
            height: 36px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
            background: rgba(255, 95, 0, 0.1);
            color: var(--primary);
        }}

        .kpi-card.blue .kpi-icon {{ background: rgba(0, 123, 195, 0.1); color: var(--blue); }}
        .kpi-card.green .kpi-icon {{ background: rgba(16, 185, 129, 0.1); color: var(--success); }}
        .kpi-card.purple .kpi-icon {{ background: rgba(139, 92, 246, 0.1); color: #8b5cf6; }}

        .kpi-value {{
            font-size: 28px;
            font-weight: 800;
            color: var(--text-main);
            line-height: 1.2;
        }}

        .kpi-subtext {{
            font-size: 12px;
            margin-top: 8px;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .badge-spike {{
            display: inline-flex;
            align-items: center;
            gap: 4px;
            background: rgba(239, 68, 68, 0.1);
            color: var(--danger);
            padding: 3px 8px;
            border-radius: 12px;
            font-weight: 700;
            font-size: 11px;
        }}

        .badge-normal {{
            display: inline-flex;
            align-items: center;
            gap: 4px;
            background: rgba(16, 185, 129, 0.1);
            color: var(--success);
            padding: 3px 8px;
            border-radius: 12px;
            font-weight: 600;
            font-size: 11px;
        }}

        /* CHART & TABLE LAYOUT */
        .chart-container {{
            background: var(--bg-card);
            border-radius: 16px;
            border: 1px solid var(--border);
            padding: 24px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-card);
        }}

        .chart-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
        }}

        .chart-header h3 {{
            font-size: 16px;
            color: var(--text-main);
        }}

        .table-card {{
            background: var(--bg-card);
            border-radius: 16px;
            border: 1px solid var(--border);
            padding: 24px;
            box-shadow: var(--shadow-card);
        }}

        .table-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
        }}

        .table-header h3 {{
            font-size: 16px;
            color: var(--text-main);
        }}

        .custom-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}

        .custom-table th {{
            background: #f8fafc;
            color: var(--text-sub);
            font-weight: 700;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 12px 16px;
            border-bottom: 1px solid var(--border);
            text-align: left;
            cursor: pointer;
        }}

        .custom-table th:hover {{
            color: var(--primary);
        }}

        .custom-table td {{
            padding: 12px 16px;
            border-bottom: 1px solid var(--border);
            color: var(--text-main);
        }}

        .custom-table tr:hover {{
            background-color: rgba(255, 95, 0, 0.02);
        }}

        .btn-export {{
            background: rgba(16, 185, 129, 0.1);
            color: var(--success);
            border: 1px solid rgba(16, 185, 129, 0.2);
            padding: 8px 16px;
            border-radius: 8px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}

        .btn-export:hover {{
            background: var(--success);
            color: white;
        }}
    </style>
</head>
<body>

    <!-- HEADER -->
    <div class="dashboard-header">
        <div class="brand-section">
            <div class="brand-icon">
                <i class="fa-solid fa-boxes-stacked"></i>
            </div>
            <div class="brand-title">
                <h1>Hệ Thống Theo Dõi Hàng Nặng (> 10kg) - Vùng NTB</h1>
                <p>Theo dõi biến động Sản lượng Tạo & Sản lượng Về Bưu cục theo Ngày</p>
            </div>
        </div>
        <div class="data-source-badge">
            <i class="fa-solid fa-database"></i> Live Google Sheet Data (Sheet: trên10kg & SL > 10kg)
        </div>
    </div>

    <!-- FILTER PANEL WITH DATE PICKER -->
    <div class="filter-panel">
        <div class="filter-header">
            <h3><i class="fa-solid fa-filter" style="color:var(--primary);"></i> Bộ Lọc & Khoảng Thời Gian</h3>
            <div class="date-presets">
                <button class="btn-preset active" onclick="setPreset('all')">Toàn bộ</button>
                <button class="btn-preset" onclick="setPreset('0803')">Đỉnh Sale 03/08</button>
                <button class="btn-preset" onclick="setPreset('0801')">Đầu tháng 01/08</button>
                <button class="btn-preset" onclick="setPreset('last7')">7 ngày gần đây</button>
            </div>
        </div>

        <div class="filter-grid">
            <!-- Mandatory Date Picker Range -->
            <div class="filter-item">
                <label><i class="fa-regular fa-calendar"></i> Từ Ngày</label>
                <input type="date" id="date-from" onchange="applyFilters()">
            </div>
            <div class="filter-item">
                <label><i class="fa-regular fa-calendar"></i> Đến Ngày</label>
                <input type="date" id="date-to" onchange="applyFilters()">
            </div>

            <!-- Warehouse Search / Selection -->
            <div class="filter-item">
                <label><i class="fa-solid fa-warehouse"></i> Tên Bưu Cục</label>
                <select id="select-bc" onchange="applyFilters()">
                    <option value="ALL">-- Tất cả Bưu Cục (84 Bưu cục) --</option>
                </select>
            </div>

            <!-- Province Filter -->
            <div class="filter-item">
                <label><i class="fa-solid fa-map-location-dot"></i> Tỉnh / Thành Phố</label>
                <select id="select-province" onchange="applyFilters()">
                    <option value="ALL">-- Tất cả Tỉnh --</option>
                </select>
            </div>

            <!-- Customer Group Filter -->
            <div class="filter-item">
                <label><i class="fa-solid fa-users"></i> Nhóm Khách Hàng</label>
                <select id="select-kh" onchange="applyFilters()">
                    <option value="ALL">-- Tất cả Nhóm KH --</option>
                    <option value="Shopee-Bulky">Shopee-Bulky (66.6%)</option>
                    <option value="SME">SME (24.5%)</option>
                    <option value="TTS-nhỏ">TTS-nhỏ (7.0%)</option>
                </select>
            </div>

            <!-- Weight Bracket Filter -->
            <div class="filter-item">
                <label><i class="fa-solid fa-weight-hanging"></i> Nhóm Khối Lượng</label>
                <select id="select-kl" onchange="applyFilters()">
                    <option value="ALL">-- Tất cả Khối Lượng --</option>
                    <option value="10-20kg">10kg - 20kg</option>
                    <option value="20-30kg">20kg - 30kg</option>
                    <option value="30-50kg">30kg - 50kg</option>
                    <option value=">50kg">> 50kg</option>
                </select>
            </div>
        </div>
    </div>

    <!-- KPI CARDS -->
    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-top">
                <span class="kpi-title">Sản Lượng Tạo (Sheet trên10kg)</span>
                <div class="kpi-icon"><i class="fa-solid fa-cart-plus"></i></div>
            </div>
            <div class="kpi-value" id="kpi-created-val">0</div>
            <div class="kpi-subtext" id="kpi-created-sub">
                <span class="badge-normal"><i class="fa-solid fa-check"></i> Đang chọn ngày</span>
            </div>
        </div>

        <div class="kpi-card blue">
            <div class="kpi-top">
                <span class="kpi-title">Sản Lượng Về Kho (Sheet SL>10kg)</span>
                <div class="kpi-icon"><i class="fa-solid fa-truck-ramp-box"></i></div>
            </div>
            <div class="kpi-value" id="kpi-volume-val">0</div>
            <div class="kpi-subtext" id="kpi-volume-sub">Hàng mới về kho ca 1 + ca 2</div>
        </div>

        <div class="kpi-card green">
            <div class="kpi-top">
                <span class="kpi-title">Giao Thành Công (GTC)</span>
                <div class="kpi-icon"><i class="fa-solid fa-circle-check"></i></div>
            </div>
            <div class="kpi-value" id="kpi-gtc-val">0</div>
            <div class="kpi-subtext" id="kpi-gtc-sub">%GTC: 0%</div>
        </div>

        <div class="kpi-card purple">
            <div class="kpi-top">
                <span class="kpi-title">Sản Lượng Tồn Kho</span>
                <div class="kpi-icon"><i class="fa-solid fa-boxes-packing"></i></div>
            </div>
            <div class="kpi-value" id="kpi-ton-val">0</div>
            <div class="kpi-subtext" id="kpi-ton-sub">Số đơn đang tồn đọng</div>
        </div>
    </div>

    <!-- CHART CONTAINER -->
    <div class="chart-container">
        <div class="chart-header">
            <h3><i class="fa-solid fa-chart-line" style="color:var(--primary);"></i> Diễn Biến Sản Lượng Tạo vs Sản Lượng Về Theo Ngày</h3>
            <span style="font-size:12px; color:var(--text-sub);">Dự báo sớm trước 24h - 48h khi có đợt đột biến</span>
        </div>
        <div id="chart-daily" style="min-height: 320px;"></div>
    </div>

    <!-- DATA TABLE CONTAINER -->
    <div class="table-card">
        <div class="table-header">
            <h3><i class="fa-solid fa-list-check" style="color:var(--blue);"></i> Bảng Chi Tiết Theo Bưu Cục & Phân Nhóm</h3>
            <button class="btn-export" onclick="exportCSV()"><i class="fa-solid fa-file-csv"></i> Xuất Báo Cáo CSV</button>
        </div>
        <div style="overflow-x: auto;">
            <table class="custom-table">
                <thead>
                    <tr>
                        <th onclick="sortTable('bc')">Tên Bưu Cục <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('province')">Tỉnh <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('am')">AM Quản Lý <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('created')" style="text-align:right;">Đơn Tạo (Sheet trên10kg) <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('volume')" style="text-align:right;">Volume Về (Sheet SL>10kg) <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('gtc')" style="text-align:right;">Giao Thành Công <i class="fa-solid fa-sort"></i></th>
                        <th onclick="sortTable('ton')" style="text-align:right;">Số Tồn <i class="fa-solid fa-sort"></i></th>
                        <th style="text-align:center;">Đánh Giá Tải</th>
                    </tr>
                </thead>
                <tbody id="table-body">
                    <!-- Rows rendered dynamically -->
                </tbody>
            </table>
        </div>
    </div>

    <script>
        // DATA LOADED FROM GOOGLE SHEETS
        const rawCreated = {json_created};
        const rawOps = {json_ops};

        let chartInstance = null;
        let currentTableData = [];

        // Initialize Options
        window.addEventListener('DOMContentLoaded', () => {{
            populateDropdowns();
            
            // Set default date range min and max from data
            const dates = rawCreated.map(r => r.date).filter(Boolean).sort();
            if (dates.length > 0) {{
                document.getElementById('date-from').value = dates[0];
                document.getElementById('date-to').value = dates[dates.length - 1];
            }}
            
            applyFilters();
        }});

        function populateDropdowns() {{
            const bcSet = new Set();
            const provSet = new Set();

            rawCreated.forEach(r => {{
                if (r.bc) bcSet.add(r.bc);
                if (r.province) provSet.add(r.province);
            }});
            rawOps.forEach(r => {{
                if (r.bc) bcSet.add(r.bc);
                if (r.province) provSet.add(r.province);
            }});

            const selectBc = document.getElementById('select-bc');
            Array.from(bcSet).sort().forEach(bc => {{
                const opt = document.createElement('option');
                opt.value = bc;
                opt.textContent = bc;
                selectBc.appendChild(opt);
            }});

            const selectProv = document.getElementById('select-province');
            Array.from(provSet).sort().forEach(p => {{
                const opt = document.createElement('option');
                opt.value = p;
                opt.textContent = p;
                selectProv.appendChild(opt);
            }});
        }}

        function setPreset(type) {{
            document.querySelectorAll('.btn-preset').forEach(b => b.classList.remove('active'));
            event.target.classList.add('active');

            if (type === 'all') {{
                const dates = rawCreated.map(r => r.date).filter(Boolean).sort();
                document.getElementById('date-from').value = dates[0] || '';
                document.getElementById('date-to').value = dates[dates.length - 1] || '';
            }} else if (type === '0803') {{
                document.getElementById('date-from').value = '2026-08-03';
                document.getElementById('date-to').value = '2026-08-03';
            }} else if (type === '0801') {{
                document.getElementById('date-from').value = '2026-08-01';
                document.getElementById('date-to').value = '2026-08-01';
            }} else if (type === 'last7') {{
                document.getElementById('date-from').value = '2026-07-30';
                document.getElementById('date-to').value = '2026-08-05';
            }}
            applyFilters();
        }}

        function applyFilters() {{
            const dateFrom = document.getElementById('date-from').value;
            const dateTo = document.getElementById('date-to').value;
            const selBc = document.getElementById('select-bc').value;
            const selProv = document.getElementById('select-province').value;
            const selKh = document.getElementById('select-kh').value;
            const selKl = document.getElementById('select-kl').value;

            // Filter Created
            let filtCreated = rawCreated.filter(r => {{
                if (dateFrom && r.date < dateFrom) return false;
                if (dateTo && r.date > dateTo) return false;
                if (selBc !== 'ALL' && r.bc !== selBc) return false;
                if (selProv !== 'ALL' && r.province !== selProv) return false;
                if (selKh !== 'ALL' && r.kh !== selKh) return false;
                if (selKl !== 'ALL' && r.kl !== selKl) return false;
                return true;
            }});

            // Filter Ops
            let filtOps = rawOps.filter(r => {{
                if (dateFrom && r.date < dateFrom) return false;
                if (dateTo && r.date > dateTo) return false;
                if (selBc !== 'ALL' && r.bc !== selBc) return false;
                if (selProv !== 'ALL' && r.province !== selProv) return false;
                return true;
            }});

            // 1. Calculate KPIs
            const totalCreated = filtCreated.reduce((acc, r) => acc + r.so_don, 0);
            const totalVolume = filtOps.reduce((acc, r) => acc + r.volume, 0);
            const totalGtc = filtOps.reduce((acc, r) => acc + r.gtc, 0);
            const totalTon = filtOps.reduce((acc, r) => acc + r.ton, 0);
            const pctGtc = totalVolume > 0 ? ((totalGtc / totalVolume) * 100).toFixed(1) : '0';

            document.getElementById('kpi-created-val').innerText = totalCreated.toLocaleString('vi-VN');
            document.getElementById('kpi-volume-val').innerText = totalVolume.toLocaleString('vi-VN');
            document.getElementById('kpi-gtc-val').innerText = totalGtc.toLocaleString('vi-VN');
            document.getElementById('kpi-ton-val').innerText = totalTon.toLocaleString('vi-VN');
            document.getElementById('kpi-gtc-sub').innerText = `%GTC Giao Thành Công: ${{pctGtc}}%`;

            const createdSub = document.getElementById('kpi-created-sub');
            if (totalCreated > 4500) {{
                createdSub.innerHTML = `<span class="badge-spike"><i class="fa-solid fa-fire"></i> CẢNH BÁO ĐỘT BIẾN TẠO (${{totalCreated.toLocaleString()}} đơn)</span>`;
            }} else {{
                createdSub.innerHTML = `<span class="badge-normal"><i class="fa-solid fa-circle-check"></i> Sản lượng bình thường</span>`;
            }}

            // 2. Render Daily Chart
            renderDailyChart(filtCreated, filtOps);

            // 3. Render Table
            renderTable(filtCreated, filtOps);
        }}

        function renderDailyChart(filtCreated, filtOps) {{
            const dateMap = {{}};

            filtCreated.forEach(r => {{
                if (!dateMap[r.date]) dateMap[r.date] = {{ created: 0, volume: 0 }};
                dateMap[r.date].created += r.so_don;
            }});

            filtOps.forEach(r => {{
                if (!dateMap[r.date]) dateMap[r.date] = {{ created: 0, volume: 0 }};
                dateMap[r.date].volume += r.volume;
            }});

            const sortedDates = Object.keys(dateMap).sort();
            const createdSeries = sortedDates.map(d => dateMap[d].created);
            const volumeSeries = sortedDates.map(d => dateMap[d].volume);

            const options = {{
                series: [
                    {{ name: 'Đơn Tạo (Sheet trên10kg)', data: createdSeries, type: 'column' }},
                    {{ name: 'Volume Về Bưu Cục (Sheet SL>10kg)', data: volumeSeries, type: 'line' }}
                ],
                chart: {{
                    height: 350,
                    type: 'line',
                    toolbar: {{ show: true }},
                    fontFamily: 'Inter, sans-serif'
                }},
                stroke: {{ width: [0, 4], curve: 'smooth' }},
                colors: ['#ff5f00', '#007bc3'],
                dataLabels: {{ enabled: false }},
                xaxis: {{ categories: sortedDates }},
                yaxis: {{
                    title: {{ text: 'Số Đơn / Volume' }}
                }},
                tooltip: {{
                    shared: true,
                    intersect: false
                }}
            }};

            if (chartInstance) chartInstance.destroy();
            chartInstance = new ApexCharts(document.querySelector("#chart-daily"), options);
            chartInstance.render();
        }}

        function renderTable(filtCreated, filtOps) {{
            const bcMap = {{}};

            filtCreated.forEach(r => {{
                if (!bcMap[r.bc]) bcMap[r.bc] = {{ bc: r.bc, province: r.province, am: '---', created: 0, volume: 0, gtc: 0, ton: 0 }};
                bcMap[r.bc].created += r.so_don;
            }});

            filtOps.forEach(r => {{
                if (!bcMap[r.bc]) bcMap[r.bc] = {{ bc: r.bc, province: r.province || '---', am: r.am || '---', created: 0, volume: 0, gtc: 0, ton: 0 }};
                if (r.am && r.am !== 'Chưa phân loại') bcMap[r.bc].am = r.am;
                if (r.province && r.province !== 'Chưa phân loại') bcMap[r.bc].province = r.province;
                bcMap[r.bc].volume += r.volume;
                bcMap[r.bc].gtc += r.gtc;
                bcMap[r.bc].ton += r.ton;
            }});

            currentTableData = Object.values(bcMap).sort((a, b) => b.created - a.created);
            drawTable(currentTableData);
        }}

        function drawTable(data) {{
            const tbody = document.getElementById('table-body');
            tbody.innerHTML = '';

            if (data.length === 0) {{
                tbody.innerHTML = '<tr><td colspan="8" style="text-align:center; padding:20px; color:#94a3b8;">Không có dữ liệu thỏa mãn bộ lọc</td></tr>';
                return;
            }}

            data.forEach((r, idx) => {{
                const tr = document.createElement('tr');
                let tag = '<span class="badge-normal">Bình thường</span>';
                if (r.created > 1500 || r.volume > 2000) {{
                    tag = '<span class="badge-spike">🔥 Tải Rất Cao</span>';
                }}

                tr.innerHTML = `
                    <td style="font-weight:600; color:var(--text-main);">${{r.bc}}</td>
                    <td>${{r.province}}</td>
                    <td>${{r.am}}</td>
                    <td style="text-align:right; font-weight:700; color:var(--primary);">${{r.created.toLocaleString('vi-VN')}}</td>
                    <td style="text-align:right; font-weight:700; color:var(--blue);">${{r.volume.toLocaleString('vi-VN')}}</td>
                    <td style="text-align:right;">${{r.gtc.toLocaleString('vi-VN')}}</td>
                    <td style="text-align:right; color:${{r.ton > 50 ? 'var(--danger)' : 'inherit'}};">${{r.ton.toLocaleString('vi-VN')}}</td>
                    <td style="text-align:center;">${{tag}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        function sortTable(key) {{
            currentTableData.sort((a, b) => {{
                if (typeof a[key] === 'string') return a[key].localeCompare(b[key]);
                return b[key] - a[key];
            }});
            drawTable(currentTableData);
        }}

        function exportCSV() {{
            let csv = 'Tên Bưu Cục,Tỉnh,AM Quản Lý,Số Đơn Tạo,Volume Về,Giao Thành Công,Số Tồn\\n';
            currentTableData.forEach(r => {{
                csv += `"${{r.bc}}","${{r.province}}","${{r.am}}",${{r.created}},${{r.volume}},${{r.gtc}},${{r.ton}}\\n`;
            }});

            const blob = new Blob(['\\uFEFF' + csv], {{ type: 'text/csv;charset=utf-8;' }});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `NTB_Hang_Nang_Bao_Cao_${{new Date().toISOString().slice(0,10)}}.csv`;
            a.click();
        }}
    </script>
</body>
</html>
"""

if __name__ == "__main__":
    main()
