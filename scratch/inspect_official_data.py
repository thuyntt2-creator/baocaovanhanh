import sys, json

sys.stdout.reconfigure(encoding='utf-8')

with open(r'c:\Users\lap4all\Documents\Auto report\scratch\web_data\rezone.json', 'r', encoding='utf-8') as f:
    rezone = json.load(f)

ntb_wards = [w for w in rezone['new_wards'] if w.get('region') == 'NTB']
print(f'Total NTB new wards in official system: {len(ntb_wards)}')

print('=' * 80)
for idx, w in enumerate(ntb_wards, 1):
    cands = ', '.join([f"{c['bc_name']} ({c['share']}%, {c['dem']} đơn)" for c in w.get('candidates', [])])
    old_names = ', '.join([f"{o['name']} ({o['dem']} đơn)" for o in w.get('olds', [])])
    print(f"{idx:2d}. [{w['new_code']}] {w['name']} ({w['province']})")
    print(f"    - Sản lượng: {w['dem']} đơn/ngày ({w['dem_kg']} kg/ngày) | Trạng thái: {w['status']} | BC Đề xuất: {w['assigned_bc_name']}")
    print(f"    - Bưu cục candidates phụ trách: {cands}")
    print(f"    - Xã cũ sáp nhập: {old_names}")
    print()
