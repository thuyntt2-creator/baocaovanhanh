import sys, json

sys.stdout.reconfigure(encoding='utf-8')

with open(r'c:\Users\lap4all\Documents\Auto report\scratch\web_data\rezone.json', 'r', encoding='utf-8') as f:
    rezone = json.load(f)

ntb = [w for w in rezone['new_wards'] if w.get('region') == 'NTB']
print(f'Total NTB new wards in rezone.json: {len(ntb)}')
for w in ntb[:15]:
    cands = ', '.join([f"{c['bc_name']} ({c['share']}%, {c['dem']}đ)" for c in w.get('candidates', [])])
    print(f"[{w['new_code']}] {w['name']} ({w['province']}) | Dem: {w['dem']}đ ({w['dem_kg']}kg) | Status: {w['status']} | Assigned: {w['assigned_bc_name']}")
    print(f"   -> Candidates: {cands}")
