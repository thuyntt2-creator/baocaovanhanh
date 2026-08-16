import sys, json
sys.stdout.reconfigure(encoding='utf-8')

with open('snapshot_aging.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Keep only the initial morning snapshot (07:59) for today's history
if 'history' in data and len(data['history']) > 0:
    snap_0759 = data['history'][0]
    EXCLUDED_AMS = ["Huỳnh Tấn Hiền", "Nguyễn Tiến Lực", "Nguyễn Minh Hoàng"]
    if 'totals' in snap_0759:
        snap_0759['totals'] = {k: v for k, v in snap_0759['totals'].items() if k not in EXCLUDED_AMS}
    data['history'] = [snap_0759]

with open('snapshot_aging.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("✅ Đã khôi phục snapshot_aging.json về mốc ban đầu 07:59 thành công!")
