import pandas as pd, sys

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel(r'C:\Users\lap4all\Downloads\NTB_Phan_Tuyen_Hanh_Chinh_Quy_Hoach_Moi.xlsx', sheet_name='Sheet1')

print("=== DETAILED MASTER WARDS & MERGED COMMUNES ===")
grouped = df.groupby(['Tỉnh, thành phố mới', 'Tên Xã mới', 'Mã Xã mới'])

for (prov, new_ward, code), group in grouped:
    old_communes = ", ".join(group['Tên Xã cũ'].astype(str).unique())
    bcs_current = ", ".join(group['Tên Bưu cục giao'].astype(str).unique())
    vol_deliv = group['Sản lượng giao/ngày (đơn)'].sum()
    vol_pick = group['Sản lượng lấy/ngày (đơn)'].sum()
    am = group['Quản lý khu vực (AM)'].iloc[0]
    proposal = group['Đánh giá & Phương án đề xuất'].iloc[0]
    print(f"[{prov}] {new_ward} (Mã: {code})")
    print(f"   Xã cũ sáp nhập: {old_communes}")
    print(f"   BC hiện tại ({len(group['Tên Bưu cục giao'].unique())}): {bcs_current}")
    print(f"   Sản lượng: {vol_deliv+vol_pick:.1f} đơn/ngày (Giao: {vol_deliv:.1f}, Lấy: {vol_pick:.1f})")
    print(f"   AM: {am} | Phương án: {proposal}")
    print("-" * 60)
