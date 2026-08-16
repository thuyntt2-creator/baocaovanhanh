import pandas as pd, sys

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel(r'C:\Users\lap4all\Downloads\NTB_Phan_Tuyen_Hanh_Chinh_Quy_Hoach_Moi.xlsx', sheet_name='Sheet1')

print("=== MASTER EXCEL NEW WARDS SUMMARY ===")
unique_wards = df[['Tỉnh, thành phố mới', 'Tên Xã mới', 'Mã Xã mới', 'ID GHN xã mới', 'Đánh giá & Phương án đề xuất', 'Quản lý khu vực (AM)']].drop_duplicates(subset=['Mã Xã mới'])

for idx, row in unique_wards.iterrows():
    prov = row['Tỉnh, thành phố mới']
    name = row['Tên Xã mới']
    code = row['Mã Xã mới']
    am = row['Quản lý khu vực (AM)']
    dexuat = str(row['Đánh giá & Phương án đề xuất'])[:60]
    print(f"Mã {code:5d} | {prov:16s} | {name:30s} | AM: {str(am):22s} | Đề xuất: {dexuat}")
