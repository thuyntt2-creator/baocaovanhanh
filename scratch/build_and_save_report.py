import sys, docx, pandas as pd, json, os

sys.stdout.reconfigure(encoding='utf-8')

# File paths
excel_path = r'C:\Users\lap4all\Downloads\NTB_Phan_Tuyen_Hanh_Chinh_Quy_Hoach_Moi.xlsx'
df = pd.read_excel(excel_path, sheet_name='Sheet1')

doc2 = docx.Document(r'C:\Users\lap4all\Downloads\Bao_Cao_Quy_Hoach_Buu_Cuc_NTB_Don_Vi_Hanh_Chinh_Moi.docx')

am_reasons = {
    24611: '''Lịch sử vận hành: Khu vực Gia Nghĩa từng bể tuyến triền miên nhiều tháng do không tuyển được nhân sự. Phương án tách 3 kho (Bắc Gia Nghĩa, Nam Gia Nghĩa 2, Nhân Cơ) vừa mới ổn định được vận hành và thu hút lực lượng lao động tại chỗ.
Xã Đắk Ha: Nhân sự hiện tại là người xã Quảng Sơn, ưu điểm gần kho Quảng Sơn. Nếu chuyển về kho Gia Nghĩa quản lý sẽ phát sinh di chuyển 25-30km/ngày, rủi ro nhân sự nghỉ việc 100%.
Đặc thù dân cư: Khu vực có tỷ lệ đồng bào dân tộc thiểu số cao. Việc tuyển nhân sự tại chỗ đã khó, nếu đổi BC mới càng không thể tuyển dụng.
Rủi ro nhân sự: Nếu sáp nhập BC theo xã mới, 100% nhân sự hiện tại các điểm xã ven đều xác nhận sẽ KHÔNG di chuyển theo bưu cục mới.
Đề xuất AM: TẠM THỜI GIỮ NGUYÊN quản lý theo các BC cũ để tránh vỡ vận hành.''',

    24615: '''Lịch sử vận hành: Giữ nguyên phân vùng giao của BC Nam Gia Nghĩa 2 và Bắc Gia Nghĩa để tránh nguy cơ gãy tuyến và biến động nhân sự. Nhân sự hiện tại tại các xã cũ quen thuộc địa bàn và kho hiện tại.
Đề xuất AM: TẠM THỜI GIỮ NGUYÊN 100% Bưu cục và phạm vi quản lý theo ranh giới Xã/Phường cũ.''',

    24781: '''Bưu cục (LDO) Xuân Hương - Đà Lạt hiện phụ trách sản lượng lớn (2,292 đơn/ngày). Địa hình Đà Lạt bị chia cắt bởi Hồ Xuân Hương, thời tiết cuối năm mưa và lạnh, thiếu nhân sự. Việc dùng 1 bưu cục sẽ gây quá tải trầm trọng.
Đề xuất AM: Tách thêm 01 bưu cục mới là Bưu cục (LDO) Xuân Hương - Đà Lạt 2 (đặt tại Phường 10, cover Phường 3 & Phường 10, sản lượng 1,050 đơn giao, 150 đơn lấy, 8 NVPTTT + 1 NVXL) và giữ BC Xuân Hương - Đà Lạt cũ (cover Phường 1 & Phường 2, sản lượng 1,800 đơn giao, 400 đơn lấy, 13 NVPTTT + 1 NVXL).''',

    24784: '''Phường Lâm Viên (Đà Lạt mới) có đặc thù địa hình bị chia cắt bởi Hồ Xuân Hương. Giữ nguyên 02 bưu cục (LDO) Lâm Viên - Đà Lạt 1 và (LDO) Lâm Viên - Đà Lạt 2 nằm ngay trong khu vực Phường 8 giúp nhân viên chỉ tập trung xử lý một phường, giảm thời gian di chuyển, đảm bảo hiệu quả giao hàng.''',

    24958: '''Trong địa bàn Đức Trọng có 2 bưu cục (LDO Đức Trọng 1 & LDO Đức Trọng 2). Nếu gộp lại 1 bưu cục, tuyến giao của nhân viên xa và dẫn đến chồng chéo tuyến. Khoảng cách giữa 2 BC trên 15km.
Lý do giữ 2 BC: Địa bàn rộng, khó tuyển dụng (đặc biệt mùa cà phê và mùa mưa cuối năm), khoảng cách di chuyển tới địa điểm giao của nhân viên xa, BC cũ không đủ không gian m² để chứa hàng, vị trí nhà tới BC xa gây ảnh hưởng xuất phát giao trễ.
Phương án: Đề xuất gộp tuyến giao theo xã mới về (LDO) Đức Trọng 2 chiếm 79.7% sản lượng (901 đơn/ngày), (LDO) Đức Trọng 1 giữ nhiệm vụ phụ trách các xã lân cận.''',

    22972: '''Vị trí đặt 2 Bưu cục tại 2 thị trấn có sản lượng hàng nhiều nhất. Khoảng cách giữa 2 thị trấn tầm 25km và địa bàn rộng nên không thể gộp lại thành 1 Bưu cục (Phan Rí Cửa và khu vực lân cận). Giữ phân vùng vận hành hiện tại.''',

    23143: '''BTH - Tân Hải phụ trách Tân Thuận, Tân Thành cũ. Đề xuất giữ nguyên phần Thuận Quý cũ cho BC BTH - Hàm Thuận Nam vì địa hình khu vực đặc thù, khoảng cách từ BC Hàm Thuận Nam đến địa điểm gần hơn, dân cư tập trung đông đúc hơn so với khoảng cách từ BC Tân Hải.''',

    23131: '''BTH - Phước Hội (mới) cover Phường Lagi, Phước Hội, Sơn Mỹ (1,200-1,500 đơn giao, 280 đơn lấy, 15 NVPTTT). BTH - Tân Hải (mới) cover Xã Tân Hải, Xã Tân Thành (400-500 đơn giao, 120 đơn lấy, 7 NVPTTT). Phân định rõ ranh giới Lagi và Tân Hải để tối ưu bán kính di chuyển.''',

    22528: '''Huyện Ninh Hòa (cũ) có diện tích lớn nhất cả nước (20 xã/phường). Bưu cục Ninh Hòa 1 được tách ra để cover các tuyến xã xa trung tâm, đặt tại các tuyến xã thường xuyên thiếu hụt nhân sự để lôi kéo nguồn lực nhân sự tại chỗ, tiện giao lấy. Chia nhỏ cụm xã nóng vừa tận dụng được nguồn lực mới vừa đảm bảo thiếu hụt nhân sự không ảnh hưởng vận hành. Diện tích BC Ninh Hòa 2 cũng không đảm bảo m² để gom gộp.''',

    25000: '''BC (LDO) Hòa Ninh hiện cover tuyến Liên Đầm của BC (LDO) Di Linh. Sau sáp nhập Liên Đầm gộp về xã Di Linh, tuy nhiên đặc thù khoảng cách xa với BC hiện tại, nhân sự chưa ổn định. Đề xuất tách Bưu cục Hàng Nhỏ / Hàng Vừa (Xã Đinh Trang Thượng) để cover Đinh Trang Thượng, Di Linh, Phúc Thọ, Liên Đầm, giảm tải cho BC Di Linh (sản lượng 1,300-1,600 đơn/ngày, bán kính xa 22-45km).''',

    22366: '''Sản lượng Phường Nha Trang rất lớn (1,592 đơn/ngày). Đề xuất giữ nguyên/gộp phân vùng về BC KHO Nha Trang chính, đồng thời giữ các bưu cục Nam Nha Trang 1 & 2 phụ trách phân đoạn phụ để tránh quá tải kho bãi và vỡ tuyến.''',

    22402: '''Sản lượng Phường Nam Nha Trang cực lớn (2,418 đơn/ngày). Giữ Bưu cục Nam Nha Trang 1 mới (gộp Nam Nha Trang 3 và phần Nam Nha Trang 1 cũ) và giữ nguyên Bưu cục Nam Nha Trang 5 (phụ trách Phước Đồng). Đóng/bỏ bưu cục Nam Nha Trang 2 & Nam Nha Trang 3 để tối ưu điểm tập kết.'''
}

ward_list = []
for code, group in df.groupby('Mã Xã mới'):
    prov = group['Tỉnh, thành phố mới'].iloc[0]
    name = group['Tên Xã mới'].iloc[0]
    n_bcs = group['Số BC'].iloc[0]
    old_communes = group['Tên Xã cũ'].tolist()
    total_giao = group['Sản lượng giao/ngày (đơn)'].sum()
    total_lay = group['Sản lượng lấy/ngày (đơn)'].sum()
    total_don = group['TỔNG ĐƠN/NGÀY (Phường mới)'].iloc[0]
    total_kg = group['TỔNG KG/NGÀY (Phường mới)'].iloc[0]
    prop = group['Đánh giá & Phương án đề xuất'].iloc[0]
    reason_excel = group['Lý do & Bố trí nhân sự'].iloc[0]

    bc_details = []
    for bc_name, bc_group in group.groupby('Tên Bưu cục giao'):
        bc_giao = bc_group['Sản lượng giao/ngày (đơn)'].sum()
        bc_lay = bc_group['Sản lượng lấy/ngày (đơn)'].sum()
        bc_total = bc_giao + bc_lay
        pct = round((bc_total / total_don * 100), 1) if total_don > 0 else 0.0
        bc_id = bc_group['ID Bưu cục giao'].iloc[0]
        am = bc_group['Quản lý khu vực (AM)'].iloc[0]
        addr = bc_group['Địa chỉ Bưu cục'].iloc[0]
        gmaps = bc_group['Google Maps Bưu cục'].iloc[0]
        bc_details.append({
            'name': bc_name, 'id': bc_id, 'am': am, 'addr': addr, 'gmaps': gmaps,
            'giao': bc_giao, 'lay': bc_lay, 'total': bc_total, 'pct': pct
        })
    bc_details.sort(key=lambda x: x['total'], reverse=True)

    if code in am_reasons:
        final_reason = am_reasons[code]
    else:
        final_reason = reason_excel

    ward_list.append({
        'code': code, 'name': name, 'prov': prov, 'n_bcs': n_bcs,
        'old_communes': old_communes, 'total_giao': total_giao, 'total_lay': total_lay,
        'total_don': total_don, 'total_kg': total_kg, 'prop': prop,
        'bc_details': bc_details, 'reason': final_reason
    })

ward_list.sort(key=lambda x: x['name'])

# Build Markdown content
md = []
md.append('# BÁO CÁO TOÀN DIỆN: RÀ SOÁT VÀ QUY HOẠCH BƯU CỤC THEO ĐƠN VỊ HÀNH CHÍNH MỚI (VÙNG NTB)\n')

md.append('## I. TỔNG QUAN HIỆN TRẠNG MẠNG LƯỚI BƯU CỤC VÙNG NTB')
md.append('- **Tổng số Bưu cục Express đang vận hành**: 83 Bưu cục (thuộc các tỉnh Khánh Hòa, Ninh Thuận, Bình Thuận, Lâm Đồng, Đắc Nông cũ).')
md.append('- **Tổng số Xã/Phường hành chính mới rà soát**: 36 Xã/Phường mới (gồm 114 xã/phường cũ sáp nhập).')
md.append('- **Số Xã/Phường mới đã quy hoạch chuẩn (01 BC phụ trách)**: 4 Xã/Phường (chiếm 11.1%).')
md.append('- **Số Xã/Phường mới bị CHIA CẮT (2-3 BC phụ trách)**: 32 Xã/Phường (chiếm 88.9%).\n')

md.append('## II. RÀ SOÁT THEO ĐƠN VỊ HÀNH CHÍNH MỚI')
md.append('### 1. Đánh giá độ phủ Bưu cục và ranh giới hành chính mới')
md.append('Sau khi sáp nhập các xã cũ thành xã/phường mới, ranh giới quản lý của các Bưu cục đang xuất hiện hiện tượng chia cắt mảnh, dẫn đến:')
md.append('- **01 Phường mới có tới 3 Bưu cục cùng giao hàng**: Gây chồng chéo tuyến đường, làm shipper di chuyển cắt ngang địa bàn của nhau.')
md.append('- **Tuyến đi chéo xa ranh giới**: Một số xã vùng ven bị giao từ Bưu cục ở huyện/tỉnh khác cách xa 30 - 40 km, trong khi Bưu cục lân cận chỉ cách 7 - 15 km.\n')

md.append('### 2. Danh sách 18 Tuyến giao chéo xa ranh giới cần Reassign ngay')
md.append('| STT | Tên Xã/Phường cũ | Tỉnh | BC hiện tại (`from`) | Khoảng cách cũ | BC tối ưu gần nhất (`to`) | Quản lý AM tiếp nhận | Khoảng cách mới | Khoảng cách tiết kiệm |')
md.append('|---|---|---|---|---|---|---|---|---|')

for r in doc2.tables[1].rows[1:]:
    cells = [c.text.strip().replace('\n', ' ') for c in r.cells]
    md.append(f'| {cells[0]} | {cells[1]} | {cells[2]} | {cells[3]} | {cells[4]} | {cells[5]} | {cells[6]} | {cells[7]} | {cells[8]} |')

md.append('\n## III. ĐÁNH GIÁ CHI TIẾT SẢN LƯỢNG VÀ ĐỀ XUẤT PHƯƠNG ÁN GỘP / GIỮ BƯU CỤC')
md.append(f'### Danh sách {len(ward_list)} Phường/Xã mới rà soát\n')

for i, w in enumerate(ward_list, 1):
    md.append(f'### {i}. {w["name"]} ({w["prov"]})')
    md.append(f'- **Mã Xã mới**: `{w["code"]}`')
    md.append(f'- **TỔNG SẢN LƯỢNG DỰ KIẾN**: {w["total_don"]:.1f} đơn/ngày ({w["total_kg"]:.1f} kg/ngày) | *Giao: {w["total_giao"]:.1f} đơn · Lấy: {w["total_lay"]:.1f} đơn*')
    
    bcs_str = ', '.join([f'"{b["name"]}"' for b in w['bc_details']])
    md.append(f'- **Các BC hiện phụ trách ({w["n_bcs"]} BC)**: {bcs_str}')
    
    old_str = ', '.join(w['old_communes'])
    md.append(f'- **Các xã cũ sáp nhập**: {old_str}')
    md.append('- **Tỷ lệ phân chia sản lượng thực tế giữa các Bưu cục**:')
    
    for b in w['bc_details']:
        gmaps_str = f' [Google Maps]({b["gmaps"]})' if b['gmaps'] and str(b['gmaps']) != 'nan' else ''
        md.append(f'  - `{b["name"]}` (ID: {b["id"]}) [AM: {b["am"]}]: {b["total"]:.1f} đơn/ngày ({b["pct"]}%) (Giao: {b["giao"]:.1f}, Lấy: {b["lay"]:.1f}) (Địa chỉ: {b["addr"]}){gmaps_str}')
    
    md.append(f'- **ĐỀ XUẤT PHƯƠNG ÁN**: **{w["prop"]}**')
    md.append(f'- **LÝ DO VÀ BỐ TRÍ NHÂN SỰ**:\n{w["reason"]}\n')

md.append('## IV. YÊU CẦU 3: ĐÁNH GIÁ NHU CẦU TÁCH BƯU CỤC VÀ TỐI ƯU CÔNG SUẤT KHO BÃI')
md.append('### 1. Đánh giá Bưu cục quá tải (Cần tách bớt Phường hoặc Mở rộng diện tích)')
md.append('| STT | ID Bưu cục | Tên Bưu cục | Tỉnh | Quản lý AM | Địa chỉ Bưu cục | Áp lực m² (`em2`) | Đề xuất giải pháp |')
md.append('|---|---|---|---|---|---|---|---|')

for r in doc2.tables[0].rows[1:]:
    cells = [c.text.strip().replace('\n', ' ') for c in r.cells]
    md.append(f'| {cells[0]} | {cells[1]} | {cells[2]} | {cells[3]} | {cells[4]} | {cells[5]} | {cells[6]} | {cells[7]} |')

md.append('\n### 2. Kế hoạch Mở mới / Tách Bưu cục & Tối ưu Mạng lưới')
md.append('1. **Mở mới Bưu cục (LDO) Xuân Hương - Đà Lạt 2**: Đặt tại khu vực Phường 10 (TP. Đà Lạt), chia tải cho Bưu cục Xuân Hương cũ (phụ trách Phường 3 & Phường 10, Vol giao: 1,050 đơn/ngày, Vol lấy: 150 đơn/ngày, 8 NVPTTT + 1 NVXL).')
md.append('2. **Tách Bưu cục Hàng Nhỏ / Hàng Vừa Di Linh (Xã Đinh Trang Thượng)**: Phụ trách Đinh Trang Thượng, Di Linh, Phúc Thọ Lâm Hà, Liên Đầm để giảm bán kính di chuyển 22-45km cho BC Di Linh.')
md.append('3. **Mở mới Bưu cục Đông Hải (Tỉnh Ninh Thuận)**: Cover khu vực ven biển Đông Hải, tối ưu điểm tập kết sản lượng lấy.')
md.append('4. **Mở mới Bưu cục (LDO) B\'Lao Mới (Bảo Lộc)**: Tối ưu mạng lưới khu vực Phường 1 & Phường 2 Bảo Lộc, đóng cửa BC (LDO) 1 Bảo Lộc cũ.\n')

md.append('## V. TỔNG HỢP BIẾN ĐỘNG MẠNG LƯỚI BƯU CỤC NTB 2026')
md.append('- **Bưu cục Mở mới (04 BC)**: BC Xuân Hương - Đà Lạt 2, BC Di Linh Hàng Nhỏ, BC Đông Hải, BC B\'Lao Mới.')
md.append('- **Bưu cục Đóng cửa (02 BC)**: BC 1 Bảo Lộc, BC Nam Nha Trang 3.')
md.append('- **Bưu cục Gộp tuyến/Điều chỉnh phân vùng (19 BC)**: Gom các tuyến xã lẻ về BC trung tâm theo ĐVHC mới.')
md.append('- **Bưu cục Giữ nguyên vận hành (13 BC)**: Duy trì tại các khu vực đặc thù Gia Nghĩa, Đà Lạt, Đức Trọng, Phan Rí Cửa, Ninh Hòa...')

full_md_text = '\n'.join(md)

# Write Markdown report to Artifact path
artifact_md_path = r'C:\Users\lap4all\.gemini\antigravity-ide\brain\d95894ad-9248-47c3-872b-72d7de66bc20\Bao_Cao_Quy_Hoach_Buu_Cuc_NTB_Don_Vi_Hanh_Chinh_Moi.md'
with open(artifact_md_path, 'w', encoding='utf-8') as f:
    f.write(full_md_text)

print(f'Saved Markdown report to: {artifact_md_path}')

# Now write Word DOCX report
new_doc = docx.Document()

# Add Title
new_doc.add_heading('BÁO CÁO TOÀN DIỆN: RÀ SOÁT VÀ QUY HOẠCH BƯU CỤC THEO ĐƠN VỊ HÀNH CHÍNH MỚI (VÙNG NTB)', level=0)

for line in md[1:]:
    if line.startswith('## '):
        new_doc.add_heading(line.replace('## ', ''), level=1)
    elif line.startswith('### '):
        new_doc.add_heading(line.replace('### ', ''), level=2)
    elif line.startswith('|') and '---' in line:
        continue
    elif line.startswith('|'):
        continue # tables handled separately below
    elif line.startswith('- '):
        new_doc.add_paragraph(line.replace('- ', ''), style='List Bullet')
    elif line.startswith('  - '):
        new_doc.add_paragraph(line.replace('  - ', ''), style='List Bullet 2')
    elif line.strip():
        new_doc.add_paragraph(line.strip())

# Add Table 2 to docx
new_doc.add_heading('Bảng 1: Danh sách 18 Tuyến giao chéo xa ranh giới cần Reassign ngay', level=2)
t2_doc = new_doc.add_table(rows=len(doc2.tables[1].rows), cols=len(doc2.tables[1].columns))
for r_idx, r in enumerate(doc2.tables[1].rows):
    for c_idx, c in enumerate(r.cells):
        t2_doc.cell(r_idx, c_idx).text = c.text.strip().replace('\n', ' ')

# Add Table 1 to docx
new_doc.add_heading('Bảng 2: Danh sách Bưu cục quá tải áp lực diện tích m² (em2)', level=2)
t1_doc = new_doc.add_table(rows=len(doc2.tables[0].rows), cols=len(doc2.tables[0].columns))
for r_idx, r in enumerate(doc2.tables[0].rows):
    for c_idx, c in enumerate(r.cells):
        t1_doc.cell(r_idx, c_idx).text = c.text.strip().replace('\n', ' ')

docx_out_path = r'C:\Users\lap4all\Downloads\Bao_Cao_Quy_Hoach_Buu_Cuc_NTB_Don_Vi_Hanh_Chinh_Moi_Hoan_Chinh.docx'
new_doc.save(docx_out_path)
print(f'Saved DOCX report to: {docx_out_path}')
