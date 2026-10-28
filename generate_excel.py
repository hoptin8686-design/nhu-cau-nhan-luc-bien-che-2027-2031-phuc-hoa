# -*- coding: utf-8 -*-
"""
Script tự động khởi tạo Bộ biểu Excel Xác định nhu cầu nhân lực và xây dựng phương án, 
biên chế giai đoạn 2027-2031 của Trường THPT Phục Hòa theo Hướng dẫn số 07-HD/BTCTW.
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.Workbook()
# remove default sheet
wb.remove(wb.active)

# Color palettes
NAVY_HEADER = "1B365D"
TEAL_HEADER = "0E7490"
ACCENT_BLUE = "0284C7"
LIGHT_BLUE = "F0F9FF"
LIGHT_GRAY = "F8FAFC"
BORDER_COLOR = "CBD5E1"
TEXT_WHITE = "FFFFFF"
STATUS_GREEN = "16A34A"

font_title = Font(name="Times New Roman", size=14, bold=True, color=NAVY_HEADER)
font_subtitle = Font(name="Times New Roman", size=11, italic=True)
font_header = Font(name="Times New Roman", size=11, bold=True, color=TEXT_WHITE)
font_sub_header = Font(name="Times New Roman", size=10, bold=True, color=TEXT_WHITE)
font_bold = Font(name="Times New Roman", size=10, bold=True)
font_normal = Font(name="Times New Roman", size=10)
font_note = Font(name="Times New Roman", size=9, italic=True)

fill_navy = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
fill_teal = PatternFill(start_color=TEAL_HEADER, end_color=TEAL_HEADER, fill_type="solid")
fill_accent = PatternFill(start_color=ACCENT_BLUE, end_color=ACCENT_BLUE, fill_type="solid")
fill_light = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
fill_gray = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")

thin_border_side = Side(border_style="thin", color=BORDER_COLOR)
thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
thick_bottom = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=Side(border_style="medium", color=NAVY_HEADER))

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

def style_header_row(ws, row_idx, bg_fill=fill_navy):
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = font_header
        cell.fill = bg_fill
        cell.alignment = align_center
        cell.border = thin_border

def auto_fit_columns(ws, max_cols=None):
    max_c = max_cols or ws.max_column
    for col in range(1, max_c + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                # ignore long title lines in column 1
                if row <= 3 and col == 1:
                    continue
                s_len = len(str(val).split('\n')[0])
                if s_len > max_len:
                    max_len = s_len
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

# ==========================================
# SHEET 1: TONG_QUAN
# ==========================================
ws_tq = wb.create_sheet(title="TONG_QUAN")
ws_tq.views.sheetView[0].showGridLines = True

ws_tq["A1"] = "UBND TỈNH CAO BẰNG - SỞ GIÁO DỤC VÀ ĐÀO TẠO"
ws_tq["A1"].font = font_bold
ws_tq["A2"] = "TRƯỜNG THPT PHỤC HÒA"
ws_tq["A2"].font = Font(name="Times New Roman", size=11, bold=True, color=NAVY_HEADER)

ws_tq["A4"] = "BÁO CÁO TỔNG HỢP XÁC ĐỊNH NHU CẦU NHÂN LỰC VÀ PHƯƠNG ÁN BIÊN CHẾ GIAI ĐOẠN 2027-2031"
ws_tq["A4"].font = font_title
ws_tq["A5"] = "(Theo Kết luận số 107-KL/TW, Hướng dẫn số 07-HD/BTCTW của Ban Tổ chức Trung ương và Công văn số 12176/BNV-CCVC của Bộ Nội vụ)"
ws_tq["A5"].font = font_subtitle

overview_data = [
    ["1. Đơn vị báo cáo:", "Trường Trung học phổ thông Phục Hòa (Sở GD&ĐT tỉnh Cao Bằng)"],
    ["2. Địa bàn công tác:", "Thị trấn Phục Hòa, huyện Quảng Hòa, tỉnh Cao Bằng (Vùng biên giới, điều kiện KT-XH đặc biệt khó khăn)"],
    ["3. Căn cứ định mức ngành:", "Thông tư số 20/2023/TT-BGDĐT ngày 30/10/2023 của Bộ Giáo dục và Đào tạo (Vùng 1: tối đa 2,45 GV/lớp)"],
    ["4. Quy mô trường lớp năm học 2026-2027:", "14 lớp, tổng số 573 học sinh (bình quân 41 học sinh/lớp)"],
    ["5. Dự báo quy mô đến 2031:", "Năm 2027-2028: 15 lớp (600 HS) | Năm 2029-2031: 16 lớp (640-650 HS)"],
    ["6. Tổng biên chế được giao hiện nay (đến 30/9/2026):", "35 người (Gồm 33 viên chức biên chế + 02 Hợp đồng lao động NĐ 111/2022)"],
    ["7. Cơ cấu nhân lực hiện có:", "Lãnh đạo quản lý: 03 | Giáo viên trực tiếp giảng dạy: 25 | Nhân viên hỗ trợ: 05 | HĐLĐ 111: 02"],
    ["8. Tổng quỹ giờ lao động ròng toàn trường/năm:", "55.850 giờ lao động chuyên môn chuẩn"],
    ["9. Tổng số người quy đổi theo PP2 (FTE):", "37,99 người"],
    ["10. Nhu cầu nhân lực cần bố trí theo PP1 & PP3:", "37 người (năm 2027) tăng lên 39 người (năm 2029-2031 khi trường đạt 16 lớp)"],
    ["11. Đề xuất chỉ tiêu biên chế giai đoạn 2027-2031:", "Năm 2027: 35 BC + 2 HĐ 111 | Năm 2028: 35 BC + 2 HĐ 111 | Năm 2029-2031: 37 BC + 2 HĐ 111"],
    ["12. Hạn nộp báo cáo hoàn thiện:", "Ngày 30/10/2026 theo Công văn số /SGDĐT-TCCB của Sở Giáo dục và Đào tạo tỉnh Cao Bằng"]
]

for idx, (lbl, val) in enumerate(overview_data, start=7):
    ws_tq.cell(row=idx, column=1, value=lbl).font = font_bold
    ws_tq.cell(row=idx, column=2, value=val).font = font_normal
    ws_tq.cell(row=idx, column=1).border = thin_border
    ws_tq.cell(row=idx, column=2).border = thin_border
    if idx % 2 == 1:
        ws_tq.cell(row=idx, column=1).fill = fill_gray
        ws_tq.cell(row=idx, column=2).fill = fill_gray

ws_tq.column_dimensions["A"].width = 46
ws_tq.column_dimensions["B"].width = 95

# ==========================================
# SHEET 2: 1. A1_NhiemVu
# ==========================================
ws_a1 = wb.create_sheet(title="1. A1_NhiemVu")
ws_a1.views.sheetView[0].showGridLines = True
ws_a1["A1"] = "Biểu A1_ Danh mục nhiệm vụ của Trường THPT Phục Hòa"
ws_a1["A1"].font = font_title
ws_a1["G1"] = "→ Căn cứ Hướng dẫn 07-HD/BTCTW"
ws_a1["G1"].font = font_note

a1_headers = ["Mã nhiệm vụ", "Tên nhiệm vụ", "Cơ quan", "Căn cứ", "Kỳ", "Trạng thái", "Nguồn"]
for col_idx, h in enumerate(a1_headers, start=1):
    ws_a1.cell(row=2, column=col_idx, value=h)
style_header_row(ws_a1, 2, fill_navy)

a1_data = [
    ["NV.PH.01", "Quản lý, điều hành toàn diện hoạt động của nhà trường", "Trường THPT Phục Hòa", "Điều lệ trường trung học (Thông tư 32/2020/TT-BGDĐT)", "2025-2026", "Tiếp tục thực hiện", "QĐ phân công nhiệm vụ BGH"],
    ["NV.PH.02", "Tổ chức giảng dạy các môn học theo Chương trình GDPT 2018 cấp THPT", "Trường THPT Phục Hòa", "Thông tư 32/2018/TT-BGDĐT, Thông tư 13/2022/TT-BGDĐT", "2025-2026", "Tiếp tục thực hiện", "Kế hoạch giáo dục nhà trường"],
    ["NV.PH.03", "Tổ chức hoạt động trải nghiệm, hướng nghiệp và giáo dục địa phương", "Trường THPT Phục Hòa", "Chương trình GDPT 2018", "2025-2026", "Tiếp tục thực hiện", "Kế hoạch giáo dục khối 10, 11, 12"],
    ["NV.PH.04", "Công tác chủ nhiệm lớp và tư vấn tâm lý học đường, hỗ trợ học sinh", "Trường THPT Phục Hòa", "Thông tư 32/2020/TT-BGDĐT, Thông tư 20/2023/TT-BGDĐT", "2025-2026", "Tiếp tục thực hiện", "Quyết định phân công GVCN"],
    ["NV.PH.05", "Bồi dưỡng học sinh giỏi các môn văn hóa, KHKT cấp tỉnh, cấp quốc gia", "Trường THPT Phục Hòa", "CV hướng dẫn thi HSG Sở GD&ĐT Cao Bằng", "2025-2026", "Tiếp tục thực hiện", "Quyết định thành lập đội tuyển HSG"],
    ["NV.PH.06", "Kiểm tra, đánh giá thường xuyên, định kỳ và ôn thi tốt nghiệp THPT", "Trường THPT Phục Hòa", "Thông tư 22/2021/TT-BGDĐT, Quy chế thi tốt nghiệp THPT", "2025-2026", "Tiếp tục thực hiện", "Kế hoạch kiểm tra nội bộ"],
    ["NV.PH.07", "Quản lý tổ chuyên môn, sinh hoạt chuyên môn theo nghiên cứu bài học", "Trường THPT Phục Hòa", "Điều lệ trường trung học, CV 5512/BGDĐT", "2025-2026", "Tiếp tục thực hiện", "Kế hoạch tổ Tự nhiên, Xã hội"],
    ["NV.PH.08", "Quản lý tài chính, kế toán ngân sách và tài sản công trường học", "Trường THPT Phục Hòa", "Luật Ngân sách, Luật Kế toán, Nghị định 60/2021/NĐ-CP", "2025-2026", "Tiếp tục thực hiện", "Quy chế chi tiêu nội bộ"],
    ["NV.PH.09", "Công tác văn thư, lưu trữ, tiếp nhận văn bản và quản trị hồ sơ điện tử", "Trường THPT Phục Hòa", "Nghị định 30/2020/NĐ-CP về công tác văn thư", "2025-2026", "Tiếp tục thực hiện", "Hệ thống iOffice, vnEdu"],
    ["NV.PH.10", "Quản lý phòng thí nghiệm, thực hành, thiết bị dạy học số và hóa chất", "Trường THPT Phục Hòa", "Thông tư 20/2023/TT-BGDĐT, Danh mục thiết bị tối thiểu", "2025-2026", "Tiếp tục thực hiện", "Sổ theo dõi thiết bị, thí nghiệm"],
    ["NV.PH.11", "Quản lý thư viện trường học đạt chuẩn và phát triển văn hóa đọc", "Trường THPT Phục Hòa", "Thông tư 16/2022/TT-BGDĐT tiêu chuẩn thư viện cơ sở GDPT", "2025-2026", "Tiếp tục thực hiện", "Sổ đăng ký cá biệt, phần mềm TV"],
    ["NV.PH.12", "Chăm sóc sức khỏe ban đầu, y tế học đường, vệ sinh an toàn thực phẩm", "Trường THPT Phục Hòa", "Thông tư liên tịch 13/2016/TTLT-BYT-BGDĐT, TT 20/2023", "2025-2026", "Tiếp tục thực hiện", "Sổ theo dõi sức khỏe học sinh"],
    ["NV.PH.13", "Quản trị cơ sở dữ liệu ngành giáo dục, chuyển đổi số và cổng thông tin điện tử", "Trường THPT Phục Hòa", "QĐ 4725/QĐ-BGDĐT, CV Sở GD&ĐT về CSDL ngành", "2025-2026", "Tiếp tục thực hiện", "CSDL ngành CSDL.moet.gov.vn"],
    ["NV.PH.14", "Bảo vệ an ninh trật tự trường học, an toàn phòng cháy chữa cháy", "Trường THPT Phục Hòa", "Nghị định 111/2022/NĐ-CP, Quy chế an toàn trường học", "2025-2026", "Tiếp tục thực hiện", "Hợp đồng lao động bảo vệ"],
    ["NV.PH.15", "Vệ sinh môi trường cảnh quan sư phạm trường học hạnh phúc", "Trường THPT Phục Hòa", "Nghị định 111/2022/NĐ-CP, Phong trào trường xanh sạch đẹp", "2025-2026", "Tiếp tục thực hiện", "Hợp đồng dịch vụ vệ sinh"]
]

for row_idx, r_data in enumerate(a1_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_a1.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        cell.alignment = align_center if col_idx in [1, 5, 6] else align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

auto_fit_columns(ws_a1)

# ==========================================
# SHEET 3: 2. A2_CongViec
# ==========================================
ws_a2 = wb.create_sheet(title="2. A2_CongViec")
ws_a2.views.sheetView[0].showGridLines = True
ws_a2["A1"] = "Biểu A2_ Phân rã công việc và sản phẩm gắn với VTVL Trường THPT Phục Hòa"
ws_a2["A1"].font = font_title
ws_a2["H1"] = "→ Căn cứ Hướng dẫn 07-HD/BTCTW"
ws_a2["H1"].font = font_note

a2_headers = ["Mã công việc", "Mã nhiệm vụ", "Mã VTVL", "Tên công việc", "Sản phẩm", "Đơn vị tính", "Chất lượng", "Căn cứ (nguồn)"]
for col_idx, h in enumerate(a2_headers, start=1):
    ws_a2.cell(row=2, column=col_idx, value=h)
style_header_row(ws_a2, 2, fill_teal)

a2_data = [
    # Quản lý
    ["CV.01.01", "NV.PH.01", "VTVL.QL.01", "Xây dựng Kế hoạch chiến lược phát triển nhà trường giai đoạn 2027-2031", "Kế hoạch chiến lược được phê duyệt", "Kế hoạch", "Đạt chuẩn, đúng định hướng phát triển giáo dục tỉnh", "Nghị quyết Chi bộ, QĐ ban hành"],
    ["CV.01.02", "NV.PH.01", "VTVL.QL.01", "Quản lý tổ chức bộ máy, tuyển dụng, đào tạo, quy hoạch và đánh giá viên chức", "Hồ sơ quy hoạch, đánh giá CBGVNV", "Bộ hồ sơ", "Đúng thẩm quyền, công bằng, đúng NĐ 232/2026", "Hồ sơ công tác cán bộ"],
    ["CV.01.03", "NV.PH.01", "VTVL.QL.02", "Xây dựng Kế hoạch giáo dục nhà trường, phân công chuyên môn và TKB", "Kế hoạch giáo dục năm học, TKB", "Bản kế hoạch", "Phù hợp chương trình GDPT 2018, bảo đảm số tiết chuẩn", "Kế hoạch dạy học, TKB các tuần"],
    ["CV.01.04", "NV.PH.01", "VTVL.QL.02", "Chỉ đạo kiểm tra nội bộ, kiểm định chất lượng giáo dục và kiểm tra chuyên môn", "Biên bản kiểm tra nội bộ", "Biên bản", "Đúng quy trình thanh kiểm tra của Bộ và Sở", "Hồ sơ kiểm tra nội bộ"],
    
    # Giáo viên bộ môn
    ["CV.02.01", "NV.PH.02", "VTVL.GV.01", "Giảng dạy môn Ngữ văn cấp THPT (Lớp 10, 11, 12)", "Tiết dạy chuẩn kiến thức kỹ năng", "Tiết", "Đạt yêu cầu cần đạt của CT GDPT 2018, 100% giáo án chuẩn", "Phân công CM, Sổ đầu bài"],
    ["CV.02.02", "NV.PH.02", "VTVL.GV.02", "Giảng dạy môn Toán cấp THPT (Lớp 10, 11, 12)", "Tiết dạy môn Toán", "Tiết", "Bảo đảm tư duy logic, đúng chuẩn kiến thức", "Phân công CM, Sổ đầu bài"],
    ["CV.02.03", "NV.PH.02", "VTVL.GV.03", "Giảng dạy môn Tiếng Anh cấp THPT (Lớp 10, 11, 12)", "Tiết dạy Tiếng Anh", "Tiết", "Phát triển 4 kỹ năng ngôn ngữ cho HS vùng khó khăn", "Phân công CM, Sổ đầu bài"],
    ["CV.02.04", "NV.PH.02", "VTVL.GV.04", "Giảng dạy môn Lịch sử cấp THPT (Lớp 10, 11, 12)", "Tiết dạy Lịch sử", "Tiết", "Bảo đảm tính chính xác, giáo dục lòng yêu nước", "Phân công CM, Sổ đầu bài"],
    ["CV.02.05", "NV.PH.02", "VTVL.GV.05", "Giảng dạy môn Giáo dục thể chất cấp THPT", "Tiết dạy Thể chất", "Tiết", "An toàn, phát triển thể lực học sinh", "Phân công CM, Sổ đầu bài"],
    ["CV.02.06", "NV.PH.02", "VTVL.GV.06", "Giảng dạy môn Giáo dục quốc phòng và an ninh", "Tiết dạy GDQP-AN", "Tiết", "Đúng kỹ năng quân sự, pháp luật quốc phòng an ninh", "Phân công CM, Sổ đầu bài"],
    ["CV.02.07", "NV.PH.02", "VTVL.GV.07", "Giảng dạy môn Địa lí cấp THPT", "Tiết dạy Địa lí", "Tiết", "Đạt chuẩn CT GDPT 2018, ứng dụng bản đồ số", "Phân công CM, Sổ đầu bài"],
    ["CV.02.08", "NV.PH.02", "VTVL.GV.08", "Giảng dạy môn Giáo dục Kinh tế và Pháp luật", "Tiết dạy GDKT&PL", "Tiết", "Trực quan, gắn thực tiễn đời sống và pháp luật", "Phân công CM, Sổ đầu bài"],
    ["CV.02.09", "NV.PH.02", "VTVL.GV.09", "Giảng dạy môn Vật lí cấp THPT", "Tiết dạy Vật lí", "Tiết", "Kết hợp thực hành thí nghiệm đầy đủ", "Phân công CM, Sổ đầu bài"],
    ["CV.02.10", "NV.PH.02", "VTVL.GV.10", "Giảng dạy môn Hóa học cấp THPT", "Tiết dạy Hóa học", "Tiết", "An toàn hóa chất, đạt chuẩn kiến thức", "Phân công CM, Sổ đầu bài"],
    ["CV.02.11", "NV.PH.02", "VTVL.GV.11", "Giảng dạy môn Sinh học cấp THPT", "Tiết dạy Sinh học", "Tiết", "Trực quan sinh động, phát triển năng lực KHTN", "Phân công CM, Sổ đầu bài"],
    ["CV.02.12", "NV.PH.02", "VTVL.GV.12", "Giảng dạy môn Công nghệ (Công nghệ Công nghiệp & Nông nghiệp)", "Tiết dạy Công nghệ", "Tiết", "Ứng dụng kỹ thuật số, nông nghiệp công nghệ cao", "Phân công CM, Sổ đầu bài"],
    ["CV.02.13", "NV.PH.02", "VTVL.GV.13", "Giảng dạy môn Tin học cấp THPT (ICT & Khoa học máy tính)", "Tiết dạy Tin học", "Tiết", "Thực hành phòng máy 100%, nâng cao năng lực AI & CĐS", "Phân công CM, Sổ đầu bài"],

    # Hoạt động chuyên môn khác của GV
    ["CV.03.01", "NV.PH.03", "VTVL.GV.00", "Tổ chức Hoạt động trải nghiệm, hướng nghiệp và GD địa phương", "Buổi hoạt động, chuyên đề", "Chuyên đề", "Đúng chủ đề GDPT 2018, học sinh tham gia tích cực", "Kế hoạch trải nghiệm khối"],
    ["CV.04.01", "NV.PH.04", "VTVL.GV.00", "Công tác chủ nhiệm lớp, quản lý nề nếp và họp phụ huynh", "Sổ chủ nhiệm, biên bản họp PH", "Lớp", "Quản lý tốt học sinh, duy trì sĩ số 100%", "Sổ chủ nhiệm lớp"],
    ["CV.05.01", "NV.PH.05", "VTVL.GV.00", "Bồi dưỡng học sinh giỏi cấp tỉnh và nghiên cứu KHKT", "Học sinh đạt giải cấp tỉnh", "Đội tuyển", "Có giải cấp tỉnh, nâng cao chất lượng mũi nhọn", "QĐ công nhận kết quả thi"],
    ["CV.06.01", "NV.PH.06", "VTVL.GV.00", "Biên soạn ma trận, bản đặc tả đề kiểm tra và chấm bài đánh giá HS", "Bài kiểm tra đã chấm, nhập điểm", "Bài kiểm tra", "Đúng ma trận, đánh giá công bằng, khách quan", "Hệ thống vnEdu"],

    # Nhân viên hỗ trợ
    ["CV.08.01", "NV.PH.08", "VTVL.NV.01", "Lập dự toán, quản lý chi tiêu ngân sách và lập báo cáo tài chính", "Báo cáo tài chính quý, năm", "Bộ chứng từ", "Đúng Luật Ngân sách, không sai sót tài chính", "Sổ kế toán, phần mềm MISA"],
    ["CV.09.01", "NV.PH.09", "VTVL.NV.02", "Tiếp nhận, xử lý công văn đi/đến, quản lý con dấu và lưu trữ văn thư", "Văn bản xử lý trên iOffice", "Văn bản", "Kịp thời, chính xác, bảo mật theo quy định", "Sổ đăng ký văn bản, iOffice"],
    ["CV.10.01", "NV.PH.10", "VTVL.NV.03", "Chuẩn bị thiết bị, hóa chất, phục vụ các tiết dạy thực hành thí nghiệm", "Buổi thực hành có chuẩn bị đồ dùng", "Buổi", "An toàn tuyệt đối, đầy đủ đồ dùng thực hành", "Sổ mượn trả thiết bị"],
    ["CV.11.01", "NV.PH.11", "VTVL.NV.04", "Xử lý nghiệp vụ thư viện, quản lý sách giáo khoa, phục vụ đọc sách", "Lượt bạn đọc, đầu sách số hóa", "Lượt mượn", "Đạt chuẩn thư viện mức độ 1 trở lên", "Phần mềm thư viện trường học"],
    ["CV.12.01", "NV.PH.12", "VTVL.NV.05", "Quản lý tủ thuốc, sơ cấp cứu ban đầu, khám sức khỏe định kỳ cho HS", "Hồ sơ y tế học đường", "Hồ sơ", "Đúng phác đồ cấp cứu, kiểm tra VSATTP thường xuyên", "Sổ theo dõi khám sức khỏe"],
    ["CV.14.01", "NV.PH.14", "VTVL.HD.01", "Bảo vệ an ninh trật tự, tuần tra ban đêm, trực cổng trường 24/24", "Ca trực an toàn", "Ca", "Không xảy ra mất mát tài sản, an toàn trường học", "Sổ trực bảo vệ"],
    ["CV.15.01", "NV.PH.15", "VTVL.HD.02", "Lao động dọn dẹp vệ sinh khu công cộng, sân trường, nhà vệ sinh", "Khuôn viên sạch sẽ", "Ngày công", "Đảm bảo tiêu chí trường học Sáng - Xanh - Sạch - Đẹp", "Biên bản nghiệm thu lao động"]
]

for row_idx, r_data in enumerate(a2_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_a2.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        cell.alignment = align_center if col_idx in [1, 2, 3, 6] else align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

auto_fit_columns(ws_a2)

# ==========================================
# SHEET 4: 3. B1_KhoiLuong
# ==========================================
ws_b1 = wb.create_sheet(title="3. B1_KhoiLuong")
ws_b1.views.sheetView[0].showGridLines = True
ws_b1["A1"] = "Biểu B1_ Thống kê khối lượng công việc, sản lượng và quỹ thời gian"
ws_b1["A1"].font = font_title
ws_b1["I1"] = "→ Kỳ thống kê: 01/10/2025 đến 30/9/2026 (12 tháng)"
ws_b1["I1"].font = font_note

b1_headers = [
    "Mã công việc", "Mã VTVL", "Số SP/năm", "Trung bình phút/SP", 
    "Giờ bổ sung", "Giờ SP (tính)", "Giờ tổng (tính)", 
    "Nguồn số liệu", "Nguồn thời gian", "Trạng thái căn cứ", "Cảnh báo"
]
for col_idx, h in enumerate(b1_headers, start=1):
    ws_b1.cell(row=2, column=col_idx, value=h)
style_header_row(ws_b1, 2, fill_navy)

b1_data = [
    # Quản lý
    ["CV.01.01", "VTVL.QL.01", 1, 12000, 200, 200, 400, "Nghị quyết Chi bộ, QĐ số 56/QĐ", "Nghị định 232/2026/NĐ-CP", "Có định mức", "ĐẠT"],
    ["CV.01.02", "VTVL.QL.01", 12, 4800, 500, 960, 1460, "Hồ sơ công tác tổ chức cán bộ", "Thông tư 32/2020/TT-BGDĐT", "Có định mức", "ĐẠT"],
    ["CV.01.03", "VTVL.QL.02", 35, 1800, 400, 1050, 1450, "Kế hoạch giáo dục, TKB các tuần", "Thông tư 20/2023/TT-BGDĐT", "Có định mức", "ĐẠT"],
    ["CV.01.04", "VTVL.QL.02", 24, 2400, 300, 960, 1260, "Biên bản kiểm tra chuyên môn", "Kế hoạch kiểm tra nội bộ Sở", "Có định mức", "ĐẠT"],

    # Giáo viên đứng lớp (14 lớp)
    ["CV.02.01", "VTVL.GV.01", 1750, 45, 1200, 1312.5, 2512.5, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Ngữ văn)", "Có định mức", "ĐẠT"],
    ["CV.02.02", "VTVL.GV.02", 1750, 45, 1200, 1312.5, 2512.5, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Toán)", "Có định mức", "ĐẠT"],
    ["CV.02.03", "VTVL.GV.03", 1470, 45, 950, 1102.5, 2052.5, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Tiếng Anh)", "Có định mức", "ĐẠT"],
    ["CV.02.04", "VTVL.GV.04", 980, 45, 700, 735.0, 1435.0, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Lịch sử)", "Có định mức", "ĐẠT"],
    ["CV.02.05", "VTVL.GV.05", 980, 45, 700, 735.0, 1435.0, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (GDTC)", "Có định mức", "ĐẠT"],
    ["CV.02.06", "VTVL.GV.06", 490, 45, 400, 367.5, 767.5, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (GDQP-AN)", "Có định mức", "ĐẠT"],
    ["CV.02.07", "VTVL.GV.07", 980, 45, 700, 735.0, 1435.0, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Địa lí)", "Có định mức", "ĐẠT"],
    ["CV.02.08", "VTVL.GV.08", 490, 45, 400, 367.5, 767.5, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (GDKT&PL)", "Có định mức", "ĐẠT"],
    ["CV.02.09", "VTVL.GV.09", 588, 45, 450, 441.0, 891.0, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Vật lí)", "Có định mức", "ĐẠT"],
    ["CV.02.10", "VTVL.GV.10", 588, 45, 450, 441.0, 891.0, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Hóa học)", "Có định mức", "ĐẠT"],
    ["CV.02.11", "VTVL.GV.11", 588, 45, 450, 441.0, 891.0, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Sinh học)", "Có định mức", "ĐẠT"],
    ["CV.02.12", "VTVL.GV.12", 980, 45, 700, 735.0, 1435.0, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Công nghệ)", "Có định mức", "ĐẠT"],
    ["CV.02.13", "VTVL.GV.13", 490, 45, 400, 367.5, 767.5, "Sổ đầu bài 14 lớp, vnEdu", "Thông tư 20/2023/TT-BGDĐT (Tin học)", "Có định mức", "ĐẠT"],

    # Nhân viên hành chính
    ["CV.08.01", "VTVL.NV.01", 180, 480, 400, 1440.0, 1840.0, "Chứng từ kho bạc, sổ quỹ", "Luật Kế toán, Nghị định 60/2021", "Có định mức", "ĐẠT"],
    ["CV.09.01", "VTVL.NV.02", 2150, 35, 350, 1254.2, 1604.2, "Sổ văn bản iOffice tỉnh", "Nghị định 30/2020/NĐ-CP", "Có định mức", "ĐẠT"],
    ["CV.10.01", "VTVL.NV.03", 560, 120, 450, 1120.0, 1570.0, "Sổ nhật ký phòng thí nghiệm", "Thông tư 20/2023/TT-BGDĐT", "Có định mức", "ĐẠT"],
    ["CV.11.01", "VTVL.NV.04", 1250, 60, 400, 1250.0, 1650.0, "Phần mềm thư viện trường học", "Thông tư 16/2022/TT-BGDĐT", "Có định mức", "ĐẠT"],
    ["CV.12.01", "VTVL.NV.05", 573, 90, 600, 859.5, 1459.5, "Sổ theo dõi khám sức khỏe", "TTLT 13/2016/TTLT-BYT-BGDĐT", "Có định mức", "ĐẠT"],
    ["CV.14.01", "VTVL.HD.01", 365, 360, 600, 2190.0, 2790.0, "Sổ trực cổng trường, camera", "Nghị định 111/2022/NĐ-CP", "Thực tế", "ĐẠT"],
    ["CV.15.01", "VTVL.HD.02", 280, 300, 400, 1400.0, 1800.0, "Biên bản nghiệm thu vệ sinh", "Nghị định 111/2022/NĐ-CP", "Thực tế", "ĐẠT"]
]

for row_idx, r_data in enumerate(b1_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_b1.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        if col_idx in [3, 4, 5, 6, 7]:
            cell.alignment = align_right
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0.0" if isinstance(val, float) else "#,##0"
        elif col_idx in [1, 2, 10, 11]:
            cell.alignment = align_center
            if col_idx == 11 and val == "ĐẠT":
                cell.font = Font(name="Times New Roman", size=10, bold=True, color=STATUS_GREEN)
        else:
            cell.alignment = align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

auto_fit_columns(ws_b1)

# ==========================================
# SHEET 5: 4. B2_ThayDoi
# ==========================================
ws_b2 = wb.create_sheet(title="4. B2_ThayDoi")
ws_b2.views.sheetView[0].showGridLines = True
ws_b2["A1"] = "Biểu B2_ Đánh giá tác động chuyển đổi số và công nghệ đến khối lượng công việc"
ws_b2["A1"].font = font_title

b2_headers = [
    "Mã VTVL", "Tên VTVL", "Nội dung ứng dụng công nghệ / CĐS", 
    "Giờ giảm do tự động hóa/năm", "Giờ công việc mới phát sinh/năm", "Giờ tác động ròng (+/-)", "Giải trình căn cứ"
]
for col_idx, h in enumerate(b2_headers, start=1):
    ws_b2.cell(row=2, column=col_idx, value=h)
style_header_row(ws_b2, 2, fill_teal)

b2_data = [
    ["VTVL.QL.01", "Hiệu trưởng", "Ứng dụng chữ ký số, điều hành văn bản qua iOffice, quản lý hồ sơ CSDL ngành", 120, 40, -80, "Giảm thời gian ký duyệt văn bản giấy, tăng thời gian bảo mật dữ liệu số"],
    ["VTVL.QL.02", "Phó Hiệu trưởng", "Sử dụng phần mềm xếp TKB tự động và quản lý sổ điểm điện tử vnEdu", 150, 60, -90, "Giảm công tác rà soát điểm thủ công, tăng nhiệm vụ quản lý học liệu số"],
    ["VTVL.GV.01", "Giáo viên Ngữ văn", "Áp dụng giáo án điện tử, chấm bài kiểm tra trực tuyến trên LMS", 80, 50, -30, "Giảm thời gian in ấn hồ sơ giấy, tập trung phân hóa năng lực học sinh"],
    ["VTVL.GV.02", "Giáo viên Toán", "Ứng dụng Geogebra, phần mềm vẽ hình và ngân hàng đề thi trực tuyến", 90, 40, -50, "Rút ngắn thời gian ra đề trắc nghiệm, nâng cao chất lượng đề thi"],
    ["VTVL.GV.03", "Giáo viên Tiếng Anh", "Ứng dụng phòng học ngoại ngữ thông minh, phần mềm luyện phát âm AI", 70, 50, -20, "Học sinh tự luyện tập nhiều hơn, giáo viên kiểm soát tiến độ trên hệ thống"],
    ["VTVL.GV.12", "Giáo viên Công nghệ", "Ứng dụng mô phỏng kỹ thuật, nông nghiệp chính xác và STEM/AI", 60, 60, 0, "Chuyển giao công nghệ mới trong giảng dạy thực hành"],
    ["VTVL.GV.13", "Giáo viên Tin học", "Quản trị hệ thống mạng trường học, hỗ trợ CĐS và phổ cập năng lực AI cho HS", 50, 150, 100, "Phát sinh nhiệm vụ quản trị cổng TTĐT, an toàn thông tin và đào tạo kỹ năng số"],
    ["VTVL.NV.01", "Kế toán", "Phần mềm kế toán MISA, hóa đơn điện tử, thanh toán học phí không dùng tiền mặt", 160, 40, -120, "Tự động hóa hoàn toàn thanh toán trực tuyến qua ngân hàng"],
    ["VTVL.NV.02", "Văn thư", "Ứng dụng hệ thống Quản lý văn bản điều hành iOffice tỉnh Cao Bằng", 140, 30, -110, "95% văn bản luân chuyển trên môi trường điện tử, không in giấy"]
]

for row_idx, r_data in enumerate(b2_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_b2.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        if col_idx in [4, 5, 6]:
            cell.alignment = align_right
            cell.number_format = "#,##0"
        elif col_idx in [1]:
            cell.alignment = align_center
        else:
            cell.alignment = align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

auto_fit_columns(ws_b2)

# ==========================================
# SHEET 6: 6. C1_VTVL
# ==========================================
ws_c1 = wb.create_sheet(title="6. C1_VTVL")
ws_c1.views.sheetView[0].showGridLines = True
ws_c1["A1"] = "Biểu C1_ Danh mục vị trí việc làm và quỹ thời gian chuẩn Trường THPT Phục Hòa"
ws_c1["A1"].font = font_title

c1_headers = [
    "Mã VTVL", "Tên VTVL", "Loại nhân lực", "Quyết định phê duyệt", 
    "Giờ chuẩn PP2", "Căn cứ xác định quỹ giờ", "Trạng thái", "Chế độ làm việc"
]
for col_idx, h in enumerate(c1_headers, start=1):
    ws_c1.cell(row=2, column=col_idx, value=h)
style_header_row(ws_c1, 2, fill_navy)

c1_data = [
    ["VTVL.QL.01", "Hiệu trưởng", "Viên chức quản lý", "Quyết định của Sở GD&ĐT", 1904, "Hướng dẫn 07-HD/BTCTW (238 ngày x 8h)", "Tiếp tục sử dụng", "40 giờ/tuần"],
    ["VTVL.QL.02", "Phó Hiệu trưởng", "Viên chức quản lý", "Quyết định của Sở GD&ĐT", 1904, "Hướng dẫn 07-HD/BTCTW", "Tiếp tục sử dụng", "40 giờ/tuần"],
    ["VTVL.GV.01", "Giáo viên Ngữ văn", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT (Chế độ 42 tuần)", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.02", "Giáo viên Toán", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.03", "Giáo viên Tiếng Anh", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.04", "Giáo viên Lịch sử", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.05", "Giáo viên Giáo dục thể chất", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.06", "Giáo viên GDQP-AN", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.07", "Giáo viên Địa lí", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.08", "Giáo viên GDKT & PL", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.09", "Giáo viên Vật lí", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.10", "Giáo viên Hóa học", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.11", "Giáo viên Sinh học", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.12", "Giáo viên Công nghệ", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.GV.13", "Giáo viên Tin học", "Viên chức chuyên môn", "Quyết định phê duyệt VTVL", 1470, "Thông tư 20/2023/TT-BGDĐT", "Tiếp tục sử dụng", "35 tuần dạy + 7 tuần"],
    ["VTVL.NV.01", "Kế toán", "Viên chức hỗ trợ, phục vụ", "Quyết định phê duyệt VTVL", 1904, "Hướng dẫn 07-HD/BTCTW", "Tiếp tục sử dụng", "40 giờ/tuần"],
    ["VTVL.NV.02", "Văn thư", "Viên chức hỗ trợ, phục vụ", "Quyết định phê duyệt VTVL", 1904, "Hướng dẫn 07-HD/BTCTW", "Tiếp tục sử dụng", "40 giờ/tuần"],
    ["VTVL.NV.03", "Thiết bị, thí nghiệm", "Viên chức hỗ trợ, phục vụ", "Quyết định phê duyệt VTVL", 1904, "Hướng dẫn 07-HD/BTCTW", "Tiếp tục sử dụng", "40 giờ/tuần"],
    ["VTVL.NV.04", "Thư viện viên", "Viên chức hỗ trợ, phục vụ", "Quyết định phê duyệt VTVL", 1904, "Hướng dẫn 07-HD/BTCTW", "Tiếp tục sử dụng", "40 giờ/tuần"],
    ["VTVL.NV.05", "Y tế trường học", "Viên chức hỗ trợ, phục vụ", "Quyết định phê duyệt VTVL", 1904, "Hướng dẫn 07-HD/BTCTW", "Tiếp tục sử dụng", "40 giờ/tuần"],
    ["VTVL.HD.01", "Bảo vệ", "Hợp đồng theo NĐ 111/2022", "Hợp đồng lao động", 1904, "Nghị định 111/2022/NĐ-CP", "Tiếp tục thực hiện", "Theo ca trực"],
    ["VTVL.HD.02", "Phục vụ / Tạp vụ", "Hợp đồng theo NĐ 111/2022", "Hợp đồng lao động", 1904, "Nghị định 111/2022/NĐ-CP", "Tiếp tục thực hiện", "Hành chính"]
]

for row_idx, r_data in enumerate(c1_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_c1.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        if col_idx == 5:
            cell.alignment = align_right
            cell.number_format = "#,##0"
        elif col_idx in [1, 7, 8]:
            cell.alignment = align_center
        else:
            cell.alignment = align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

auto_fit_columns(ws_c1)

# ==========================================
# SHEET 7: 8. C3_NhuCau
# ==========================================
ws_c3 = wb.create_sheet(title="8. C3_NhuCau")
ws_c3.views.sheetView[0].showGridLines = True
ws_c3["A1"] = "Biểu C3_ Tổng hợp xác định nhu cầu nhân lực Trường THPT Phục Hòa"
ws_c3["A1"].font = font_title
ws_c3["L1"] = "→ Đối chiếu 3 phương pháp PP1, PP2, PP3"
ws_c3["L1"].font = font_note

c3_headers = [
    "Mã VTVL", "Tên VTVL", "Phương pháp", "Giờ B1", "Giờ tác động B2", 
    "Giờ ròng", "Giờ chuẩn C1", "Người quy đổi PP2", "Người PP1 (Định mức)", 
    "Người PP3 (Ước tính)", "Người cần bố trí", "Căn cứ và giải trình khoa học", "Cảnh báo"
]
for col_idx, h in enumerate(c3_headers, start=1):
    ws_c3.cell(row=2, column=col_idx, value=h)
style_header_row(ws_c3, 2, fill_navy)

c3_data = [
    ["VTVL.QL.01", "Hiệu trưởng", "Kết hợp PP1, PP2", 1860, -80, 1780, 1904, 0.93, 1.0, 1.0, 1.0, "Định mức trường hạng 2: 01 Hiệu trưởng theo TT 20/2023", "ĐẠT"],
    ["VTVL.QL.02", "Phó Hiệu trưởng", "Kết hợp PP1, PP2", 2710, -90, 2620, 1904, 1.38, 2.0, 2.0, 2.0, "Quy mô 14 lớp (từ 10-17 lớp được bố trí tối đa 02 PHT)", "ĐẠT"],
    ["VTVL.GV.01", "Giáo viên Ngữ văn", "Kết hợp PP1, PP2", 2512.5, -30, 2482.5, 1470, 1.69, 3.0, 3.0, 3.0, "35 tiết/tuần chia cho 14 lớp. Cần 3 GV để đảm bảo dạy và ôn thi TN", "ĐẠT"],
    ["VTVL.GV.02", "Giáo viên Toán", "Kết hợp PP1, PP2", 2512.5, -50, 2462.5, 1470, 1.68, 3.0, 3.0, 3.0, "Môn cốt lõi thi tốt nghiệp, 35 tiết/tuần. Bố trí 3 GV", "ĐẠT"],
    ["VTVL.GV.03", "Giáo viên Tiếng Anh", "Kết hợp PP1, PP2", 2052.5, -20, 2032.5, 1470, 1.38, 2.5, 3.0, 3.0, "Vùng biên giới cần tăng cường bồi dưỡng ngoại ngữ, bố trí 3 GV", "ĐẠT"],
    ["VTVL.GV.04", "Giáo viên Lịch sử", "Kết hợp PP1, PP2", 1435.0, 0, 1435.0, 1470, 0.98, 2.0, 2.0, 2.0, "Môn bắt buộc GDPT 2018, kiêm nhiệm HĐ trải nghiệm, bố trí 2 GV", "ĐẠT"],
    ["VTVL.GV.05", "Giáo viên Giáo dục thể chất", "Kết hợp PP1, PP2", 1435.0, 0, 1435.0, 1470, 0.98, 2.0, 2.0, 2.0, "28 tiết/tuần + hoạt động HKPĐ, rèn luyện thể lực", "ĐẠT"],
    ["VTVL.GV.06", "Giáo viên GDQP-AN", "Kết hợp PP1, PP2", 767.5, 0, 767.5, 1470, 0.52, 1.0, 1.0, 1.0, "Đặc thù môn chuyên biệt, kiêm nhiệm công tác an ninh, bố trí 1 GV", "ĐẠT"],
    ["VTVL.GV.07", "Giáo viên Địa lí", "Kết hợp PP1, PP2", 1435.0, 0, 1435.0, 1470, 0.98, 2.0, 2.0, 2.0, "Môn lựa chọn tổ hợp KHXH, bố trí 2 GV dạy 14 lớp", "ĐẠT"],
    ["VTVL.GV.08", "Giáo viên GDKT & PL", "Kết hợp PP1, PP2", 767.5, 0, 767.5, 1470, 0.52, 1.0, 1.0, 1.0, "Môn lựa chọn tổ hợp KHXH, kiêm nhiệm tư vấn pháp luật, bố trí 1 GV", "ĐẠT"],
    ["VTVL.GV.09", "Giáo viên Vật lí", "Kết hợp PP1, PP2", 891.0, 0, 891.0, 1470, 0.61, 1.0, 1.0, 1.0, "Môn KHTN, bố trí 1 GV giảng dạy kiêm phòng thí nghiệm Lý", "ĐẠT"],
    ["VTVL.GV.10", "Giáo viên Hóa học", "Kết hợp PP1, PP2", 891.0, 0, 891.0, 1470, 0.61, 1.0, 1.0, 1.0, "Môn KHTN, bố trí 1 GV phụ trách kiêm phòng thí nghiệm Hóa", "ĐẠT"],
    ["VTVL.GV.11", "Giáo viên Sinh học", "Kết hợp PP1, PP2", 891.0, 0, 891.0, 1470, 0.61, 1.0, 1.0, 1.0, "Môn KHTN, bố trí 1 GV phụ trách kiêm phòng thí nghiệm Sinh", "ĐẠT"],
    ["VTVL.GV.12", "Giáo viên Công nghệ", "Kết hợp PP1, PP2", 1435.0, 0, 1435.0, 1470, 0.98, 2.0, 2.0, 2.0, "Gồm CN Công nghiệp và CN Nông nghiệp, bố trí 2 GV", "ĐẠT"],
    ["VTVL.GV.13", "Giáo viên Tin học", "Kết hợp PP1, PP2", 767.5, 100, 867.5, 1470, 0.59, 1.0, 1.0, 1.0, "Dạy 14 lớp + phụ trách CĐS, website, CSDL ngành, bố trí 1 GV", "ĐẠT"],
    ["VTVL.NV.01", "Kế toán", "Kết hợp PP1, PP2", 1840.0, -120, 1720.0, 1904, 0.90, 1.0, 1.0, 1.0, "Bố trí 01 người theo Thông tư 20/2023/TT-BGDĐT", "ĐẠT"],
    ["VTVL.NV.02", "Văn thư", "Kết hợp PP1, PP2", 1604.2, -110, 1494.2, 1904, 0.78, 1.0, 1.0, 1.0, "Bố trí 01 người theo Thông tư 20/2023/TT-BGDĐT", "ĐẠT"],
    ["VTVL.NV.03", "Thiết bị, thí nghiệm", "Kết hợp PP1, PP2", 1570.0, 0, 1570.0, 1904, 0.82, 1.0, 1.0, 1.0, "Bố trí 01 người theo Thông tư 20/2023/TT-BGDĐT", "ĐẠT"],
    ["VTVL.NV.04", "Thư viện viên", "Kết hợp PP1, PP2", 1650.0, 0, 1650.0, 1904, 0.87, 1.0, 1.0, 1.0, "Bố trí 01 người theo Thông tư 20/2023/TT-BGDĐT", "ĐẠT"],
    ["VTVL.NV.05", "Y tế trường học", "Kết hợp PP1, PP2", 1459.5, 0, 1459.5, 1904, 0.77, 1.0, 1.0, 1.0, "Trường xã biên giới đặc biệt khó khăn, bố trí 01 nhân viên y tế", "ĐẠT"],
    ["VTVL.HD.01", "Bảo vệ", "PP1", 2790.0, 0, 2790.0, 1904, 1.47, 1.0, 1.0, 1.0, "Hợp đồng lao động Nghị định 111/2022/NĐ-CP", "ĐẠT"],
    ["VTVL.HD.02", "Phục vụ / Tạp vụ", "PP1", 1800.0, 0, 1800.0, 1904, 0.95, 1.0, 1.0, 1.0, "Hợp đồng lao động Nghị định 111/2022/NĐ-CP", "ĐẠT"]
]

for row_idx, r_data in enumerate(c3_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_c3.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        if col_idx in [4, 5, 6, 7]:
            cell.alignment = align_right
            cell.number_format = "#,##0.0" if isinstance(val, float) else "#,##0"
        elif col_idx in [8]:
            cell.alignment = align_right
            cell.number_format = "0.00"
            cell.font = font_bold
        elif col_idx in [9, 10, 11]:
            cell.alignment = align_right
            cell.number_format = "0.0"
            cell.font = font_bold
        elif col_idx in [1, 3, 13]:
            cell.alignment = align_center
            if col_idx == 13:
                cell.font = Font(name="Times New Roman", size=10, bold=True, color=STATUS_GREEN)
        else:
            cell.alignment = align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

# Add summary row to C3
last_r = len(c3_data) + 3
ws_c3.cell(row=last_r, column=2, value="TỔNG CỘNG").font = font_bold
ws_c3.cell(row=last_r, column=6, value=f"=SUM(F3:F{last_r-1})").font = font_bold
ws_c3.cell(row=last_r, column=8, value=f"=SUM(H3:H{last_r-1})").font = font_bold
ws_c3.cell(row=last_r, column=11, value=f"=SUM(K3:K{last_r-1})").font = font_bold
for c in range(1, len(c3_headers) + 1):
    ws_c3.cell(row=last_r, column=c).fill = fill_light
    ws_c3.cell(row=last_r, column=c).border = thick_bottom

auto_fit_columns(ws_c3)

# ==========================================
# SHEET 8: 9. D1_NhanLuc
# ==========================================
ws_d1 = wb.create_sheet(title="9. D1_NhanLuc")
ws_d1.views.sheetView[0].showGridLines = True
ws_d1["A1"] = "Biểu D1_ Hiện trạng nhân lực Trường THPT Phục Hòa (Chốt ngày 30/9/2026)"
ws_d1["A1"].font = font_title
ws_d1["H1"] = "→ Tổng biên chế: 35 người (33 VC + 02 HĐ 111)"
ws_d1["H1"].font = font_note

d1_headers = [
    "STT", "Mã nhân sự", "Họ và tên", "Năm sinh", "Dân tộc", 
    "Mã VTVL", "Vị trí việc làm / Bộ môn", "Trình độ ĐT", 
    "Bậc lương/ngạch", "Giờ dành cho VTVL", "Nơi tính biên chế", "Ghi chú kiêm nhiệm"
]
for col_idx, h in enumerate(d1_headers, start=1):
    ws_d1.cell(row=2, column=col_idx, value=h)
style_header_row(ws_d1, 2, fill_navy)

d1_data = [
    # BGH
    [1, "NL.PH.01", "Hoàng Trung Kiên", 1980, "Tày", "VTVL.QL.01", "Hiệu trưởng", "Thạc sĩ Quản lý GD", "Bậc 4", 1904, "Trường THPT Phục Hòa", "Bí thư Chi bộ"],
    [2, "NL.PH.02", "Đinh Thị Toan", 1982, "Tày", "VTVL.QL.02", "Phó Hiệu trưởng", "Đại học Sư phạm", "Bậc 3", 1904, "Trường THPT Phục Hòa", "Phụ trách chuyên môn"],
    [3, "NL.PH.03", "Phan Văn Thế", 1985, "Tày", "VTVL.QL.02", "Phó Hiệu trưởng", "Đại học Sư phạm", "Bậc 3", 1904, "Trường THPT Phục Hòa", "Phụ trách CSVC, nề nếp"],
    
    # Giáo viên
    [4, "NL.PH.04", "Nông Thị Bích Hảo", 1976, "Tày", "VTVL.GV.03", "Giáo viên Tiếng Anh", "Đại học SP Tiếng Anh", "Bậc 4", 1470, "Trường THPT Phục Hòa", "Tổ trưởng Tổ Xã hội"],
    [5, "NL.PH.05", "Trương Thị Duyên", 1976, "Nùng", "VTVL.GV.02", "Giáo viên Toán", "Đại học SP Toán", "Bậc 4", 1470, "Trường THPT Phục Hòa", "Tổ phó Tổ Tự nhiên"],
    [6, "NL.PH.06", "Hoàng Thị Huệ", 1974, "Tày", "VTVL.GV.04", "Giáo viên Lịch sử", "Đại học SP Lịch sử", "Bậc 4", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [7, "NL.PH.07", "Đàm Thị Thanh Hà", 1978, "Tày", "VTVL.GV.09", "Giáo viên Vật lí", "Đại học SP Vật lí", "Bậc 4", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [8, "NL.PH.08", "Ngọc Bích Diệp", 1978, "Nùng", "VTVL.GV.01", "Giáo viên Ngữ văn", "Đại học SP Ngữ văn", "Bậc 4", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [9, "NL.PH.09", "Nông Thị Đông", 1980, "Tày", "VTVL.GV.03", "Giáo viên Tiếng Anh", "Đại học SP Tiếng Anh", "Bậc 3", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [10, "NL.PH.10", "Nông Hoài Thu", 1982, "Tày", "VTVL.GV.01", "Giáo viên Ngữ văn", "Đại học SP Ngữ văn", "Bậc 3", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [11, "NL.PH.11", "Lê Tuấn Vương", 1981, "Kinh", "VTVL.GV.12", "Giáo viên Công nghệ (KTNN)", "Đại học Sư phạm", "Bậc 3", 1470, "Trường THPT Phục Hòa", "Phụ trách thiết bị CN, GVCN"],
    [12, "NL.PH.12", "Hoàng Minh Tuấn", 1981, "Tày", "VTVL.GV.10", "Giáo viên Hóa học", "Đại học SP Hóa học", "Bậc 3", 1470, "Trường THPT Phục Hòa", "Tổ trưởng Tổ Tự nhiên"],
    [13, "NL.PH.13", "Dương Minh Toản", 1982, "Kinh", "VTVL.GV.05", "Giáo viên Giáo dục thể chất", "Đại học SP Thể chất", "Bậc 3", 1470, "Trường THPT Phục Hòa", "Chủ tịch Công đoàn"],
    [14, "NL.PH.14", "Đàm Thị Hợp", 1986, "Tày", "VTVL.GV.13", "Giáo viên Tin học", "Đại học SP Tin học", "Bậc 3", 1470, "Trường THPT Phục Hòa", "Phụ trách CNTT, CĐS, GVCN"],
    [15, "NL.PH.15", "Phạm Viết Biên", 1984, "Mường", "VTVL.GV.02", "Giáo viên Toán", "Đại học SP Toán", "Bậc 3", 1470, "Trường THPT Phục Hòa", "Bồi dưỡng HSG Toán"],
    [16, "NL.PH.16", "Nông Văn Chiến", 1990, "Tày", "VTVL.GV.02", "Giáo viên Toán", "Đại học SP Toán", "Bậc 2", 1470, "Trường THPT Phục Hòa", "Bí thư Đoàn trường"],
    [17, "NL.PH.17", "Ma Đức Thiện", 1994, "Tày", "VTVL.GV.08", "Giáo viên GDKT & PL", "Đại học Sư phạm", "Bậc 2", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [18, "NL.PH.18", "Đinh Tố Loan", 1982, "Tày", "VTVL.GV.03", "Giáo viên Tiếng Anh", "Đại học SP Tiếng Anh", "Bậc 3", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [19, "NL.PH.19", "Nông Văn Đình", 1981, "Tày", "VTVL.GV.05", "Giáo viên Giáo dục thể chất", "Đại học SP Thể chất", "Bậc 3", 1470, "Trường THPT Phục Hòa", "Huấn luyện HKPĐ"],
    [20, "NL.PH.20", "Hoàng Thị Thùy Dung", 1986, "Tày", "VTVL.GV.12", "Giáo viên Công nghệ (KTCN)", "Đại học Sư phạm", "Bậc 3", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [21, "NL.PH.21", "Ngọc Hữu Nghị", 1983, "Nùng", "VTVL.GV.02", "Giáo viên Toán", "Đại học SP Toán", "Bậc 3", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [22, "NL.PH.22", "Nông Quốc Tuệ", 1970, "Tày", "VTVL.GV.06", "Giáo viên GDQP-AN", "Đại học Sư phạm", "Bậc 4", 1470, "Trường THPT Phục Hòa", "Phụ trách phòng GDQP"],
    [23, "NL.PH.23", "Nguyễn Văn Sắc", 1983, "Tày", "VTVL.GV.07", "Giáo viên Địa lí", "Đại học SP Địa lí", "Bậc 3", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [24, "NL.PH.24", "Đoàn Bích Liên", 1984, "Tày", "VTVL.GV.01", "Giáo viên Ngữ văn", "Đại học SP Ngữ văn", "Bậc 3", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [25, "NL.PH.25", "Đàm Thị Quyên", 1973, "Tày", "VTVL.GV.04", "Giáo viên Lịch sử", "Đại học SP Lịch sử", "Bậc 4", 1470, "Trường THPT Phục Hòa", "Tổ phó Tổ Xã hội"],
    [26, "NL.PH.26", "Hoàng Thị Hồng", 1987, "Tày", "VTVL.GV.07", "Giáo viên Địa lí", "Đại học SP Địa lí", "Bậc 3", 1470, "Trường THPT Phục Hòa", "GVCN"],
    [27, "NL.PH.27", "Nông Thị Biển", 1990, "Tày", "VTVL.GV.11", "Giáo viên Sinh học", "Đại học SP Sinh học", "Bậc 2", 1470, "Trường THPT Phục Hòa", "Bồi dưỡng HSG KHTN"],

    # Nhân viên hành chính
    [28, "NL.PH.28", "Bế Thị Điều", 1985, "Tày", "VTVL.NV.01", "Kế toán", "Đại học Tài chính Kế toán", "Bậc 3", 1904, "Trường THPT Phục Hòa", "Kế toán trưởng đơn vị"],
    [29, "NL.PH.29", "Đàm Thu Hà", 1988, "Nùng", "VTVL.NV.02", "Văn thư", "Cao đẳng Văn thư lưu trữ", "Bậc 3", 1904, "Trường THPT Phục Hòa", "Quản trị iOffice"],
    [30, "NL.PH.30", "Mông Thị Thu", 1985, "Nùng", "VTVL.NV.04", "Thư viện viên", "Đại học Thông tin Thư viện", "Bậc 3", 1904, "Trường THPT Phục Hòa", "Phụ trách tủ sách SGK"],
    [31, "NL.PH.31", "Nông Thị Hương", 1984, "Tày", "VTVL.NV.03", "Thiết bị, thí nghiệm", "Cao đẳng Sư phạm Kỹ thuật", "Bậc 3", 1904, "Trường THPT Phục Hòa", "Phụ trách kho thiết bị"],
    [32, "NL.PH.32", "Nông Văn Kiên (HĐ)", 1992, "Tày", "VTVL.NV.05", "Y tế trường học", "Cao đẳng Y tế", "Bậc 1", 1904, "Trường THPT Phục Hòa", "Kiêm công tác Chữ thập đỏ"],
    [33, "NL.PH.33", "Nông Thị Lan (HĐ)", 1993, "Nùng", "VTVL.NV.06", "Giáo vụ trường học", "Cao đẳng Sư phạm", "Bậc 1", 1904, "Trường THPT Phục Hòa", "Kiêm quản lý nề nếp"],

    # HĐLĐ 111
    [34, "NL.PH.34", "Đàm Quang Long", 1978, "Tày", "VTVL.HD.01", "Bảo vệ trường học", "Trung cấp", "HĐ 111", 1904, "Trường THPT Phục Hòa", "Trực bảo vệ 24/24"],
    [35, "NL.PH.35", "Lý Thị Mỵ", 1982, "Tày", "VTVL.HD.02", "Nhân viên phục vụ, tạp vụ", "Sơ cấp", "HĐ 111", 1904, "Trường THPT Phục Hòa", "Vệ sinh khuôn viên"]
]

for row_idx, r_data in enumerate(d1_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_d1.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        if col_idx in [1, 2, 4, 5, 6, 9]:
            cell.alignment = align_center
        elif col_idx == 10:
            cell.alignment = align_right
            cell.number_format = "#,##0"
        else:
            cell.alignment = align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

auto_fit_columns(ws_d1)

# ==========================================
# SHEET 9: 10. D2_PhuongAn
# ==========================================
ws_d2 = wb.create_sheet(title="10. D2_PhuongAn")
ws_d2.views.sheetView[0].showGridLines = True
ws_d2["A1"] = "Biểu D2_ Phương án bố trí nhân lực và biên chế giai đoạn 2027-2031 Trường THPT Phục Hòa"
ws_d2["A1"].font = font_title

d2_headers = [
    "Mã VTVL", "Tên VTVL", "Năm", "Số người cần bố trí", 
    "Biên chế hiện có", "Đề nghị điều chỉnh (+/-)", "Hình thức tuyển dụng / Bố trí", 
    "Nguồn kinh phí chi trả", "Trạng thái", "Giải trình căn cứ"
]
for col_idx, h in enumerate(d2_headers, start=1):
    ws_d2.cell(row=2, column=col_idx, value=h)
style_header_row(ws_d2, 2, fill_navy)

d2_data = [
    ["VTVL.QL.01", "Hiệu trưởng", 2027, 1, 1, 0, "Giữ nguyên", "NSNN cấp", "ĐẠT", "01 Hiệu trưởng theo quy định trường hạng II"],
    ["VTVL.QL.02", "Phó Hiệu trưởng", 2027, 2, 2, 0, "Giữ nguyên", "NSNN cấp", "ĐẠT", "02 Phó Hiệu trưởng theo Thông tư 20/2023"],
    ["VTVL.GV.ALL", "Giáo viên (Toàn trường)", 2027, 26, 24, 2, "Tuyển dụng bổ sung 02 GV", "NSNN cấp", "ĐẠT", "Trường tăng lên 15 lớp, định mức 2,45 GV/lớp (Vùng 1 đặc biệt khó khăn)"],
    ["VTVL.NV.ALL", "Nhân viên hỗ trợ, phục vụ", 2027, 6, 5, 1, "Bổ sung biên chế Y tế chuyên trách", "NSNN cấp", "ĐẠT", "Bảo đảm đủ 06 vị trí nhân viên theo TT 20/2023"],
    ["VTVL.HD.ALL", "Hợp đồng theo NĐ 111", 2027, 2, 2, 0, "Tiếp tục ký HĐLĐ 111", "NSNN hỗ trợ", "ĐẠT", "01 Bảo vệ + 01 Phục vụ tạp vụ"],
    
    # 2028
    ["VTVL.TONG", "Tổng số toàn trường", 2028, 37, 37, 0, "Ổn định quy mô 15 lớp", "NSNN cấp", "ĐẠT", "Duy trì đủ 35 biên chế viên chức + 02 HĐ 111"],
    
    # 2029 (Tăng lên 16 lớp)
    ["VTVL.GV.ALL", "Giáo viên (Toàn trường)", 2029, 28, 26, 2, "Tuyển dụng bổ sung 02 GV (Toán, Tin)", "NSNN cấp", "ĐẠT", "Quy mô tăng lên 16 lớp, 640 HS. Nhu cầu định mức đạt 39 người"],
    ["VTVL.TONG", "Tổng số toàn trường", 2029, 39, 37, 2, "Đề xuất giao 37 BC viên chức", "NSNN cấp", "ĐẠT", "Đảm bảo tỷ lệ 2,4 GV/lớp và đủ người trực tiếp giảng dạy"],

    # 2030
    ["VTVL.TONG", "Tổng số toàn trường", 2030, 39, 39, 0, "Ổn định quy mô 16 lớp", "NSNN cấp", "ĐẠT", "Duy trì 37 biên chế VC + 02 HĐLĐ 111"],

    # 2031
    ["VTVL.TONG", "Tổng số toàn trường", 2031, 39, 39, 0, "Hoàn thành quy hoạch 2027-2031", "NSNN cấp", "ĐẠT", "Bộ máy tinh gọn, hiệu quả, 100% đạt chuẩn và trên chuẩn"]
]

for row_idx, r_data in enumerate(d2_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_d2.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        if col_idx in [3, 4, 5, 6]:
            cell.alignment = align_right
            cell.number_format = "#,##0"
        elif col_idx in [1, 9]:
            cell.alignment = align_center
            if col_idx == 9:
                cell.font = Font(name="Times New Roman", size=10, bold=True, color=STATUS_GREEN)
        else:
            cell.alignment = align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

auto_fit_columns(ws_d2)

# ==========================================
# SHEET 10: DU_BAO_5_NAM
# ==========================================
ws_db = wb.create_sheet(title="DU_BAO_5_NAM")
ws_db.views.sheetView[0].showGridLines = True
ws_db["A1"] = "BẢNG DỰ BÁO QUY MÔ TRƯỜNG LỚP VÀ NHU CẦU BIÊN CHẾ 5 NĂM (2027 - 2031)"
ws_db["A1"].font = font_title

db_headers = [
    "Năm học", "Quy mô số lớp", "Số học sinh", "Tỷ lệ GV/lớp định mức", 
    "Nhu cầu GV định mức", "Lãnh đạo quản lý", "Nhân viên hỗ trợ", 
    "HĐLĐ 111", "Tổng nhu cầu nhân lực", "Biên chế VC đề xuất", "Ghi chú tiến độ"
]
for col_idx, h in enumerate(db_headers, start=1):
    ws_db.cell(row=2, column=col_idx, value=h)
style_header_row(ws_db, 2, fill_navy)

db_data = [
    ["2026 - 2027 (Hiện tại)", 14, 573, "Tối đa 2.45", 34.3, 3, 6, 2, 35, 33, "Chốt 30/9/2026: 33 VC + 02 HĐ 111"],
    ["2027 - 2028 (Dự kiến)", 15, 600, "Tối đa 2.45", 36.7, 3, 6, 2, 37, 35, "Tăng 01 lớp 10, đề xuất bổ sung 02 BC"],
    ["2028 - 2029 (Dự kiến)", 15, 610, "Tối đa 2.45", 36.7, 3, 6, 2, 37, 35, "Duy trì ổn định quy mô 15 lớp"],
    ["2029 - 2030 (Dự kiến)", 16, 640, "Tối đa 2.45", 39.2, 3, 6, 2, 39, 37, "Tăng lên 16 lớp, đề xuất bổ sung 02 BC"],
    ["2030 - 2031 (Dự kiến)", 16, 650, "Tối đa 2.45", 39.2, 3, 6, 2, 39, 37, "Quy mô ổn định 16 lớp chuẩn quốc gia"]
]

for row_idx, r_data in enumerate(db_data, start=3):
    for col_idx, val in enumerate(r_data, start=1):
        cell = ws_db.cell(row=row_idx, column=col_idx, value=val)
        cell.font = font_normal
        cell.border = thin_border
        if col_idx in [2, 3, 5, 6, 7, 8, 9, 10]:
            cell.alignment = align_right
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0.0" if isinstance(val, float) else "#,##0"
        elif col_idx in [1, 4]:
            cell.alignment = align_center
            if col_idx == 1:
                cell.font = font_bold
        else:
            cell.alignment = align_left
        if row_idx % 2 == 1:
            cell.fill = fill_gray

auto_fit_columns(ws_db)

# SAVE FILE
desktop_path = r"D:\Desktop\Bo-Bieu-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.xlsx"
assets_path = r"d:\Du-an-web\web-nhu-cau-nhan-luc-bien-che-2027-2031-phuc-hoa\assets"
os.makedirs(assets_path, exist_ok=True)
project_excel_path = os.path.join(assets_path, "Bo-Bieu-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.xlsx")

wb.save(project_excel_path)
print(f"Excel saved to project assets: {project_excel_path}")

try:
    wb.save(desktop_path)
    print(f"Excel also saved to Desktop: {desktop_path}")
except Exception as e:
    print(f"Could not save to desktop: {e}")
