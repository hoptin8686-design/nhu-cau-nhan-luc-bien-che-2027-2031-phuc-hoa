# -*- coding: utf-8 -*-
"""
Script tự động tạo Đề án Xác định nhu cầu nhân lực và xây dựng phương án, 
biên chế giai đoạn 2027-2031 của Trường THPT Phục Hòa định dạng Word (.docx) chuẩn.
"""
import os
import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')

doc = Document()

# Set page margins to standard A4 (2cm left/right/top/bottom, left 2.5cm)
sections = doc.sections
for s in sections:
    s.top_margin = Inches(0.79) # 2cm
    s.bottom_margin = Inches(0.79)
    s.left_margin = Inches(0.98) # 2.5cm
    s.right_margin = Inches(0.79) # 2cm

# Color constants
COLOR_NAVY = RGBColor(27, 54, 93)
COLOR_TEXT = RGBColor(30, 41, 59)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_title(text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.bold = True
    if level == 1:
        run.font.size = Pt(14)
        run.font.color.rgb = COLOR_NAVY
    elif level == 2:
        run.font.size = Pt(13)
        run.font.color.rgb = COLOR_NAVY
    else:
        run.font.size = Pt(12)
    return p

def add_para(text, bold_prefix="", italic=False, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = "Times New Roman"
        r_pre.font.size = Pt(11.5)
        r_pre.font.bold = True
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(11.5)
    r.font.italic = italic
    return p

# ---------------------------------------------
# PHẦN 1: TỜ TRÌNH
# ---------------------------------------------
p_header = doc.add_table(rows=1, cols=2)
p_header.alignment = WD_TABLE_ALIGNMENT.CENTER
p_header.autofit = False
p_header.columns[0].width = Inches(3.2)
p_header.columns[1].width = Inches(3.5)

cell_left = p_header.cell(0, 0)
cell_right = p_header.cell(0, 1)

p_l = cell_left.paragraphs[0]
p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_l.add_run("SỞ GIÁO DỤC VÀ ĐÀO TẠO CAO BẰNG\n")
r.font.name = "Times New Roman"
r.font.size = Pt(10)
r = p_l.add_run("TRƯỜNG THPT PHỤC HÒA\n")
r.font.name = "Times New Roman"
r.font.size = Pt(11)
r.font.bold = True
r = p_l.add_run("Số: 26/TTr-THPT")
r.font.name = "Times New Roman"
r.font.size = Pt(10)
r.font.italic = True

p_r = cell_right.paragraphs[0]
p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_r.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
r.font.name = "Times New Roman"
r.font.size = Pt(10)
r.font.bold = True
r = p_r.add_run("Độc lập - Tự do - Hạnh phúc\n")
r.font.name = "Times New Roman"
r.font.size = Pt(11)
r.font.bold = True
r = p_r.add_run("Quảng Hòa, ngày 28 tháng 10 năm 2026")
r.font.name = "Times New Roman"
r.font.size = Pt(10)
r.font.italic = True

p_space = doc.add_paragraph()
p_space.paragraph_format.space_after = Pt(10)

p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_t = p_title.add_run("TỜ TRÌNH\n")
r_t.font.name = "Times New Roman"
r_t.font.size = Pt(14)
r_t.font.bold = True
r_sub = p_title.add_run("Về việc đề nghị thẩm định và phê duyệt Đề án xác định nhu cầu nhân lực\nvà phương án biên chế giai đoạn 2027 - 2031 của Trường THPT Phục Hòa")
r_sub.font.name = "Times New Roman"
r_sub.font.size = Pt(12)
r_sub.font.bold = True

p_kinhgui = doc.add_paragraph()
p_kinhgui.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_kg = p_kinhgui.add_run("Kính gửi: ")
r_kg.font.name = "Times New Roman"
r_kg.font.size = Pt(11.5)
r_kg.font.bold = True
r_kg_dest = p_kinhgui.add_run("Sở Giáo dục và Đào tạo tỉnh Cao Bằng")
r_kg_dest.font.name = "Times New Roman"
r_kg_dest.font.size = Pt(11.5)
r_kg_dest.font.bold = True

add_para("Căn cứ Kết luận số 107-KL/TW ngày 17/10/2026 của Bộ Chính trị về hoàn thiện vị trí việc làm và biên chế;", italic=True)
add_para("Căn cứ Hướng dẫn số 07-HD/BTCTW ngày 19/10/2026 của Ban Tổ chức Trung ương về rà soát, hoàn thiện vị trí việc làm, xác định nhu cầu nhân lực và xây dựng phương án biên chế giai đoạn 2027 - 2031;", italic=True)
add_para("Căn cứ Công văn số 12176/BNV-CCVC ngày 23/10/2026 của Bộ Nội vụ về triển khai rà soát VTVL và phương án biên chế;", italic=True)
add_para("Căn cứ Kế hoạch số 201-KH/TU ngày 23/10/2026 của Tỉnh ủy Cao Bằng và Công văn số 5406/SNV-TCBM&CCVC ngày 26/10/2026 của Sở Nội vụ tỉnh Cao Bằng;", italic=True)
add_para("Căn cứ Công văn số /SGDĐT-TCCB ngày 27/10/2026 của Sở Giáo dục và Đào tạo tỉnh Cao Bằng về việc rà soát, hoàn thiện vị trí việc làm, xác định nhu cầu nhân lực và xây dựng phương án biên chế giai đoạn 2027 - 2031;", italic=True)
add_para("Căn cứ Thông tư số 20/2023/TT-BGDĐT ngày 30/10/2023 của Bộ Giáo dục và Đào tạo hướng dẫn về vị trí việc làm, cơ cấu viên chức theo chức danh nghề nghiệp và định mức số lượng người làm việc trong các cơ sở giáo dục phổ thông công lập;", italic=True)

add_para("Trường THPT Phục Hòa kính trình Sở Giáo dục và Đào tạo tỉnh Cao Bằng xem xét, thẩm định và phê duyệt Đề án xác định nhu cầu nhân lực và phương án biên chế giai đoạn 2027 - 2031 của nhà trường với các nội dung trọng tâm như sau:")

add_para("1. Đặc điểm tình hình: Trường THPT Phục Hòa đóng trên địa bàn thị trấn Phục Hòa, huyện Quảng Hòa, tỉnh Cao Bằng (khu vực xã biên giới, thuộc Vùng 1 có điều kiện kinh tế - xã hội đặc biệt khó khăn theo Thông tư 20/2023/TT-BGDĐT). Năm học 2026-2027, nhà trường có quy mô 14 lớp với 573 học sinh. Tổng số biên chế viên chức được giao là 35 người (hiện có 33 viên chức biên chế và 02 hợp đồng lao động theo Nghị định 111/2022/NĐ-CP).")

add_para("2. Kết quả xác định nhu cầu nhân lực: Áp dụng phương pháp PP1, PP2, PP3 theo Hướng dẫn số 07-HD/BTCTW:")
add_para("- Tổng số giờ lao động chuyên môn ròng toàn trường: 55.850 giờ/năm.")
add_para("- Số lao động quy đổi theo PP2 (FTE): 37,99 người.")
add_para("- Nhu cầu nhân lực cần bố trí: 37 người (giai đoạn 2027-2028 quy mô 15 lớp) và 39 người (giai đoạn 2029-2031 khi quy mô đạt 16 lớp).")

add_para("3. Phương án biên chế giai đoạn 2027 - 2031 đề xuất:")
add_para("- Năm 2027: Đề nghị giao 35 biên chế viên chức + 02 HĐLĐ 111 (bổ sung 02 biên chế giáo viên trực tiếp giảng dạy).")
add_para("- Năm 2028: Duy trì 35 biên chế viên chức + 02 HĐLĐ 111.")
add_para("- Năm 2029: Đề nghị giao 37 biên chế viên chức + 02 HĐLĐ 111 (bổ sung 02 biên chế giáo viên khi quy mô trường đạt 16 lớp).")
add_para("- Năm 2030 - 2031: Duy trì ổn định 37 biên chế viên chức + 02 HĐLĐ 111.")

add_para("Hồ sơ gửi kèm Tờ trình gồm có:")
add_para("1. Bản thuyết minh Đề án xác định nhu cầu nhân lực và phương án biên chế giai đoạn 2027-2031.")
add_para("2. Trọn bộ 09 Biểu mẫu nghiệp vụ Excel theo Hướng dẫn số 07-HD/BTCTW (A1, A2, B1, B2, C1, C2, C3, D1, D2).")
add_para("3. Bản mô tả công việc và Khung năng lực từng vị trí việc làm.")

add_para("Kính đề nghị Sở Giáo dục và Đào tạo tỉnh Cao Bằng xem xét, thẩm định và phê duyệt để nhà trường có căn cứ triển khai thực hiện./.")

# Chữ ký tờ trình
t_sign = doc.add_table(rows=1, cols=2)
t_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
t_sign.columns[0].width = Inches(3.2)
t_sign.columns[1].width = Inches(3.5)

c_s_l = t_sign.cell(0, 0).paragraphs[0]
r = c_s_l.add_run("Nơi nhận:\n")
r.font.name = "Times New Roman"
r.font.size = Pt(9.5)
r.font.bold = True
r = c_s_l.add_run("- Như kính gửi;\n- Lãnh đạo Sở (báo cáo);\n- Phòng TCCB Sở;\n- Chi bộ, BGH trường;\n- Lưu: VT, HS VTVL.")
r.font.name = "Times New Roman"
r.font.size = Pt(9)
r.font.italic = True

c_s_r = t_sign.cell(0, 1).paragraphs[0]
c_s_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = c_s_r.add_run("HIỆU TRƯỞNG\n\n\n\n\n")
r.font.name = "Times New Roman"
r.font.size = Pt(11.5)
r.font.bold = True
r = c_s_r.add_run("Hoàng Trung Kiên")
r.font.name = "Times New Roman"
r.font.size = Pt(12)
r.font.bold = True

doc.add_page_break()

# ---------------------------------------------
# PHẦN 2: THUYẾT MINH ĐỀ ÁN CHI TIẾT
# ---------------------------------------------
p_de_an_title = doc.add_paragraph()
p_de_an_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_de_an_title.add_run("ĐỀ ÁN\nXÁC ĐỊNH NHU CẦU NHÂN LỰC VÀ XÂY DỰNG PHƯƠNG ÁN BIÊN CHẾ\nGIAI ĐOẠN 2027 - 2031 CỦA TRƯỜNG THPT PHỤC HÒA\n")
r.font.name = "Times New Roman"
r.font.size = Pt(14)
r.font.bold = True
r.font.color.rgb = COLOR_NAVY

add_title("PHẦN I: CĂN CỨ PHÁP LÝ VÀ SỰ CẦN THIẾT XÂY DỰNG ĐỀ ÁN", 1)
add_title("1. Căn cứ pháp lý", 2)
add_para("- Kết luận số 107-KL/TW ngày 17/10/2026 của Bộ Chính trị về hoàn thiện vị trí việc làm và biên chế trong hệ thống chính trị;")
add_para("- Hướng dẫn số 07-HD/BTCTW ngày 19/10/2026 của Ban Tổ chức Trung ương về rà soát, hoàn thiện vị trí việc làm, xác định nhu cầu nhân lực và xây dựng phương án biên chế giai đoạn 2027 - 2031;")
add_para("- Công văn số 12176/BNV-CCVC ngày 23/10/2026 của Bộ Nội vụ về việc rà soát hoàn thiện hệ thống vị trí việc làm và phương án biên chế;")
add_para("- Nghị định số 232/2026/NĐ-CP ngày 26/6/2026 của Chính phủ quy định về vị trí việc làm của viên chức;")
add_para("- Nghị định số 111/2022/NĐ-CP ngày 30/12/2022 của Chính phủ về hợp đồng đối với một số loại công việc trong cơ quan hành chính và đơn vị sự nghiệp công lập;")
add_para("- Thông tư số 20/2023/TT-BGDĐT ngày 30/10/2023 của Bộ Giáo dục và Đào tạo hướng dẫn về vị trí việc làm, cơ cấu viên chức theo chức danh nghề nghiệp và định mức số lượng người làm việc trong các cơ sở giáo dục phổ thông công lập;")
add_para("- Kế hoạch số 201-KH/TU ngày 23/10/2026 của Tỉnh ủy Cao Bằng; Công văn số 4142/UBND-NC ngày 25/10/2026 của UBND tỉnh Cao Bằng; Công văn số 5406/SNV-TCBM&CCVC ngày 26/10/2026 của Sở Nội vụ tỉnh Cao Bằng;")
add_para("- Công văn số /SGDĐT-TCCB ngày 27/10/2026 của Sở Giáo dục và Đào tạo tỉnh Cao Bằng về việc rà soát, hoàn thiện vị trí việc làm, xác định nhu cầu nhân lực và xây dựng phương án biên chế giai đoạn 2027 - 2031 (thời hạn nộp báo cáo ngày 30/10/2026).")

add_title("2. Sự cần thiết xây dựng Đề án", 2)
add_para("Trường THPT Phục Hòa là đơn vị sự nghiệp công lập trực thuộc Sở GD&ĐT tỉnh Cao Bằng, thực hiện nhiệm vụ giáo dục trung học phổ thông cho con em đồng bào các dân tộc thuộc địa bàn huyện Quảng Hòa và khu vực biên giới lân cận. Trong bối cảnh triển khai toàn diện Chương trình Giáo dục phổ thông 2018 với định hướng phân hóa sâu, dạy học môn lựa chọn, chuyên đề học tập và tăng cường chuyển đổi số, khối lượng công việc của đội ngũ cán bộ quản lý, giáo viên và nhân viên tăng lên đáng kể. Việc rà soát, xác định nhu cầu nhân lực theo phương pháp khoa học của Hướng dẫn 07-HD/BTCTW là cơ sở vững chắc để xây dựng phương án biên chế 5 năm 2027-2031 sát thực tiễn, đáp ứng yêu cầu nâng cao chất lượng giáo dục vùng khó khăn.")

add_title("PHẦN II: THỰC TRẠNG TỔ CHỨC BỘ MÁY, BIÊN CHẾ VÀ NHÂN LỰC HIỆN CÓ (CHỐT 30/9/2026)", 1)
add_title("1. Khái quát quy mô trường lớp, học sinh", 2)
add_para("- Vị trí địa lý: Thị trấn Phục Hòa, huyện Quảng Hòa, tỉnh Cao Bằng (địa bàn xã biên giới, thuộc Vùng 1 theo phân loại tại Thông tư 20/2023/TT-BGDĐT).")
add_para("- Quy mô năm học 2026-2027: Tổng số 14 lớp với 573 học sinh (Khối 10: 5 lớp - 212 HS; Khối 11: 5 lớp - 198 HS; Khối 12: 4 lớp - 163 HS). Bình quân 41 học sinh/lớp.")
add_para("- Cơ sở vật chất: Đang được đầu tư kiên cố hóa, có 14 phòng học văn hóa, phòng thực hành bộ môn (Vật lí, Hóa học, Sinh học, Tin học, Ngoại ngữ), thư viện trường học và khu giáo dục thể chất.")

add_title("2. Thực trạng biên chế và nhân lực hiện có", 2)
add_para("Tổng chỉ tiêu biên chế được cơ quan có thẩm quyền giao: 35 người. Hiện trạng có mặt đến ngày 30/9/2026: 35 người, gồm:")
add_para("a) Viên chức quản lý (03 người):", bold_prefix="• ")
add_para("- 01 Hiệu trưởng: Thầy Hoàng Trung Kiên (Thạc sĩ Quản lý giáo dục, Bậc 4).")
add_para("- 02 Phó Hiệu trưởng: Cô Đinh Thị Toan (Đại học SP, Bậc 3, phụ trách chuyên môn); Thầy Phan Văn Thế (Đại học SP, Bậc 3, phụ trách CSVC, nề nếp).")

add_para("b) Viên chức chuyên môn, nghiệp vụ - Giáo viên trực tiếp giảng dạy (24 người):", bold_prefix="• ")
add_para("- Môn Ngữ văn: 03 giáo viên (Cô Ngọc Bích Diệp, Cô Nông Hoài Thu, Cô Đoàn Bích Liên).")
add_para("- Môn Toán: 04 giáo viên (Cô Trương Thị Duyên, Thầy Phạm Viết Biên, Thầy Nông Văn Chiến, Thầy Ngọc Hữu Nghị).")
add_para("- Môn Tiếng Anh: 03 giáo viên (Cô Nông Thị Bích Hảo, Cô Nông Thị Đông, Cô Đinh Tố Loan).")
add_para("- Môn Lịch sử: 02 giáo viên (Cô Hoàng Thị Huệ, Cô Đàm Thị Quyên).")
add_para("- Môn Giáo dục thể chất: 02 giáo viên (Thầy Dương Minh Toản, Thầy Nông Văn Đình).")
add_para("- Môn GDQP-AN: 01 giáo viên (Thầy Nông Quốc Tuệ).")
add_para("- Môn Địa lí: 02 giáo viên (Thầy Nguyễn Văn Sắc, Cô Hoàng Thị Hồng).")
add_para("- Môn Giáo dục Kinh tế & Pháp luật: 01 giáo viên (Thầy Ma Đức Thiện).")
add_para("- Môn Vật lí: 01 giáo viên (Cô Đàm Thị Thanh Hà).")
add_para("- Môn Hóa học: 01 giáo viên (Thầy Hoàng Minh Tuấn).")
add_para("- Môn Sinh học: 01 giáo viên (Cô Nông Thị Biển).")
add_para("- Môn Công nghệ: 02 giáo viên (Thầy Lê Tuấn Vương - CN Nông nghiệp; Cô Hoàng Thị Thùy Dung - CN Công nghiệp).")
add_para("- Môn Tin học: 01 giáo viên (Cô Đàm Thị Hợp - phụ trách chuyên môn Tin học kiêm phụ trách CNTT và CĐS).")

add_para("c) Nhân viên hỗ trợ, phục vụ (06 người):", bold_prefix="• ")
add_para("- Kế toán: 01 người (Cô Bế Thị Điều - Bậc 3).")
add_para("- Văn thư: 01 người (Cô Đàm Thu Hà - Bậc 3).")
add_para("- Thiết bị, thí nghiệm: 01 người (Cô Nông Thị Hương - Bậc 3).")
add_para("- Thư viện: 01 người (Cô Mông Thị Thu - Bậc 3).")
add_para("- Y tế học đường: 01 người (Hợp đồng chuyên môn theo vị trí).")
add_para("- Giáo vụ: 01 người (Hợp đồng hỗ trợ giáo vụ).")

add_para("d) Hợp đồng lao động theo Nghị định 111/2022/NĐ-CP (02 người):", bold_prefix="• ")
add_para("- Bảo vệ: 01 người (Chú Đàm Quang Long).")
add_para("- Phục vụ, tạp vụ: 01 người (Cô Lý Thị Mỵ).")

add_title("PHẦN III: PHƯƠNG PHÁP VÀ KẾT QUẢ XÁC ĐỊNH NHU CẦU NHÂN LỰC GIAI ĐOẠN 2027 - 2031", 1)
add_title("1. Phương pháp áp dụng theo Hướng dẫn 07-HD/BTCTW", 2)
add_para("Nhà trường thực hiện xác định nhu cầu nhân lực độc lập với số biên chế hiện có, kết hợp khoa học 03 phương pháp:")
add_para("- Phương pháp 1 (PP1 - Theo cơ cấu chức danh và định mức): Áp dụng định mức theo Thông tư 20/2023/TT-BGDĐT. Đối với cấp THPT tại Vùng 1 (xã biên giới đặc biệt khó khăn), định mức tối đa là 2,45 giáo viên/lớp; Hiệu trưởng 01 người; Phó Hiệu trưởng tối đa 02 người (đối với trường từ 10 đến 17 lớp); nhân viên hỗ trợ 06 người.")
add_para("- Phương pháp 2 (PP2 - Theo khối lượng công việc và quỹ giờ quy đổi): Thống kê sản phẩm đầu ra trong 12 tháng (01/10/2025 - 30/9/2026), tính số giờ sản phẩm, cộng giờ công việc bổ trợ, trừ giờ giảm do chuyển đổi số, chia cho Quỹ giờ chuẩn (1.904 giờ/năm đối với hành chính và 1.470 giờ/năm đối với giáo viên THPT theo chế độ 42 tuần).")
add_para("- Phương pháp 3 (PP3 - Ước lượng chuyên gia và yêu cầu chất lượng): Đánh giá đặc thù phân môn chuyên sâu, yêu cầu đảm bảo 100% các môn học đều có giáo viên cơ hữu giảng dạy và ôn thi tốt nghiệp THPT.")

add_title("2. Kết quả tính toán nhu cầu nhân lực (Biểu C3)", 2)
add_para("- Tổng quỹ giờ lao động ròng của toàn đơn vị: 55.850 giờ lao động chuyên môn/năm.")
add_para("- Số người quy đổi theo PP2 (FTE): 37,99 người.")
add_para("- Kết luận số người cần bố trí năm cơ sở 2027: 37 người (gồm 35 viên chức biên chế và 02 hợp đồng lao động theo Nghị định 111/2022/NĐ-CP).")

add_title("PHẦN IV: PHƯƠNG ÁN BIÊN CHẾ GIAI ĐOẠN 2027 - 2031 VÀ GIẢI PHÁP THỰC HIỆN", 1)
add_title("1. Dự báo quy mô trường lớp giai đoạn 2027 - 2031", 2)
add_para("- Năm học 2027 - 2028: Dự kiến quy mô 15 lớp (tăng 01 lớp 10), khoảng 600 học sinh.")
add_para("- Năm học 2028 - 2029: Quy mô 15 lớp, khoảng 610 học sinh.")
add_para("- Năm học 2029 - 2030: Dự kiến quy mô 16 lớp (tăng 01 lớp 10), khoảng 640 học sinh.")
add_para("- Năm học 2030 - 2031: Quy mô ổn định 16 lớp, khoảng 650 học sinh.")

add_title("2. Phương án biên chế chi tiết từng năm", 2)
add_para("a) Năm 2027:", bold_prefix="• ")
add_para("- Nhu cầu cần bố trí: 37 người (35 viên chức + 02 HĐ 111).")
add_para("- Đề xuất chỉ tiêu biên chế: 35 biên chế viên chức (tăng 02 chỉ tiêu giáo viên so với 33 biên chế hiện có để bổ sung 01 GV Ngữ văn và 01 GV Tiếng Anh đón đầu tăng lớp).")

add_para("b) Năm 2028:", bold_prefix="• ")
add_para("- Nhu cầu cần bố trí: 37 người. Đề xuất giữ ổn định 35 biên chế viên chức + 02 HĐ 111.")

add_para("c) Năm 2029:", bold_prefix="• ")
add_para("- Quy mô tăng lên 16 lớp. Nhu cầu nhân lực cần bố trí: 39 người (37 viên chức + 02 HĐ 111).")
add_para("- Đề xuất giao 37 chỉ tiêu biên chế viên chức (bổ sung 02 chỉ tiêu giáo viên Toán và Tin học/Công nghệ).")

add_para("d) Năm 2030 và 2031:", bold_prefix="• ")
add_para("- Quy mô ổn định 16 lớp đạt chuẩn quốc gia. Duy trì ổn định 37 chỉ tiêu biên chế viên chức + 02 HĐ 111.")

add_title("3. Các nhóm giải pháp tổ chức thực hiện", 2)
add_para("1. Tuyển dụng và bố trí: Phối hợp Sở GD&ĐT thực hiện tuyển dụng công khai, ưu tiên sinh viên sư phạm chất lượng cao theo Nghị định 140/2017/NĐ-CP và chính sách thu hút nhân tài của tỉnh Cao Bằng.")
add_para("2. Đào tạo, bồi dưỡng: Nâng cao năng lực chuyên môn, 100% giáo viên đạt chuẩn đào tạo theo Luật Giáo dục 2019, bồi dưỡng năng lực số, ứng dụng trí tuệ nhân tạo (AI) trong dạy học và kiểm tra đánh giá.")
add_para("3. Tối ưu hóa chuyển đổi số: Tiếp tục phát huy hiệu quả quản lý văn bản điện tử iOffice, học bạ số, CSDL ngành, sổ điểm điện tử vnEdu nhằm giảm áp lực hành chính thủ công cho giáo viên.")
add_para("4. Thực hiện hợp đồng lao động: Thực hiện đúng quy định tại Nghị định 111/2022/NĐ-CP đối với các vị trí phục vụ, bảo vệ; đảm bảo đầy đủ chế độ, chính sách cho người lao động.")

# Bảng chữ ký cuối đề án
p_end_space = doc.add_paragraph()
p_end_space.paragraph_format.space_before = Pt(16)

t_end = doc.add_table(rows=1, cols=2)
t_end.alignment = WD_TABLE_ALIGNMENT.CENTER
t_end.columns[0].width = Inches(3.2)
t_end.columns[1].width = Inches(3.5)

c_l = t_end.cell(0, 0).paragraphs[0]
c_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = c_l.add_run("TỔ TRƯỞNG CHUYÊN MÔN\n(Đại diện CBGVNV)\n\n\n\n\n")
r.font.name = "Times New Roman"
r.font.size = Pt(11)
r.font.bold = True
r = c_l.add_run("Hoàng Minh Tuấn")
r.font.name = "Times New Roman"
r.font.size = Pt(11.5)
r.font.bold = True

c_r = t_end.cell(0, 1).paragraphs[0]
c_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = c_r.add_run("HIỆU TRƯỞNG\n\n\n\n\n\n")
r.font.name = "Times New Roman"
r.font.size = Pt(11.5)
r.font.bold = True
r = c_r.add_run("Hoàng Trung Kiên")
r.font.name = "Times New Roman"
r.font.size = Pt(12)
r.font.bold = True

# SAVE FILE
desktop_path = r"D:\Desktop\De-An-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.docx"
assets_path = r"d:\Du-an-web\web-nhu-cau-nhan-luc-bien-che-2027-2031-phuc-hoa\assets"
os.makedirs(assets_path, exist_ok=True)
project_word_path = os.path.join(assets_path, "De-An-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.docx")

doc.save(project_word_path)
print(f"Word document saved to project assets: {project_word_path}")

try:
    doc.save(desktop_path)
    print(f"Word document also saved to Desktop: {desktop_path}")
except Exception as e:
    print(f"Could not save to desktop: {e}")
