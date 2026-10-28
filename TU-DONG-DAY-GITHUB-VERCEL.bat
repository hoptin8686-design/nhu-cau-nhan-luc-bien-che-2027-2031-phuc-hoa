@echo off
chcp 65001 >nul
title TỰ ĐỘNG XÂY DỰNG ĐỀ ÁN VTVL VÀ ĐẨY LÊN GITHUB & VERCEL - THPT PHỤC HÒA
color 0B

echo ======================================================================
echo   HỆ THỐNG TỰ ĐỘNG ĐỒNG BỘ GITHUB & VERCEL
echo   Xác Định Nhu Cầu Nhân Lực & Xây Dựng Phương Án Biên Chế 2027-2031
echo   Trường THPT Phục Hòa - Tỉnh Cao Bằng
echo   Căn cứ Hướng dẫn số 07-HD/BTCTW & Công văn Sở GD&ĐT Cao Bằng
echo ======================================================================
echo.

set "PATH=C:\Program Files\Python39;C:\Users\DMX HOA THUAN\AppData\Local\Programs\Python\Python39;C:\Users\DMX HOA THUAN\AppData\Local\Microsoft\WindowsApps;C:\Users\DMX HOA THUAN\AppData\Local\Programs\Git\cmd;C:\Users\DMX HOA THUAN\AppData\Local\Programs\gh\bin;%PATH%"
cd /d "d:\Du-an-web\web-nhu-cau-nhan-luc-bien-che-2027-2031-phuc-hoa"

echo [1/4] Đang đồng bộ tạo tệp Excel & Word chuẩn ra Desktop...
python generate_excel.py
python generate_word.py
echo       ✓ Đã cập nhật file Excel & Word ra màn hình Desktop thành công!

echo.
echo [2/4] Kiểm tra kho lưu trữ Git cục bộ...
if not exist ".git" (
    echo       - Đang khởi tạo kho git...
    git init -b main
)

echo.
echo [3/4] Đang thêm tập tin và tạo bản lưu (Commit)...
git add .
git commit -m "feat: Xac dinh nhu cau nhan luc va xay dung phuong an bien che 2027-2031 THPT Phuc Hoa theo HD 07-HD/BTCTW va CV So GDDT Cao Bang"

echo.
echo [4/4] Đang tự động đẩy lên GitHub (Tự động kích hoạt Vercel & GitHub Pages)...
gh repo view nhu-cau-nhan-luc-bien-che-2027-2031-phuc-hoa >nul 2>&1
if %errorlevel% neq 0 (
    echo       - Đang tạo repository mới trên GitHub...
    gh repo create nhu-cau-nhan-luc-bien-che-2027-2031-phuc-hoa --public --source=. --remote=origin --push
) else (
    git push -u origin main
)

if %errorlevel% equ 0 (
    echo.
    echo ======================================================================
    echo   ✓ ĐẨY LÊN GITHUB THÀNH CÔNG!
    echo   ✓ GitHub Pages & Vercel đang tự động triển khai (CI/CD)...
    echo ======================================================================
    echo.
    echo [LINK TRUY CẬP TRỰC TUYẾN]:
    echo   1. Kho mã nguồn GitHub: https://github.com/hoptin8686-design/nhu-cau-nhan-luc-bien-che-2027-2031-phuc-hoa
    echo   2. Vercel 1-Click Deploy: https://vercel.com/new/clone?repository-url=https%%3A%%2F%%2Fgithub.com%%2Fhoptin8686-design%%2Fnhu-cau-nhan-luc-bien-che-2027-2031-phuc-hoa
    echo.
    echo [FILE ĐÃ LƯU TRÊN MÀN HÌNH MÁY TÍNH (DESKTOP)]:
    echo   - D:\Desktop\Bo-Bieu-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.xlsx
    echo   - D:\Desktop\De-An-Xac-Dinh-Nhu-Cau-Nhan-Luc-Bien-Che-2027-2031-THPT-Phuc-Hoa.docx
    echo.
) else (
    echo.
    echo [!] Có lỗi xảy ra trong quá trình push. Vui lòng kiểm tra lại kết nối mạng.
)

echo Bấm phím bất kỳ để đóng cửa sổ này...
pause >nul
