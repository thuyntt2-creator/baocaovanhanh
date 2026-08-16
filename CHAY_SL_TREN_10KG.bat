@echo off
chcp 65001 > nul
echo =======================================================
echo     TỰ ĐỘNG TẢI BÁO CÁO GIAO TRONG NGÀY (KHỐI LƯỢNG ^> 10KG)
echo =======================================================
echo.

echo Đang chạy quy trình tự động tải báo cáo SL trên 10kg...
python "c:\Users\lap4all\Documents\Auto report\SL trên 10kg.py"
if %ERRORLEVEL% neq 0 (
    echo.
    echo [LỖI] Quy trình tự động tải báo cáo thất bại. Code: %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

echo.
echo =======================================================
echo     HOÀN THÀNH TẢI BÁO CÁO SL TRÊN 10KG VÀ CẬP NHẬT SHEET!
echo =======================================================
echo.
