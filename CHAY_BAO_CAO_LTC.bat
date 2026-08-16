@echo off
chcp 65001 > nul
echo =======================================================
echo     TỰ ĐỘNG TẢI BÁO CÁO LẤY VÀ LUÂN CHUYỂN (TABLE A)
echo =======================================================
echo.

echo Đang chạy quy trình tự động đăng nhập, tải dữ liệu Table A và ghi vào rawLTC...
python "C:\Users\lap4all\.gemini\antigravity-ide\scratch\download_report_kpi_ltc.py"
if %ERRORLEVEL% NEQ 0 (
echo.
echo ❌ [LỖI] Quy trình tự động tải báo cáo thất bại.
pause
goto end
)

echo.
echo =======================================================
echo   🎉 HOÀN THÀNH TẢI BÁO CÁO LTC VÀ CẬP NHẬT SHEET!
echo =======================================================
echo.
pause

:end

The above content shows the entire, complete file contents of the requested file.