@echo off
chcp 65001 > nul
echo =======================================================
echo        TỰ ĐỘNG TẢI DỮ LIỆU VÀ CẬP NHẬT BÁO CÁO GÁN
echo =======================================================
echo.

echo [Bước 1/2] Đang kích hoạt tải dữ liệu Last Mile mới nhất từ GHN...
python "c:\Users\lap4all\.gemini\antigravity-ide\scratch\download_report_thuy.py"
if %ERRORLEVEL% NEQ 0 (
echo.
echo ❌ [LỖI] Tải dữ liệu từ GHN thất bại. Vui lòng kiểm tra Telegram bot / API.
pause
goto end
)

echo.
echo [Bước 2/2] Đang tính toán thống kê và cập nhật Google Sheets báo cáo gán...
python "c:\Users\lap4all\Documents\Auto report\calculate_report_gan.py"
if %ERRORLEVEL% NEQ 0 (
echo.
echo ❌ [LỖI] Tính toán và ghi đè Google Sheets thất bại.
pause
goto end
)

echo.
echo =======================================================
echo   🎉 CẬP NHẬT THÀNH CÔNG!
echo   Báo cáo đã được cập nhật và gửi ảnh bảng màu lên GTalk.
echo =======================================================
echo.
pause

:end

The above content shows the entire, complete file contents of the requested file.