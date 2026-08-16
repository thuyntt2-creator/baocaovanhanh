@echo off
chcp 65001 > nul
echo =======================================================
echo     TỰ ĐỘNG CẬP NHẬT BÁO CÁO ĐƠN GIAO CHƯA GIAO LẦN NÀO
echo =======================================================
echo.

echo [Bước 1/1] Đang tính toán thống kê và cập nhật Google Sheets báo cáo...
python "c:\Users\lap4all\Documents\Auto report\calculate_report_no_attempt.py" %*
if %ERRORLEVEL% NEQ 0 (
echo.
echo ❌ [LỖI] Tính toán và cập nhật báo cáo thất bại.
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