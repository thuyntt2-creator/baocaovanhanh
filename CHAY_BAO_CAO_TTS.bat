@echo off
chcp 65001 > nul
echo =======================================================
echo     TỰ ĐỘNG TẢI BÁO CÁO GIAO TRONG NGÀY (CHỈ TTS)
echo =======================================================
echo.

echo Đang chạy quy trình tự động đăng nhập, lọc TTS, tải dữ liệu và ghi vào rawtts...
python "C:\Users\lap4all\.gemini\antigravity-ide\scratch\download_report_kpi_tts.py"
if %ERRORLEVEL% NEQ 0 (
echo.
echo ❌ [LỖI] Quy trình tự động tải báo cáo thất bại.
pause
goto end
)

echo.
echo =======================================================
echo   🎉 HOÀN THÀNH TẢI BÁO CÁO TTS VÀ CẬP NHẬT SHEET!
echo =======================================================
echo.
pause

:end

The above content shows the entire, complete file contents of the requested file.