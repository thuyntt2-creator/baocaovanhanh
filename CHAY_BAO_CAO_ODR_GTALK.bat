@echo off
chcp 65001 > nul
echo =======================================================
echo     TỰ ĐỘNG CHẠY BÁO CÁO ODR & TỒN ODR TTS SANG GTALK
echo =======================================================
echo.

echo Đang tính toán dữ liệu ODR, chụp ảnh bảng và gửi GTalk...
python "c:\Users\lap4all\Documents\Auto report\calculate_report_odr_tts.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ [LỖI] Quy trình tự động chạy báo cáo ODR thất bại.
    pause
    goto end
)

echo.
echo =======================================================
echo   🎉 HOÀN THÀNH GỬI BÁO CÁO ODR & TỒN TTS LÊN GTALK!
echo =======================================================
echo.
pause

:end