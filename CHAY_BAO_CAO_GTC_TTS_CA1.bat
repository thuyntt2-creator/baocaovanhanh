@echo off
chcp 65001 > nul
echo =======================================================================
echo   QUY TRÌNH TỰ ĐỘNG TẢI BÁO CÁO GTC CA 1 TTS VÀ GỬI BÁO CÁO SANG GTALK
echo =======================================================================
echo.

echo 🚀 Bước 1: Đăng nhập, lọc TTS + Ca 1, tải dữ liệu và ghi vào rawGTCTTS...
python "c:\Users\lap4all\Documents\Auto report\download_report_gtc_ca1_tts.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ [LỖI] Quy trình tự động tải báo cáo từ Looker Studio thất bại.
    pause
    goto end
)

echo.
echo 🚀 Bước 2: Tính toán tỷ lệ GTC theo AM, dựng ảnh bảng biểu và gửi sang GTalk...
python "c:\Users\lap4all\Documents\Auto report\calculate_gtc_ca1_tts.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ [LỖI] Tính toán báo cáo hoặc phát sóng sang GTalk thất bại.
    pause
    goto end
)

echo.
echo =======================================================================
echo   🎉 HOÀN THÀNH TOÀN BỘ QUY TRÌNH BÁO CÁO GTC CA 1 TTS!
echo =======================================================================
echo.
pause

:end
