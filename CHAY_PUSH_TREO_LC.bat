@echo off
chcp 65001 > nul
echo =======================================================
echo     GỬI CẢNH BÁO ĐƠN TREO LUÂN CHUYỂN TOÀN AM (GTALK)
echo =======================================================
echo.

cd /d "%~dp0"
python pushtreolc.py --force-send %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ [LỖI] Tiến trình gửi tin nhắn thất bại.
    pause
    goto end
)

echo.
echo =======================================================
echo   🎉 ĐÃ GỬI THÀNH CÔNG CẢNH BÁO ĐƠN TREO LUÂN CHUYỂN!
echo =======================================================
echo.
pause

:end
