@echo off
echo =======================================================
echo   TU DONG GUY CANH BAO TRANG THAI OFF TUYEN QUA TELEGRAM
echo =======================================================
echo.
cd /d "c:\Users\lap4all\Documents\Auto report"

echo Dang kiem tra va gui tin nhan canh bao...
python send_telegram_alerts.py

echo.
echo =======================================================
echo   HOAN THANH CANH BAO!
echo =======================================================
echo.
pause