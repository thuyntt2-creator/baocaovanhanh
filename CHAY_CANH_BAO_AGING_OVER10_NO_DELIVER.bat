@echo off
echo =======================================================
echo   CANH BAO DON AGING > 10 NGAY CHUA TUNG DI GIAO (NUM_DELIVER = 0)
echo =======================================================
echo.

echo Dang chay kiem tra don hang chua giao...
python check_unattempted_aging_10.py %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [LOI] Quy trinh kiem tra that bai.
    pause
    goto end
)

echo.
echo =======================================================
echo   KIEM TRA VA GOI CANH BAO THANH CONG!
echo =======================================================
echo.
pause

:end