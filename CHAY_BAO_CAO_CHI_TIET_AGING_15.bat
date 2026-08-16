@echo off
echo =======================================================
echo   TU DONG GUI CHI TIET MA DON AGING > 15 NGAY CUA TUNG AM
echo =======================================================
echo.

echo Dang chay quy trinh tong hop va gui bao cao chi tiet...
python report_aging_15_details.py %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [LOI] Quy trinh chay that bai.
    pause
    goto end
)

echo.
echo =======================================================
echo   TONG HOP VA GUI BAO CAO CHI TIET THANH CONG!
echo =======================================================
echo.
pause

:end