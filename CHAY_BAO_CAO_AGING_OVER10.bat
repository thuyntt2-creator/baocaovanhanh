@echo off
echo =======================================================
echo   TU DONG TAI DU LIEU VA CAP NHAT BAO CAO AGING TREN 10 NGAY
echo =======================================================
echo.

echo Dang chay quy trinh tai du lieu va tinh toan bao cao aging tren 10 ngay...
python calculate_report_aging_over10.py %*
if %ERRORLEVEL% NEQ 0 (
echo.
echo [LOI] Quy trinh tinh toan va cap nhat that bai.
pause
goto end
)

echo.
echo =======================================================
echo   CAP NHAT THANH CONG BAO CAO AGING TREN 10 NGAY!
echo   Xem chi tiet tai sheet: PIVOT va cac tab AM
echo =======================================================
echo.
pause

:end

The above content shows the entire, complete file contents of the requested file.