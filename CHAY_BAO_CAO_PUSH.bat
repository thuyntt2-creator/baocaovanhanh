@echo off
echo =======================================================
echo   TU DONG TAI DU LIEU VA CAP NHAT LUOT GAN (AGING TREN 5)
echo =======================================================
echo.

echo Dang chay quy trinh tai du lieu va tinh toan gan don...
python update_aging_assignments.py %*
if %ERRORLEVEL% NEQ 0 (
echo.
echo [LOI] Quy trinh tinh toan va cap nhat that bai.
pause
goto end
)

echo.
echo Dang chay phan tich don chua xu ly va gui GTalk...
python scratch_check_unprocessed.py

echo.
echo =======================================================
echo   CAP NHAT THANH CONG!
echo   Xem chi tiet tai sheet: PUSH REGION va Luot gan
echo =======================================================
echo.
pause

:end

The above content shows the entire, complete file contents of the requested file.