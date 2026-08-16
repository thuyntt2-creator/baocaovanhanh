@echo off
echo =======================================================
echo     TU DONG TAI BAO CAO LAY VA LUAN CHUYEN (TABLE A)
echo =======================================================
echo.

echo Dang chay quy trinh tu dong tai bao cao LTC...
python "C:\Users\lap4all\.gemini\antigravity-ide\scratch\download_report_kpi_ltc.py"
if %ERRORLEVEL% neq 0 (
    echo.
    echo [LOI] Quy trinh tu dong tai bao cao that bai. Code: %ERRORLEVEL%
    exit /b %ERRORLEVEL%
)

echo.
echo =======================================================
echo     HOAN THANH TAI BAO CAO LTC VA CAP NHAT SHEET!
echo =======================================================
echo.
