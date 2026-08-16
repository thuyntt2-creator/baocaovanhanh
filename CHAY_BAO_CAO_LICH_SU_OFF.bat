@echo off
echo =======================================================
echo   TU DONG PHAN TICH TAN SUAT OFF TUYEN LICH SU (NTB)
echo =======================================================
echo.
cd /d "c:\Users\lap4all\Documents\Auto report"

echo Dang tinh toan tan suat va viet tab historyoff...
python populate_historyoff.py

echo.
echo =======================================================
echo   HOAN THANH PHAN TICH!
echo   Xem chi tiet tai sheet: historyoff
echo =======================================================
echo.
pause