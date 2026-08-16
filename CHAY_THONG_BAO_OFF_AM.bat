@echo off
echo =======================================================
echo   TU DONG GUI THONG BAO KET QUA OFF TUYEN CHO AM
echo =======================================================
echo.
cd /d "c:\Users\lap4all\Documents\Auto report"

echo Dang lay du lieu va gui thong bao GTalk...
python push_off_tuyen_am.py --send

echo.
echo =======================================================
echo   HOAN THANH GUI THONG BAO!
echo =======================================================
echo.
pause
