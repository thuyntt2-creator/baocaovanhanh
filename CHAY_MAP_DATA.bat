@echo off
chcp 65001 > nul
echo =======================================================
1: echo    TỰ ĐỘNG ÁNH XẠ DỮ LIỆU VÀ CẬP NHẬT SHEET DATA
echo =======================================================
echo.

echo [Bước 1/1] Đang ánh xạ dữ liệu và cập nhật lên Google Sheets...
python "c:\Users\lap4all\Documents\Auto report\calculate_report_data_mapping.py"
if %ERRORLEVEL% NEQ 0 (
echo.
echo ❌ [LỖI] Ánh xạ và cập nhật Google Sheets thất bại.
pause
goto end
)

echo.
echo =======================================================
echo   🎉 CẬP NHẬT THÀNH CÔNG!
echo   Dữ liệu đã được ánh xạ và đẩy thành công vào sheet 'Data'.
echo =======================================================
echo.
pause

:end

The above content shows the entire, complete file contents of the requested file.