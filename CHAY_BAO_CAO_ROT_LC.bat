@echo off
chcp 65001 > nul
echo =======================================================
echo     TỰ ĐỘNG CẬP NHẬT BÁO CÁO RỚT LUÂN CHUYỂN (ROT LC)
echo =======================================================
echo.

echo Đang chạy quy trình tính toán và cập nhật báo cáo rớt LC...
python calculate_report_rot_lc.py %*
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ [LỖI] Quy trình tính toán và cập nhật thất bại.
    pause
    goto end
)

echo.
echo =======================================================
echo   🎉 CẬP NHẬT THÀNH CÔNG BÁO CÁO RỚT LUÂN CHUYỂN!
echo =======================================================
echo.
pause

:end