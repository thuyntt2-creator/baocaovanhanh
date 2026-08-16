@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo =======================================================
echo   TỰ ĐỘNG TẢI DỮ LIỆU VÀ CẬP NHẬT ĐƠN LUÂN CHUYỂN (>36H)
echo =======================================================
echo.

echo [Bắt đầu] Đang kích hoạt tải dữ liệu và chạy báo cáo...
python -u download_report_luanchuyen.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ [LỖI] Chạy báo cáo Luân Chuyển thất bại. Vui lòng kiểm tra log hoặc session.
    pause
    goto end
)

echo.
echo =======================================================
echo   🎉 CẬP NHẬT THÀNH CÔNG!
echo   Xem chi tiết tại sheet: PIVOT, stuck và các tab AM
echo =======================================================
echo.
pause

:end