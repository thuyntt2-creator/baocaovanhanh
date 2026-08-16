@echo off
chcp 65001 > nul
echo =======================================================
echo       TỰ ĐỘNG CHẠY BÁO CÁO TỔNG QUAN VÀ TOP BƯU CỤC
echo =======================================================
echo.

echo Đang chạy quy trình tính toán, chụp ảnh và gửi báo cáo tổng quan...
python report_tongquan.py %*
if %ERRORLEVEL% NEQ 0 (
echo.
echo ❌ [LỖI] Quy trình chạy báo cáo tổng quan thất bại.
pause
goto end
)

echo.
echo =======================================================
echo   🎉 GỬI BÁO CÁO TỔNG QUAN THÀNH CÔNG!
echo   Báo cáo và Top bưu cục đã được gửi qua Telegram và GTalk.
echo =======================================================
echo.
pause

:end

The above content shows the entire, complete file contents of the requested file.