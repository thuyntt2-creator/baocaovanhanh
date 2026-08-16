@echo off
chcp 65001 > nul
title Gửi cảnh báo Ticket tồn cho AMs (Tồn thấp <=10 phiếu)
cd /d "%~dp0"
echo ============================================================
echo GỬI THÔNG BÁO TICKET TỒN HỐI GIAO/LẤY/TRẢ TỚI GTALK AM
echo ============================================================
python push_ticket_low_am.py --send
echo.
echo Đã hoàn tất!
pause
