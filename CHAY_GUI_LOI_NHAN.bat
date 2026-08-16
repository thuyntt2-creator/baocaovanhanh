@echo off
chcp 65001 > nul
title Gửi Lời Nhắn Hàng Loạt Qua GTalk
cd /d "%~dp0"
echo ============================================================
echo GỬI LỜI NHẮN HÀNG LOẠT THEO GROUP ID (SHEET 'lời nhắn')
echo ============================================================
python send_loi_nhan.py --send
echo.
echo Đã hoàn tất!
pause
