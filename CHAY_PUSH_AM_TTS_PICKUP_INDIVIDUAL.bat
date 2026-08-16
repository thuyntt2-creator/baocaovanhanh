@echo off
chcp 65001 > nul
echo =======================================================
echo     CẢNH BÁO GÁN ĐƠN LẤY TTS RIÊNG TỪNG AM QUA GTALK
echo =======================================================
echo.
python "%~dp0push_am_tts_pickup_individual.py"
pause
