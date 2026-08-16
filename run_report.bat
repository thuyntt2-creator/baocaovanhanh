@echo off
title GHN Daily Auto Report Download ^& Broadcast
cd /d "c:\Users\lap4all\Documents\Auto report"

echo 🚀 Step 1/2: Running GHN Report Download ^& Copy to Data...
python download_report_gtc.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ Step 1 failed with error code %ERRORLEVEL%. Aborting Step 2.
    exit /b %ERRORLEVEL%
)

echo.
echo 🚀 Step 2/2: Calculating, Rendering and Broadcasting Report...
python calculate_and_render_report.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ Step 2 failed with error code %ERRORLEVEL%.
    exit /b %ERRORLEVEL%
)

echo.
echo 🎉 All steps completed successfully!