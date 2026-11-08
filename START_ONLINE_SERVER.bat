@echo off
cd /d "%~dp0"
title CEAT SAFETY WORKPLACE PORTAL - ONLINE SERVER
echo ========================================================
echo   Starting CEAT Safety Reporting System with Public URL
echo ========================================================
where py >nul 2>nul
if not errorlevel 1 (
    py -3 start_online.py
) else (
    python start_online.py
)
pause
