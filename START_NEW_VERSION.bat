@echo off
cd /d "%~dp0"
title CEAT SAFETY PORTAL - VERSION 4.0 - PORT 8091
echo =========================================
echo CEAT SAFETY - WORKPLACE SAFETY REPORTING
 echo URL: http://127.0.0.1:8091/login
 echo Old app may be on 8080. This app uses 8091.
echo =========================================
where py >nul 2>nul
if not errorlevel 1 (
 start "" "http://127.0.0.1:8091/login"
 py -3 app.py serve --port 8091
) else (
 start "" "http://127.0.0.1:8091/login"
 python app.py serve --port 8091
)
echo If startup failed, inspect the error above.
pause
