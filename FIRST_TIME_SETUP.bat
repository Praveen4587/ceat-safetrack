@echo off
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (py -3 app.py setup) else (python app.py setup)
pause
