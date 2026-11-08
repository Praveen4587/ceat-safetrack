@echo off
cd /d "%~dp0"
echo ===============================================
echo  SafeTrack - Create or Reset Admin Login
 echo ===============================================
where py >nul 2>nul
if not errorlevel 1 (
  py -3 reset_admin.py
) else (
  python reset_admin.py
)
echo.
echo If SUCCESS appears above, run START_SERVER.bat and sign in.
pause
