@echo off
title Uploading SafeTrack to GitHub...
cd /d "%~dp0"
echo ========================================================
echo   UPLOADING SAFETRACK TO GITHUB (Praveen4587/ceat-safetrack)
echo ========================================================
echo.
echo If a browser window opens asking to Sign in or Authorize,
echo please click 'Authorize GitCredentialManager'.
echo.
git push -u origin main
echo.
if %ERRORLEVEL% EQU 0 (
    echo ========================================================
    echo   UPLOAD SUCCESSFUL! ALL FILES ARE NOW ON GITHUB!
    echo ========================================================
) else (
    echo [!] If you see an error, check the message above.
)
pause
