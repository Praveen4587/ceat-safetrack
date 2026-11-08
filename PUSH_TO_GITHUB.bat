@echo off
setlocal
title SafeTrack - Push to GitHub
cd /d "%~dp0"

echo ===================================================
echo     PUSH SAFETRACK TO GITHUB (FOR 24/7 CLOUD)
echo ===================================================
echo.
echo 1. Open https://github.com/new in your browser.
echo 2. Repository name: ceat-safetrack
echo 3. Choose Private.
echo 4. Leave "Add a README" UNCHECKED.
echo 5. Click "Create repository".
echo.
echo Copy the HTTPS URL from GitHub (e.g. https://github.com/Praveen4587/ceat-safetrack.git)
echo.
set /p REPO_URL="Paste your GitHub repository URL: "
if "%REPO_URL%"=="" (
    echo [!] No URL entered. Exiting.
    pause
    exit /b 1
)

git remote remove origin >nul 2>&1
git remote add origin %REPO_URL%
echo.
echo [*] Pushing main branch to %REPO_URL% ...
git push -u origin main
echo.
if %ERRORLEVEL% EQU 0 (
    echo ===================================================
    echo [SUCCESS] Code successfully pushed to GitHub!
    echo.
    echo Next Step:
    echo  - Open https://dashboard.render.com
    echo  - Click New + -> Web Service
    echo  - Connect your "ceat-safetrack" GitHub repository
    echo  - Render will deploy your site 24/7 automatically!
    echo ===================================================
) else (
    echo ===================================================
    echo [!] Push failed.
    echo Please make sure your GitHub credentials are valid
    echo and the repository URL is correct.
    echo ===================================================
)
pause
