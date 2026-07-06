@echo off
echo ===============================================
echo   LET Dashboard - One-Click Refresh
echo ===============================================
echo.
echo Pulling fresh data from SharePoint...
cd /d "%~dp0"
py refresh_data.py
if errorlevel 1 (
    echo.
    echo ERROR: Could not pull data. Check your credentials.
    pause
    exit /b 1
)
echo.
echo Rebuilding region map from latest LaMP download...
py build_region_map.py
if errorlevel 1 (
    echo.
    echo WARNING: Region map rebuild failed - continuing with previous map.
)
echo.
echo Building dashboard...
py generate_dashboard.py
if errorlevel 1 (
    echo.
    echo ERROR: Dashboard generation failed.
    pause
    exit /b 1
)
echo.
echo Syncing .aspx copy (for SharePoint embeds)...
copy /Y "%~dp0LET_Dashboard.html" "%~dp0LET_Dashboard.aspx" >nul
echo.
echo Pushing to GitHub Pages...
set "REPO_DIR=%TEMP%\LET-Dashboard-push"
if not exist "%REPO_DIR%\.git" (
    git clone https://github.com/ylahat-intc/LET-Dashboard.git "%REPO_DIR%"
)
git -C "%REPO_DIR%" pull --quiet
copy /Y "%~dp0LET_Dashboard.html" "%REPO_DIR%\LET_Dashboard.html" >nul
copy /Y "%~dp0let_data.json"      "%REPO_DIR%\let_data.json"      >nul
copy /Y "%~dp0Designer (14).png"  "%REPO_DIR%\favicon.png"        >nul
git -C "%REPO_DIR%" add LET_Dashboard.html let_data.json favicon.png
git -C "%REPO_DIR%" commit -m "Refresh data %DATE:~-4%-%DATE:~4,2%-%DATE:~7,2%"
git -C "%REPO_DIR%" push origin main
if errorlevel 1 (
    echo.
    echo ERROR: GitHub push failed. Check your credentials.
    pause
    exit /b 1
)
echo.
echo Done! Live dashboard updated at:
echo   https://ylahat-intc.github.io/LET-Dashboard/LET_Dashboard.html
timeout /t 5
