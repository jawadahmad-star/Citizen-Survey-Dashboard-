@echo off
REM ============================================================
REM  Citizen Property Tax & RS Assessment Survey (v25)
REM  Daily dashboard update — Research Solutions
REM
REM  Steps:
REM    1. Download the fresh SurveyCTO WIDE export and save it
REM       over "Citizen Property Tax & RS Assessment Survey (v25)_WIDE.csv"
REM       in this folder.
REM    2. Double-click this file.
REM  It rebuilds index.html and pushes it to GitHub Pages.
REM ============================================================
cd /d "%~dp0"

echo.
echo [1/3] Rebuilding dashboard from latest data...
python build_dashboard.py
if errorlevel 1 (
    echo BUILD FAILED — dashboard NOT updated. Check the error above.
    pause
    exit /b 1
)

echo.
echo [2/3] Committing to git...
git add index.html build_dashboard.py dashboard_template.html README.md update_dashboard.bat .gitignore
git commit -m "Daily data refresh: %date% %time%"

echo.
echo [3/3] Pushing to GitHub...
git push origin main
if errorlevel 1 (
    echo PUSH FAILED — check your internet connection / GitHub login.
    pause
    exit /b 1
)

echo.
echo DONE — dashboard updated and live.
pause
