@echo off
title Job Application Tracker
REM ===========================================================
REM  Double-click this file to build / refresh your tracker.
REM  It scans this folder, updates job_applications.xlsx (+ .csv),
REM  keeps your edits, and saves a backup of the previous version.
REM ===========================================================

cd /d "%~dp0"

REM --- make sure Python is installed ---
where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo  ====================================================================
    echo   Python is not installed ^(or not on your PATH^).
    echo.
    echo   1. Go to:  https://www.python.org/downloads/
    echo   2. Download and run the installer.
    echo   3. IMPORTANT: tick the box "Add Python to PATH" on the first screen.
    echo   4. After it finishes, double-click this file again.
    echo  ====================================================================
    echo.
    pause
    exit /b
)

REM --- make sure the Excel library is available (quietly install if missing) ---
python -m pip install --quiet --disable-pip-version-check openpyxl

REM --- run the tracker ---
python "update_job_report.py"

echo.
echo  Done! Open  job_applications.xlsx  to see your tracker.
echo  You can close this window.
pause
