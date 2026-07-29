@REM MIS work-in-process RPA launcher
@REM Opens the production plan/results screen and samples WIP data.
@echo off
echo ============================================
echo  MIS Work-in-Process RPA
echo  Date range: auto D-1 (month start to D-1)
echo  Data file : configured RawDB WIP workbook
echo ============================================
echo.

cd /d "%~dp0"

REM Activate virtual environment if available.
if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

echo [START] Running MIS work-in-process RPA...
python wip_daily_rpa.py %*

echo.
if %errorlevel% equ 0 (
    echo [OK] RPA completed successfully.
) else (
    echo [ERROR] RPA failed. Exit code: %errorlevel%
)
echo.
pause
