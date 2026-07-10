@REM MIS utility usage RPA launcher
@echo off
echo ============================================
echo  MIS Utility Usage RPA
echo  Date range: auto D-2
echo ============================================
echo.

cd /d "%~dp0"

REM Activate virtual environment if available.
if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

echo [START] Running MIS utility usage RPA...
python utility_daily_rpa.py %*

echo.
if %errorlevel% equ 0 (
    echo [OK] RPA completed successfully.
) else (
    echo [ERROR] RPA failed. Exit code: %errorlevel%
)
echo.
pause
