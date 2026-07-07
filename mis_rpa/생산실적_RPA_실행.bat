@REM MIS production results RPA launcher
@REM Samples raw production data and builds the production DW workbook.
@REM Use --skip-dw-build to skip the DW build step.
@echo off
echo ============================================
echo  MIS Production Results RPA
echo  Date range: auto D-2 (month start to D-2)
echo  Steps     : 1) Raw sampling  2) Build production DW workbook
echo ============================================
echo.

cd /d "%~dp0"

REM Activate virtual environment if available.
if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

echo [START] Running MIS production results RPA...
python production_daily_rpa.py %*

echo.
if %errorlevel% equ 0 (
    echo [OK] RPA completed successfully.
) else (
    echo [ERROR] RPA failed. Exit code: %errorlevel%
)
echo.
pause
