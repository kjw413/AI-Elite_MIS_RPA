@REM MIS energy (utility) RPA launcher - unit-input screen
@REM   No args  : collect previous day's month + recover missing previous date
@REM   With args: passed straight through, e.g. --from 2024-01 --to 2026-06
@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ============================================
echo  MIS Energy (Utility) RPA
echo  Screen: unit-input (daily)
echo ============================================
echo.

cd /d "%~dp0"

REM Activate virtual environment if available.
if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

REM Explicit args win - used by the orchestrator and for past-month collection.
if not "%~1"=="" (
    echo [START] utility_daily_rpa.py %*
    python utility_daily_rpa.py %*
    goto :done
)

echo Leave the start month blank for the daily run ^(with missing-date recovery^).
set "YM_FROM="
set "YM_TO="
set /p YM_FROM=Start month for past collection (YYYY-MM, blank = daily run):

if "!YM_FROM!"=="" (
    echo.
    echo [START] Daily run - previous day's month + missing-date recovery...
    python utility_daily_rpa.py
    goto :done
)

set /p YM_TO=End month (YYYY-MM, blank = previous day's month):

echo.
echo [START] Past-month collection from !YM_FROM!...
if "!YM_TO!"=="" (
    python utility_daily_rpa.py --from !YM_FROM!
) else (
    python utility_daily_rpa.py --from !YM_FROM! --to !YM_TO!
)

:done
echo.
if %errorlevel% equ 0 (
    echo [OK] RPA completed successfully.
) else (
    echo [ERROR] RPA failed. Exit code: %errorlevel%
)
echo.
pause
