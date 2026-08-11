@REM RawDB_에너지_수집 -> RawDB_에너지 build - processing only, no MIS
@REM   No args  : prompts for period (blank = whole range)
@REM   With args: passed straight through, e.g. --from 2024-04 --to 2024-08 --dry-run
@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ============================================
echo  Energy dataset build (processing only)
echo  Source: RawDB_energy_collection.xlsx + production actuals
echo  Target: RawDB_energy.xlsx
echo ============================================
echo.

cd /d "%~dp0"

REM Activate virtual environment if available.
if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

REM Explicit args win.
if not "%~1"=="" (
    echo [START] build_energy_dataset.py %*
    python build_energy_dataset.py %*
    goto :done
)

echo Close RawDB_energy_collection.xlsx and RawDB_energy.xlsx before running.
echo Leave both months blank to process the whole range.
set "YM_FROM="
set "YM_TO="
set "PREVIEW="
set /p YM_FROM=Start month (YYYY-MM, blank = from the beginning):
set /p YM_TO=End month (YYYY-MM, blank = to the end):
set /p PREVIEW=Preview only without saving? [y/N]:

set "ARGS="
if not "!YM_FROM!"=="" set "ARGS=!ARGS! --from !YM_FROM!"
if not "!YM_TO!"=="" set "ARGS=!ARGS! --to !YM_TO!"
if /i "!PREVIEW!"=="y" set "ARGS=!ARGS! --dry-run"

echo.
echo [START] build_energy_dataset.py!ARGS!
python build_energy_dataset.py!ARGS!

:done
echo.
if %errorlevel% equ 0 (
    echo [OK] Build completed successfully.
) else (
    echo [ERROR] Build failed. Exit code: %errorlevel%
)
echo.
pause
