@REM RawDB_wip -> DB_wip build - processing only, no MIS
@REM   No args  : prompts for plants and preview mode
@REM   With args: passed straight through, e.g. --plants F20,F30 --dry-run
@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ============================================
echo  WIP dataset build (processing only)
echo  RawDB_wip.xlsx -^> DB_wip.xlsx
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
    echo [START] wip_refactoring.py %*
    python wip_refactoring.py %*
    goto :done
)

echo Close DB_wip.xlsx in Excel before running.
echo Leave plants blank to process every plant.
set "PLANTS="
set "PREVIEW="
set /p PLANTS=Plants (e.g. F20,F30 - blank = all):
set /p PREVIEW=Preview only without saving? [y/N]:

set "ARGS="
if not "!PLANTS!"=="" set "ARGS=!ARGS! --plants !PLANTS!"
if /i "!PREVIEW!"=="y" set "ARGS=!ARGS! --dry-run"

echo.
echo [START] wip_refactoring.py!ARGS!
python wip_refactoring.py!ARGS!

:done
echo.
if %errorlevel% equ 0 (
    echo [OK] Build completed successfully.
) else (
    echo [ERROR] Build failed. Exit code: %errorlevel%
)
echo.
pause
