@REM RawDB_production -> DB_production build - processing only, no MIS
@REM   No args  : prompts for preview mode
@REM   With args: passed straight through, e.g. --dry-run -v
@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

echo ============================================
echo  Production dataset build (processing only)
echo  RawDB_production.xlsx -^> DB_production.xlsx
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
    echo [START] build_production_dataset.py %*
    python build_production_dataset.py %*
    goto :done
)

echo Close DB_production.xlsx in Excel before running.
set "PREVIEW="
set /p PREVIEW=Preview only without saving? [y/N]:

set "ARGS="
if /i "!PREVIEW!"=="y" set "ARGS= --dry-run"

echo.
echo [START] build_production_dataset.py!ARGS!
python build_production_dataset.py!ARGS!

:done
echo.
if %errorlevel% equ 0 (
    echo [OK] Build completed successfully.
) else (
    echo [ERROR] Build failed. Exit code: %errorlevel%
)
echo.
pause
