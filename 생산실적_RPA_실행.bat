@REM MIS production results RPA launcher
@REM Samples raw production data and builds the production DW workbook.
@REM Use --skip-dw-build to skip the DW build step.
@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
echo ============================================
echo  MIS Production Results RPA
echo  Date range: prompt (blank = month start through yesterday)
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

REM Explicit command-line args bypass the interactive prompt.
if not "%~1"=="" (
    echo [START] production_daily_rpa.py %*
    python production_daily_rpa.py %*
    goto :done
)

echo 조회 기간을 입력하세요. 둘 다 비우면 이번 달 1일 ~ 어제로 실행됩니다.
echo 예: 2026년 전체 복구 시작일 = 2026-01-01
set "DATE_FROM="
set "DATE_TO="
set /p DATE_FROM=시작일 (YYYY-MM-DD, Enter=기본값):
set /p DATE_TO=종료일 (YYYY-MM-DD, Enter=어제):

set "DATE_ARGS="
if not "!DATE_FROM!"=="" set "DATE_ARGS=--from !DATE_FROM!"
if not "!DATE_TO!"=="" set "DATE_ARGS=!DATE_ARGS! --to !DATE_TO!"

echo.
echo [START] production_daily_rpa.py !DATE_ARGS!
python production_daily_rpa.py !DATE_ARGS!

:done
set "FINAL=!errorlevel!"
echo.
if !FINAL! equ 0 (
    echo [OK] RPA completed successfully.
) else (
    echo [ERROR] RPA failed. Exit code: !FINAL!
)
echo.
pause
endlocal & exit /b %FINAL%
