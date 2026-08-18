@REM MIS work-in-process RPA launcher
@REM Opens the production plan/results screen and samples WIP data.
@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
echo ============================================
echo  MIS Work-in-Process RPA
echo  Date range: prompt (blank = month start through yesterday)
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

REM Explicit command-line args bypass the interactive prompt.
if not "%~1"=="" (
    echo [START] wip_daily_rpa.py %*
    python wip_daily_rpa.py %*
    goto :done
)

echo 조회 기간을 입력하세요. 둘 다 비우면 이번 달 1일 ~ 어제로 실행됩니다.
echo 예: 2026년 전체 복구 시작일 = 2026-01-01
set "DATE_FROM="
set "DATE_TO="
set /p DATE_FROM=시작일 (YYYY-MM-DD, Enter=기본값):
set /p DATE_TO=종료일 (YYYY-MM-DD, Enter=어제):
set "FACTORIES="
set /p FACTORIES=공장 (남양주/김해/광주/논산/경산 또는 F10~F50, Enter=전체):

set "DATE_ARGS="
if not "!DATE_FROM!"=="" set "DATE_ARGS=--from !DATE_FROM!"
if not "!DATE_TO!"=="" set "DATE_ARGS=!DATE_ARGS! --to !DATE_TO!"

echo.
echo [START] wip_daily_rpa.py !DATE_ARGS! --factories "!FACTORIES!"
python wip_daily_rpa.py !DATE_ARGS! --factories "!FACTORIES!"

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
