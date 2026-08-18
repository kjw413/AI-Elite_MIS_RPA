@echo off
setlocal EnableDelayedExpansion

REM ============================================================
REM  MIS 3-RPA full auto-run wrapper (calls run_all_rpa.py)
REM
REM  Pipeline:
REM    [Collect] production -> utility -> wip
REM    [Process] production -> utility -> wip
REM
REM  Real-time progress is printed to the console and mirrored
REM  to logs\auto_run_*.log. Per-RPA detail logs remain in logs\.
REM
REM  Recommended: register in Windows Task Scheduler.
REM  For manual double-click, uncomment the "pause" line below
REM  if you want the window to stay open after completion.
REM ============================================================

set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
chcp 65001 >nul

cd /d "%~dp0"

REM Activate venv (optional)
if exist "..\venv\Scripts\activate.bat" call "..\venv\Scripts\activate.bat"
if exist "..\.venv\Scripts\activate.bat" call "..\.venv\Scripts\activate.bat"

REM Verify python on PATH
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] python not found in PATH.
    pause
    exit /b 9
)

REM Explicit command-line args bypass the interactive prompt.
if not "%~1"=="" goto :run_with_args

echo 조회 기간을 입력하세요. 입력한 기간이 생산실적·에너지·재공품에 공통 적용됩니다.
echo 둘 다 비우면 이번 달 1일 ~ 어제로 실행됩니다.
echo 예: 2026년 전체 복구 시작일 = 2026-01-01
set "DATE_FROM="
set "DATE_TO="
set /p DATE_FROM=시작일 (YYYY-MM-DD, Enter=기본값):
set /p DATE_TO=종료일 (YYYY-MM-DD, Enter=어제):

set "DATE_ARGS="
if not "!DATE_FROM!"=="" set "DATE_ARGS=--from !DATE_FROM!"
if not "!DATE_TO!"=="" set "DATE_ARGS=!DATE_ARGS! --to !DATE_TO!"
python -u "%~dp0run_all_rpa.py" !DATE_ARGS!
set "FINAL=!errorlevel!"
goto :after_run

:run_with_args
python -u "%~dp0run_all_rpa.py" %*
set "FINAL=!errorlevel!"

:after_run

echo.
echo ============================================================
echo  RPA 작업 완료 (exit code: %FINAL%)
echo  로그 확인 후 아무 키나 누르면 창이 닫힙니다.
echo ============================================================
pause

endlocal & exit /b %FINAL%
