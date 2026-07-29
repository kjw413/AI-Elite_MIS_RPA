@echo off
setlocal EnableDelayedExpansion

REM ============================================================
REM  MIS 3-RPA full auto-run wrapper (calls run_all_rpa.py)
REM
REM  Pipeline:
REM    production UI -> utility UI -> wip UI -> wait
REM                  \                       \
REM                   production DW (BG)      wip DB (BG)
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

python -u "%~dp0run_all_rpa.py" %*
set FINAL=%errorlevel%

echo.
echo ============================================================
echo  RPA 작업 완료 (exit code: %FINAL%)
echo  로그 확인 후 아무 키나 누르면 창이 닫힙니다.
echo ============================================================
pause

endlocal & exit /b %FINAL%
