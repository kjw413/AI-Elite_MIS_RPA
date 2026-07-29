@REM MIS 좌클릭 좌표 기록기 실행 배치 파일
@echo off
chcp 65001 >nul
echo ============================================
echo  MIS 마우스 좌클릭 좌표 기록기
echo  좌클릭마다 절대/상대 좌표 자동 기록
echo  종료: Ctrl+C
echo  로그: logs/click_log_*.txt
echo ============================================
echo.

cd /d "%~dp0"

if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

python mouse_click_logger.py

echo.
pause
