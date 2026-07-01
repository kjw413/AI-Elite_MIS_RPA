@REM MIS 좌표 측정 도구 실행 배치 파일
@echo off
chcp 65001 >nul
echo ============================================
echo  MIS 좌표 측정 도구 (mouse_pos)
echo  창 기준 상대 좌표를 실시간 표시합니다.
echo  종료: Ctrl+C
echo ============================================
echo.

cd /d "%~dp0"

if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

python mouse_pos.py

echo.
pause
