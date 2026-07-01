@REM MIS 유틸리티 일자별 사용량 추이 RPA 실행 배치 파일
@echo off
echo ============================================
echo  MIS 유틸리티 일자별 사용량 추이 RPA
echo  기준일: D-2 자동 계산
echo ============================================
echo.

cd /d "%~dp0"

REM 가상환경 활성화 (있을 경우)
if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

echo [시작] MIS RPA 실행 중...
python utility_daily_rpa.py %*

echo.
if %errorlevel% equ 0 (
    echo [완료] RPA 정상 종료
) else (
    echo [오류] RPA 실행 중 오류 발생 (코드: %errorlevel%)
)
echo.
pause
