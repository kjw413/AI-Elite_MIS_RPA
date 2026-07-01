@REM MIS 재공품 RPA 실행 배치 파일
@REM  '생산계획 대비 실적현황(완제품/재공품)' 화면 -> RawDB_재공품.xlsx
@echo off
echo ============================================
echo  MIS 재공품 RPA
echo  기준일: D-2 자동 계산 (월 첫일 ~ D-2)
echo  출력  : E:\Sampled DB\RawDB_재공품.xlsx
echo ============================================
echo.

cd /d "%~dp0"

REM 가상환경 활성화 (있을 경우)
if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

echo [시작] MIS 재공품 RPA 실행 중...
python wip_daily_rpa.py %*

echo.
if %errorlevel% equ 0 (
    echo [완료] RPA 정상 종료
) else (
    echo [오류] RPA 실행 중 오류 발생 (코드: %errorlevel%)
)
echo.
pause