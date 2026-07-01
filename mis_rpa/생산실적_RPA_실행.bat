@REM MIS 기간별 제품 생산실적 RPA 실행 배치 파일
@REM  Raw_생산실적 샘플링 → DB_생산실적_DW.xlsx 통합까지 일괄 수행
@REM  (DW 통합 생략은 --skip-dw-build 인자 전달)
@echo off
echo ============================================
echo  MIS 기간별 제품 생산실적 RPA
echo  기준일: D-2 자동 계산 (월 첫일 ~ D-2)
echo  단계  : 1) Raw 샘플링  2) DB_생산실적_DW.xlsx 통합
echo ============================================
echo.

cd /d "%~dp0"

REM 가상환경 활성화 (있을 경우)
if exist "..\venv\Scripts\activate.bat" (
    call "..\venv\Scripts\activate.bat"
) else if exist "..\.venv\Scripts\activate.bat" (
    call "..\.venv\Scripts\activate.bat"
)

echo [시작] MIS 생산실적 RPA 실행 중...
python production_daily_rpa.py %*

echo.
if %errorlevel% equ 0 (
    echo [완료] RPA 정상 종료
) else (
    echo [오류] RPA 실행 중 오류 발생 (코드: %errorlevel%)
)
echo.
pause
