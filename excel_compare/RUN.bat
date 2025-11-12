@echo off
chcp 65001 >nul
echo 엑셀 비교 뷰어 실행 중...
echo.

REM 가상환경이 있으면 활성화
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)

REM Python 스크립트 실행
python excel_compare_app.py

if errorlevel 1 (
    echo.
    echo [오류] 실행 실패
    echo 패키지가 설치되어 있는지 확인하세요: pip install -r requirements.txt
    pause
)

