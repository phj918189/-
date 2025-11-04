@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 엑셀 비교 뷰어 실행 중...
echo ========================================
python excel_compare_app.py
if errorlevel 1 (
    echo.
    echo 오류가 발생했습니다!
    echo Python이 설치되어 있는지 확인하세요.
    echo.
    pause
) else (
    echo.
    echo 프로그램이 종료되었습니다.
)

