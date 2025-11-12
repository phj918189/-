@echo off
chcp 65001 >nul
echo ========================================
echo 엑셀 비교 뷰어 설치 스크립트
echo ========================================
echo.

REM Python 설치 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo Python을 먼저 설치해주세요: https://www.python.org/downloads/
    echo 설치 시 "Add Python to PATH" 옵션을 체크하세요.
    pause
    exit /b 1
)

echo [확인] Python이 설치되어 있습니다.
python --version
echo.

REM pip 업그레이드
echo [1/3] pip 업그레이드 중...
python -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo [경고] pip 업그레이드 실패, 계속 진행합니다...
)
echo.

REM 패키지 설치
echo [2/3] 필요한 패키지 설치 중...
pip install -r requirements.txt
if errorlevel 1 (
    echo [오류] 패키지 설치 실패
    pause
    exit /b 1
)
echo.

REM 설치 확인
echo [3/3] 설치 확인 중...
pip list | findstr /i "pandas openpyxl xlrd tksheet" >nul
if errorlevel 1 (
    echo [경고] 일부 패키지가 제대로 설치되지 않았을 수 있습니다.
) else (
    echo [성공] 모든 패키지가 설치되었습니다.
)
echo.

echo ========================================
echo 설치 완료!
echo ========================================
echo.
echo 실행 방법:
echo   python excel_compare_app.py
echo.
echo 또는 dist\excel_compare_app.exe 파일을 더블클릭하세요.
echo.
pause

