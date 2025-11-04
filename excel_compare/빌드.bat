@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 엑셀 비교 뷰어 EXE 빌드 중...
echo ========================================
echo.

REM PyInstaller가 설치되어 있는지 확인
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller가 설치되어 있지 않습니다.
    echo 설치 중...
    pip install pyinstaller
    if errorlevel 1 (
        echo.
        echo PyInstaller 설치 실패!
        pause
        exit /b 1
    )
)

echo.
echo 빌드 시작...
python -m PyInstaller excel_compare_app.spec --clean

if errorlevel 1 (
    echo.
    echo 빌드 실패!
    pause
    exit /b 1
) else (
    echo.
    echo ========================================
    echo 빌드 완료!
    echo 실행 파일 위치: dist\excel_compare_app.exe
    echo ========================================
    echo.
    pause
)

