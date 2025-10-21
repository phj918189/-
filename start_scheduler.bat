@echo off
chcp 65001 >nul
REM EIMS 자동화 스케줄러 실행 배치 파일

echo ========================================
echo    EIMS 자동화 스케줄러 시작
echo ========================================

REM 현재 디렉토리를 스크립트 위치로 변경
cd /d "%~dp0"

REM Python 환경 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python이 설치되지 않았습니다.
    pause
    exit /b 1
)

REM schedule 라이브러리 확인
python -c "import schedule" >nul 2>&1
if errorlevel 1 (
    echo [INFO] schedule 라이브러리 설치 중...
    pip install schedule
)

REM 스케줄러 실행
echo [INFO] EIMS 자동화 스케줄러 시작...
echo [INFO] 매일 오전 9시에 자동 실행됩니다.
echo [INFO] 중단하려면 Ctrl+C를 누르세요.
echo.

python scripts/scheduler.py

pause
