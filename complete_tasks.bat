@echo off
chcp 65001 >nul
REM EIMS 작업 완료 처리 배치 파일

echo ========================================
echo    EIMS 작업 완료 처리 시스템
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

REM 작업 완료 처리 실행
echo [INFO] 작업 완료 처리 시작...
python scripts/complete_tasks.py

echo.
echo 작업이 완료되었습니다.
pause
