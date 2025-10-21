@echo off
chcp 65001 >nul
REM EIMS 자동화 시스템 부팅 시 자동 시작 스크립트

echo ========================================
echo    EIMS 자동화 시스템 부팅 시 시작
echo ========================================

REM 현재 디렉토리를 스크립트 위치로 변경
cd /d "%~dp0"

REM 로그 파일 생성
set LOG_FILE=storage\logs\boot_start_%date:~0,4%%date:~5,2%%date:~8,2%.log
echo [%date% %time%] 부팅 시 EIMS 자동화 시작 >> %LOG_FILE%

REM Python 환경 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python이 설치되지 않았습니다. >> %LOG_FILE%
    exit /b 1
)

REM 스케줄러 실행
echo [%date% %time%] EIMS 자동화 스케줄러 시작 >> %LOG_FILE%
python scripts/scheduler.py >> %LOG_FILE% 2>&1
