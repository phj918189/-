@echo off
chcp 65001 >nul
REM EIMS 자동화 시스템 실행 스크립트 (자동 실행용)

echo ========================================
echo    EIMS 자동화 시스템 시작
echo    실행시간: %date% %time%
echo ========================================

REM 현재 디렉토리를 스크립트 위치로 변경
cd /d "%~dp0"

REM 로그 파일 생성
set LOG_FILE=storage\logs\auto_run_%date:~0,4%%date:~5,2%%date:~8,2%.log
echo [%date% %time%] EIMS 자동화 시작 >> %LOG_FILE%

REM Python 환경 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python이 설치되지 않았습니다. >> %LOG_FILE%
    exit /b 1
)

REM 메인 프로세스 실행
echo [%date% %time%] EIMS 자동화 프로세스 시작 >> %LOG_FILE%
python sql/job.py >> %LOG_FILE% 2>&1

if errorlevel 1 (
    echo [ERROR] EIMS 자동화 실행 실패 >> %LOG_FILE%
    exit /b 1
) else (
    echo [SUCCESS] EIMS 자동화 완료 >> %LOG_FILE%
)

echo [%date% %time%] 작업 완료 >> %LOG_FILE%
