#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EIMS 자동화 스케줄러
매일 정해진 시간에 EIMS 자동화 시스템을 실행
"""

import schedule
import time
import subprocess
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('storage/logs/scheduler.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def run_eims_automation():
    """EIMS 자동화 시스템 실행"""
    try:
        logger.info("=== EIMS 자동화 시작 ===")
        
        # 현재 디렉토리를 프로젝트 루트로 변경
        project_root = Path(__file__).parent.parent
        os.chdir(project_root)
        
        # job.py 실행
        result = subprocess.run([sys.executable, 'sql/job.py'], 
                              capture_output=True, 
                              text=True, 
                              encoding='cp949',  # Windows 기본 인코딩 사용
                              errors='replace')   # 인코딩 오류 시 대체 문자 사용
        
        if result.returncode == 0:
            logger.info("✅ EIMS 자동화 성공")
            logger.info(f"출력: {result.stdout}")
        else:
            logger.error("❌ EIMS 자동화 실패")
            logger.error(f"오류: {result.stderr}")
            
    except Exception as e:
        logger.error(f"EIMS 자동화 실행 중 오류: {e}")

def show_next_runs():
    """다음 실행 시간 표시"""
    jobs = schedule.get_jobs()
    if jobs:
        logger.info("📅 다음 실행 예정:")
        for job in jobs:
            next_run = job.next_run
            logger.info(f"  - {job.job_func.__name__}: {next_run}")
    else:
        logger.info("등록된 작업이 없습니다.")

def main():
    """메인 스케줄러 함수"""
    logger.info("🚀 EIMS 자동화 스케줄러 시작")
    
    # 1) 부팅/시작 직후 1회 즉시 실행 (시간에 상관없이 즉시 수행)
    logger.info("⚡ 시작 즉시 1회 실행")
    run_eims_automation()
    
    # 2) 이후에는 고정 간격으로 반복 실행 (예: 60분마다)
    INTERVAL_MINUTES = int(os.environ.get("EIMS_INTERVAL_MINUTES", "60"))
    schedule.every(INTERVAL_MINUTES).minutes.do(run_eims_automation)
    logger.info(f"📋 스케줄 설정 완료: {INTERVAL_MINUTES}분마다 자동 실행")
    
    # 다음 실행 시간 표시
    show_next_runs()
    
    logger.info("⏰ 스케줄러가 실행 중입니다... (Ctrl+C로 중단)")
    
    # 스케줄러 실행
    try:
        while True:
            schedule.run_pending()
            time.sleep(10)  # 더 촘촘하게 체크
    except KeyboardInterrupt:
        logger.info("스케줄러가 중단되었습니다.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"스케줄러 오류: {e}")
