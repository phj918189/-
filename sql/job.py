
import sys, pathlib
import logging

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path: sys.path.insert(0, str(ROOT))

from scripts.scrape_eims import fetch_excel_df
from scripts.normalize import normalize
from scripts import db
from scripts.assign import run_assign
from scripts.notify import notify
from scripts.cleanup import cleanup_old_files
from scripts.structured_logger import logger
from scripts.database_manager import DatabaseManager
from scripts.sync_to_shared import SharedFolderSync

def main():
    try:
        # 구조화된 로깅 시작
        logger.log_event("process_started", version="1.0")
        
        # 데이터베이스 관리자 초기화
        db_manager = DatabaseManager()
        
        # 1) 데이터베이스 초기화
        try:
            logger.log_event("database_init_started")
            db.init_db()
            logger.log_event("database_init_completed")
        except Exception as e:
            logger.log_error("database_init_failed", e)
            raise
        
        # 2) 엑셀 다운로드
        try:
            logger.log_event("excel_download_started")
            raw = fetch_excel_df()
            src = raw.attrs.get("source_path","")
            logger.log_download(src.split('/')[-1], len(raw), len(raw))
        except Exception as e:
            logger.log_error("excel_download_failed", e, 
                           suggestion="로그인 문제일 가능성이 높습니다. .env 파일과 네트워크 연결을 확인하세요.")
            raise
        
        # 3) 데이터 표준화
        try:
            logger.log_event("data_normalization_started", rows=len(raw))
            df = normalize(raw)
            logger.log_event("data_normalization_completed", rows=len(df))
        except Exception as e:
            logger.log_error("data_normalization_failed", e)
            raise
        
        # 4) 데이터베이스 저장
        try:
            logger.log_event("database_save_started", rows=len(df))
            db.upsert_samples(df, src)
            logger.log_event("database_save_completed")
        except Exception as e:
            logger.log_error("database_save_failed", e)
            raise
        
        # 5) 작업 배정
        try:
            logger.log_event("assignment_started", rows=len(df))
            assigned = run_assign(df)
            logger.log_assignment("system", len(assigned), [a["item"] for a in assigned])
        except Exception as e:
            logger.log_error("assignment_failed", e)
            raise
        
        # 6) 알림 발송
        try:
            logger.log_event("notification_started", count=len(assigned))
            notify(assigned)
            logger.log_event("notification_completed")
        except Exception as e:
            logger.log_error("notification_failed", e)
            logger.log_event("notification_failed_but_continuing")
        
        # 7) 데이터베이스 백업
        try:
            logger.log_event("backup_started")
            backup_path = db_manager.backup_database()
            logger.log_event("backup_completed", backup_path=str(backup_path))
        except Exception as e:
            logger.log_error("backup_failed", e)
        
        # 8) 파일 정리
        try:
            logger.log_event("cleanup_started")
            cleanup_old_files()
            logger.log_event("cleanup_completed")
        except Exception as e:
            logger.log_error("cleanup_failed", e)
        
        # 9) 공유폴더 동기화
        try:
            logger.log_event("shared_folder_sync_started")
            shared_sync = SharedFolderSync()
            shared_sync.sync_all()
            logger.log_event("shared_folder_sync_completed")
        except Exception as e:
            logger.log_error("shared_folder_sync_failed", e)
            logger.log_event("shared_folder_sync_failed_but_continuing")
        
        logger.log_event("process_completed", success=True)
        
    except Exception as e:
        logger.log_error("process_failed", e)
        sys.exit(1)

if __name__ == "__main__":
    main()
