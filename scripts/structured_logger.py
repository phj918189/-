import logging
import json
import os
from datetime import datetime
from pathlib import Path
from logging.handlers import RotatingFileHandler

class StructuredLogger:
    """구조화된 로깅 시스템"""
    
    def __init__(self, name="eims_automation"):
        self.logger = logging.getLogger(name)
        self._setup_logger()
    
    def _setup_logger(self):
        """로거 설정"""
        if self.logger.handlers:
            return  # 이미 설정됨
        
        # 로그 디렉토리 생성
        log_dir = Path("storage/logs")
        log_dir.mkdir(exist_ok=True)
        
        # 로그 레벨 설정
        self.logger.setLevel(logging.INFO)
        
        # 파일 핸들러 (회전 로그)
        file_handler = RotatingFileHandler(
            log_dir / "app.log",
            maxBytes=10*1024*1024,  # 10MB
            backupCount=5,
            encoding='utf-8'
        )
        
        # 콘솔 핸들러
        console_handler = logging.StreamHandler()
        
        # 포맷터
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
    
    def log_event(self, event_type, **kwargs):
        """구조화된 이벤트 로깅"""
        log_data = {
            "event_type": event_type,
            "timestamp": datetime.now().isoformat(),
            **kwargs
        }
        
        self.logger.info(json.dumps(log_data, ensure_ascii=False))
    
    def log_assignment(self, researcher, count, items):
        """배정 로그"""
        self.log_event(
            "assignment_completed",
            researcher=researcher,
            count=count,
            items=items
        )
    
    def log_download(self, filename, size, rows):
        """다운로드 로그"""
        self.log_event(
            "excel_downloaded",
            filename=filename,
            size_bytes=size,
            rows_count=rows
        )
    
    def log_error(self, error_type, error_message, **context):
        """에러 로그"""
        self.log_event(
            "error_occurred",
            error_type=error_type,
            error_message=str(error_message),
            **context
        )

# 전역 로거 인스턴스
logger = StructuredLogger()

# 사용 예시
if __name__ == "__main__":
    logger.log_event("system_startup", version="1.0")
    logger.log_assignment("김연구원", 5, ["수은", "페놀"])
    logger.log_download("test.xlsx", 12345, 100)
    logger.log_error("validation_error", "필수 컬럼 누락", column="item")
