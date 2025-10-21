import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path

def cleanup_old_files():
    """오래된 파일들을 정리하는 함수"""
    
    # 설정
    DOWNLOADS_DIR = Path("storage/downloads")
    ARCHIVE_DIR = Path("storage/archive")
    CSV_DIR = Path("storage/csv_exports")
    
    # 아카이브 디렉토리 생성
    ARCHIVE_DIR.mkdir(exist_ok=True)
    CSV_DIR.mkdir(exist_ok=True)
    
    # 7일 이상 된 엑셀 파일을 아카이브로 이동
    cutoff_date = datetime.now() - timedelta(days=7)
    
    moved_files = []
    for file_path in DOWNLOADS_DIR.glob("*.xlsx"):
        if file_path.stat().st_mtime < cutoff_date.timestamp():
            archive_path = ARCHIVE_DIR / file_path.name
            shutil.move(str(file_path), str(archive_path))
            moved_files.append(file_path.name)
    
    # CSV 파일들을 별도 폴더로 이동
    for csv_file in Path(".").glob("assignments_*.csv"):
        csv_dest = CSV_DIR / csv_file.name
        shutil.move(str(csv_file), str(csv_dest))
        moved_files.append(csv_file.name)
    
    print(f"✅ {len(moved_files)}개 파일을 정리했습니다.")
    if moved_files:
        print("이동된 파일들:")
        for file in moved_files:
            print(f"  - {file}")

if __name__ == "__main__":
    cleanup_old_files()
