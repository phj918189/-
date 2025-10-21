import sqlite3
import shutil
import gzip
from datetime import datetime, timedelta
from pathlib import Path

class DatabaseManager:
    """데이터베이스 관리 클래스"""
    
    def __init__(self, db_path="storage/lab.db"):
        self.db_path = Path(db_path)
        self.backup_dir = Path("storage/backups")
        self.backup_dir.mkdir(exist_ok=True)
    
    def backup_database(self, compress=True):
        """데이터베이스 백업"""
        if not self.db_path.exists():
            raise FileNotFoundError(f"데이터베이스 파일을 찾을 수 없습니다: {self.db_path}")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if compress:
            backup_filename = f"lab_backup_{timestamp}.db.gz"
            backup_path = self.backup_dir / backup_filename
            
            # 압축 백업
            with open(self.db_path, 'rb') as f_in:
                with gzip.open(backup_path, 'wb') as f_out:
                    shutil.copyfileobj(f_in, f_out)
        else:
            backup_filename = f"lab_backup_{timestamp}.db"
            backup_path = self.backup_dir / backup_filename
            shutil.copy2(self.db_path, backup_path)
        
        print(f"✅ 데이터베이스 백업 완료: {backup_path}")
        return backup_path
    
    def cleanup_old_backups(self, retention_days=30):
        """오래된 백업 파일 정리"""
        cutoff_date = datetime.now() - timedelta(days=retention_days)
        
        deleted_count = 0
        for backup_file in self.backup_dir.glob("lab_backup_*"):
            if backup_file.stat().st_mtime < cutoff_date.timestamp():
                backup_file.unlink()
                deleted_count += 1
        
        if deleted_count > 0:
            print(f"🗑️ 오래된 백업 파일 {deleted_count}개를 삭제했습니다.")
    
    def get_backup_info(self):
        """백업 파일 정보 조회"""
        backups = []
        for backup_file in self.backup_dir.glob("lab_backup_*"):
            backups.append({
                "filename": backup_file.name,
                "size": backup_file.stat().st_size,
                "created": datetime.fromtimestamp(backup_file.stat().st_mtime),
                "compressed": backup_file.suffix == '.gz'
            })
        
        return sorted(backups, key=lambda x: x["created"], reverse=True)
    
    def restore_database(self, backup_filename):
        """데이터베이스 복원"""
        backup_path = self.backup_dir / backup_filename
        
        if not backup_path.exists():
            raise FileNotFoundError(f"백업 파일을 찾을 수 없습니다: {backup_path}")
        
        # 현재 DB 백업
        current_backup = self.backup_database()
        
        try:
            if backup_path.suffix == '.gz':
                # 압축 해제 후 복원
                with gzip.open(backup_path, 'rb') as f_in:
                    with open(self.db_path, 'wb') as f_out:
                        shutil.copyfileobj(f_in, f_out)
            else:
                # 직접 복사
                shutil.copy2(backup_path, self.db_path)
            
            print(f"✅ 데이터베이스 복원 완료: {backup_filename}")
            print(f"📋 복원 전 백업: {current_backup}")
            
        except Exception as e:
            print(f"❌ 복원 실패: {e}")
            # 복원 전 상태로 되돌리기
            if current_backup.exists():
                shutil.copy2(current_backup, self.db_path)
                print("🔄 복원 전 상태로 되돌렸습니다.")

# 사용 예시
if __name__ == "__main__":
    db_manager = DatabaseManager()
    
    # 백업 생성
    backup_path = db_manager.backup_database()
    
    # 백업 정보 조회
    backups = db_manager.get_backup_info()
    print("\n📋 백업 파일 목록:")
    for backup in backups[:5]:  # 최근 5개만 표시
        print(f"  - {backup['filename']} ({backup['size']} bytes, {backup['created']})")
    
    # 오래된 백업 정리
    db_manager.cleanup_old_backups()
