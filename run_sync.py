import sys
import pathlib
sys.path.append('scripts')

from sync_to_shared import SharedFolderSync

print("=== 공유폴더 동기화 실행 ===")
try:
    sync = SharedFolderSync()
    sync.sync_all()
    print("공유폴더 동기화 완료!")
except Exception as e:
    print(f"공유폴더 동기화 오류: {e}")
    import traceback
    traceback.print_exc()
