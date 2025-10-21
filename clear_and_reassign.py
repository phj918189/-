import sqlite3
import sys
import pathlib
sys.path.append('scripts')

# 1. 이전 배정 데이터 삭제
print("=== 이전 배정 데이터 삭제 ===")
conn = sqlite3.connect('storage/lab.db')
cursor = conn.cursor()

# 현재 배정 현황 확인
print("삭제 전 배정 현황:")
cursor.execute('SELECT DISTINCT researcher, COUNT(*) as count FROM assignments GROUP BY researcher ORDER BY researcher')
results = cursor.fetchall()
for researcher, count in results:
    print(f"  {researcher}: {count}개")

# 모든 배정 데이터 삭제
cursor.execute('DELETE FROM assignments')
conn.commit()
print("이전 배정 데이터 삭제 완료!")

conn.close()

# 2. 새로운 배정 실행
print()
print("=== 새로운 배정 실행 ===")
from assign import run_assign
from db import get_samples_df

samples_df = get_samples_df()
if not samples_df.empty:
    result = run_assign(samples_df)
    print(f"새로운 배정 완료: {len(result)}개")
    
    # 배정 결과 요약
    from collections import Counter
    researcher_counts = Counter([a['researcher'] for a in result])
    for researcher, count in researcher_counts.items():
        print(f"  {researcher}: {count}개")
else:
    print("샘플 데이터가 없습니다.")

print()
print("=== 완료! 이제 공유폴더 동기화를 실행하세요 ===")
