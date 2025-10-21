import sqlite3
import sys
import pathlib
sys.path.append('scripts')

from assign import run_assign
from db import get_samples_df

# 1. 이전 배정 데이터 삭제
print("이전 배정 데이터 삭제 중...")
conn = sqlite3.connect('storage/lab.db')
cursor = conn.cursor()
cursor.execute('DELETE FROM assignments')
conn.commit()
conn.close()
print("삭제 완료!")

# 2. 새로운 배정 실행
print()
print("새로운 배정 실행 중...")
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
