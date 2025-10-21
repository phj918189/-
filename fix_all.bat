@echo off
echo === EIMS 시스템 전체 진단 및 수정 ===
echo.

echo [1단계] 데이터베이스 현황 확인
python -c "
import sqlite3
conn = sqlite3.connect('storage/lab.db')
cursor = conn.cursor()

print('=== 배정 현황 ===')
cursor.execute('SELECT DISTINCT researcher, COUNT(*) as count FROM assignments GROUP BY researcher ORDER BY researcher')
results = cursor.fetchall()
for researcher, count in results:
    print(f'{researcher}: {count}개')

print()
print('=== 샘플 데이터 ===')
cursor.execute('SELECT COUNT(*) as count FROM samples')
sample_count = cursor.fetchone()[0]
print(f'총 샘플 수: {sample_count}개')
conn.close()
"

echo.
echo [2단계] 배정 로직 테스트
python -c "
import sys
sys.path.append('scripts')
from assign import _load_rules, _load_people
import pandas as pd

print('=== 연구원 목록 ===')
people = _load_people()
names = [p['name'] for p in people]
print(f'연구원: {names}')

print()
print('=== 배정 규칙 ===')
rules = _load_rules()
for i, rule in enumerate(rules):
    print(f'규칙 {i+1}: {rule[\"preferred\"]} - {rule[\"item_pattern\"][:50]}...')

print()
print('=== 테스트 배정 ===')
test_data = pd.DataFrame({
    'sample_no': ['S001', 'S002', 'S003'],
    'item': ['총질소', '총인', '부유물질']
})

for _, row in test_data.iterrows():
    item = str(row['item'])
    chosen = None
    
    for rule in sorted(rules, key=lambda x: int(x.get('priority','1'))):
        if rule['item_pattern']:
            patterns = rule['item_pattern'].split('|')
            for pattern in patterns:
                if pattern.strip() in item:
                    pref = rule.get('preferred')
                    if pref in names:
                        chosen = pref
                        break
            if chosen:
                break
    
    if not chosen:
        chosen = names[0] if names else 'system'
    
    print(f'{item} -> {chosen}')
"

echo.
echo [3단계] 공유폴더 동기화 테스트
python -c "
import sys
sys.path.append('scripts')
from sync_to_shared import SharedFolderSync

try:
    sync = SharedFolderSync()
    researchers = sync.get_researchers()
    print(f'공유폴더에 생성될 연구원 폴더: {researchers}')
except Exception as e:
    print(f'공유폴더 동기화 오류: {e}')
"

echo.
echo [4단계] 문제 해결 실행
python -c "
import sqlite3
import sys
sys.path.append('scripts')
from assign import run_assign
from db import get_samples_df

print('이전 배정 데이터 삭제 중...')
conn = sqlite3.connect('storage/lab.db')
cursor = conn.cursor()
cursor.execute('DELETE FROM assignments')
conn.commit()
conn.close()
print('삭제 완료!')

print()
print('새로운 배정 실행 중...')
samples_df = get_samples_df()
if not samples_df.empty:
    result = run_assign(samples_df)
    print(f'새로운 배정 완료: {len(result)}개')
    
    # 배정 결과 요약
    from collections import Counter
    researcher_counts = Counter([a['researcher'] for a in result])
    for researcher, count in researcher_counts.items():
        print(f'  {researcher}: {count}개')
else:
    print('샘플 데이터가 없습니다.')
"

echo.
echo [5단계] 공유폴더 동기화 실행
python scripts/sync_to_shared.py

echo.
echo === 완료! ===
pause
