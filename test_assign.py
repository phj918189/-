import sys
import pathlib
sys.path.append('scripts')

from assign import run_assign
from db import get_samples_df
import pandas as pd

# 간단한 테스트 데이터
test_data = pd.DataFrame({
    'sample_no': ['S001', 'S002', 'S003'],
    'item': ['총질소', '총인', '부유물질']
})

print("=== 배정 로직 테스트 ===")
result = run_assign(test_data)
print(f"배정 결과: {len(result)}개")
for assignment in result:
    print(f"  {assignment['item']} -> {assignment['researcher']}")
