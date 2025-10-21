#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EIMS 시스템 전체 진단 및 수정 스크립트
"""

import sqlite3
import sys
import pathlib
from collections import Counter

# 프로젝트 루트 경로 추가
ROOT = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

def main():
    print("=== EIMS 시스템 전체 진단 및 수정 ===")
    print()
    
    # 1단계: 데이터베이스 현황 확인
    print("[1단계] 데이터베이스 현황 확인")
    try:
        conn = sqlite3.connect('storage/lab.db')
        cursor = conn.cursor()
        
        print("=== 배정 현황 ===")
        cursor.execute('SELECT DISTINCT researcher, COUNT(*) as count FROM assignments GROUP BY researcher ORDER BY researcher')
        results = cursor.fetchall()
        for researcher, count in results:
            print(f"{researcher}: {count}개")
        
        print()
        print("=== 샘플 데이터 ===")
        cursor.execute('SELECT COUNT(*) as count FROM samples')
        sample_count = cursor.fetchone()[0]
        print(f"총 샘플 수: {sample_count}개")
        conn.close()
        
    except Exception as e:
        print(f"데이터베이스 확인 오류: {e}")
        return
    
    print()
    
    # 2단계: 배정 로직 테스트
    print("[2단계] 배정 로직 테스트")
    try:
        from scripts.assign import _load_rules, _load_people
        import pandas as pd
        
        print("=== 연구원 목록 ===")
        people = _load_people()
        names = [p['name'] for p in people]
        print(f"연구원: {names}")
        
        print()
        print("=== 배정 규칙 ===")
        rules = _load_rules()
        for i, rule in enumerate(rules):
            pattern_preview = rule['item_pattern'][:50] + "..." if len(rule['item_pattern']) > 50 else rule['item_pattern']
            print(f"규칙 {i+1}: {rule['preferred']} - {pattern_preview}")
        
        print()
        print("=== 테스트 배정 ===")
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
            
            print(f"{item} -> {chosen}")
            
    except Exception as e:
        print(f"배정 로직 테스트 오류: {e}")
        return
    
    print()
    
    # 3단계: 공유폴더 동기화 테스트
    print("[3단계] 공유폴더 동기화 테스트")
    try:
        from scripts.sync_to_shared import SharedFolderSync
        
        sync = SharedFolderSync()
        researchers = sync.get_researchers()
        print(f"공유폴더에 생성될 연구원 폴더: {researchers}")
        
    except Exception as e:
        print(f"공유폴더 동기화 오류: {e}")
        return
    
    print()
    
    # 4단계: 문제 해결 실행
    print("[4단계] 문제 해결 실행")
    try:
        print("이전 배정 데이터 삭제 중...")
        conn = sqlite3.connect('storage/lab.db')
        cursor = conn.cursor()
        cursor.execute('DELETE FROM assignments')
        conn.commit()
        conn.close()
        print("삭제 완료!")
        
        print()
        print("새로운 배정 실행 중...")
        from scripts.assign import run_assign
        from scripts.db import get_samples_df
        
        samples_df = get_samples_df()
        if not samples_df.empty:
            result = run_assign(samples_df)
            print(f"새로운 배정 완료: {len(result)}개")
            
            # 배정 결과 요약
            researcher_counts = Counter([a['researcher'] for a in result])
            for researcher, count in researcher_counts.items():
                print(f"  {researcher}: {count}개")
        else:
            print("샘플 데이터가 없습니다.")
            
    except Exception as e:
        print(f"문제 해결 실행 오류: {e}")
        return
    
    print()
    
    # 5단계: 공유폴더 동기화 실행
    print("[5단계] 공유폴더 동기화 실행")
    try:
        from scripts.sync_to_shared import SharedFolderSync
        
        sync = SharedFolderSync()
        sync.sync_all()
        print("공유폴더 동기화 완료!")
        
    except Exception as e:
        print(f"공유폴더 동기화 실행 오류: {e}")
        return
    
    print()
    print("=== 모든 작업 완료! ===")

if __name__ == "__main__":
    main()
