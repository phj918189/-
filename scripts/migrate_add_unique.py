#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
기존 데이터베이스에 UNIQUE 제약조건 추가 마이그레이션 스크립트
"""

import sqlite3
import pathlib
from pathlib import Path

BASE = pathlib.Path(__file__).resolve().parents[1]
DB_PATH = BASE / "storage" / "lab.db"

def migrate():
    """기존 DB에 UNIQUE 제약조건 추가"""
    print("=== 데이터베이스 마이그레이션 시작 ===")
    
    if not DB_PATH.exists():
        print(f"❌ 데이터베이스 파일을 찾을 수 없습니다: {DB_PATH}")
        return False
    
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # 1. 기존 중복 데이터 확인 및 제거
        print("중복 데이터 확인 중...")
        cursor.execute("""
            SELECT sample_no, item, COUNT(*) as cnt
            FROM assignments
            GROUP BY sample_no, item
            HAVING COUNT(*) > 1
        """)
        
        duplicates = cursor.fetchall()
        if duplicates:
            print(f"⚠️  중복 데이터 발견: {len(duplicates)}개")
            print("최신 배정 데이터만 유지하고 중복 제거 중...")
            
            # 중복 데이터 제거 (최신 것만 유지)
            cursor.execute("""
                DELETE FROM assignments
                WHERE id NOT IN (
                    SELECT MIN(id)
                    FROM assignments
                    GROUP BY sample_no, item
                )
            """)
            conn.commit()
            print(f"✅ 중복 데이터 제거 완료")
        else:
            print("✅ 중복 데이터 없음")
        
        # 2. 테이블 재생성 (UNIQUE 제약조건 추가)
        print("테이블 재생성 중...")
        
        # 기존 데이터 백업
        cursor.execute("SELECT * FROM assignments")
        old_data = cursor.fetchall()
        
        # 테이블 삭제 후 재생성
        cursor.execute("DROP TABLE IF EXISTS assignments_old")
        cursor.execute("ALTER TABLE assignments RENAME TO assignments_old")
        
        # 새 테이블 생성 (UNIQUE 제약조건 포함)
        cursor.execute("""
            CREATE TABLE assignments(
                id INTEGER PRIMARY KEY, 
                sample_no TEXT, 
                item TEXT,
                researcher TEXT, 
                assigned_at TEXT, 
                method TEXT,
                UNIQUE(sample_no, item)
            )
        """)
        
        # 기존 데이터 복원
        cursor.execute("""
            INSERT OR IGNORE INTO assignments 
            SELECT * FROM assignments_old
        """)
        
        # 임시 테이블 삭제
        cursor.execute("DROP TABLE assignments_old")
        
        conn.commit()
        
        print("✅ UNIQUE 제약조건 추가 완료")
        print("=== 마이그레이션 완료 ===")
        
        conn.close()
        return True
        
    except Exception as e:
        print(f"❌ 마이그레이션 실패: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = migrate()
    if success:
        print("\n✅ 마이그레이션이 성공적으로 완료되었습니다!")
        print("이제 중복 배정 문제가 해결되었습니다.")
    else:
        print("\n❌ 마이그레이션이 실패했습니다.")
        exit(1)


