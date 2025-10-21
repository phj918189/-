#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
작업 완료 처리 스크립트
연구원이 작업을 완료했을 때 실행하는 스크립트
"""

import os
import sys
import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime
import logging

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class TaskCompletionHandler:
    def __init__(self, shared_path=r"\\samyang\homes\SAMYANG\Drive\SAMYANG\연구분석\공통분석실\데이터정리\1_TEST"):
        self.shared_path = Path(shared_path)
        self.local_project = Path(__file__).parent.parent
        self.db_path = self.local_project / "storage" / "lab.db"
    
    def get_researcher_name(self):
        """연구원 이름을 입력받거나 자동 감지"""
        # 공유폴더에서 현재 사용자 폴더 찾기
        researcher_folders = [f for f in self.shared_path.iterdir() if f.is_dir() and f.name != "system"]
        
        if len(researcher_folders) == 1:
            return researcher_folders[0].name
        
        print("=== 작업 완료 처리 ===")
        print("연구원을 선택하세요:")
        for i, folder in enumerate(researcher_folders, 1):
            print(f"{i}. {folder.name}")
        
        try:
            choice = int(input("번호를 입력하세요: ")) - 1
            if 0 <= choice < len(researcher_folders):
                return researcher_folders[choice].name
            else:
                print("잘못된 선택입니다.")
                return None
        except ValueError:
            print("숫자를 입력해주세요.")
            return None
    
    def get_today_assignments(self, researcher):
        """오늘 배정된 작업 목록 조회"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            assignments = pd.read_sql_query("""
                SELECT 
                    a.sample_no,
                    a.item,
                    a.assigned_at,
                    s.site_name,
                    s.collected_at
                FROM assignments a
                JOIN samples s ON a.sample_no = s.sample_no
                WHERE a.researcher = ? 
                AND date(a.assigned_at) = date('now', 'localtime')
                ORDER BY a.assigned_at
            """, conn, params=(researcher,))
            
            conn.close()
            return assignments
            
        except Exception as e:
            logger.error(f"작업 목록 조회 실패: {e}")
            return pd.DataFrame()
    
    def show_assignments(self, assignments):
        """작업 목록 표시"""
        if assignments.empty:
            print("오늘 배정된 작업이 없습니다.")
            return []
        
        print(f"\n=== 오늘 배정된 작업 목록 ({len(assignments)}건) ===")
        print("번호 | 샘플번호 | 측정항목 | 사업장명")
        print("-" * 60)
        
        for i, (_, row) in enumerate(assignments.iterrows(), 1):
            print(f"{i:2d}   | {row['sample_no']:10s} | {row['item'][:20]:20s} | {row['site_name'][:15]:15s}")
        
        return assignments
    
    def select_completed_tasks(self, assignments):
        """완료할 작업 선택"""
        if assignments.empty:
            return []
        
        print("\n완료할 작업의 번호를 입력하세요 (여러 개는 쉼표로 구분, 예: 1,3,5):")
        print("전체 완료: 'all' 입력")
        print("취소: 'cancel' 입력")
        
        choice = input("선택: ").strip().lower()
        
        if choice == 'cancel':
            return []
        elif choice == 'all':
            return assignments.index.tolist()
        else:
            try:
                indices = [int(x.strip()) - 1 for x in choice.split(',')]
                valid_indices = [i for i in indices if 0 <= i < len(assignments)]
                return [assignments.index[i] for i in valid_indices]
            except ValueError:
                print("잘못된 입력입니다.")
                return []
    
    def mark_tasks_completed(self, researcher, completed_indices, assignments):
        """작업 완료 처리"""
        if not completed_indices:
            print("완료할 작업이 선택되지 않았습니다.")
            return
        
        try:
            conn = sqlite3.connect(self.db_path)
            
            completed_count = 0
            for idx in completed_indices:
                row = assignments.loc[idx]
                
                # 데이터베이스에 완료 상태 업데이트
                conn.execute("""
                    UPDATE assignments 
                    SET status = 'completed', completed_at = ?
                    WHERE sample_no = ? AND item = ? AND researcher = ?
                """, (datetime.now(), row['sample_no'], row['item'], researcher))
                
                completed_count += 1
                print(f"✅ 완료 처리: {row['sample_no']} - {row['item']}")
            
            conn.commit()
            conn.close()
            
            print(f"\n🎉 {completed_count}건의 작업이 완료 처리되었습니다!")
            
            # 공유폴더에 완료 파일 생성
            self.create_completion_report(researcher, completed_indices, assignments)
            
        except Exception as e:
            logger.error(f"작업 완료 처리 실패: {e}")
            print(f"오류가 발생했습니다: {e}")
    
    def create_completion_report(self, researcher, completed_indices, assignments):
        """완료 보고서 생성"""
        try:
            researcher_folder = self.shared_path / researcher
            completed_folder = researcher_folder / "completed"
            completed_folder.mkdir(exist_ok=True)
            
            # 오늘 날짜 폴더 생성
            today_folder = completed_folder / datetime.now().strftime('%Y%m%d')
            today_folder.mkdir(exist_ok=True)
            
            # 완료된 작업만 추출
            completed_tasks = assignments.loc[completed_indices]
            
            # 완료 보고서 생성
            report_path = today_folder / f"completed_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
            completed_tasks.to_csv(report_path, index=False, encoding='utf-8-sig')
            
            # 완료 요약 생성
            summary_path = today_folder / f"completion_summary_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
            with open(summary_path, 'w', encoding='utf-8') as f:
                f.write(f"=== {researcher} 작업 완료 보고서 ===\n")
                f.write(f"완료일: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
                f.write(f"완료 작업 수: {len(completed_tasks)}건\n\n")
                
                f.write("📋 완료된 작업 목록:\n")
                for _, row in completed_tasks.iterrows():
                    f.write(f"  - {row['sample_no']}: {row['item']} ({row['site_name']})\n")
            
            print(f"📄 완료 보고서가 생성되었습니다: {report_path}")
            
        except Exception as e:
            logger.error(f"완료 보고서 생성 실패: {e}")
    
    def run(self):
        """메인 실행 함수"""
        try:
            print("=== EIMS 작업 완료 처리 시스템 ===")
            
            # 1. 연구원 선택
            researcher = self.get_researcher_name()
            if not researcher:
                print("연구원을 선택하지 않았습니다.")
                return
            
            print(f"선택된 연구원: {researcher}")
            
            # 2. 오늘 작업 목록 조회
            assignments = self.get_today_assignments(researcher)
            
            # 3. 작업 목록 표시
            self.show_assignments(assignments)
            
            # 4. 완료할 작업 선택
            completed_indices = self.select_completed_tasks(assignments)
            
            # 5. 작업 완료 처리
            self.mark_tasks_completed(researcher, completed_indices, assignments)
            
        except Exception as e:
            logger.error(f"작업 완료 처리 실패: {e}")
            print(f"오류가 발생했습니다: {e}")

def main():
    """메인 실행 함수"""
    try:
        handler = TaskCompletionHandler()
        handler.run()
        
    except KeyboardInterrupt:
        print("\n작업이 취소되었습니다.")
    except Exception as e:
        logger.error(f"스크립트 실행 실패: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())
