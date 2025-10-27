#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
공유폴더 동기화 스크립트
로컬에서 생성된 연구원별 작업 파일들을 공유폴더에 복사
"""

import os
import shutil
import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime
import logging

# 로깅 설정
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class SharedFolderSync:
    def __init__(self, shared_path=r"\\samyang\homes\SAMYANG\Drive\SAMYANG\연구분석\공통분석실\데이터정리\1_TEST"):
        self.shared_path = Path(shared_path)
        self.local_project = Path(__file__).parent.parent
        self.db_path = self.local_project / "storage" / "lab.db"
        
        # 공유폴더 연결 테스트
        self._test_shared_folder()
    
    def _test_shared_folder(self):
        """공유폴더 접근 가능 여부 테스트"""
        try:
            # 공유폴더 경로 존재 확인
            if not self.shared_path.exists():
                logger.warning(f"공유폴더 경로가 존재하지 않습니다: {self.shared_path}")
                logger.info("공유폴더를 생성합니다...")
                self.shared_path.mkdir(parents=True, exist_ok=True)
            
            # 쓰기 권한 테스트
            test_file = self.shared_path / "test_write.tmp"
            test_file.write_text("test")
            test_file.unlink()
            
            logger.info(f"공유폴더 접근 성공: {self.shared_path}")
            
        except Exception as e:
            logger.error(f"공유폴더 접근 실패: {e}")
            raise
    
    def get_researchers(self):
        """item_rules.csv에서 연구원 목록 조회 (주요 소스)"""
        researchers = set()
        
        # 1) item_rules.csv에서 연구원 목록 조회 (주요 소스)
        try:
            rules_path = self.local_project / "sql" / "item_rules.csv"
            if rules_path.exists():
                rules_df = pd.read_csv(rules_path)
                for _, row in rules_df.iterrows():
                    if pd.notna(row.get('preferred')):
                        researchers.add(row['preferred'])
                logger.info(f"item_rules.csv에서 연구원 조회: {len(researchers)}명")
        except Exception as e:
            logger.error(f"item_rules.csv 조회 실패: {e}")
        
        # 2) 데이터베이스에서 배정된 연구원도 추가 (보조) - 테스트용 이름들 제외
        try:
            conn = sqlite3.connect(self.db_path)
            db_researchers = pd.read_sql_query("""
                SELECT DISTINCT researcher 
                FROM assignments 
                WHERE researcher IS NOT NULL 
                AND researcher != 'system'
                AND researcher NOT IN ('김연구원', '박연구원', '이연구원')
                ORDER BY researcher
            """, conn)
            conn.close()
            
            for researcher in db_researchers['researcher'].tolist():
                researchers.add(researcher)
            
        except Exception as e:
            logger.error(f"데이터베이스 연구원 목록 조회 실패: {e}")
        
        researcher_list = sorted(list(researchers))
        logger.info(f"최종 연구원 목록: {researcher_list}")
        return researcher_list
    
    def create_researcher_folders(self):
        """연구원별 폴더 구조 생성"""
        researchers = self.get_researchers()
        
        # 테스트용 폴더들 삭제 (필요시)
        test_names = ['김연구원', '박연구원', '이연구원']
        for test_name in test_names:
            test_folder = self.shared_path / test_name
            if test_folder.exists():
                logger.info(f"테스트 폴더 삭제: {test_folder}")
                shutil.rmtree(test_folder, ignore_errors=True)
        
        for researcher in researchers:
            researcher_folder = self.shared_path / researcher
            researcher_folder.mkdir(parents=True, exist_ok=True)
            
            # 하위 폴더들 생성
            subfolders = ["today", "pending", "completed", "reports"]
            for subfolder in subfolders:
                (researcher_folder / subfolder).mkdir(exist_ok=True)
            
            # 연구원별 안내 파일 생성
            readme_path = researcher_folder / "README.md"
            if not readme_path.exists():
                readme_content = f"""# {researcher} 담당 작업 폴더

## 📁 폴더 구조
- **today/**: 오늘 배정된 작업
- **pending/**: 대기 중인 작업  
- **completed/**: 완료된 작업
- **reports/**: 분석 보고서

## 📋 사용 방법
1. `today/` 폴더에서 오늘 배정된 작업 확인
2. 작업 파일 다운로드하여 분석 작업 수행
3. 완료된 작업은 `completed/` 폴더에 보고서 저장
4. 필요시 `reports/` 폴더에 추가 보고서 저장

## 📝 작업 완료 처리
- 작업 완료 상태는 관리자가 처리합니다
- 완료된 작업 보고서를 `completed/` 폴더에 저장해주세요

## 📞 문의
작업 관련 문의사항이 있으면 담당 관리자에게 연락하세요.

---
*자동 생성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}*
"""
                readme_path.write_text(readme_content, encoding='utf-8')
            
            logger.info(f"연구원 폴더 생성 완료: {researcher}")
    
    def create_today_assignments(self):
        """오늘 배정된 작업을 연구원별로 파일 생성"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            # 오늘 배정된 작업 조회
            assignments = pd.read_sql_query("""
                SELECT 
                    a.researcher,
                    a.sample_no,
                    a.item,
                    a.assigned_at,
                    s.site_name,
                    s.collected_at,
                    s.status
                FROM assignments a
                JOIN samples s ON a.sample_no = s.sample_no
                WHERE date(a.assigned_at) = date('now', 'localtime')
                ORDER BY a.researcher, a.assigned_at
            """, conn)
            
            conn.close()
            
            if assignments.empty:
                logger.info("오늘 배정된 작업이 없습니다.")
                return
            
            # 연구원별로 그룹화하여 파일 생성
            for researcher in assignments['researcher'].unique():
                researcher_data = assignments[assignments['researcher'] == researcher]
                researcher_folder = self.shared_path / researcher / "today"
                
                # 오늘 날짜로 폴더 정리
                today_folder = researcher_folder / datetime.now().strftime('%Y%m%d')
                today_folder.mkdir(parents=True, exist_ok=True)
                
                # CSV 파일 생성
                csv_path = today_folder / f"assignments_{datetime.now().strftime('%Y%m%d')}.csv"
                researcher_data.to_csv(csv_path, index=False, encoding='utf-8-sig')
                
                # Excel 파일 생성
                excel_path = today_folder / f"assignments_{datetime.now().strftime('%Y%m%d')}.xlsx"
                researcher_data.to_excel(excel_path, index=False, engine='openpyxl')
                
                # 요약 보고서 생성
                summary_path = today_folder / f"summary_{datetime.now().strftime('%Y%m%d')}.txt"
                with open(summary_path, 'w', encoding='utf-8') as f:
                    f.write(f"=== {researcher} 담당 작업 요약 ===\n")
                    f.write(f"배정일: {datetime.now().strftime('%Y-%m-%d')}\n")
                    f.write(f"총 작업 수: {len(researcher_data)}건\n\n")
                    
                    # 측정항목별 통계
                    item_counts = researcher_data['item'].value_counts()
                    f.write("📊 측정항목별 작업 수:\n")
                    for item, count in item_counts.items():
                        f.write(f"  - {item}: {count}건\n")
                    
                    f.write("\n📋 상세 작업 목록:\n")
                    for _, row in researcher_data.iterrows():
                        f.write(f"  - {row['sample_no']}: {row['item']} ({row['site_name']})\n")
                
                logger.info(f"{researcher} 작업 파일 생성 완료: {len(researcher_data)}건")
            
        except Exception as e:
            logger.error(f"오늘 작업 파일 생성 실패: {e}")
    
    def create_dashboard(self):
        """전체 현황 대시보드 생성"""
        try:
            conn = sqlite3.connect(self.db_path)
            
            # 전체 현황 조회
            dashboard_data = pd.read_sql_query("""
                SELECT 
                    researcher,
                    COUNT(*) as total_assignments,
                    COUNT(CASE WHEN date(assigned_at) = date('now', 'localtime') THEN 1 END) as today_assignments
                FROM assignments 
                GROUP BY researcher
                ORDER BY total_assignments DESC
            """, conn)
            
            conn.close()
            
            # HTML 대시보드 생성
            html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>EIMS 작업 배정 현황</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th {{ text-align: center; }}
        td {{ text-align: center; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; }}
        .today {{ background-color: #e8f5e8; }}
        .header {{ background-color: #4CAF50; color: white; padding: 20px; text-align: center; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 EIMS 작업 배정 현황</h1>
        <p>Environmental Information Management System</p>
        <p>업데이트: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
    </div>
    
    <h2>👥 담당자별 현황</h2>
    <table>
        <tr>
            <th>담당자</th>
            <th>총 작업 수</th>
            <th>오늘 작업 수</th>
            <th>작업 폴더</th>
        </tr>
"""
            
            for _, row in dashboard_data.iterrows():
                researcher = row['researcher']
                html_content += f"""
        <tr class="{'today' if row['today_assignments'] > 0 else ''}">
            <td>{researcher}</td>
            <td>{row['total_assignments']}건</td>
            <td>{row['today_assignments']}건</td>
            <td><a href="{researcher}/">📁 폴더 열기</a></td>
        </tr>
"""
            
            html_content += """
    </table>
    
    <h2>📋 사용 안내</h2>
    <ul>
        <li>각 담당자 폴더에서 오늘 배정된 작업을 확인하세요</li>
        <!-- <li>작업 완료 후 completed 폴더로 이동하세요</li> -->
        <li>문의사항이 있으면 담당 관리자에게 연락하세요</li>
    </ul>
</body>
</html>
"""
            
            # HTML 파일 저장
            dashboard_path = self.shared_path / "dashboard.html"
            with open(dashboard_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            logger.info("대시보드 생성 완료")
            
        except Exception as e:
            logger.error(f"대시보드 생성 실패: {e}")
    
    def sync_all(self):
        """전체 동기화 실행"""
        logger.info("=== 공유폴더 동기화 시작 ===")
        
        try:
            # 1. 연구원별 폴더 생성
            self.create_researcher_folders()
            
            # 2. 오늘 작업 파일 생성
            self.create_today_assignments()
            
            # 3. 대시보드 생성
            self.create_dashboard()
            
            logger.info("=== 공유폴더 동기화 완료 ===")
            
        except Exception as e:
            logger.error(f"동기화 실패: {e}")
            raise

def main():
    """메인 실행 함수"""
    try:
        sync = SharedFolderSync()
        sync.sync_all()
        
    except Exception as e:
        logger.error(f"스크립트 실행 실패: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())
