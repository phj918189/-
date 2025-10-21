# 삼양자동화 프로젝트

EIMS 시스템에서 현장수질 데이터를 자동으로 다운로드하고 분석하는 자동화 프로젝트입니다.

## 🚀 주요 기능

- **자동 로그인**: EIMS 시스템에 자동으로 로그인
- **데이터 다운로드**: 현장수질 데이터를 엑셀 파일로 자동 다운로드
- **데이터 정규화**: 다운로드된 데이터를 표준 형식으로 변환
- **작업 배정**: 연구원별로 작업을 자동 배정
- **알림 발송**: 배정된 작업을 이메일로 알림
- **자동 백업**: 데이터베이스 자동 백업 및 복원
- **구조화된 로깅**: JSON 형태의 체계적 로그 관리
- **파일 관리**: 오래된 파일 자동 정리 및 아카이브

## 📁 프로젝트 구조

```
삼양자동화/
├── scripts/                    # 핵심 스크립트
│   ├── scrape_eims.py          # 웹 스크래핑 (로그인, 다운로드)
│   ├── normalize.py            # 데이터 정규화
│   ├── assign.py               # 작업 배정
│   ├── notify.py               # 알림 발송
│   ├── db.py                   # 데이터베이스 관리
│   ├── cleanup.py              # 파일 정리
│   ├── config_manager.py       # 설정 관리
│   ├── structured_logger.py    # 구조화된 로깅
│   └── database_manager.py     # 데이터베이스 백업
├── sql/                        # SQL 관련 파일
│   ├── job.py                  # 메인 실행 파일
│   ├── schema.sql              # 데이터베이스 스키마
│   ├── item_rules.csv          # 항목별 배정 규칙
│   ├── researchers.csv         # 연구원 정보
│   └── requirements.txt        # 의존성 패키지
├── storage/                    # 저장소
│   ├── downloads/              # 다운로드 파일
│   ├── logs/                   # 로그 파일
│   ├── backups/                # 데이터베이스 백업
│   ├── archive/                # 압축된 오래된 파일
│   ├── csv_exports/            # CSV 내보내기 파일
│   └── lab.db                  # SQLite 데이터베이스
├── config.yaml                 # 중앙화된 설정 파일
├── .env                        # 환경 변수 (보안상 중요)
└── README.md                   # 프로젝트 문서
```

## 🛠️ 설치 및 실행

### 1. 의존성 설치
```bash
pip install -r sql/requirements.txt
```

### 2. 환경 변수 설정
`.env` 파일을 생성하고 다음 정보를 입력하세요:
```env
EIMS_ID=your_username
EIMS_PW=your_password
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
MAIL_FROM=your_email@gmail.com
```

### 3. 설정 파일 생성 (선택사항)
`config.yaml` 파일을 생성하여 세부 설정을 관리할 수 있습니다:
```yaml
database:
  path: "storage/lab.db"
  backup_retention_days: 30

file_management:
  excel_retention_days: 7
  csv_retention_days: 30
  log_retention_days: 30
```

### 4. 실행
```bash
python sql/job.py
```

## 📋 사용된 기술

- **Python 3.x**
- **Playwright**: 웹 자동화
- **Pandas**: 데이터 처리
- **SQLite**: 데이터 저장
- **SMTP**: 이메일 알림
- **PyYAML**: 설정 파일 관리
- **Logging**: 구조화된 로깅

## 🔧 유틸리티 명령어

### 데이터베이스 관리
```bash
# 데이터베이스 백업
python scripts/database_manager.py

# 설정 확인
python scripts/config_manager.py

# 구조화된 로깅 테스트
python scripts/structured_logger.py
```

### 파일 정리
```bash
# 수동 파일 정리
python scripts/cleanup.py
```

## 📊 데이터 확인 방법

### 1. 데이터베이스 직접 조회
```bash
# SQLite 명령어로 접근
sqlite3 storage/lab.db

# SQLite 내에서 사용할 명령어들:
.tables                    # 테이블 목록 확인
.schema                    # 테이블 구조 확인
SELECT * FROM samples LIMIT 5;     # 샘플 데이터 확인
SELECT * FROM assignments LIMIT 5; # 배정 데이터 확인
.quit                      # 종료
```

### 2. 연구원별 배정 현황 확인
```python
import sqlite3
import pandas as pd

conn = sqlite3.connect('storage/lab.db')

# 연구원별 배정 현황
df = pd.read_sql_query("""
    SELECT researcher, COUNT(*) as 배정건수
    FROM assignments 
    GROUP BY researcher
    ORDER BY 배정건수 DESC
""", conn)
print(df)

conn.close()
```

## 🚨 문제 해결

### 데이터 표준화 실패
**증상**: `'DataFrame' object has no attribute 'dtype'` 에러
**해결**: `scripts/normalize.py`의 동의어 매핑 중복 문제 해결됨

### 이메일 발송 실패
**증상**: SMTP 연결 실패
**해결**: `.env` 파일의 SMTP 설정 확인 및 이메일 서비스 업그레이드 권장

### 파일 누적 문제
**증상**: 다운로드 파일과 CSV 파일이 계속 누적
**해결**: 자동 정리 시스템 구현됨 (`scripts/cleanup.py`)

## 📈 개선사항 로드맵

### 1단계 (완료) - 무료 개선사항
- ✅ 구조화된 로깅 시스템
- ✅ 설정 파일 관리 (YAML)
- ✅ 데이터베이스 자동 백업
- ✅ 파일 자동 정리

### 2단계 (3-6개월 후) - 저비용 개선사항
- 🔄 이메일 서비스 개선 (SendGrid)
- 🔄 간단한 웹 대시보드 구축
- 🔄 모니터링 시스템 구축

### 3단계 (6개월 후) - 중비용 개선사항
- 🔄 클라우드 백업 (필요 시)
- 🔄 고급 분석 기능 추가

## ⚠️ 주의사항

- `.env` 파일에는 실제 로그인 정보가 포함되므로 절대 공개하지 마세요
- 웹 스크래핑 시 해당 사이트의 이용약관을 준수하세요
- 개인정보 보호를 위해 로그인 정보는 안전하게 관리하세요
- 데이터베이스 백업은 정기적으로 확인하세요
- 파일 정리 시스템이 자동으로 작동하므로 중요한 파일은 별도 보관하세요

## 📞 문의

프로젝트 관련 문의사항이 있으시면 이슈를 등록해주세요.
