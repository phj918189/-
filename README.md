# 삼양자동화 프로젝트

EIMS 시스템에서 현장수질 데이터를 자동으로 다운로드하고 분석하는 자동화 프로젝트입니다.

## 🚀 주요 기능

- **자동 로그인**: EIMS 시스템에 자동으로 로그인
- **데이터 다운로드**: 현장수질 데이터를 엑셀 파일로 자동 다운로드
- **데이터 정규화**: 다운로드된 데이터를 표준 형식으로 변환
- **작업 배정**: 연구원별로 작업을 자동 배정
- **알림 발송**: 배정된 작업을 이메일로 알림

## 📁 프로젝트 구조

```
삼양자동화/
├── scripts/           # 핵심 스크립트
│   ├── scrape_eims.py # 웹 스크래핑 (로그인, 다운로드)
│   ├── normalize.py   # 데이터 정규화
│   ├── assign.py      # 작업 배정
│   ├── notify.py      # 알림 발송
│   └── db.py          # 데이터베이스 관리
├── sql/               # SQL 관련 파일
│   ├── job.py         # 메인 실행 파일
│   ├── schema.sql     # 데이터베이스 스키마
│   ├── item_rules.csv # 항목별 배정 규칙
│   └── researchers.csv# 연구원 정보
├── storage/           # 저장소
│   ├── downloads/     # 다운로드 파일
│   ├── logs/          # 로그 파일
│   └── lab.db         # SQLite 데이터베이스
└── .env               # 환경 변수 (보안상 중요)
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

### 3. 실행
```bash
python sql/job.py
```

## 📋 사용된 기술

- **Python 3.x**
- **Playwright**: 웹 자동화
- **Pandas**: 데이터 처리
- **SQLite**: 데이터 저장
- **SMTP**: 이메일 알림

## ⚠️ 주의사항

- `.env` 파일에는 실제 로그인 정보가 포함되므로 절대 공개하지 마세요
- 웹 스크래핑 시 해당 사이트의 이용약관을 준수하세요
- 개인정보 보호를 위해 로그인 정보는 안전하게 관리하세요

## 📞 문의

프로젝트 관련 문의사항이 있으시면 이슈를 등록해주세요.
