# ZIP 파일 배포 가이드

외부 사용자가 ZIP 파일을 다운로드받아 사용할 수 있도록 준비하는 방법입니다.

## 📦 ZIP 파일에 포함할 파일 목록

### 필수 파일
- ✅ `excel_compare_app.py` - 메인 애플리케이션
- ✅ `requirements.txt` - 패키지 목록
- ✅ `README.md` - 사용 설명서
- ✅ `설치_및_실행_가이드.md` - 상세 설치 가이드
- ✅ `INSTALL.bat` - 자동 설치 스크립트 (Windows)
- ✅ `RUN.bat` - 실행 스크립트 (Windows)

### 선택 파일
- ⚠️ `dist/excel_compare_app.exe` - 실행 파일 (Python 설치 불필요)
- ⚠️ `.gitignore` - Git 사용자용
- ⚠️ `.gitattributes` - Git 사용자용

### 제외할 파일
- ❌ `__pycache__/` - Python 캐시
- ❌ `build/` - 빌드 임시 파일
- ❌ `.venv/` - 가상환경
- ❌ `excel_compare_config.json` - 개인 설정 (선택)
- ❌ `결과파일/` - 사용자 결과 파일
- ❌ `원본파일/` - 예제 파일 (용량이 크면 제외)

## 🗜️ ZIP 파일 생성 방법

### 방법 1: Windows 탐색기 사용

1. `excel_compare` 폴더로 이동
2. 필요한 파일만 선택 (Ctrl + 클릭)
3. 우클릭 → "압축" → "ZIP 폴더로 압축"

### 방법 2: PowerShell 사용

```powershell
# excel_compare 폴더에서 실행
Compress-Archive -Path excel_compare_app.py,requirements.txt,README.md,설치_및_실행_가이드.md,INSTALL.bat,RUN.bat,dist -DestinationPath excel_compare_v1.0.zip
```

### 방법 3: Git 사용 (GitHub Releases)

```bash
# Git 태그 생성
git tag v1.0.0
git push origin v1.0.0

# GitHub에서 Releases 탭 → New release
# 태그 선택 후 "Source code (zip)" 자동 생성됨
```

## 📋 ZIP 파일 구조 예시

```
excel_compare_v1.0.zip
├── excel_compare_app.py
├── requirements.txt
├── README.md
├── 설치_및_실행_가이드.md
├── INSTALL.bat
├── RUN.bat
└── dist/
    └── excel_compare_app.exe
```

## 🚀 사용자 배포 시 안내사항

ZIP 파일과 함께 다음 안내를 제공하세요:

### 간단한 안내 (README.md에 포함됨)

```
1. ZIP 파일을 원하는 위치에 압축 해제
2. INSTALL.bat 더블클릭 (자동 설치)
3. RUN.bat 더블클릭 (앱 실행)
또는
dist/excel_compare_app.exe 더블클릭 (Python 설치 불필요)
```

### 상세 안내 (설치_및_실행_가이드.md 참조)

- Python 설치 방법
- 패키지 설치 방법
- 문제 해결 방법

## ✅ 배포 전 체크리스트

- [ ] 모든 필수 파일이 포함되어 있는가?
- [ ] 불필요한 파일(캐시, 빌드 파일 등)이 제외되었는가?
- [ ] README.md가 최신 상태인가?
- [ ] EXE 파일이 최신 빌드인가? (포함하는 경우)
- [ ] ZIP 파일 크기가 적절한가? (100MB 이하 권장)
- [ ] 다른 PC에서 테스트해봤는가?

## 🔄 버전 관리

ZIP 파일명에 버전 번호를 포함하세요:
- `excel_compare_v1.0.zip`
- `excel_compare_v1.1.zip`

버전 변경 시:
1. `README.md`의 버전 정보 업데이트
2. Git 태그 생성 (`git tag v1.1`)
3. 새 ZIP 파일 생성
4. GitHub Releases에 업로드

