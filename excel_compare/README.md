# 엑셀 비교 뷰어 (Excel Compare Viewer)

두 개의 엑셀 파일을 비교하여 차이점을 시각적으로 확인할 수 있는 GUI 애플리케이션입니다.

## 🎯 주요 기능

- **양방향 비교**: 두 엑셀 파일을 나란히 비교
- **차이점 하이라이트**: 다른 셀을 색상으로 표시
- **스크롤 동기화**: 양쪽 시트를 동시에 스크롤
- **선택 동기화**: 한쪽 셀 선택 시 다른 쪽도 자동 선택
- **다크 모드 지원**: 라이트/다크 테마 전환 가능
- **설정 저장**: 창 크기, 위치, 테마 등 설정 자동 저장
- **오래된 .xls 파일 지원**: 자동으로 .xlsx로 변환하여 처리

## 📋 시스템 요구사항

- **운영체제**: Windows 10 이상
- **Python**: 3.8 이상
- **필수 폰트**: 맑은 고딕 (Windows 기본 포함)

## 🚀 설치 방법

### 1. ZIP 파일 다운로드 및 압축 해제

GitHub에서 프로젝트를 다운로드하거나 ZIP 파일을 받아 원하는 위치에 압축을 해제하세요.

```
C:\Projects\excel_compare\
```

### 2. Python 설치 확인

PowerShell 또는 명령 프롬프트를 열고 다음 명령어로 Python이 설치되어 있는지 확인하세요:

```powershell
python --version
```

Python이 설치되어 있지 않다면 [python.org](https://www.python.org/downloads/)에서 다운로드하여 설치하세요.

### 3. 가상환경 생성 및 활성화 (권장)

```powershell
# 프로젝트 폴더로 이동
cd C:\Projects\excel_compare

# 가상환경 생성
python -m venv .venv

# 가상환경 활성화 (PowerShell)
.\.venv\Scripts\Activate.ps1

# 만약 실행 정책 오류가 발생하면:
Set-ExecutionPolicy -Scope Process RemoteSigned
```

### 4. 패키지 설치

```powershell
# pip 업그레이드
python -m pip install --upgrade pip

# 필요한 패키지 설치
pip install -r requirements.txt
```

## 💻 실행 방법

### 방법 1: Python 스크립트로 실행

```powershell
# 가상환경이 활성화된 상태에서
python excel_compare_app.py
```

### 방법 2: 실행 파일 사용 (EXE)

`dist/excel_compare_app.exe` 파일을 더블클릭하여 실행할 수 있습니다.

> **참고**: EXE 파일은 별도로 Python 설치 없이 실행 가능하지만, 첫 실행 시 Windows Defender가 차단할 수 있습니다. "추가 정보" → "실행"을 클릭하세요.

## 📖 사용 방법

1. **첫 번째 파일 선택**: 상단의 "📂 첫 번째 파일" 버튼 클릭
2. **두 번째 파일 선택**: "📂 두 번째 파일" 버튼 클릭
3. **비교 시작**: "비교 시작" 버튼 클릭
4. **결과 확인**: 
   - 다른 셀은 초록색(라이트 모드) 또는 빨간색(다크 모드)으로 표시됩니다
   - 좌우 스크롤바를 움직이면 양쪽이 동시에 스크롤됩니다
   - 한쪽 셀을 클릭하면 다른 쪽도 자동으로 해당 위치로 이동합니다

### 추가 기능

- **테마 변경**: 상단의 "🌙 다크 모드" / "☀️ 라이트 모드" 버튼으로 테마 전환
- **차이점 내보내기**: 비교 후 "차이점 내보내기" 버튼으로 결과를 엑셀 파일로 저장
- **비교 취소**: 비교 중 "비교 취소" 버튼으로 중단 가능

## ⚙️ 설정 파일

앱을 실행하면 `excel_compare_config.json` 파일이 자동으로 생성됩니다. 이 파일에서 다음 설정을 변경할 수 있습니다:

- 창 크기 및 위치
- 테마 (라이트/다크)
- 색상 설정
- 폰트 설정

## 🐛 문제 해결

### "모듈을 찾을 수 없습니다" 오류

```powershell
# 패키지가 제대로 설치되었는지 확인
pip list

# requirements.txt의 패키지들을 다시 설치
pip install -r requirements.txt --force-reinstall
```

### Windows Defender가 EXE 파일을 차단하는 경우

1. Windows Defender 알림에서 "추가 정보" 클릭
2. "실행" 버튼 클릭
3. 또는 Windows 보안 설정에서 예외 추가

### 폰트 관련 오류

Windows에서 기본 제공되는 "맑은 고딕" 폰트가 필요합니다. 다른 폰트를 사용하려면 `excel_compare_config.json` 파일에서 `font.family` 값을 변경하세요.

### .xls 파일 읽기 오류

오래된 .xls 파일은 자동으로 .xlsx로 변환됩니다. 변환된 파일은 원본 파일과 같은 폴더에 `_converted.xlsx` 접미사가 붙어 저장됩니다.

## 📁 프로젝트 구조

```
excel_compare/
├── excel_compare_app.py          # 메인 애플리케이션 파일
├── requirements.txt               # 필요한 패키지 목록
├── README.md                      # 이 파일
├── excel_compare_config.json     # 설정 파일 (자동 생성)
├── dist/                          # 빌드된 실행 파일
│   └── excel_compare_app.exe
└── build/                         # 빌드 임시 파일
```

## 🔄 업데이트 방법

1. GitHub에서 최신 버전 다운로드
2. 기존 폴더를 새 버전으로 교체 (설정 파일은 유지)
3. 필요시 패키지 재설치:
   ```powershell
   pip install -r requirements.txt --upgrade
   ```

## 📝 라이선스

이 프로젝트는 개인 사용 목적으로 제작되었습니다.

## 💬 문의

문제가 발생하거나 개선 사항이 있으면 GitHub Issues에 등록해주세요.

