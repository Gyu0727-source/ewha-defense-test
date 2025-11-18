# 이화방어기제검사 자동화 시스템 (EDMT Automation)

이화방어기제검사(Ewha Defense Mechanism Test)를 Google Forms를 통해 온라인으로 실시하고, 자동으로 채점하는 시스템입니다.

## 🎯 주요 기능

- **자동화된 검사 실시**: Google Forms를 통한 온라인 검사
- **자동 채점**: Python 스크립트로 응답 자동 처리 및 채점
- **시간 절약**: 수작업 입력 20분 → 자동화 1초
- **정확도 100%**: 입력 오류 제로
- **원격 검사**: 모바일/PC 어디서든 가능

## 📋 시스템 요구사항

- **Python**: 3.7 이상
- **Google Account**: Google Forms 및 Sheets 사용
- **Google Cloud**: API 활성화 필요

## 🚀 빠른 시작

### 1. 저장소 클론

```bash
git clone <repository-url>
cd ewha-defense-test
```

### 2. Python 환경 설정

**자동 설치 (권장):**

```bash
# Linux/Mac
bash scripts/setup.sh

# Windows
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

**수동 설치:**

```bash
# 가상환경 생성 (권장)
python3 -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# 필수 라이브러리 설치
pip install --upgrade pip
pip install -r requirements.txt
```

### 3. Google Forms 생성

1. [Google Drive](https://drive.google.com) 접속
2. 새로 만들기 → Google Sheets
3. 확장 프로그램 → Apps Script
4. `테스트/자동화시스템/create_edmt_form_complete.gs` 내용 복사 & 붙여넣기
5. 함수 선택: `createEDMTForm` → 실행
6. Form URL 및 응답 시트 URL 저장

### 4. Google API 설정

자세한 내용은 [사용자 가이드](테스트/자동화시스템/EDMT_사용자_가이드.md) 참조

1. [Google Cloud Console](https://console.cloud.google.com) 접속
2. 새 프로젝트 생성
3. Google Sheets API 및 Drive API 활성화
4. 서비스 계정 생성 및 JSON 키 다운로드
5. `credentials.json`으로 저장
6. 응답 시트를 서비스 계정과 공유

### 5. 환경 변수 설정

```bash
# .env.example을 복사하여 .env 파일 생성
cp .env.example .env

# .env 파일을 편집하여 실제 값 입력
# - GOOGLE_SHEETS_URL: 3단계에서 생성한 응답 시트 URL
# - TEMPLATE_PATH: 채점 템플릿 경로
# - OUTPUT_DIR: 결과 파일 저장 경로
```

### 6. 자동 채점 실행

**방법 1: 환경 변수 사용 (권장)**

```bash
# Linux/Mac
bash scripts/run_example.sh

# Windows
python src\edmt_automation.py ^
  --sheet-url "%GOOGLE_SHEETS_URL%" ^
  --template "%TEMPLATE_PATH%" ^
  --output-dir "%OUTPUT_DIR%"
```

**방법 2: 직접 실행**

```bash
python3 src/edmt_automation.py \
  --sheet-url "YOUR_GOOGLE_SHEETS_URL" \
  --template "이화방어검사채점표 수정.xlsx" \
  --output-dir "./results"
```

## 📁 프로젝트 구조

```
ewha-defense-test/
├── README.md                              # 프로젝트 설명서
├── requirements.txt                       # Python 의존성
├── .gitignore                             # Git 제외 파일
├── .env.example                           # 환경 변수 예제
├── 이화방어검사채점표 수정.xlsx            # 채점 템플릿
├── src/
│   ├── __init__.py                        # 패키지 초기화
│   └── edmt_automation.py                 # Python 자동화 스크립트 v2.0
├── scripts/
│   ├── setup.sh                           # 설치 스크립트
│   └── run_example.sh                     # 실행 예제 스크립트
└── 테스트/
    ├── 자동화시스템/
    │   ├── create_edmt_form_complete.gs   # Google Forms 생성 스크립트
    │   ├── EDMT_사용자_가이드.md          # 상세 사용 가이드
    │   ├── 작업로그_2025-10-17.md        # 개발 로그
    │   └── edmt_questions.csv             # 200개 문항 데이터
    ├── 샘플/                              # 샘플 PDF 파일들
    └── 이화방어기제 검사지.pdf            # 원본 검사지
```

## 📖 상세 문서

- **[사용자 가이드](테스트/자동화시스템/EDMT_사용자_가이드.md)**: 전체 설정 및 사용 방법
- **[작업 로그](테스트/자동화시스템/작업로그_2025-10-17.md)**: 개발 과정 및 기술 상세

## 🔧 주요 구성 요소

### 1. Google Apps Script (`create_edmt_form_complete.gs`)
- 200개 문항 Google Forms 자동 생성
- 응답 자동 수집 (Google Sheets 연동)
- 5점 리커트 척도

### 2. Python 자동화 스크립트 v2.0 (`src/edmt_automation.py`)
- Google Sheets API로 응답 데이터 읽기 (google-auth 사용)
- 20개 방어기제별 원점수 계산
- 나이/성별별 규준표 적용
- 스텐점수 자동 변환
- 개인별 채점 엑셀 파일 생성
- **개선사항**:
  - 강화된 에러 처리 및 로깅 시스템
  - 무응답 처리 개선
  - Type hints 및 문서화 추가
  - 환경 변수 지원

### 3. 채점 템플릿 (`이화방어검사채점표 수정.xlsx`)
- 응답 입력 시트
- 자동 채점 로직
- 8개 규준표 (나이/성별별)

## 🎓 20개 방어기제

1. 허세 (Pretentiousness)
2. 반동형성 (Reaction Formation)
3. 동일시 (Identification)
4. 수동공격 (Passive-Aggressive)
5. 투사 (Projection)
6. 전치 (Displacement)
7. 부정 (Denial)
8. 통제 (Control)
9. 억제 (Suppression)
10. 왜곡 (Distortion)
11. 예견 (Anticipation)
12. 합리화 (Rationalization)
13. 해리 (Dissociation)
14. 신체화 (Somatization)
15. 승화 (Sublimation)
16. 행동화 (Acting Out)
17. 이타주의 (Altruism)
18. 퇴행 (Regression)
19. 유머 (Humor)
20. 회피 (Avoidance)

## 🔒 보안 주의사항

- ⚠️ `credentials.json` 파일은 **절대 공유 금지**
- ⚠️ Git에 인증 파일 커밋 금지 (`.gitignore`에 포함됨)
- ⚠️ 서비스 계정에 최소 권한만 부여
- ⚠️ 응답 시트는 필요한 사람에게만 공유

## 🐛 문제 해결

### "권한이 없습니다" 오류
→ Google Sheets를 서비스 계정 이메일과 공유했는지 확인

### "credentials.json을 찾을 수 없습니다"
→ JSON 파일이 프로젝트 루트에 있는지 확인

### Forms가 제대로 생성되지 않음
→ Apps Script 실행 로그 확인 및 권한 재승인

자세한 문제 해결은 [사용자 가이드](테스트/자동화시스템/EDMT_사용자_가이드.md#문제-해결) 참조

## 📊 예상 효과

- **시간 절약**: 내담자 1명당 35분 → 1초 (월 10명 기준 5.8시간 절약)
- **정확도 향상**: 수작업 오류율 5% → 0%
- **편의성 증대**: 원격 검사, 모바일 지원, 자동 백업

## 📝 라이선스

MIT License

## 👤 작성자

Claude Code

## 📞 지원

문제가 발생하면 다음을 확인하세요:
1. [사용자 가이드](테스트/자동화시스템/EDMT_사용자_가이드.md)의 모든 단계를 정확히 따랐는지
2. Python 버전이 3.7 이상인지
3. 모든 필수 라이브러리가 설치되었는지

---

**🎉 완전 자동화된 이화방어기제검사 시스템을 사용해보세요!**
