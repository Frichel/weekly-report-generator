# 📊 주간 업무보고서 자동 생성 도구

Excel 템플릿을 기반으로 주간 업무보고서를 자동 생성하는 Python 도구입니다.  
연도별 폴더 자동 생성, 이전 주 데이터 복사, Excel 행 높이 자동 조정 등 업무보고서 작성 시간을 획기적으로 단축시켜 줍니다.

## ✨ 주요 기능

- 📅 **연도별 폴더 자동 생성**: 입력한 날짜의 연도를 자동 감지하여 폴더 생성
- 🔄 **이전 주 데이터 자동 복사**: 지난 주 "차주일정"을 이번 주로 자동 반영
- 📏 **행 높이 자동 맞춤**: pywin32를 활용한 Excel 행 높이 자동 조정
- 📋 **데이터 유효성 검사**: 드롭다운 목록 자동 추가
- 🎯 **동적 경로 인식**: 폴더 위치가 바뀌어도 코드 수정 불필요
- ⏰ **자동화 실행 지원**: 작업 스케줄러 연동 가능

## 🚀 빠른 시작

### 요구사항

- **Python 3.7 이상**
- **Microsoft Excel** (행 높이 자동 맞춤 기능)
- **Windows OS** (pywin32 라이브러리 사용)

### 설치 방법

1. **저장소 복제**
```bash
git clone https://github.com/사용자명/weekly-report-generator.git
cd weekly-report-generator
```

2. **가상환경 생성 및 활성화**
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
```

3. **필수 패키지 설치**
```bash
pip install -r requirements.txt
```

4. **템플릿 파일 준비**
- `res/` 폴더에 Excel 템플릿 파일 배치
- 파일명 형식: `WW00_부서명_업무보고서_이름.xlsx`

## 📖 사용 방법

### 수동 실행 (대화형)

사용자가 직접 입력하며 실행하는 방식입니다.

```bash
python autorepocreate.py
```

**실행 예시:**
```
작업경로: C:\Users\...\WorkReport
작성 날짜를 입력(YYYYMMDD): 20270115
2027년 폴더를 생성했습니다.
이미 해당 주의 파일이 존재합니다. 덮어쓰시겠습니까?(Y/N): y
이전 주의 파일이 존재합니다. 기록된 내용을 가져오시겠습니까?(Y/N): y
Excel을 열어 행 높이를 자동 맞춤 중...
행 높이 자동 맞춤 완료!
파일 생성 완료: 2027\WW02_부서명_업무보고서_이름.xlsx
Job Done!
```

### 자동 실행 (무인)

스케줄러 등록이나 자동화에 적합한 방식입니다.

```bash
# 특정 날짜로 실행
python autorepocreate_pad.py 20270115

# 날짜 생략 시 이번 주 금요일 자동 계산
python autorepocreate_pad.py
```

**특징:**
- 파일이 이미 존재하면 자동으로 건너뜀
- 이전 주 데이터 자동 복사 (오류 시 스킵)
- 무인 실행에 최적화

## 🗂️ 프로젝트 구조

```
WorkReport/
├── autorepocreate.py          # 수동 실행용 스크립트 (대화형)
├── autorepocreate_pad.py      # 자동화 실행용 스크립트 (무인)
├── res/                       # 리소스 폴더
│   └── WW00_부서명_업무보고서_이름.xlsx  # Excel 템플릿
├── 2026/                      # 연도별 자동 생성 폴더
│   ├── WW01_...xlsx
│   └── WW02_...xlsx
└── 2027/                      # 새 연도 입력 시 자동 생성
```

## ⚙️ 작업 스케줄러 등록 (선택사항)

매주 자동으로 보고서를 생성하려면 Windows 작업 스케줄러에 등록하세요.

1. **작업 스케줄러** 실행
2. **새 작업 만들기**
   - 이름: "주간 업무보고서 자동 생성"
   - 트리거: 매주 금요일 오전 9시
3. **동작 설정**
   - 프로그램: `C:\...\WorkReport\.venv\Scripts\python.exe`
   - 인수: `C:\...\WorkReport\autorepocreate_pad.py`
   - 시작 위치: `C:\...\WorkReport`

## 🔧 문제 해결

### pywin32 설치 오류

```bash
pip install pywin32
python C:\Python3xx\Scripts\pywin32_postinstall.py -install
```

### 행 높이 자동 맞춤이 작동하지 않을 때

- Excel이 설치되어 있는지 확인
- 가상환경 내에서 실행 중인지 확인
- pywin32가 올바르게 설치되었는지 확인

### 템플릿 파일을 찾을 수 없다는 오류

- `res/` 폴더가 스크립트와 같은 위치에 있는지 확인
- Excel 템플릿 파일(.xlsx)이 `res/` 폴더에 있는지 확인
