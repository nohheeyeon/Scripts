# Automate Tacs Validation Certificate

## 개요
TACS 패치셋 검증 시 작성하는 검증서를 자동화시킨 Python 스크립트입니다.

## 프로젝트 구조
```
certificate_automation/
├── .gitignore
├── README.md
├── requirements.txt
├── data/
│   ├── 증분 패치 검증 QA 결과서.docx
│   ├── ms_patch_list.xlsx
│   ├── sw_patch_list.xlsx
│   └── <패치셋 폴더>
└── src/
    ├── automate_tacs_validation_certificate.bat
    ├── automate_tacs_validation_certificate.py
    ├── document_updater.py
    ├── excel_data_processor.py
    └── patch_directory_handler.py
```
## 세팅 방법
1. **레포지토리 클론**:
```bash
git clone <repository_url>
cd certificate_automation
```

2. **가상환경 생성 및 활성화**:
```bash
python -m venv .venv
.venv\Scripts\activate
```

3. requirements.txt 설치
```bash
pip install -r requirements.txt
```

## 실행 방법
1. `/data` 디렉토리에 필요한 Excel 파일(`ms_patch_list.xlsx`, `sw_patch_list.xlsx`)과 패치셋 폴더를 배치합니다.
3. 스크립트 실행
```bash
automate_tacs_validation_certificate.bat 클릭
```

## 스크립트 상세 설명
### `automate_tacs_validation_certificate.py`
- 각 모듈을 초기화하고 문서 처리 워크플로우를 실행합니다.

### `document_updater.py`
- Word 문서 수정, 패치 세부사항 업데이트, 검증 섹션 삽입 및 서식 설정을 처리합니다.

### `excel_data_processor.py`
- Excel 파일에서 패치 제목, KB ID 및 SW 버전을 추출하고 처리합니다.

### `patch_directory_handler.py`
- 디렉토리 구조를 관리하고 검증에 필요한 패치 파일을 검색합니다.

## 출력 결과
- 스크립트 실행 후 `/data` 디렉토리에 아래 형식의 .docx가 생성됩니다.
```
MM월 증분 패치 검증 QA 결과서_YYMMDD.docx
```

EX:
```
02월 증분 패치 검증 QA 결과서_250207.docx
```

## .gitignore 설정
```
.venv/
/data/
```

## requirements.txt
```
docx
et_xmlfile
lxml
numpy
openpyxl
pandas
pillow
python-dateutil
python-docx
pytz
pywin32
six
typing_extensions
tzdata
```