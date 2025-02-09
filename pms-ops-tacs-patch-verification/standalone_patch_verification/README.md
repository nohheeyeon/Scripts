# standalone_patch_verification

Windows VM에서 SW 및 MS 독립 실행 패치를 자동으로 설치하는 Python 스크립트입니다.

## 기능
- SW 및 MS 패치 리스트를 기반으로 자동 설치
- 설치된 패치 확인은 수동으로 수행
- 가상 환경을 이용한 실행

## 구조
```
standalone_patch_verification/
│── data/
│   ├── sw_patch_list.xlsx
│   ├── ms_patch_list.xlsx
│── logs/
│── package/
│── standalone_patch_auto_install.py
│── standalone_patch_auto_install_X86.bat
│── standalone_patch_auto_install_AMD64.bat
│── .gitignore
│── requirements.txt
```

## 설치 및 실행 방법

### 1. 환경 설정

#### 1.1 프로젝트 클론 및 가상환경 설정
```sh
# 프로젝트 클론
git clone <repository-url>
cd standalone_patch_verification

# 가상환경 생성 및 활성화
python -m venv .venv
call .venv\Scripts\activate.bat
```

#### 1.2 requirements.txt 설치
```sh
pip install -r requirements.txt
```

### 2. 데이터 준비
1. `standalone_patch_verification/data/` 폴더에 `sw_patch_list.xlsx` 및 `ms_patch_list.xlsx` 파일 배치
2. `standalone_patch_verification/package/` 폴더에 독립실행패치 디렉토리(`standalone`) 배치

### 3. 실행 방법

#### 3.1 X86 아키텍처에서 실행
**관리자 모드로 실행**
```sh
standalone_patch_auto_install_X86.bat
```

#### 3.2 AMD64 아키텍처에서 실행
**관리자 모드로 실행**
```sh
standalone_patch_auto_install_AMD64.bat
```

### 4. 패치 설치 및 확인
- **SW 패치 확인**:
  - 제어판 > 프로그램 및 기능 > 프로그램 제거 또는 변경에서 확인
  - 패치는 자동으로 설치되며, 확인은 수동으로 수행해야 합니다.
- **MS 패치 확인**:
  - 제어판 > 프로그램 및 기능 > 설치된 업데이트에서 확인
  - `C:\Users\사용자이름\AppData\Local\Temp\package` 하위 경로에 KBID 값의 폴더가 생성되었는지 확인
  - 패치는 자동으로 설치되며, 확인은 수동으로 수행해야 합니다.

## .gitignore
```
/data
/logs
.venv/
```

## requirements.txt
```
openpyxl
pandas
python
```

## 로깅
- 실행 로그는 `logs/` 폴더 내에 저장됩니다.
- 오류 발생 시 `logs/*.log` 파일을 확인하세요.