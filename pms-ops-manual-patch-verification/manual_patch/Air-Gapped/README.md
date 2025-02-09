# manual_patch

폐쇄망 환경에서 수동 패치셋의 무결성을 검증하는 Python 스크립트입니다.

## 기능
- 로컬 및 원격 패치셋 파일 비교 및 누락 파일 탐지
- 패치셋 내 압축 파일(.zip) 자동 탐색 및 내부 파일 검증
- 원격 서버에서 이 달에 업로드된 패치 파일 목록 조회
- `.ayt` 파일 검증을 통한 패치셋의 완전성 확인
- Office 패치셋 포함 여부 검사
- 로그 파일을 통해 검증 과정 기록

## 프로젝트 구조
```
C:\pms-ops-manual-patch-verification\manual_patch
│── check_office_patch_inclusion.json
│── verify_manual_patchset_integrity.py
│── verify_manual_patchset_integrity.bat
│── {YYYY-MM-DD}/
│   ├── log.txt
│   ├── local_file_list.txt
│   ├── remote_file_list.txt
│   ├── remote_modified_file_list.txt
```

## .txt 파일 설명
- **log.txt**: 스크립트 실행 과정 및 오류 메시지 기록
- **local_file_list.txt**: 로컬에서 수집한 패치셋 내 모든 파일 목록
- **remote_file_list.txt**: 원격 서버 패치셋 경로 내의 모든 파일 목록
- **remote_modified_file_list.txt**: 이 달에 원격 서버에 업로드된 패치 파일 목록을 필터링하여 저장한 목록

## 설치 및 실행 방법

### 1. 환경 설정

#### 1.1 프로젝트 클론
```sh
git clone <repository-url>
```

### 2. 데이터 준비
1. `C:/ftp_root/manual/{YYYY}/{MM}/ms` 및 `C:/ftp_root/manual/{YYYY}/{MM}/sw` 폴더에 패치셋이 있는지 확인
2. 수동 업데이트 툴로 패치셋 업로드 중이면 사용 불가
  - 해당 스크립트는 123번 서버에 업로드된 패치셋과 수동 업데이트 툴을 통해 원격 서버에 업로드된 패치셋을 비교하는 스크립트이므로, 원격 서버에 패치셋 업로드가 완료되기 전에 스크립트를 실행하면, 원격 서버에서의 패치 파일 리스트가 완전히 추출되지 않아 정상적인 검증이 불가능합니다.

### 3. 실행 방법

#### 3.1 배치 파일 실행
```sh
verify_manual_patchset_integrity.bat 클릭
```

### 4. 패치 검증 및 확인
- **로컬 및 원격 파일 비교**:
  - `./{YYYY-MM-DD}/local_file_list.txt` 및 `remote_file_list.txt`을 비교
- **업로드된 패치 파일 확인**:
  - `./{YYYY-MM-DD}/remote_modified_file_list.txt`에서 이 달에 업로드된 파일 목록 확인
- **Office 패치셋 확인**:
  - `./check_office_patch_inclusion.json`을 기준으로 검증
- **.ayt 파일 검증**:
  - `.cab`, `.exe` 파일에 대한 `.ayt` 파일 존재 여부 확인
- **로그 파일 확인**:
  - `./{YYYY-MM-DD}/log.txt`에서 검증 과정 확인 가능

## 로깅
- 실행 로그는 `./{YYYY-MM-DD}/log.txt` 파일에 저장
- 오류 발생 시 `./{YYYY-MM-DD}/log.txt` 파일을 확인
- 각 서버에 포함된 패치에 대해서 분석하기 위해서는 해당되는 .txt 파일 확인
