# manual_patch

폐쇄망 환경에서 수동 패치셋의 무결성을 검증하는 Python 스크립트입니다.

## 기능
- 로컬 및 원격 패치셋 파일 비교 및 누락 파일 탐지
- 패치셋 내 압축 파일(.zip) 자동 탐색 및 내부 파일 검증
- 원격 서버에서 특정 기간 동안 수정된 파일 목록 조회
- `.ayt` 파일 검증을 통한 패치셋의 완전성 확인
- Office 패치셋 포함 여부 검사
- 로그 파일을 통해 검증 과정 기록

## 프로젝트 구조
```
Desktop/
│── check_office_patch_inclusion.json
│── verify_manual_patchset_integrity.py
│── verify_manual_patchset_integrity.bat
│── {YYYY-MM-DD}/
│   ├── log.txt
│   ├── local_file_list.txt
│   ├── remote_file_list.txt
│   ├── remote_modified_file_list.txt
```

## 설치 및 실행 방법

### 1. 환경 설정

#### 1.1 프로젝트 클론
```sh
git clone <repository-url>
```

### 2. 데이터 준비
1. `C:/ftp_root/manual/ms/{YYYY}/{MM}` 및 `C:/ftp_root/manual/sw/{YYYY}/{MM}` 폴더에 패치셋이 있는지 확인
2. 수동 업데이트 툴로 패치셋 업로드 중이면 사용 불가

### 3. 실행 방법

#### 3.1 배치 파일 실행 (Windows 환경)
```sh
verify_manual_patchset_integrity.bat
```

### 4. 패치 검증 및 확인
- **로컬 및 원격 파일 비교**:
  - `Desktop/{YYYY-MM-DD}/local_file_list.txt` 및 `remote_file_list.txt`에서 검증
- **수정된 패치 파일 확인**:
  - `Desktop/{YYYY-MM-DD}/remote_modified_file_list.txt`에서 최근 변경된 파일 확인
- **Office 패치셋 확인**:
  - `Desktop/check_office_patch_inclusion.json`을 기준으로 검증
- **.ayt 파일 검증**:
  - `.cab`, `.exe` 파일에 대한 `.ayt` 파일 존재 여부 확인
- **로그 파일 확인**:
  - `Desktop/{YYYY-MM-DD}/log.txt`에서 검증 과정 확인 가능

## 로깅
- 실행 로그는 `Desktop/{YYYY-MM-DD}/log.txt` 파일에 저장됩니다.
- 오류 발생 시 `Desktop/{YYYY-MM-DD}/log.txt` 파일을 확인하세요.

