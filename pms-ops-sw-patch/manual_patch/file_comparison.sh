#!/bin/bash

# 에러 처리 함수
exit_with_error() {
    echo "$1"
    exit 1
}

# 시스템 날짜를 사용하여 디렉토리 경로 설정
year=$(date +%Y)
month=$(date +%-m)

# 이전 달로 설정
if [ $month -eq 1 ]; then
    month=12
else
    month=$((month - 1))
fi

# 베이스 디렉토리 설정
BASE_DIRECTORY="BASE_DIRECTORY"

# patches 디렉토리를 찾기 위한 함수
find_patches_directory() {
    local base_dir=$1
    local patches_dir=""

    # depth 1 디렉토리를 검색
    for dir in "$base_dir"/*/; do
        if [ -d "$dir/patches" ]; then
            patches_dir="$dir/patches"
            break
        fi
    done

    echo "$patches_dir"
}

# patches 디렉토리를 찾기
LOCAL_DIRECTORY=$(find_patches_directory "$BASE_DIRECTORY")

if [ -z "$LOCAL_DIRECTORY" ]; then
    exit_with_error "patches 디렉토리를 찾을 수 없습니다."
fi

# SSH 서버 정보
SSH_SERVER="SSH_SERVER"
PORT="PORT"
USERNAME="USERNAME"
REMOTE_DIRECTORY="DIRECTORY"

# 원격 서버의 패치셋에 있는 파일 이름 추출
ssh -p $PORT $USERNAME@SSH_SERVER "find $REMOTE_DIRECTORY -type f -o -type d" >remote_files.txt || exit_with_error "원격 서버에서 파일/폴더 목록을 가져오는 데 실패"
echo "원격 서버의 파일/폴더 목록 추출 완료"

# 로컬 디렉토리의 파일/폴더 추출
find "$LOCAL_DIRECTORY" -type f -o -type d >local_files.txt || exit_with_error "로컬 디렉토리에서 파일/폴더 목록을 가져오는 데 실패"
echo "로컬 디렉토리의 파일/폴더 목록 추출 완료"

# 필요한 부분만 출력하여 새로운 파일에 저장
awk -v prefix="$REMOTE_DIRECTORY/" '{sub(prefix, ""); print}' remote_files.txt >remote_files_substr.txt || exit_with_error "원격 파일 목록에서 필요한 부분 추출에 실패"

awk -v prefix="$(echo $LOCAL_DIRECTORY | sed 's|]]|/|g')/" '{sub(prefix, ""); print}' local_files.txt >local_files_substr.txt || exit_with_error "로컬 파일 목록에서 필요한 부분 추출에 실패"

# 파일 경로를 딕셔너리로 저장
declare -A remote_files_dict
while IFS= read -r line; do
    remote_files_dict["$line"]=1
done <remote_files_substr.txt

# 비교 결과 출력 및 .ayt 파일 존재 여부 확인
echo "로컬 디렉토리에만 있는 파일 :"
while IFS= read -r file; do
    if [[ -z "${remote_files_dict[$file]}" ]]; then
        echo "$file"
    fi
done <local_files_substr.txt

no_ayt_files=()
echo "동일한 이름의 .ayt 파일이 존재하지 않는 파일:"
while IFS= read -r file; do
    if [[ "$file" == ms_files/* ]]; then
        if [[ "${file##*.}" == "cab" || "${file##*.}" == "exe" ]]; then
            ayt_file="${file}.ayt"
            if [[ ! -f "$LOCAL_DIRECTORY/$ayt_file" ]]; then
                echo "$file"
                no_ayt_files+=("$file")
            fi
        fi
    fi
done <local_files_substr.txt

# .ayt 파일이 존재하지 않는 파일이 없다면 메시지 출력
if [ ${#no_ayt_files[@]} -eq 0 ]; then
    echo "모든 파일에 대해 동일한 이름의 .ayt 파일이 존재합니다."
fi

# 임시 파일 삭제
rm remote_files_txt local_files.txt

echo "파일 비교 완료"

# 스크립트 성공 여부 확인 및 종료 코드
if [ ${#no_ayt_files[@]} -eq 0 ]; then
    echo "스크립트 실행 성공"
    exit_code=0
else
    echo "스크립트 실행 실패"
    exit_code=1
fi

bash

# 스크립트 종료 코드 반환
exit $exit_code
