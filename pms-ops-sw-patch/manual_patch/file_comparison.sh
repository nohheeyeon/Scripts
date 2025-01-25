#!/bin/bash

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
    echo "patches 디렉토리를 찾을 수 없습니다."
    exit 1
fi

# SSH 서버 정보
SSH_SERVER="SSH_SERVER"
PORT="PORT"
USERNAME="USERNAME"
REMOTE_DIRECTORY="DIRECTORY"

# 원격 서버의 패치셋에 있는 파일 이름 추출
ssh -p $PORT $USERNAME@SSH_SERVER "find $REMOTE_DIRECTORY -type f -o -type d" >remote_files.txt
if [ $? -ne 0 ]; then
    echo "원격 서버에서 파일/폴더 목록을 가져오는 데 실패"
    exit 1
fi
echo "원격 서버의 파일/폴더 목록 추출 완료"

# 로컬 디렉토리의 파일/폴더 추출
find "$LOCAL_DIRECTORY" -type f -o -type d >local_files.txt
if [ $? -ne 0 ]; then
    echo "로컬 디렉토리에서 파일/폴더 목록을 가져오는 데 실패"
    exit 1
fi
echo "로컬 디렉토리의 파일/폴더 목록 추출 완료"

# 추출된 파일/폴더 목록 출력
echo "원격 서버의 파일/폴더 목록 :"
cat remote_files.txt

echo "로컬 디렉토리의 파일/폴더 목록 :"
cat local_files.txt

bash
