#! /bin/sh

# 디렉토리 경로 설정
target_directory="패치셋 경로"

# 오늘 날짜를 YYYYMMDD 형식으로 얻기
current_date=$(date +'%Y%m%d')

# unzip된 파일들을 모아놓을 폴더
new_folder="$target_directory/$current_date"

# 디렉토리로 이동
cd "$target_directory" || exit
