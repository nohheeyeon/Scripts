#! /bin/sh

# 디렉토리 경로 설정
target_directory="패치셋 경로"

# 오늘 날짜를 YYYYMMDD 형식으로 얻기
current_date=$(date +'%Y%m%d')

# unzip된 파일들을 모아놓을 폴더
new_folder="$target_directory/$current_date"

# 디렉토리로 이동
cd "$target_directory" || exit

# zip 파일이 있는 지 확인하는 If 문
if [ -n "$(find "$target_directory" -maxdepth 1 -type f -name '*.zip')" ]; then
    echo "zip 파일이 존재합니다."

    # 압축을 푸는 작업
    for zip_file in "$target_directory"/*.zip; do
        unip -o "$zip_file" -d "$new_folder"
        echo "$zip_file의 압축을 풀었습니다."

        # 기존의 zip 파일 삭제
        rm -r "$zip_file"
    done

    mv "$target_directory"/*.zip "$new_folder"
    echo "zip 파일을 '$new_folder' 폴더로 이동했습니다."
else
    echo "zip 파일이 존재하지 않습니다."
fi

# patches.zip이 있는 지 확인하는 if 문
if [ -e "$new_folder/patches.zip" ]; then
    unzip -o "$new_folder/patches.zip" -d "$new_folder"
    echo "patches.zipdmf '$new_folder' 폴더로 압축 해제했습니다."
else
    echo "patches.zip이 없습니다."
fi
