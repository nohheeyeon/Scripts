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
    for zip_file in *.zip; do
        unip -d "${zip_file%.zip}" "$zip_file"
        echo "$zip_file의 압축을 풀었습니다."

        # patches.zip 파일 압축 해제 작업
        patches_zip_file="$zip_file%.zip}/patches.zip"
        if [ -f "$patches_zip_file" ]; then
            mkdir -p "${zip_file%.zip}/patches"
            unzip -d "${zip_file%.zip}/patches" "$patches_zip_file"
            echo "$patches_zip_file의 패치 파일 압축을 풀었습니다."
        fi
    done

    mkdir "$new_folder"
    for folder in "$target_directory"/*/; do
        if [ -d "$folder" ]; then
            find "$folder" -mindepth 1 -maxdepth 1 -exec cp -rf {} "$new_folder/" \;
            rm -r "$folder"
            echo "$folder 내의 모든 파일과 폴더를 $new_folder/로 복사하고 원본 폴더를 삭제했습니다."
        fi
    done

fi
