#! /bin/sh

# 에러 발생 시 메시지를 출력하고 스크립트를 종료하는 함수
exit_with_error() {
    echo "$1" # 에러 메시지를 출력
    exit 1    # 에러 코드와 함께 스크립트 종료
}

# 이전 달과 연도 계산
year=$(date +%Y)
month=$(date +%m)

if [ "$month" -eq 1 ]; then
    month=12
    year=$(($year - 1))
else
    month=$((10#$month - 1))
fi

# 디렉토리 경로 설정
target_directory="C:ftp_root/manual/ms/$year/$month"

# 오늘 날짜를 YYYYMMDD 형식으로 얻기
current_date=$(date +'%Y%m%d')

# unzip된 파일들을 모아놓을 폴더
new_folder="$target_directory/$current_date"

# 디렉토리로 이동
cd "$target_directory" || exit_with_error "디렉토리로 이동 실패"

# 압축 해제 작업 전에 대상 폴더 생성
mkdir -p "$new_folder" || exit_with_error "폴더 생성 실패: $new_folder"

# zip 파일이 있는 지 확인하는 If 문
if [ -n "$(find "$target_directory" -maxdepth 1 -type f -name '*.zip')" ]; then
    echo "zip 파일이 존재합니다."

    # 압축을 푸는 작업
    for zip_file in *.zip; do
        unzip_dir="${zip_file%.zip}"
        mkdir -p "$unzip_dir" || exit_with_error "폴더 생성 실패: $unzip_dir"
        unzip -d "$unzip_dir" "$zip_file" || exit_with_error "압축 해제 실패: $zip_file"
        # patches.zip 파일 압축 해제 작업
        patches_zip_file="${unzip_dir}/patches.zip"
        if [ -f "$patches_zip_file" ]; then
            unzip -d "${zip_file%.zip}/patches" "$patches_zip_file" || exit_with_error "패치 압축 해제 실패"
            echo "$patches_zip_file의 패치 파일 압축을 풀었습니다."
        fi

        # unzip된 파일들을 new_folder로 복사
        cp -rf "$unzip_dir"/* "$new_folder/" || exit_with_error "파일 복사 실패"
        echo "$unzip_dir의 파일들을 $new_folder로 복사했습니다."

        # unzip된 폴더 삭제
        rm -rf "$unzip_dir"
        echo "$unzip_dir를 삭제했습니다."

    done
fi

# 스크립트 실행 완료
echo "스크립트 실행 완료"
bash
