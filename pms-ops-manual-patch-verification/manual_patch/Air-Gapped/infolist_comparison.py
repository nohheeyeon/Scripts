import datetime
import os
import sys
import zipfile
from io import BytesIO

import paramiko

sys.stdout.reconfigure(encoding="utf-8")

now = datetime.datetime.now()
year = now.year
month = now.month

if month == 1:
    month = 12
    year -= 1
else:
    month -= 1

BASE_DIRECTORY = f"C:/ftp_root/manual/ms/{year}/{month}"
local_output_txt_path = os.path.join(BASE_DIRECTORY, "local_file_list.txt")
remote_output_txt_path = os.path.join(BASE_DIRECTORY, "remote_file_list.txt")


def log(message):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}")


def exit_with_error(message):
    print(message)
    exit(1)


def find_patches_directory(base_dir):
    for root, dirs, _ in os.walk(base_dir):
        if "patches" in dirs:
            return os.path.join(root, "patches")
    return None


def process_zip(zip_file, parent_path=""):
    for file_info in zip_file.infolist():
        current_path = (
            os.path.normpath(f"{parent_path}/{file_info.filename}")
            if parent_path
            else file_info.filename
        )
        all_file_names.append(current_path)
        log(f"파일 발견: {current_path}")

        if file_info.filename.endswith(".zip"):
            log(f"내부 Zip 파일 처리 중: {current_path}")
            with zip_file.open(file_info) as inner_file:
                with zipfile.ZipFile(BytesIO(inner_file.read())) as nested_zip:
                    process_zip(nested_zip, parent_path=current_path)


LOCAL_DIRECTORY = find_patches_directory(BASE_DIRECTORY)
if LOCAL_DIRECTORY is None:
    exit_with_error("patches 디렉토리를 찾을 수 없습니다.")

all_file_names = []
log("로컬 파일 및 Zip 파일 내부 목록 작성 중")
for root, _, files in os.walk(LOCAL_DIRECTORY):
    for file in files:
        if file.endswith(".zip"):
            zip_path = os.path.join(root, file)
            log(f"Zip 파일 처리 중: {zip_path}")
            with zipfile.ZipFile(zip_path, "r") as zip_file:
                process_zip(zip_file)
        else:
            file_path = os.path.join(root, file).replace("\\", "/")
            all_file_names.append(file_path.replace(f"{LOCAL_DIRECTORY}/", ""))
            log(f"파일 발견: {file_path}")

log(f"로컬 파일 목록 저장 중: {local_output_txt_path}")
with open(local_output_txt_path, "w", encoding="utf-8") as output_file:
    for file_name in all_file_names:
        output_file.write(file_name + "\n")
log("로컬 파일 목록 작성 완료")
