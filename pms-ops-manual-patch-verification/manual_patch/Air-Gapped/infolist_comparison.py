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
SW_DIRECTORY = f"C:/ftp_root/manual/sw/{year}/{month}"
local_output_txt_path = os.path.join(BASE_DIRECTORY, "local_file_list.txt")
remote_output_txt_path = os.path.join(BASE_DIRECTORY, "remote_file_list.txt")


def log(message):
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}")


def exit_with_error(message):
    print(message)
    exit(1)


def process_zip(zip_file, parent_path=""):
    for file_info in zip_file.infolist():
        current_path = (
            os.path.normpath(f"{parent_path}/{file_info.filename}")
            if parent_path
            else file_info.filename
        )

        if "patches.zip" in current_path:
            relative_path = current_path.split("patches.zip", 1)[1].lstrip("\\/")
            all_file_names.append(relative_path)
            log(f"파일 발견: {relative_path}")
        else:
            all_file_names.append(current_path)
            log(f"파일 발견: {current_path}")

        if file_info.filename.endswith(".zip") and "patches.zip" not in current_path:
            log(f"Zip 파일 이름만 기록: {current_path}")
            continue
        elif file_info.filename.endswith(".zip"):
            log(f"내부 patches.zip 파일 처리 중: {current_path}")
            with zip_file.open(file_info) as inner_file:
                with zipfile.ZipFile(BytesIO(inner_file.read())) as nested_zip:
                    process_zip(nested_zip, parent_path=current_path)


def process_directory(directory_path):
    for root, _, files in os.walk(directory_path):
        for file in files:
            if file.endswith(".zip"):
                zip_path = os.path.join(root, file)
                log(f"Zip 파일 처리 중: {zip_path}")
                with zipfile.ZipFile(zip_path, "r") as zip_file:
                    process_zip(zip_file)


all_file_names = []
log("MS 디렉토리 파일 및 patches.zip 내부 목록 작성 중")
process_directory(BASE_DIRECTORY)

log("SW 디렉토리 파일 및 patches.zip 내부 목록 작성 중")
process_directory(SW_DIRECTORY)

log(f"로컬 파일 목록 저장 중: {local_output_txt_path}")
with open(local_output_txt_path, "w", encoding="utf-8") as output_file:
    for file_name in all_file_names:
        output_file.write(file_name + "\n")
log("로컬 파일 목록 작성 완료")


def fetch_remote_files(ssh_server, port, username, password, remote_dir):
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(ssh_server, port=port, username=username, password=password)

        stdin, stdout, stderr = ssh.exec_command(
            f"find {remote_dir} -type f -o -type d"
        )
        remote_files = stdout.read().decode("utf-8").splitlines()

        log("원격 서버의 파일/폴더 목록 추출 완료")
        return [file.replace(f"{remote_dir}/", "") for file in remote_files]
    except Exception as e:
        exit_with_error(f"원격 서버에서 파일/폴더 목록을 가져오는 데 실패: {str(e)}")
    finally:
        ssh.close()


SSH_SERVER = "10.119.13.30"
PORT = 6879
USERNAME = "patch_test"
PASSWORD = "PASSWORD"
REMOTE_DIRECTORY = "/usr/local/ASM4/fileset/pms_patch/v1"

remote_files = fetch_remote_files(
    SSH_SERVER, PORT, USERNAME, PASSWORD, REMOTE_DIRECTORY
)
log(f"원격 서버 파일 목록 저장 중: {remote_output_txt_path}")
with open(remote_output_txt_path, "w", encoding="utf-8") as output_file:
    for file in remote_files:
        output_file.write(file + "\n")
log("원격 서버 파일 목록 작성 완료")

remote_files_set = set(remote_files)
local_only_files = [file for file in all_file_names if file not in remote_files_set]
log("로컬 디렉토리에만 있는 파일:")
for file in local_only_files:
    print(file)

local_files_set = set(all_file_names)
remote_only_files = [file for file in remote_files if file not in local_files_set]
log("원격 서버에만 있는 파일:")
for file in remote_only_files:
    print(file)

no_ayt_files = []
log("동일한 이름의 .ayt 파일이 존재하지 않는 파일:")
for file in all_file_names:
    if file.startswith("ms_files/") and file.split(".")[-1] in {"cab", "exe"}:
        ayt_file_path = os.path.join(BASE_DIRECTORY, f"{file}.ayt")
        if not os.path.isfile(ayt_file_path):
            print(file)
            no_ayt_files.append(file)

if not no_ayt_files:
    log("모든 파일에 대해 동일한 이름의 .ayt 파일이 존재합니다.")

if not no_ayt_files:
    log("스크립트 실행 성공")
    exit_code = 0
else:
    log("스크립트 실행 실패")
    exit_code = 1

exit(exit_code)
