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
remote_modified_output_txt_path = os.path.join(
    BASE_DIRECTORY, "remote_modified_file_list.txt"
)
log_file_path = os.path.join(BASE_DIRECTORY, "log.txt")

with open(log_file_path, "w", encoding="utf-8") as log_file:
    log_file.write("로그 파일 시작\n")


def log(message):
    timestamped_message = (
        f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    )
    print(timestamped_message)
    with open(log_file_path, "a", encoding="utf-8") as log_file:
        log_file.write(timestamped_message + "\n")


def exit_with_error(message):
    log(message)
    exit(1)


def process_zip(zip_file, parent_path="", level=0):
    for file_info in zip_file.infolist():
        relative_path = file_info.filename.replace("\\", "/").rstrip("/")
        all_file_names.add(relative_path)
        log(f"Zip 내부 파일 발견: {relative_path}")

        if file_info.filename.endswith(".zip"):
            if level >= 1:
                log(f"{relative_path}")
                continue

            log(f"patches.zip 파일 처리 중: {relative_path}")
            with zip_file.open(file_info) as inner_file:
                with zipfile.ZipFile(BytesIO(inner_file.read())) as nested_zip:
                    process_zip(nested_zip, parent_path=relative_path, level=1)
        else:
            log(f"파일 추가: {relative_path}")


def process_top_level_zips(directory_path):
    for file_name in os.listdir(directory_path):
        if file_name.endswith(".zip"):
            zip_path = os.path.join(directory_path, file_name)
            log(f"최상위 Zip 파일 처리 중: {zip_path}")
            try:
                with zipfile.ZipFile(zip_path, "r") as zip_file:
                    process_zip(zip_file, parent_path="", level=0)
            except zipfile.BadZipFile:
                log(f"잘못된 Zip 파일: {zip_path} - zip 형식이 아님")


def fetch_remote_files(ssh_server, port, username, password, remote_dir):
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(ssh_server, port=port, username=username, password=password)

        stdin, stdout, stderr = ssh.exec_command(
            f"find {remote_dir} -type f -o -type d"
        )
        all_files = set(
            line.strip().replace(f"{remote_dir}/", "")
            for line in stdout.read().decode("utf-8").splitlines()
        )

        log("원격 서버의 v1 하위 경로부터 모든 파일 경로 목록 추출 완료")

        month_filter = now.strftime("%Y-%m")
        stdin, stdout, stderr = ssh.exec_command(
            f"find {remote_dir} -type f -newermt {month_filter}-01 ! -newermt {month_filter}-31"
        )
        modified_files = set(
            line.strip().replace(f"{remote_dir}/", "")
            for line in stdout.read().decode("utf-8").splitlines()
        )

        log("원격 서버에서 이번 달에 수정된 파일 목록 추출 완료")
        return all_files, modified_files
    except Exception as e:
        exit_with_error(f"원격 서버에서 파일/폴더 목록을 가져오는 데 실패: {str(e)}")
    finally:
        ssh.close()


SSH_SERVER = "SSH_SERVER"
PORT = 6879
USERNAME = "USERNAME"
PASSWORD = "PASSWORD"
REMOTE_DIRECTORY = "REMOTE_DIRECTORY"

all_file_names = set()
log("MS 디렉토리의 patches.zip 파일 하위 파일 목록 작성 중")
process_top_level_zips(BASE_DIRECTORY)

log("SW 디렉토리의 patches.zip 파일 하위 파일 목록 작성 중")
process_top_level_zips(SW_DIRECTORY)

with open(local_output_txt_path, "w", encoding="utf-8") as output_file:
    for file_name in sorted(all_file_names):
        output_file.write(file_name + "\n")
log("로컬 파일 목록 작성 완료")

remote_files, modified_files = fetch_remote_files(
    SSH_SERVER, PORT, USERNAME, PASSWORD, REMOTE_DIRECTORY
)
with open(remote_output_txt_path, "w", encoding="utf-8") as output_file:
    for file in sorted(remote_files):
        output_file.write(file + "\n")
log("원격 서버의 v1 하위 경로부터 전체 파일 목록 작성 완료")

with open(remote_modified_output_txt_path, "w", encoding="utf-8") as output_file:
    for file in sorted(modified_files):
        output_file.write(file + "\n")
log("이번 달에 수정된 원격 서버 파일 목록 작성 완료")


def normalize_path(file_path, is_windows=True):
    return file_path.replace("\\", "/")


def check_inclusion(reference_file, target_file, description):
    with open(reference_file, "r", encoding="utf-8") as ref_file:
        reference_list = set(normalize_path(line.strip()) for line in ref_file)

    with open(target_file, "r", encoding="utf-8") as tgt_file:
        target_list = set(normalize_path(line.strip()) for line in tgt_file)

    not_included = reference_list - target_list
    if not_included:
        log(f"{description} 포함되지 않은 파일 목록:")
        for file in sorted(not_included):
            log(file)
    else:
        log(f"{description} 모든 파일이 포함되어 있습니다.")


check_inclusion(
    remote_modified_output_txt_path, local_output_txt_path, "원격 수정된 파일이 로컬에"
)
check_inclusion(
    local_output_txt_path, remote_output_txt_path, "로컬 파일이 원격 전체 파일에"
)
