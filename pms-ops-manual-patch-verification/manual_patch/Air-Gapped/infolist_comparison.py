import calendar
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

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
current_date_folder = os.path.join(
    desktop_path, datetime.datetime.now().strftime("%Y-%m-%d")
)
os.makedirs(current_date_folder, exist_ok=True)

BASE_DIRECTORY = f"C:/ftp_root/manual/ms/{year}/{month}"
SW_DIRECTORY = f"C:/ftp_root/manual/sw/{year}/{month}"

local_output_txt_path = os.path.join(current_date_folder, "local_file_list.txt")
remote_output_txt_path = os.path.join(current_date_folder, "remote_file_list.txt")
remote_modified_output_txt_path = os.path.join(
    current_date_folder, "remote_modified_file_list.txt"
)
log_file_path = os.path.join(current_date_folder, "log.txt")

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


def process_top_level_zips(directory_path):
    for file_name in os.listdir(directory_path):
        if file_name.endswith(".zip"):
            zip_path = os.path.join(directory_path, file_name)
            print(f"최상위 Zip 파일 처리 중: {zip_path}")
            try:
                with zipfile.ZipFile(zip_path, "r") as zip_file:
                    for file_info in zip_file.infolist():
                        relative_path = file_info.filename.replace("\\", "/").rstrip(
                            "/"
                        )
                        all_file_names.add(relative_path)
                        if file_info.filename.endswith(".zip"):
                            with zip_file.open(file_info) as inner_file:
                                with zipfile.ZipFile(
                                    BytesIO(inner_file.read())
                                ) as nested_zip:
                                    for nested_file_info in nested_zip.infolist():
                                        nested_path = nested_file_info.filename.replace(
                                            "\\", "/"
                                        ).rstrip("/")
                                        all_file_names.add(nested_path)
            except zipfile.BadZipFile:
                with open(log_file_path, "a", encoding="utf-8") as log_file:
                    log_file.write(
                        f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 잘못된 Zip 파일: {zip_path} - zip 형식이 아님\n"
                    )


def process_zip(zip_file, parent_path="", level=0):
    for file_info in zip_file.infolist():
        relative_path = file_info.filename.replace("\\", "/").rstrip("/")
        all_file_names.add(relative_path)
        if file_info.filename.endswith(".zip"):
            if level >= 1:
                log(f"{relative_path}")
                continue
            log(f"patches.zip")
            with zip_file.open(file_info) as inner_file:
                with zipfile.ZipFile(BytesIO(inner_file.read())) as nested_zip:
                    process_zip(nested_zip, parent_path=relative_path, level=1)


def fetch_remote_files(ssh_server, port, username, password, remote_dir):
    try:
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(ssh_server, port=port, username=username, password=password)

        stdin, stdout, stderr = ssh.exec_command(
            f"find {remote_dir} -type f -o -type d"
        )
        all_files_raw = stdout.read().decode("utf-8").strip()
        all_files = set(
            line.strip().replace(f"{remote_dir}/", "")
            for line in all_files_raw.splitlines()
        )

        month_filter = now.strftime("%Y-%m")
        last_day = calendar.monthrange(now.year, now.month)[1]
        command = f"find {remote_dir} -type f -newermt {month_filter}-01 ! -newermt {month_filter}-{last_day}T23:59:59"
        stdin, stdout, stderr = ssh.exec_command(command)

        result = stdout.read().decode("utf-8").strip()
        error = stderr.read().decode("utf-8").strip()

        if error:
            log(f"명령 에러: {error}")

        modified_files = set(
            line.strip().replace(f"{remote_dir}/", "")
            for line in result.splitlines()
            if not line.strip().endswith(".json")
        )

        with open(
            remote_modified_output_txt_path, "w", encoding="utf-8"
        ) as output_file:
            if not modified_files:
                log("수정된 파일이 없습니다. 빈 파일을 생성합니다.")
            for file in sorted(modified_files):
                output_file.write(file + "\n")
        log(f"수정된 파일 목록이 {remote_modified_output_txt_path}에 저장되었습니다.")

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
process_top_level_zips(BASE_DIRECTORY)

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


def check_office_patch_inclusion(file_list_path):
    patterns = {
        # Windows 11
        "11v24H2-x64": "Windows 11 버전 24H2 (X64)",
        "11v23H2-x64": "Windows 11 버전 23H2 (X64)",
        "11v22H2-x64": "Windows 11 버전 22H2 (x64)",
        "11v21H2-x64": "Windows 11 버전 21H2 (x64)",
        # Windows 10
        "10v22H2-x64": "Windows 10 버전 22H2 (x64)",
        "10v22H2-x86": "Windows 10 버전 22H2 (x86)",
        "10v21H2-x64": "Windows 10 버전 21H2 (x64)",
        "10v21H2-x86": "Windows 10 버전 21H2 (x86)",
        "10v21H1-x64": "Windows 10 버전 21H1 (x64)",
        "10v21H1-x86": "Windows 10 버전 21H1 (x86)",
        "10v20H2-x64": "Windows 10 버전 20H2 (x64)",
        "10v20H2-x86": "Windows 10 버전 20H2 (x86)",
        "10v2004-x64": "Windows 10 버전 2004 (x64)",
        "10v2004-x86": "Windows 10 버전 2004 (x86)",
        "10v1909ent-x64": "Windows 10 버전 1909 Enterprise (x64)",
        "10v1909ent-x86": "Windows 10 버전 1909 Enterprise (x86)",
        "10v1909-x64": "Windows 10 버전 1909 (x64)",
        "10v1909-x86": "Windows 10 버전 1909 (x86)",
        "10v1903v2-x64": "Windows 10 버전 1903 v2 (x64)",
        "10v1903v2-x86": "Windows 10 버전 1903 v2 (x86)",
        "10v1903-x64": "Windows 10 버전 1903 (x64)",
        "10v1903-x86": "Windows 10 버전 1903 (x86)",
        "10v1809ltsc-x64": "Windows 10 버전 1809 LTSC (x64)",
        "10v1809ltsc-x86": "Windows 10 버전 1809 LTSC (x86)",
        "10v1809ent-x64": "Windows 10 버전 1809 ENT (X64)",
        "10v1809ent-x86": "Windows 10 버전 1809 ENT (X86)",
        "10v1809-x64": "Windows 10 버전 1809 (x64)",
        "10v1809-x86": "Windows 10 버전 1809 (x86)",
        "10v1803gdrdu-x64": "Windows 10 버전 1809 GDRDU (X64)",
        "10v1803gdrdu-x86": "Windows 10 버전 1809 GDRDU (X86)",
        "10v1803-x64": "Windows 10 버전 1803 (x64)",
        "10v1803-x86": "Windows 10 버전 1803 (x86)",
        "10v1709-x64": "Windows 10 버전 1709 (x64)",
        "10v1709-x86": "Windows 10 버전 1709 (x86)",
        "10v1703-x64": "Windows 10 버전 1703 (x64)",
        "10v1703-x86": "Windows 10 버전 1703 (x86)",
        "10v1607-x64": "Windows 10 버전 1607 (x64)",
        "10v1607-x86": "Windows 10 버전 1607 (x86)",
        "10v1511-x64": "Windows 10 버전 1511 (x64)",
        "10v1511-x86": "Windows 10 버전 1511 (x86)",
        "10v1507-x64": "Windows 10 버전 1507 (x64)",
        "10v1507-x86": "Windows 10 버전 1507 (x86)",
        # Windows 8, 7
        "8.1-x64": "Windows 8.1 (x64)",
        "8.1-x86": "Windows 8.1 (x86)",
        "8-x64": "Windows 8 (x64)",
        "8-x86": "Windows 8 (x86)",
        "7sp1-x64": "Windows 7 SP1 (x64)",
        "7sp1-x86": "Windows 7 SP1 (x86)",
        "7rtm-x64": "Windows 7 RTM (x64)",
        "7rtm-x86": "Windows 7 RTM (x86)",
        # Windows Upgrade 10 -> 10 비즈니스에디션
        "10edunvl-x86": "Windows 10 EDUNVL (X86)",
        "10eduvl-x86": "Windows 10 EDUVL (X86)",
        "10provl-x86": "Windows 10 PROVL (X86)",
        "10pronvl-x86": "Windows 10 PRONVL (X86)",
        "10ent-x86": "Windows 10 ENT (X86)",
        "10entn-x86": "Windows 10 ENTN (X86)",
        "10edunvl-x64": "Windows 10 EDUNVL (x64)",
        "10eduvl-x64": "Windows 10 EDUVL (X64)",
        "10provl-x64": "Windows 10 PROVL (X64)",
        "10pronvl-x64": "Windows 10 PRONVL (X64)",
        "10ent-x64": "Windows 10 ENT (X64)",
        "10entn-x64": "Windows 10 ENTN (X64)",
        # Windows Upgrade 10 -> 소비자에디션
        "10cloud-x86": "Windows 10 CLOUD (X86)",
        "10cloudn-x86": "Windows 10 CLOUDN (X86)",
        "10core-x86": "Windows 10 CORE (X86)",
        "10coren-x86": "Windows 10 COREN (X86)",
        "10edunonvl-x86": "Windows 10 EDUNONVL (X86)",
        "10edunonvln-x86": "Windows 10 EDUNOWNVLN (X86)",
        "10prononvl-x86": "Windows 10 PRONONVL (X86)",
        "10prononvln-x86": "Windows 10 PRONONVLN (X86)",
        "10cloud-x64": "Windows 10 CLOUD (X64)",
        "10cloudn-x64": "Windows 10 CLOUDN (X64)",
        "10core-x64": "Windows 10 CORE (X64)",
        "10coren-x64": "Windows 10 COREN (X64)",
        "10edunonvl-x64": "Windows 10 EDUNONVL (X64)",
        "10edunonvln-x64": "Windows 10 EDUNONVLN (X64)",
        "10prononvl-x64": "Windows 10 PRONONVL (X64)",
        "10prononvln-x64": "Windows 10 PRONONVLN (X64)",
        # Windows Upgrade 7, 8.1 -> 10
        "7-x86": "Windows 7 (x86)",
        "7-x64": "Windows 7 (x64)",
        "8.1-x86": "Windows 8.1 (x86)",
        "8.1-x64": "Windows 8.1 (x64)",
        # Windows Server
        "2022-all": "Windows Server 2022",
        "2019-all": "Windows Server 2019",
        "2016-all": "Windows Server 2016",
        "2012r2-all": "Windows Server 2012 R2",
        "2012-all": "Windows Server 2012",
        "2008r2sp1-all": "Windows Server 2008 R2 SP1",
        "2008r2rtm-all": "Windows Server 2008 R2 RTM",
        "2008sp2-x64": "Windows Server 2008 SP2 (x64)",
        "2008sp2-x86": "Windows Server 2008 SP2 (x86)",
        "2008rtm-x64": "Windows Server 2008 RTM (X64)",
        "2008rtm-x86": "Windows Server 2008 RTM (X86)",
        "2003sp2-x64": "Windows Server 2003 SP2 (x64)",
        "2003sp2-x86": "Windows Server 2003 SP2 (x86)",
        "2003sp1-x64": "Windows Server 2003 SP1 (X64)",
        "2003sp1-x86": "Windows Server 2003 SP1 (X86)",
        "2003rtm-x64": "Windows Server 2003 RTM (x64)",
        "2003rtm-x86": "Windows Server 2003 RTM (x86)",
        # Windows Embedded
        "w8embed-x64": "Windows Embedded 8 (x64)",
        "w8embed-x86": "Windows Embedded 8 (x86)",
        "w7sp1embed-x64": "Windows Embedded 7 SP1 (x64)",
        "w7sp1embed-x86": "Windows Embedded 7 SP1 (x86)",
        "w7rtmembed-x64": "Windows Embedded 7 RTM (X64)",
        "w7rtmembed-x86": "Windows Embedded 7 RTM (X86)",
        "xp3embed-x64": "Windows Embedded XP SP3 (x64)",
        "xp3embed-x86": "Windows Embedded XP SP3 (x86)",
        # Office
        "2019-all": "Office 2019 패치셋에 포함",
        "2016-all": "Office 2016 패치셋에 포함",
        "2013-all": "Office 2013 패치셋에 포함",
        "2010-all": "Office 2010 패치셋에 포함",
        "2007-all": "Office 2007 패치셋에 포함",
        "2003-all": "Office 2003 패치셋에 포함",
        # Microsoft Edge
        "Edge-x64": "Microsoft Edge (x64)",
        "Edge-x86": "Microsoft Edge (x86)",
    }

    found_patterns = set()
    missing_patterns = set(patterns.keys())

    with open(file_list_path, "r", encoding="utf-8") as file:
        for line in file:
            stripped_line = normalize_path(line.strip()).lower()
            for pattern in patterns:
                if pattern.lower() in stripped_line:
                    found_patterns.add(pattern)

    missing_patterns -= found_patterns

    if missing_patterns:
        for pattern in sorted(missing_patterns):
            log(f"{pattern}가 미포함되어 있습니다")
    else:
        log("모든 패치셋이 포함되어 있습니다.")


check_office_patch_inclusion(local_output_txt_path)

no_ayt_files = []
log("동일한 이름의 .ayt 파일이 존재하지 않는 파일:")

for file in all_file_names:
    if file.startswith("ms_files/") and file.endswith((".cab", ".exe")):
        ayt_file = f"{file}.ayt"
        if ayt_file not in all_file_names:
            print(file)
            no_ayt_files.append(file)

if not no_ayt_files:
    log("모든 파일에 대해 동일한 이름의 .ayt 파일이 존재합니다.")
else:
    log("동일한 이름의 .ayt 파일이 존재하지 않는 파일 목록 작성 완료")
