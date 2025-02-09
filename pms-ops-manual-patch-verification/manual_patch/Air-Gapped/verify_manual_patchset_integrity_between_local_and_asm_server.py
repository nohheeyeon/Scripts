import calendar
import datetime
import os
import sys
import zipfile
from io import BytesIO

import paramiko

sys.stdout.reconfigure(encoding="utf-8")


class ManualPatchsetIntegrityVerifier:
    """
    패쇄망 패치셋 무결성 검증을 위한 클래스입니다.
    로컬 및 원격 서버의 패치 파일을 비교하고, 특정 파일이 누락되었는지 확인합니다.
    """

    def __init__(self):
        now = datetime.datetime.now()
        self.now = now
        self.year = now.year
        self.month = now.month
        self.setup_directories()
        self.setup_remote_server()
        self.all_file_names = set()
        self.no_ayt_files = []

    def setup_directories(self):
        if self.month == 1:
            self.month = 12
            self.year -= 1
        else:
            self.month -= 1

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.current_date_folder = os.path.join(
            desktop_path, datetime.datetime.now().strftime("%Y-%m-%d")
        )
        os.makedirs(self.current_date_folder, exist_ok=True)

        self.BASE_DIRECTORY = f"C:/ftp_root/manual/ms/{self.year}/{self.month}"
        self.SW_DIRECTORY = f"C:/ftp_root/manual/sw/{self.year}/{self.month}"

        self.local_output_txt_path = os.path.join(
            self.current_date_folder, "local_file_list.txt"
        )
        self.remote_output_txt_path = os.path.join(
            self.current_date_folder, "remote_file_list.txt"
        )
        self.remote_modified_output_txt_path = os.path.join(
            self.current_date_folder, "remote_modified_file_list.txt"
        )
        self.log_file_path = os.path.join(self.current_date_folder, "log.txt")

        with open(self.log_file_path, "w", encoding="utf-8") as log_file:
            log_file.write("로그 파일 시작\n")

    def setup_remote_server(self):
        self.SSH_SERVER = "SSH_SERVER"
        self.PORT = 6879
        self.USERNAME = "USERNAME"
        self.PASSWORD = "PASSWORD"
        self.REMOTE_DIRECTORY = "REMOTE_DIRECTORY"

    def log(self, message):
        timestamped_message = (
            f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
        )
        print(timestamped_message)
        with open(self.log_file_path, "a", encoding="utf-8") as log_file:
            log_file.write(timestamped_message + "\n")

    def process_top_level_zips(self, directory_path):
        for file_name in os.listdir(directory_path):
            if file_name.endswith(".zip"):
                zip_path = os.path.join(directory_path, file_name)
                self.log(f"최상위 Zip 파일 처리 중: {zip_path}")
                try:
                    with zipfile.ZipFile(zip_path, "r") as zip_file:
                        self.extract_zip_files(zip_file)
                except zipfile.BadZipFile:
                    self.log(f"잘못된 Zip 파일: {zip_path} - zip 형식이 아님")

    def extract_zip_files(self, zip_file, level=0):
        for file_info in zip_file.infolist():
            relative_path = file_info.filename.replace("\\", "/").rstrip("/")
            self.all_file_names.add(relative_path)
            if file_info.filename.endswith(".zip") and level < 1:
                self.log(f"내부 Zip 파일 처리: {relative_path}")
                with zip_file.open(file_info) as inner_file:
                    with zipfile.ZipFile(BytesIO(inner_file.read())) as nested_zip:
                        self.extract_zip_files(nested_zip, level=level + 1)

    def fetch_remote_files(self):
        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(
                self.SSH_SERVER,
                port=self.PORT,
                username=self.USERNAME,
                password=self.PASSWORD,
            )

            command = f"find {self.REMOTE_DIRECTORY} -type f -o -type d"
            stdin, stdout, stderr = ssh.exec_command(command)
            all_files = set(
                line.strip().replace(f"{self.REMOTE_DIRECTORY}/", "")
                for line in stdout.read().decode("utf-8").splitlines()
            )

            month_filter = self.now.strftime("%Y-%m")
            last_day = calendar.monthrange(self.now.year, self.now.month)[1]
            command = f"find {self.REMOTE_DIRECTORY} -type f -newermt {month_filter}-01 ! -newermt {month_filter}-{last_day}T23:59:59"
            stdin, stdout, stderr = ssh.exec_command(command)
            modified_files = set(
                line.strip().replace(f"{self.REMOTE_DIRECTORY}/", "")
                for line in stdout.read().decode("utf-8").splitlines()
            )

            self.write_file_list(modified_files, self.remote_modified_output_txt_path)
            self.log("수정된 파일 목록 작성 완료")
            return all_files, modified_files
        except Exception as e:
            self.log(f"원격 서버 접근 실패: {e}")
        finally:
            ssh.close()

    def write_file_list(self, file_list, output_path):
        with open(output_path, "w", encoding="utf-8") as output_file:
            for file in sorted(file_list):
                output_file.write(file + "\n")

    def check_inclusion(self, reference_file, target_file, description):
        with open(reference_file, "r", encoding="utf-8") as ref_file:
            reference_list = set(line.strip() for line in ref_file)

        with open(target_file, "r", encoding="utf-8") as tgt_file:
            target_list = set(line.strip() for line in tgt_file)

        not_included = reference_list - target_list
        if not_included:
            self.log(f"{description} 포함되지 않은 파일 목록:")
            for file in sorted(not_included):
                self.log(f"- {file}")
        else:
            self.log(f"{description} 모든 파일이 포함되어 있습니다.")

    def check_office_patch_inclusion(self):
        patterns = {  # Windows 11
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
            # "xp3embed-x64": "Windows Embedded XP SP3 (x64)", 지원 종류 범위에 따라 빌드 설정에서 제외
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

        missing_patterns = set(patterns.keys())

        with open(self.local_output_txt_path, "r", encoding="utf-8") as file:
            for line in file:
                for pattern in patterns:
                    if pattern in line:
                        missing_patterns.discard(pattern)

        if missing_patterns:
            for pattern in missing_patterns:
                self.log(f"{pattern} 패치셋이 누락되었습니다.")
        else:
            self.log("모든 Office 패치셋이 포함되어 있습니다.")

    def validate_ayt_files(self):
        for file in self.all_file_names:
            if file.startswith("ms_files/") and file.endswith((".cab", ".exe")):
                ayt_file = f"{file}.ayt"
                if ayt_file not in self.all_file_names:
                    self.no_ayt_files.append(file)

        if not self.no_ayt_files:
            self.log("모든 파일에 대한 .ayt 파일이 존재합니다.")
        else:
            self.log(f".ayt 파일이 없는 파일: {self.no_ayt_files}")

    def run(self):
        self.process_top_level_zips(self.BASE_DIRECTORY)
        self.process_top_level_zips(self.SW_DIRECTORY)
        self.write_file_list(self.all_file_names, self.local_output_txt_path)
        self.log("로컬 파일 목록 작성 완료")

        remote_files, _ = self.fetch_remote_files()
        self.write_file_list(remote_files, self.remote_output_txt_path)
        self.log("원격 서버 파일 목록 작성 완료")

        self.check_inclusion(
            self.local_output_txt_path,
            self.remote_output_txt_path,
            "로컬 파일 목록이 원격 파일 목록에",
        )
        self.check_inclusion(
            self.remote_modified_output_txt_path,
            self.local_output_txt_path,
            "원격 수정된 파일 목록이 로컬 파일 목록에",
            exclude_ext=[".json"],
        )

        self.check_office_patch_inclusion()
        self.validate_ayt_files()
        self.log("검증 완료.")


def main():
    verifier = ManualPatchsetIntegrityVerifier()
    verifier.run()


if __name__ == "__main__":
    main()
