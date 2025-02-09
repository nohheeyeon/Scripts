import calendar
import datetime
import json
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

        script_path = os.path.dirname(os.path.abspath(__file__))
        self.current_date_folder = os.path.join(
            script_path, datetime.datetime.now().strftime("%Y-%m-%d")
        )
        os.makedirs(self.current_date_folder, exist_ok=True)

        self.BASE_DIRECTORY = f"C:/ftp_root/manual_patch/{self.year}/{self.month}/ms"
        self.SW_DIRECTORY = f"C:/ftp_root/manual_patch/{self.year}/{self.month}/sw"

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

        self.patterns_json_path = os.path.join(
            script_path, "check_office_patch_inclusion.json"
        )

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

    def check_inclusion(
        self, reference_file, target_file, description, exclude_ext=None
    ):
        with open(reference_file, "r", encoding="utf-8") as ref_file:
            reference_list = set(line.strip() for line in ref_file)

        with open(target_file, "r", encoding="utf-8") as tgt_file:
            target_list = set(line.strip() for line in tgt_file)

        if exclude_ext:
            reference_list = {
                f
                for f in reference_list
                if not any(f.endswith(ext) for ext in exclude_ext)
            }
            target_list = {
                f
                for f in target_list
                if not any(f.endswith(ext) for ext in exclude_ext)
            }

        not_included = reference_list - target_list
        if not_included:
            self.log(f"{description} 포함되지 않은 파일 목록:")
            for file in sorted(not_included):
                self.log(f"- {file}")
        else:
            self.log(f"{description} 모든 파일이 포함되어 있습니다.")

    def check_office_patch_inclusion(self):

        try:
            with open(self.patterns_json_path, "r", encoding="utf-8") as json_file:
                patterns = json.load(json_file)
        except Exception as e:
            self.log(f"JSON 파일을 불러오는 중 오류 발생: {e}")
            return

        missing_patterns = set(patterns.keys())
        found_patterns = set()

        with open(self.local_output_txt_path, "r", encoding="utf-8") as file:
            for line in file:
                stripped_line = line.strip().lower()
                for pattern in patterns:
                    if pattern.lower() in stripped_line:
                        found_patterns.add(pattern)

        missing_patterns -= found_patterns

        if missing_patterns:
            for pattern in sorted(missing_patterns):
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
