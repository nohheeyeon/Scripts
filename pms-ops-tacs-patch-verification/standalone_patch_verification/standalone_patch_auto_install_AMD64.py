import os
import subprocess
import time

import pandas as pd


class StandAlonePatchInstaller:
    def __init__(self, sw_patch_file, ms_patch_file):
        self.sw_patch_file = sw_patch_file
        self.ms_patch_file = ms_patch_file
        self.patch_folder = os.path.join(
            os.path.expanduser("~"), r"AppData\Local\Temp\package\standalone"
        )

    def find_file_in_folder(self, base_folder, target_name):
        for root, dirs, files in os.walk(base_folder):
            for file in files:
                if (
                    file.lower() == target_name.lower()
                    or file.lower() == f"{target_name.lower()}.exe"
                ):
                    return os.path.join(root, file)
        return None

    def silent_install_patch(self, patch_path):
        try:
            cmd = [patch_path, "/quiet", "/norestart"]
            result = subprocess.run(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
            )

            if result.returncode == 0:
                print(f"설치 성공 : {patch_path}")
                return True
            else:
                print(f"설치 실패 : {patch_path}")
                return False
        except Exception as e:
            print(f"패치 설치 중 오류 발생 : {e}")
            return False

    def run_as_admin(self, file_path):
        try:
            print(f"관리자 권한으로 파일 실행 시도: {file_path}")
            result = subprocess.run(
                ["powershell", "Start-Process", f"'{file_path}'", "-Verb", "runAs"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            if result.returncode == 0:
                print(f"실행 성공: {file_path}")
                return True
            else:
                print(f"실행 실패: {file_path}")
                return False
        except Exception as e:
            print(f"관리자 실행 중 오류 발생: {e}")
            return False

    def process_patch_list(self, excel_file, column_name, keyword, admin=False):
        try:
            df = pd.read_excel(excel_file, engine="openpyxl")
            filtered_data = df[df[column_name].str.contains(keyword, na=False)]
            patch_files = filtered_data["패치파일"].dropna().tolist()

            for patch in patch_files:
                try:
                    patch_path = self.find_file_in_folder(self.patch_folder, patch)
                    if patch_path:
                        print(f"패치 파일 경로: {patch_path}")
                        if admin:
                            success = self.run_as_admin(patch_path)
                        else:
                            success = self.silent_install_patch(patch_path)
                        if not success:
                            print(f"{patch} 처리 실패, 다음 패치로 이동")
                    else:
                        print(f"패치 파일을 찾을 수 없음: {patch}")
                except Exception as e:
                    print(f"{patch} 처리 중 오류 발생: {e}")
        except Exception as e:
            print(f"패치 리스트 처리 중 오류 발생: {e}")

    def process_sw_patch_list(self):
        print("SW 패치 리스트 처리 중")
        self.process_patch_list(self.sw_patch_file, "비트", "AMD64", admin=True)

    def process_ms_patch_list(self):
        print("MS 패치 리스트 처리 중")
        self.process_patch_list(self.ms_patch_file, "제목", "64비트", admin=False)


if __name__ == "__main__":
    user_home = os.path.expanduser("~")
    desktop = os.path.join(user_home, "Desktop")

    sw_patch_list = os.path.join(desktop, "sw_patch_list.xlsx")
    ms_patch_list = os.path.join(desktop, "ms_patch_list.xlsx")

    installer = StandAlonePatchInstaller(sw_patch_list, ms_patch_list)

    installer.process_sw_patch_list()
    print("SW 패치 처리 후 대기 중...")
    time.sleep(180)
    installer.process_ms_patch_list()
