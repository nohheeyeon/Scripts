import logging
import os
import subprocess
import sys
import time
from datetime import datetime
from enum import Enum
from pathlib import Path

import pandas as pd

# 실행 순서

# 1. 패치셋의 standalone 폴더를 각 VM의 "C:\Users\사용자이름\AppData\Local\Temp\package" 경로로 가져오기 (/package 폴더는 새롭게 생성해야 합니다.)
# 2. standalone_patch_verification/data 폴더에 sw_patch_list.xlsx와 ms_patch_list.xlsx 파일 배치
# 3. standalone_patch_auto_install_X86.bat 또는 standalone_patch_auto_install_AMD64.bat를 관리자 권한으로 실행
# 4. 스크립트 실행 종료 후 업데이트 확인
#   - SW 패치 : "제어판\프로그램\프로그램 및 기능\프로그램 제거 또는 변경"에서 업데이트가 설치되었는지 확인
#   - MS 패치 : "제어판\프로그램\프로그램 및 기능\설치된 업데이트"에서 KB ID가 설치되었는지 확인
#                   또한 "C:\Users\사용자이름\AppData\Local\Temp\package" 위치에 해당 ID 값의 폴더가 생성되었는지 확인
module_name = os.path.splitext(os.path.basename(__file__))[0]
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

script_dir = Path(__file__).resolve().parent
data_dir = os.path.join(script_dir, "data")
if not os.path.exists(data_dir):
    logging.error(
        "'data' 폴더가 존재하지 않습니다. 데이터를 배치할 'data' 폴더를 생성하고 파일을 준비해주세요."
    )
    sys.exit(1)

log_dir = os.path.join(script_dir, "logs")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"{module_name}_{timestamp}.log")

logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


class PatchType(Enum):
    SW = "SW"
    MS = "MS"


class StandAlonePatchInstaller:
    def __init__(self, sw_patch_list_file, ms_patch_list_file, target_architecture):
        self.sw_patch_list_file = sw_patch_list_file
        self.ms_patch_list_file = ms_patch_list_file
        self.target_architecture = target_architecture
        self.temp_package_dir = os.path.join(
            os.environ.get("USERPROFILE", ""), r"AppData\Local\Temp\package"
        )

    def find_file_in_path(self, base_path, target_name):
        for root, dirs, files in os.walk(base_path):
            for file in files:
                if (
                    file.lower() == target_name.lower()
                    or file.lower() == f"{target_name.lower()}.exe"
                ):
                    return os.path.join(root, file)
        return None

    def install_ms_patch(self, patch_path):
        cmd = [patch_path, "/quiet", "/norestart"]

        try:
            result = subprocess.run(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
            )

            if result.returncode != 0:
                logging.error(f"설치 실패 : {patch_path}, 오류: {result.stderr}")
                return False

            logging.info(f"설치 성공 : {patch_path}")
            return True

        except subprocess.CalledProcessError as e:
            logging.error(f"패치 설치 중 subprocess.CalledProcessError 발생: {e}")
        except Exception as e:
            logging.error(f"패치 설치 중 Exception 발생: {e}")

        return False

    def install_sw_patch(self, file_path):
        try:
            logging.info(f"관리자 권한으로 파일 실행 시도: {file_path}")
            result = subprocess.run(
                ["powershell", "Start-Process", f"'{file_path}'", "-Verb", "runAs"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            if result.returncode == 0:
                logging.info(f"실행 성공: {file_path}")
                return True
            else:
                logging.error(f"실행 실패: {file_path}")
                return False
        except Exception as e:
            logging.error(f"관리자 실행 중 오류 발생: {e}")
            return False

    def parse_patch_list_from_excel(self, excel_file, column_name, keyword):
        try:
            df = pd.read_excel(excel_file, engine="openpyxl")
            filtered_data = df[df[column_name].str.contains(keyword, na=False)]
            patch_file_names = filtered_data["패치파일"].dropna().tolist()
            return patch_file_names
        except Exception as e:
            logging.error(f"엑셀 파일 파싱 중 오류 발생: {e}")
            return []

    def install_patches(self, patch_file_names, patch_type):
        for patch_file_name in patch_file_names:
            try:
                patch_file_path = self.find_file_in_path(
                    self.temp_package_dir, patch_file_name
                )
                if patch_file_path:
                    logging.info(f"패치 파일 경로: {patch_file_path}")
                    if patch_type == PatchType.SW:
                        success = self.install_sw_patch(patch_file_path)
                    elif patch_type == PatchType.MS:
                        success = self.install_ms_patch(patch_file_path)
                    if not success:
                        logging.error(f"{patch_file_name} 처리 실패, 다음 패치로 이동")
                else:
                    logging.error(f"패치 파일을 찾을 수 없음: {patch_file_name}")
            except Exception as e:
                logging.error(f"{patch_file_name} 처리 중 오류 발생: {e}")

    def process_sw_patch_list(self):
        logging.info("SW 패치 리스트 처리 중")
        keyword = "X86" if self.target_architecture == "X86" else "AMD64"
        patch_file_names = self.parse_patch_list_from_excel(
            self.sw_patch_list_file, "비트", keyword
        )
        self.install_patches(patch_file_names, PatchType.SW)

    def process_ms_patch_list(self):
        logging.info("MS 패치 리스트 처리 중")
        keyword = "32비트" if self.target_architecture == "X86" else "64비트"
        patch_file_names = self.parse_patch_list_from_excel(
            self.ms_patch_list_file, "제목", keyword
        )
        self.install_patches(patch_file_names, PatchType.MS)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        logging.error("사용법: python standalone_patch_auto_install.py [X86|AMD64]")
        sys.exit(1)

    target_architecture = sys.argv[1].upper()
    if target_architecture not in ["X86", "AMD64"]:
        logging.error("잘못된 인자 : X86 또는 AMD64만 지원됩니다.")
        sys.exit(1)

    sw_patch_list_file = os.path.join(data_dir, "sw_patch_list.xlsx")
    ms_patch_list_file = os.path.join(data_dir, "ms_patch_list.xlsx")

    if not os.path.exists(sw_patch_list_file) or not os.path.exists(ms_patch_list_file):
        logging.error(
            "data 디렉토리에 sw_patch_list.xlsx 또는 ms_patch_list.xlsx 파일이 없습니다."
        )
        sys.exit(1)

    installer = StandAlonePatchInstaller(
        sw_patch_list_file, ms_patch_list_file, target_architecture
    )

    installer.process_sw_patch_list()
    time.sleep(180)
    installer.process_ms_patch_list()

    logging.info("스크립트 종료")
