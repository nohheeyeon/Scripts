import os
import subprocess

import pandas as pd

user_home = os.path.expanduser("~")
desktop = os.path.join(user_home, "Desktop")

sw_patch_list = os.path.join(desktop, "sw_patch_list.xlsx")
ms_patch_list = os.path.join(desktop, "ms_patch_list.xlsx")

patch_folder = os.path.join(user_home, r"AppData\Local\Temp\package\standalone")


def find_file_in_folder(base_folder, target_name):
    for root, dirs, files in os.walk(base_folder):
        for file in files:
            if (
                file.lower() == target_name.lower()
                or file.lower() == f"{target_name.lower()}.exe"
            ):
                return os.path.join(root, file)
    return None


def silent_install_patch(patch_path):
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


def run_as_admin(file_path):
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


def process_sw_patch_list(excel_file):
    try:
        df = pd.read_excel(excel_file, engine="openpyxl")
        filtered_data = df[df["비트"].str.contains("AMD64", na=False)]
        patch_files = filtered_data["패치파일"].dropna().tolist()

        for patch in patch_files:
            try:
                patch_path = find_file_in_folder(patch_folder, patch)
                if patch_path:
                    print(f"{patch_path}")
                    success = run_as_admin(patch_path)
                    if not success:
                        print(f"{patch} 처리 실패, 다음 패치로 이동")
                else:
                    print(f"패치 파일을 찾을 수 없음: {patch}")
            except Exception as e:
                print(f"{patch} 처리 중 오류 발생: {e}")
    except Exception as e:
        print(f"sw_patch_list 처리 중 오류 발생: {e}")
