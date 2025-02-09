import os
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
from docx import Document
from docx.shared import Pt


class PatchFileManager:
    def __init__(self, base_path):
        self.base_path = base_path

    def get_patch_dir(self):
        subdirs = [
            f
            for f in os.listdir(self.base_path)
            if os.path.isdir(os.path.join(self.base_path, f))
        ]
        if len(subdirs) == 1:
            return subdirs[0]
        else:
            raise Exception("data 폴더 하위에 폴더가 하나만 있어야 합니다.")

    def find_v1_dir(self):
        for root, dirs, _ in os.walk(self.base_path):
            if "pms_patch" in dirs:
                pms_patch_path = os.path.join(root, "pms_patch")
                v1_path = os.path.join(pms_patch_path, "v1")
                if os.path.exists(v1_path):
                    return v1_path
        raise Exception("pms_patch/v1 폴더를 찾을 수 없습니다.")

    def map_ms_patch_titles(self, file_list, ms_excel_file):
        df = pd.read_excel(ms_excel_file)
        ms_patch_title_map = dict(zip(df["패치파일"], df["제목"]))

        mapped_files = []

        for file in file_list:
            title = ms_patch_title_map.get(file, None)
            if title:
                mapped_files.append(f"{title}\n{file}")
            else:
                mapped_files.append(file)

        return mapped_files

    def list_files_in_v1(self, v1_dir_path, ms_excel_file):
        if not os.path.exists(v1_dir_path) or not os.path.isdir(v1_dir_path):
            return "v1 폴더가 존재하지 않습니다."

        result = []
        section_count = 0
        sw_files_grouped = defaultdict(list)

        for root, dirs, files in os.walk(v1_dir_path):
            if "sw_files" in dirs:
                dirs.remove("sw_files")
            relative_path = os.path.relpath(root, v1_dir_path).replace(os.sep, "/")
            filtered_files = [
                file
                for file in files
                if file.endswith((".cab", ".exe")) and not file.endswith(".zip")
            ]
            if filtered_files:
                section_count += 1
                prefix = chr(96 + section_count) + "."
                mapped_files = list(
                    dict.fromkeys(
                        self.map_ms_patch_titles(filtered_files, ms_excel_file)
                    )
                )
                file_list = [f"{i + 1}) {file}" for i, file in enumerate(mapped_files)]
                section = (
                    f"{prefix} {relative_path} 하위 {len(filtered_files)}개 파일 확인\n"
                    + "\n".join(file_list)
                )
                result.append(section)

        sw_files_path = os.path.join(v1_dir_path, "sw_files")
        if os.path.exists(sw_files_path):
            for root, _, files in os.walk(sw_files_path):
                relative_path = os.path.relpath(root, sw_files_path).replace(
                    os.sep, "/"
                )
                base_dir = relative_path.split("/")[0] if relative_path != "." else ""
                for file in files:
                    file_path = os.path.relpath(
                        os.path.join(root, file), os.path.join(sw_files_path, base_dir)
                    ).replace(os.sep, "/")
                    sw_files_grouped[base_dir].append(file_path)

            for base_dir, grouped_files in sw_files_grouped.items():
                section_count += 1
                prefix = chr(96 + section_count) + "."
                mapped_files = list(
                    dict.fromkeys(
                        self.map_ms_patch_titles(grouped_files, ms_excel_file)
                    )
                )
                file_list = [f"{i + 1}) {file}" for i, file in enumerate(mapped_files)]
                section = (
                    f"{prefix} sw_files/{base_dir} 하위 {len(grouped_files)}개 파일 확인\n"
                    + "\n".join(file_list)
                )
                result.append(section)

        return "\n\n".join(result)

    def get_numeric_dirs(self):
        patch_dir = self.get_patch_dir()
        standalone_path = os.path.join(
            self.base_path, patch_dir, "pms_patch", "standalone"
        )

        if not os.path.exists(standalone_path):
            raise FileNotFoundError(
                f"'standalone' 폴더를 찾을 수 없습니다: {standalone_path}"
            )

        numeric_dirs = [
            f
            for f in os.listdir(standalone_path)
            if os.path.isdir(os.path.join(standalone_path, f)) and f.isdigit()
        ]
        numeric_dirs.sort()
        return numeric_dirs
