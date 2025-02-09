import re

import pandas as pd


class ExcelDataProcessor:
    def __init__(self):
        pass

    def map_patch_titles(self, file_list, ms_excel_file):
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

    def count_21h2_titles(self, ms_excel_file):
        df = pd.read_excel(ms_excel_file)
        return df["제목"].str.contains("21H2", na=False).sum()

    def count_22h2_titles(self, ms_excel_file):
        df = pd.read_excel(ms_excel_file)
        return df["제목"].str.contains("22H2", na=False).sum()

    def get_kb_ids_grouped(self, ms_excel_file):
        df = pd.read_excel(ms_excel_file)

        kb_21h2 = (
            df[df["제목"].str.contains("21H2", na=False)]["KB ID"].dropna().unique()
        )
        kb_22h2 = (
            df[df["제목"].str.contains("22H2", na=False)]["KB ID"].dropna().unique()
        )

        kb_21h2 = sorted(set(f"KB{str(kb_id)}" for kb_id in kb_21h2))
        kb_22h2 = sorted(set(f"KB{str(kb_id)}" for kb_id in kb_22h2))

        return ", ".join(kb_21h2), ", ".join(kb_22h2)

    def get_office_kb_ids(self, ms_excel_file):
        df = pd.read_excel(ms_excel_file)
        office_kb_ids = (
            df[df["제목"].str.contains("보안 업데이트", na=False)]["KB ID"]
            .dropna()
            .unique()
        )
        office_kb_ids = sorted(set(f"KB{str(kb_id)}" for kb_id in office_kb_ids))
        return len(office_kb_ids), ", ".join(office_kb_ids)

    def get_software_versions(self, ms_excel_file, software_names):
        df = pd.read_excel(ms_excel_file)
        versions = {}
        for software in software_names:
            matched_versions = (
                df[df["패치명"].str.contains(software, na=False)]["버전"]
                .dropna()
                .unique()
            )
            versions[software] = ", ".join(sorted(set(matched_versions)))
        return versions

    def get_edge_version(self, excel_file):
        df = pd.read_excel(excel_file)
        edge_versions = (
            df[df["제목"].str.contains("Edge", na=False)]["제목"].dropna().tolist()
        )
        version_pattern = r"빌드\s*([\d\.]+)"
        for title in edge_versions:
            match = re.search(version_pattern, title)
            if match:
                return match.group(1)
        return "버전 정보 없음"
