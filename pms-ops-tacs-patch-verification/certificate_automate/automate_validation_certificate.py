import os
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
from docx import Document
from docx.shared import Pt


def find_v1_folder(base_path):
    for root, dirs, files in os.walk(base_path):
        if "pms_patch" in dirs:
            pms_patch_path = os.path.join(root, "pms_patch")
            v1_path = os.path.join(pms_patch_path, "v1")
            if os.path.exists(v1_path):
                return v1_path
    raise Exception("pms_patch/v1 폴더를 찾을 수 없습니다.")


def get_unique_subfolder(base_path):
    subfolders = [
        f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))
    ]
    if len(subfolders) == 1:
        return subfolders[0]
    else:
        raise Exception("data 폴더 하위에 폴더가 하나만 있어야 합니다.")


def map_patch_titles(file_list, ms_excel_file):
    df = pd.read_excel(ms_excel_file)
    patch_title_map = dict(zip(df["패치파일"], df["제목"]))
    mapped_files = []
    for file in file_list:
        title = patch_title_map.get(file, None)
        if title:
            mapped_files.append(f"{title}\n{file}")
        else:
            mapped_files.append(file)
    return mapped_files


def count_21h2_titles(ms_excel_file):
    df = pd.read_excel(ms_excel_file)
    return df["제목"].str.contains("21H2").sum()


def count_22h2_titles(ms_excel_file):
    df = pd.read_excel(ms_excel_file)
    return df["제목"].str.contains("22H2").sum()


def list_files_in_v1(v1_folder_path, ms_excel_file):
    if not os.path.exists(v1_folder_path) or not os.path.isdir(v1_folder_path):
        return "v1 폴더가 존재하지 않습니다."

    result = []
    section_count = 0
    sw_files_grouped = defaultdict(list)

    for root, dirs, files in os.walk(v1_folder_path):
        if "sw_files" in dirs:
            dirs.remove("sw_files")
        relative_path = os.path.relpath(root, v1_folder_path).replace(os.sep, "/")
        filtered_files = [
            file
            for file in files
            if file.endswith((".cab", ".exe")) and not file.endswith(".zip")
        ]
        if filtered_files:
            section_count += 1
            prefix = chr(96 + section_count) + "."
            mapped_files = map_patch_titles(filtered_files, excel_file)
            file_list = [f"{i + 1}) {file}" for i, file in enumerate(mapped_files)]
            section = (
                f"{prefix} {relative_path} 하위 {len(filtered_files)}개 파일 확인\n"
                + "\n".join(file_list)
            )
            result.append(section)

    sw_files_path = os.path.join(v1_folder_path, "sw_files")
    if os.path.exists(sw_files_path):
        for root, dirs, files in os.walk(sw_files_path):
            relative_path = os.path.relpath(root, sw_files_path).replace(os.sep, "/")
            base_dir = relative_path.split("/")[0] if relative_path != "." else ""
            for file in files:
                file_path = os.path.relpath(
                    os.path.join(root, file), os.path.join(sw_files_path, base_dir)
                ).replace(os.sep, "/")
                sw_files_grouped[base_dir].append(file_path)

        for base_dir, grouped_files in sw_files_grouped.items():
            section_count += 1
            prefix = chr(96 + section_count) + "."
            mapped_files = map_patch_titles(grouped_files, ms_excel_file)
            file_list = [f"{i + 1}) {file}" for i, file in enumerate(mapped_files)]
            section = (
                f"{prefix} sw_files/{base_dir} 하위 {len(grouped_files)}개 파일 확인\n"
                + "\n".join(file_list)
            )
            result.append(section)

    return "\n\n".join(result)


def get_kb_ids_grouped(ms_excel_file):
    df = pd.read_excel(ms_excel_file)

    kb_21h2 = df[df["제목"].str.contains("21H2", na=False)]["KB ID"].dropna().unique()
    kb_22h2 = df[df["제목"].str.contains("22H2", na=False)]["KB ID"].dropna().unique()

    kb_21h2 = sorted(set(f"KB{str(kb_id)}" for kb_id in kb_21h2))
    kb_22h2 = sorted(set(f"KB{str(kb_id)}" for kb_id in kb_22h2))

    return ", ".join(kb_21h2), ", ".join(kb_22h2)


def get_office_kb_ids(ms_excel_file):
    df = pd.read_excel(ms_excel_file)

    office_kb_ids = (
        df[df["제목"].str.contains("보안 업데이트", na=False)]["KB ID"]
        .dropna()
        .unique()
    )

    office_kb_ids = sorted(set(f"KB{str(kb_id)}" for kb_id in office_kb_ids))

    return len(office_kb_ids), ", ".join(office_kb_ids)


def get_software_versions(ms_excel_file, software_names):
    df = pd.read_excel(ms_excel_file)

    versions = {}
    for software in software_names:
        matched_versions = (
            df[df["패치명"].str.contains(software, na=False)]["버전"].dropna().unique()
        )
        versions[software] = ", ".join(sorted(set(matched_versions)))

    return versions


def get_edge_version(excel_file):
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


def get_previous_year_and_month():
    now = datetime.now()
    if now.month == 1:
        previous_year = now.year - 1
        previous_month = 12
    else:
        previous_year = now.year
        previous_month = now.month - 1

    return previous_year, previous_month


def set_font_size(document, size):
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(size)


def update_docx_with_content(source_file_path, content_to_add, new_file_path, month):
    document = Document(source_file_path)
    try:
        data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
        folder_name = get_unique_subfolder(data_dir)
        current_year = datetime.now().year
        previous_year, previous_month = get_previous_year_and_month()
        ms_excel_file = os.path.join(data_dir, "ms_patch_list.xlsx")
        sw_excel_file = os.path.join(data_dir, "sw_patch_list.xlsx")
        count_21h2 = count_21h2_titles(ms_excel_file)
        count_22h2 = count_22h2_titles(ms_excel_file)
        kb_ids_21h2, kb_ids_22h2 = get_kb_ids_grouped(ms_excel_file)
        office_count, office_kb_ids = get_office_kb_ids(ms_excel_file)
        edge_version = get_edge_version(ms_excel_file)

        software_list = ["Adobe Acrobat Reader", "BandiZip", "Chrome"]
        software_versions = get_software_versions(sw_excel_file, software_list)

        table = document.tables[1]
        for row in table.rows:
            for cell in row.cells:
                if ".tar" in cell.text:
                    if not any(
                        f"{folder_name}.tar" in word for word in cell.text.split()
                    ):
                        cell.text = cell.text.replace(".tar", f"{folder_name}.tar")
                if "월 증분 패치" in cell.text and not cell.text.startswith(
                    f"{month}월"
                ):
                    cell.text = cell.text.replace(
                        "월 증분 패치", f"{month}월 증분 패치"
                    )
            target_numbers = {"4", "6", "7", "8", "9"}
            for row in table.rows:
                if row.cells[0].text.strip() in target_numbers:
                    for cell in row.cells:
                        if row.cells[0].text.strip() == "9":
                            header_text = f"하위 서버 PMS 클라이언트> 최신 패치 탭에서 발표일 {current_year}년 {month}월까지 노출 확인\n\n"
                            if len(row.cells) > 1:
                                right_cell = row.cells[1]
                                content = right_cell.text if right_cell.text else ""

                                content = (
                                    f"21H2({count_21h2}) - {kb_ids_21h2}\n"
                                    f"22H2({count_22h2}) - {kb_ids_22h2}\n\n"
                                    f"Office 2016({office_count}) - {office_kb_ids}\n\n"
                                )

                                for software in software_list:
                                    version = software_versions.get(
                                        software, "버전 정보 없음"
                                    )
                                    content += f"{software} - {version}\n"

                                right_cell.text = header_text + content
                        if "확인 사항" in cell.text:
                            if content_to_add not in cell.text:
                                cell.text = (
                                    cell.text.split("확인 사항")[0]
                                    + "확인 사항\n"
                                    + content_to_add
                                    + "\n"
                                    + "\n".join(cell.text.split("확인 사항")[1:])
                                )

            for row in table.rows:
                if row.cells[0].text.strip() == "10":
                    for cell in row.cells:
                        if "Edge 브라우저 업데이트 확인" in cell.text:
                            updated_text = re.sub(
                                r"Edge 브라우저 업데이트 확인\(\)",
                                f"Edge 브라우저 업데이트 확인({edge_version})",
                                cell.text,
                            )
                            cell.text = updated_text

            for row in table.rows:
                if row.cells[0].text.strip() == "13":
                    for cell in row.cells:
                        cell.text = re.sub(
                            r"(?<!\d)년", f"{previous_year}년", cell.text
                        )
                        cell.text = re.sub(
                            r"(?<!\d)월", f"{previous_month}월", cell.text
                        )

        set_font_size(document, 10)
        document.save(new_file_path)
        print(f"파일이 성공적으로 생성되었습니다: {new_file_path}")

    except Exception as e:
        print(f"오류 발생: {e}")


data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
folder_name = get_unique_subfolder(data_dir)
v1_folder_path = find_v1_folder(os.path.join(data_dir, folder_name))
print(f"v1 폴더 경로: {v1_folder_path}")
excel_file = os.path.join(data_dir, "ms_patch_list.xlsx")
docx_file_path = os.path.join(data_dir, "증분 패치 검증 QA 결과서.docx")

now = datetime.now()
month = now.strftime("%m")
date_str = now.strftime("%y%m%d")
new_docx_file_path = os.path.join(
    data_dir, f"{month}월 증분 패치 검증 QA 결과서_{date_str}.docx"
)

previous_output = list_files_in_v1(v1_folder_path, excel_file)
update_docx_with_content(docx_file_path, previous_output, new_docx_file_path, month)
