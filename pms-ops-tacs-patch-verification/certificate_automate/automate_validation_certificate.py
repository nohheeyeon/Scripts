import os
from collections import defaultdict
from datetime import datetime

import pandas as pd
from docx import Document


def create_docx_file_with_v1_content(folder_name, excel_file):
    data_dir = os.path.dirname(os.path.abspath(__file__))
    downloads_path = os.path.join(data_dir, "data")
    now = datetime.now()
    month = now.strftime("%m")
    date_str = now.strftime("%y%m%d")
    docx_file_name = os.path.join(
        downloads_path, f"{month}월 증분 패치 검증 QA 결과서_{date_str}.docx"
    )

    content = list_files_in_v1(folder_name, excel_file)

    document = Document()
    document.add_heading(f"{month}월 증분 패치 검증 QA 결과서", level=1)

    for line in content.split("\n"):
        document.add_paragraph(line)

    document.save(docx_file_name)

    print(f"'{docx_file_name}' 파일이 생성되고 내용이 추가되었습니다.")
    return docx_file_name


def map_patch_titles(file_list, excel_file):
    df = pd.read_excel(excel_file)
    patch_title_map = dict(zip(df["패치파일"], df["제목"]))

    mapped_files = []
    for file in file_list:
        title = patch_title_map.get(file, None)
        if title:
            mapped_files.append(f"{title}\n{file}")
        else:
            mapped_files.append(file)

    return mapped_files


def list_files_in_v1(folder_name, excel_file):
    data_dir = os.path.dirname(os.path.abspath(__file__))
    downloads_path = os.path.join(data_dir, "data")
    folder_path = os.path.join(downloads_path, folder_name)
    v1_folder_path = os.path.join(folder_path, "v1")

    if not os.path.exists(v1_folder_path) or not os.path.isdir(v1_folder_path):
        return "v1 폴더가 존재하지 않습니다."

    result = []
    section_count = 0
    sw_files_grouped = defaultdict(list)

    for root, dirs, files in os.walk(v1_folder_path):
        relative_path = os.path.relpath(root, v1_folder_path).replace(os.sep, "/")

        if "sw_files" in relative_path:
            parts = relative_path.split("/")
            if len(parts) > 1:
                base_dir = parts[1]
                for file in files:
                    file_path = os.path.relpath(
                        os.path.join(root, file),
                        os.path.join(v1_folder_path, "sw_files", base_dir),
                    ).replace(os.sep, "/")
                    sw_files_grouped[base_dir].append(file_path)
        else:
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

    for base_dir, grouped_files in sw_files_grouped.items():
        section_count += 1
        prefix = chr(96 + section_count) + "."
        mapped_files = map_patch_titles(grouped_files, excel_file)
        file_list = [f"{i + 1}) {file}" for i, file in enumerate(mapped_files)]
        section = (
            f"{prefix} sw_files/{base_dir} 하위 {len(grouped_files)}개 파일 확인\n"
            + "\n".join(file_list)
        )
        result.append(section)

    return "\n\n".join(result)


folder_name = "pms_patch/pms_patch"
data_dir = os.path.dirname(os.path.abspath(__file__))
downloads_path = os.path.join(data_dir, "data")
excel_file = os.path.join(downloads_path, "ms_patch_list.xlsx")

create_docx_file_with_v1_content(folder_name, excel_file)
