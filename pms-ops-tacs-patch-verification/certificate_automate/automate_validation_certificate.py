import os
from collections import defaultdict
from datetime import datetime

import pandas as pd
from docx import Document


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


def update_docx_with_content(source_file_path, content_to_add, new_file_path, month):
    document = Document(source_file_path)

    try:
        table = document.tables[1]
        for row in table.rows:
            for cell in row.cells:
                if "월 증분 패치" in cell.text and not cell.text.startswith(
                    f"{month}월"
                ):
                    cell.text = cell.text.replace(
                        "월 증분 패치", f"{month}월 증분 패치"
                    )

        target_numbers = {"4", "6", "7", "8"}
        for row in table.rows:
            if row.cells[0].text.strip() in target_numbers:
                for cell in row.cells:
                    if "확인 사항" in cell.text:
                        cell.text = (
                            cell.text.split("확인 사항")[0]
                            + "확인 사항\n"
                            + content_to_add
                            + "\n"
                            + "\n".join(cell.text.split("확인 사항")[1:])
                        )

        document.save(new_file_path)
        print(f"파일이 성공적으로 생성되었습니다: {new_file_path}")
    except Exception as e:
        print(f"오류 발생: {e}")


data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
docx_file_path = os.path.join(data_dir, "증분 패치 검증 QA 결과서.docx")
excel_file = os.path.join(data_dir, "ms_patch_list.xlsx")
folder_name = "pms_patch/pms_patch"

now = datetime.now()
month = now.strftime("%m")
date_str = now.strftime("%y%m%d")
new_docx_file_path = os.path.join(
    data_dir, f"{month}월 증분 패치 검증 QA 결과서_{date_str}.docx"
)

previous_output = list_files_in_v1(folder_name, excel_file)

update_docx_with_content(docx_file_path, previous_output, new_docx_file_path, month)
