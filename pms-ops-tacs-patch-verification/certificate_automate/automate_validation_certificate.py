import os
from datetime import datetime

import win32com.client as win32


def create_hwp_file_with_v1_content(folder_name, excel_file):
    data_dir = os.path.dirname(os.path.abspath(__file__))
    downloads_path = os.path.join(data_dir, "data")
    now = datetime.now()
    month = now.strftime("%m")
    date_str = now.strftime("%y%m%d")
    hwp_file_name = os.path.join(
        downloads_path, f"{month}월 증분 패치 검증 QA 결과서_{date_str}.hwp"
    )

    content = list_files_in_v1(folder_name, excel_file)

    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.HAction.Run("FileNew")

    for line in content.split("\n"):
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = line
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        hwp.HAction.Run("BreakPara")

    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.TextSize = 1000
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

    hwp.SaveAs(hwp_file_name)
    hwp.Quit()

    print(f"'{hwp_file_name}' 파일이 생성되고 내용이 추가되었습니다.")
    return hwp_file_name
