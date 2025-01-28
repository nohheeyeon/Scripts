import openpyxl
import pandas as pd

# 작업할 엑셀 시트 불러오기
wb = openpyxl.local_workbook("SW_현황_및_검증서".xlsx)

# 첫 번째 시트와 네 번째 시트 가져오기
ws1 = wb.worksheets[0]  # 첫 번째 시트
ws4 = wb.worksheets[3]  # 네 번째 시트

# 첫 번째 시트의 D열 데이터를 딕셔너리로 저장 (값 : 행 번호)
d_column_values = {cell.value for cell in ws1["D"] if cell.value is not None}

# 네번째 시트의 E열 데이터와 비교하여 동일한 값이 있는 행의 D열 데이터를 첫번째 시트의 F열(발표일)로 복사
for cell in ws4["E"]:
    if cell.value in d_column_values:
        # 동일한 값이 있는 행 번호
        row_in_ws1 = d_column_values[cell.value]
        row_in_ws4 = cell.row

        # 네번째 시트의 D열 값을 첫번째 시트의 F열로 복사
        ws1.cell(row=row_in_ws1, column=6, value=ws4.cell(row_in_ws4, column=4).value)
        print(
            f"동일한 값 발견: {cell.value} 첫 번째 시트의 {row_in_ws4}행, 네번째 시트이 {row_in_ws4}헹"
        )
        print(
            f"네번째 시트의 D열 값 '{ws4.cell(row=row_in_ws4, column=4).value}'을 첫번째 시트의 F열로 복사"
        )

# 엑셀 파일 저장 및 닫기
wb.save("SW_현황_및_검증서.xlsx")
wb.close()
