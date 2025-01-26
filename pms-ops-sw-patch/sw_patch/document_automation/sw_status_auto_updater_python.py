import openpyxl
import pandas as pd

# 작업할 엑셀 시트 불러오기
wb = openpyxl.local_workbook("SW_현황_및_검증서".xlsx)

# 첫 번째 시트와 네 번째 시트 가져오기
ws1 = wb.worksheets[0]  # 첫 번째 시트
ws4 = wb.worksheets[3]  # 네 번째 시트

# 첫 번째 시트의 D열 데이터 리스트로 저장
d_column_values = [cell.value for cell in ws1["D"]]

# 네 번째 시트와 E열 데이터와 비교하여 동일한 값 찾기
for cell in ws4["E"]:
    if cell.value in d_column_values:
        print(f"동일한 값 발견: {cell.value}")

# 엑셀 파일 닫기
wb.close()
