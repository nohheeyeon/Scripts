import openpyxl
import pandas as pd

# "SW 현황 및 검증서" 시트 불러오기
wb = openpyxl.local_workbook("SW_현황_및_검증서".xlsx)

# 첫 번째 시트(SW현황)를 가져오기
ws = wb.active

# D열(제품)의 모든 값을 출력하기
for cell in ws["D"]:
    print(cell.value)

# 엑셀 파일 닫기
wb.close()
