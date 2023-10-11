from openpyxl import Workbook

wb = Workbook() # 엑셀 파일 생성
ws = wb.active # 활성화된 sheet 가져 옴
ws.title = "Nado Sheet" # sheet의 이름을 변경

wb.save("sample.xlsx") # 파일 저장
wb.close()
