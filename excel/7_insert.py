from openpyxl import load_workbook

# row, column 추가 기능

wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.insert_rows(8) # 8번째 줄에 빈 row 추가 됨
# ws.insert_rows(8, 5) # 8번째 줄 위치에서 5줄 추가
# wb.save("sample_insert_rows.xlsx")

# column 1줄 추가
# ws.insert_cols(2) # B번째 빈 column 추가 됨
ws.insert_cols(2, 3) # B번째 column 부터 3column 추가
wb.save("sample_insert_cols.xlsx")

wb.close()