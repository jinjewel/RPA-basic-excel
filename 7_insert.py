from openpyxl import load_workbook

wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.insert_rows(8) # 8번째 줄에 1줄이 삽입됨
# ws.insert_rows(8,5) # 8번째 줄에 5줄이 삽입됨

# wb.save("sample_insert_rows.xlsx")

# ws.insert_cols(2) # 2번째 열에 1열이 삽입됨
ws.insert_cols(2, 2) # 2번째 열에 2열이 삽입됨

wb.save("sample_insert_cols.xlsx")