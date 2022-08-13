from openpyxl import load_workbook

wb = load_workbook("sample.xlsx")
ws = wb.active

# ws.delete_rows(8) # 8번째 줄에 있는 7번 학생 데이터 삭제
# ws.delete_rows(8, 2) # 8번째 줄에 있는 7, 8번 학생 데이터 삭제

# wb.save("sample_delete_rows.xlsx")

# ws.delete_cols(2) # 2열에 있는 수학 데이터 삭제
ws.delete_cols(2, 2) # 2, 3열에 있는 수학, 영어 데이터 삭제

wb.save("sample_delete_cols.xlsx")