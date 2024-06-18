# import openpyxl
# from openpyxl.formula.translate import Translator

# def update_prev_month_col(workbook, sheet_name='3. 사용현황'):
#     worksheet = workbook[sheet_name]
#     prev_month_col = None
#     curr_month_col = None

#     # 모든 행을 탐색하며 '전월', '당월' 셀 찾기
#     for row in range(1, worksheet.max_row + 1):
#         for col in range(1, worksheet.max_column + 1):
#             cell_value = worksheet.cell(row=row, column=col).value
#             if cell_value == '전월':
#                 prev_month_col = col
#             elif cell_value == '당월':
#                 curr_month_col = col

#             # '전월'과 '당월' 열이 모두 찾아졌으면 데이터 덮어쓰기 작업 수행
#             if prev_month_col and curr_month_col:
#                 for data_row in range(row+1, worksheet.max_row + 1):
#                     prev_cell = worksheet.cell(row=data_row, column=prev_month_col)
#                     curr_cell = worksheet.cell(row=data_row, column=curr_month_col)
                    
#                     if isinstance(prev_cell, openpyxl.cell.MergedCell):
#                         continue  # 병합된 셀은 건너뛰기
#                     elif prev_cell.data_type == 's':  # 문자열인 경우
#                         prev_cell.value = curr_cell.value
#                     elif prev_cell.data_type == 'f':  # 함수나 수식인 경우
#                         if curr_cell.data_type == 'f':  # 수식인 경우
#                             prev_cell.value = '=' + str(Translator(curr_cell.value, origin='A1').translate_formula('A1'))
#                         else:  # 함수인 경우
#                             prev_cell.value = curr_cell.value
#                             prev_cell.data_type = 'n'  # 데이터 타입을 숫자로 설정
#                             prev_cell.value = '=' + str(Translator(prev_cell.value, origin='A1').translate_formula('A1'))
#                     else:  # 그 외의 경우 (숫자, 날짜 등)
#                         prev_cell.value = curr_cell.value
                        
#                 print(f"'{sheet_name}' 시트의 {row}번째 행에서 '당월' 열의 값이 '전월' 열로 업데이트되었습니다.")
#                 return  # 작업 완료 후 함수 종료

#     print(f"'{sheet_name}' 시트에서 '전월'과 '당월' 열을 찾을 수 없습니다.")