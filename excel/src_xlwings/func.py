import xlwings as xw

# prev_month_keyword: 지켜야져야하는 열 current_month_keyword: curr -> prev 로 업데이트할 열
# 업데이트할 열이 없으면 NameError 반환

def update_prev_month_col(wb, sheet_name, previous_month_keyword, current_month_keyword):
    if not previous_month_keyword or not current_month_keyword:
        return NameError("Keyword not found")
    worksheet = wb.sheets[sheet_name]
    strip_keyword = lambda x: x.replace(' ', '').replace('\n', '').replace('\t', '')
    previous_month_keyword = strip_keyword(previous_month_keyword)
    current_month_keyword = strip_keyword(current_month_keyword)
    if previous_month_keyword == current_month_keyword:
        return NameError("키워드가 같음")
    elif previous_month_keyword == '' or current_month_keyword == '':
        return NameError("키워드가 비어있음")
    
    prev_month_col = None
    curr_month_col = None
    print("Start update_prev_month_col")
    data_range = worksheet.range('A1').current_region
    last_row = data_range.last_cell.row
    last_col = data_range.last_cell.column
    prev_low = 0
    curr_low = 0
    # 모든 행을 탐색하며 '전월', '당월' 셀 찾기
    for row in range(1, last_row + 1):
        for col in range(1, last_col + 1):
            cell_value = worksheet.cells(row, col).value
            if cell_value is None:
                continue
            tmp_cell_value = strip_keyword(str(cell_value))
            if tmp_cell_value == previous_month_keyword:
                prev_month_col = col
                prev_low = row
            elif tmp_cell_value == current_month_keyword:
                curr_month_col = col
                curr_low = row
            elif not prev_month_col or not curr_month_col:
                continue
                
                

            # '전월'과 '당월' 열이 모두 찾아졌으면 데이터 덮어쓰기 작업 수행
            if prev_month_col and curr_month_col:
                for data_row in range(max(prev_low + 1, curr_low + 1), last_row + 1):
                    prev_cell = worksheet.cells(data_row, prev_month_col)
                    curr_cell = worksheet.cells(data_row, curr_month_col)
                    
                    if prev_cell.formula.startswith('='):  # 셀에 수식이 있는 경우
                        # if curr_cell.formula.startswith('='):
                        #     continue
                        # else:
                        #     continue  # 함수인 경우
                        #     prev_cell.value = curr_cell.value
                        continue
                    elif curr_cell.formula.startswith('='):
                        continue
                    elif curr_cell is None:
                        continue
                    else:
                        print(f"prev_cell: {prev_cell.value}")
                        prev_cell.value = curr_cell.value
                print(f"'{sheet_name}' 시트의 {row}번째 행에서 '{current_month_keyword}' 열의 값이 '{previous_month_keyword}' 열로 업데이트되었습니다.")
                return  None

    print(f"'{sheet_name}' 시트에서 '{previous_month_keyword}'과 '{current_month_keyword}' 열을 찾을 수 없습니다.")
    return NameError("키워드를 찾을 수 없음")

def merge_sheets(wb1, wb2, sheet_name=0, range_=None, _range_="All", prev_month_keyword=None, curr_month_keyword=None):
    statuses = True
    status = None
    # 전월 업데이트 작업 수행

    if _range_ == "All" and prev_month_keyword and curr_month_keyword:
        status = update_prev_month_col(wb1, sheet_name, prev_month_keyword, curr_month_keyword)
        print(f"status: {status}")
    try:
        strip_keyword = lambda x: x.replace(' ', '').replace('\n', '').replace('\t', '')
        if prev_month_keyword and curr_month_keyword:
            prev_month_keyword = strip_keyword(prev_month_keyword)
            curr_month_keyword = strip_keyword(curr_month_keyword)
        # 시트 가져오기
        sheet1 : xw.main.Sheet = wb1.sheets[sheet_name]
        sheet2 : xw.main.Sheet = wb2.sheets[sheet_name]
        data_region = sheet1.range('A1').current_region
        last_row = data_region.last_cell.row
        last_col = data_region.last_cell.column
        
        # 병합 범위 설정
        if range_ is None:
            merge_range = sheet1.range((1, 1), (last_row, last_col))
        else:
            merge_range = sheet1.range(range_)
            
        merge_range = sheet2.range(merge_range.address)
        
        for col in range(merge_range.column, merge_range.last_cell.column + 1):
            has_keyword = False
            for row in range(merge_range.row, merge_range.last_cell.row + 1):
                cell = sheet1.cells(row, col)
        
                if prev_month_keyword and strip_keyword(str(cell.formula)).find(prev_month_keyword) != -1:
                    has_keyword = True
                    break
            

            if has_keyword:
                continue  # '전월' 키워드가 있는 열은 건너뛰기
            for row in range(merge_range.row, merge_range.last_cell.row + 1):
                cell = sheet1.cells(row, col)
                if cell.formula:  # 셀에 수식이나 함수가 있는 경우
                    if cell.formula.startswith('='):
                        continue
                    cell.value = sheet2.cells(row, col).value

    except Exception as e:
        print(f"오류 발생: {str(e)}")
        print(f"오류 발생 위치: {e.__traceback__.tb_lineno}")
        statuses = False
    finally:
        # wb1.app.quit()
        # wb2.app.quit()
        return statuses, status