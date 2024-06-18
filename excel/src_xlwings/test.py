import unittest
import xlwings as xw
from xlwings.constants import CellType
from unittest.mock import patch, mock_open
import pandas as pd

def update_prev_month_col(workbook, sheet_name, previous_month_keyword, current_month_keyword):
    if not previous_month_keyword or not current_month_keyword:
        return NameError("Keyword not found")
    worksheet = workbook.sheets[sheet_name]
    strip_keyword = lambda x: x.replace(' ', '').replace('\n', '').replace('\t', '')
    previous_month_keyword = strip_keyword(previous_month_keyword)
    current_month_keyword = strip_keyword(current_month_keyword)
    if previous_month_keyword == current_month_keyword:
        return NameError("키워드가 같음")
    elif previous_month_keyword == '' or current_month_keyword == '':
        return NameError("키워드가 비어있음")

    prev_month_col = None
    curr_month_col = None

    # 모든 행을 탐색하며 '전월', '당월' 셀 찾기
    for row in range(1, worksheet.used_range.last_cell.row + 1):
        for col in range(1, worksheet.used_range.last_cell.column + 1):
            cell_value = worksheet.range((row, col)).value
            tmp_cell_value = strip_keyword(str(cell_value))
            if tmp_cell_value == previous_month_keyword:
                prev_month_col = col
            elif tmp_cell_value == current_month_keyword:
                curr_month_col = col

            # '전월'과 '당월' 열이 모두 찾아졌으면 데이터 덮어쓰기 작업 수행
            if prev_month_col and curr_month_col:
                for data_row in range(row+1, worksheet.used_range.last_cell.row + 1):
                    prev_cell = worksheet.range((data_row, prev_month_col))
                    curr_cell = worksheet.range((data_row, curr_month_col))

                    if prev_cell.merge_area != prev_cell:
                        continue  # 병합된 셀은 건너뛰기
                    elif prev_cell.value is None:
                        prev_cell.value = curr_cell.value
                    elif prev_cell.value_type == CellType.string:  # 문자열인 경우
                        prev_cell.value = curr_cell.value
                    elif prev_cell.value_type == CellType.formula:  # 함수나 수식인 경우
                        if curr_cell.value_type == CellType.formula:  # 수식인 경우
                            continue
                        else:
                            continue  # 함수인 경우
                    else:  # 그 외의 경우 (숫자, 날짜 등)
                        prev_cell.value = curr_cell.value

                return None  # 작업 완료 후 함수 종료

    return NameError("키워드를 찾을 수 없음")

def merge_sheets(wb1, wb2, sheet_name=0, range_=None, prev_month_keyword=None, curr_month_keyword=None):
    if prev_month_keyword and curr_month_keyword:
        status = update_prev_month_col(wb1, sheet_name, prev_month_keyword, curr_month_keyword)
        if status:
            return False, status
    try:
        strip_keyword = lambda x: x.replace(' ', '').replace('\n', '').replace('\t', '')
        if prev_month_keyword and curr_month_keyword:
            prev_month_keyword = strip_keyword(prev_month_keyword)
            curr_month_keyword = strip_keyword(curr_month_keyword)
        # 시트 이름으로 시트 가져오기
        sheet1 = wb1.sheets[sheet_name]
        sheet2 = wb2.sheets[sheet_name]

        # 병합 범위 설정
        if range_ is None:
            merge_range = sheet2.used_range
        else:
            merge_range = xw.Range(sheet2, range_)

        # 병합 범위의 데이터를 리스트로 변환
        data2 = merge_range.value

        # 리스트를 DataFrame으로 변환
        df2 = pd.DataFrame(data2)

        # wb2의 병합된 데이터를 wb1의 시트에 저장
        for col in range(merge_range.column, merge_range.last_cell.column + 1):
            has_keyword = False
            for row in range(merge_range.row, merge_range.last_cell.row + 1):
                cell = sheet1.range((row, col))
                if cell.merge_area != cell:
                    continue  # 병합된 셀은 건너뛰기
                elif prev_month_keyword and strip_keyword(str(cell.value)).find(prev_month_keyword) != -1:
                    has_keyword = True
                    break

            if has_keyword:
                continue  # '전월' 키워드가 있는 열은 건너뛰기

            for row in range(merge_range.row, merge_range.last_cell.row + 1):
                cell = sheet1.range((row, col))
                if cell.merge_area != cell:
                    continue  # 병합된 셀은 건너뛰기
                elif cell.value_type == CellType.formula:  # 셀에 수식이나 함수가 있는 경우
                    formula_value = cell.formula
                    cell.value = df2.at[row - merge_range.row, col - merge_range.column]
                    cell.formula = formula_value
                else:
                    cell.value = df2.at[row - merge_range.row, col - merge_range.column]

    except Exception as e:
        return False, e
    finally:
        wb2.close()
        return True, None

# class TestMergeSheets(unittest.TestCase):
#     @classmethod
#     def setUpClass(cls):
#         # Excel 인스턴스 시작
#         cls.app = xw.App(visible=False)

#     @classmethod
#     def tearDownClass(cls):
#         # Excel 인스턴스 종료
#         cls.app.quit()

#     def setUp(self):
#         # 새로운 워크북 생성
#         self.wb1 = self.app.books.add()
#         self.wb2 = self.app.books.add()
#         self.sheet1 = self.wb1.sheets[0]
#         self.sheet2 = self.wb2.sheets[0]

#         # 테스트 데이터 설정
#         self.sheet1.range('A1').value = '이전달'
#         self.sheet1.range('B1').value = '이번달'
#         self.sheet1.range('A2').value = 10
#         self.sheet1.range('B2').value = 20
#         self.sheet1.range('A3').value = '=SUM(A2:A2)'
#         self.sheet1.range('B3').value = '=SUM(B2:B2)'

#         self.sheet2.range('A1').value = '이번달'
#         self.sheet2.range('B1').value = 30
#         self.sheet2.range('B2').value = 40

#     def tearDown(self):
#         # 워크북 닫기
#         self.wb1.close()
#         self.wb2.close()

#     # 다른 테스트 케이스 생략

# if __name__ == '__main__':
#     unittest.main()