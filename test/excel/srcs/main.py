
import sys
import os
from ver1 import excel

if __name__ == "__main__":
	
	client = excel()
	client.run()

import xlwings as xw

# 엑셀 애플리케이션 실행
excel_app = xw.App(visible=True)

# 새로운 엑셀 워크북 생성
workbook = excel_app.books.add()
worksheet = workbook.sheets.active

while True:
    try:
        # 사용자가 선택한 셀 좌표 출력
        selected_range = xw.apps.active.selection
        if selected_range:
            selected_cell = selected_range[0]
            print(f"선택한 셀 좌표: {selected_cell.address}")
        else:
            print("선택한 셀이 없습니다.")
        
        # 사용자 입력 대기
        input("엔터를 누르면 계속됩니다...")
    except KeyboardInterrupt:
        print("프로그램을 종료합니다.")
        break

# 엑셀 파일 저장 및 종료
workbook.save("cell_coordinates_xlwings.xlsx")
workbook.close()
excel_app.quit()