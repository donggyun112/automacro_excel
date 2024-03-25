from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
import pandas as pd
import os
import tkinter
from tkinter import filedialog


import pandas as pd

def find_sheet_name(file_name):
    try:
        # 전체 데이터프레임 읽기
        df = pd.read_excel(file_name, header=None)
        max_rows = df.shape[0]  # 전체 행 수
        
        column_names = input("열 이름을 입력하세요: ").split()
        for i in range(max_rows):
            try:
                # 각 행을 헤더로 지정하여 데이터프레임 읽기
                df = pd.read_excel(file_name, header=i)
                
                for column_name in column_names:
                    if column_name in df.columns:
                        print(f"'{i}'번째 '{column_name}' 열이 존재합니다")
                        return i
            except ValueError:
                continue
        
        print("유효한 헤더 행을 찾을 수 없습니다.")
        return -1
    except FileNotFoundError:
        print(f"'{file_name}' 파일을 찾을 수 없습니다.")
        return -1

print("Excel 파일을 선택하세요.")
file_name = filedialog.askopenfilename(parent=None, filetypes=[("Excel files", "*.xlsx")])

if not file_name:
    print("파일을 선택하지 않았습니다. 프로그램을 종료합니다.")
    exit()



try:
	col_idx = find_sheet_name(file_name)
	if col_idx == -1:
		print("해당 열을 찾을 수 없습니다. 프로그램을 종료합니다.")
		exit()
	df = pd.read_excel(file_name, header=col_idx)
except Exception as e:
    print(f"파일을 읽는 중 오류가 발생했습니다: {str(e)}")
    print("올바른 Excel 파일을 선택해 주세요. 프로그램을 종료합니다.")
    exit()



# 이미지를 삽입할 열 이름
image_column = '사진'

# 엑셀 파일 로드
workbook = load_workbook(file_name)
sheet = workbook.active

# 'D' 열의 너비 가져오기
column_width = sheet.column_dimensions['D'].width

# 각 행의 높이를 저장할 리스트
row_heights = []

# 리스트 초기화
for i, cell in enumerate(sheet['D'], start=1):
    if i >= 4:
        row_heights.append(sheet.row_dimensions[cell.row].height)

column_index = df.columns.get_loc(image_column)

# 데이터프레임의 각 행을 반복하면서 이미지 삽입
try:
	for i, row in df.iterrows():
		cell = sheet.cell(row=i+4, column=column_index+1)  # openpyxl은 1부터 시작
		
		# 이미 이미지가 있는 경우 다음 행으로 넘어감
		if cell.value:
			print(f"이미지가 이미 있습니다. 행: {i+4}")
			continue
		else:
			print(f"이미지를 삽입합니다. 행: {i+4}")
			
			# 파일 탐색기 창 열기
			file_path = filedialog.askopenfilename(parent=None, filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.gif")])
			
			# 선택한 파일 경로가 있으면 진행
			if file_path:
				image_file = os.path.basename(file_path)
				image_dir = os.path.dirname(file_path)
				img = Image(file_path)
				
				# 이미지 크기 조정
				img_width = column_width * 7.8
				img_height = row_heights[i] * 1.28
				img.width = img_width
				img.height = img_height
				
				# 이미지를 셀의 중앙에 삽입하기 위해 alignment 설정
				cell.alignment = Alignment(horizontal='center', vertical='center')
				sheet.add_image(img, cell.coordinate)
				cell.value = " "
			else:
				# 파일을 선택하지 않았으면 작업 종료
				print("작업을 취소합니다.")
				break
except Exception as e:
      print (f"이미지 삽입 중 오류가 발생했습니다: {str(e)}")
      print("프로그램을 종료합니다.")
      exit()

# 수정된 엑셀 파일 저장
print("현재 작업 내용을 저장합니다.")
output_file = 'output.xlsx'
workbook.save(output_file)