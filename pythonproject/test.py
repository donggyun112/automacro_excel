from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
import pandas as pd
import os
import tkinter
from tkinter import filedialog

# 엑셀 파일 읽기
file_name = 'output.xlsx'
df = pd.read_excel(file_name, header=2)

# 이미지를 삽입할 열 이름
image_column = '사진'

# 엑셀 파일 로드
workbook = load_workbook(file_name)
sheet = workbook.active

# 'D' 열의 너비 가져오기
column_width = sheet.column_dimensions['D'].width

# 각 행의 높이를 저장할 리스트
row_heights = [] # 리스트 초기화

for i, cell in enumerate(sheet['D'], start=1):
    if i >= 4:
        row_heights.append(sheet.row_dimensions[cell.row].height)

column_index = df.columns.get_loc(image_column)

# 데이터프레임의 각 행을 반복하면서 이미지 삽입
for i, row in df.iterrows():
    cell = sheet.cell(row=i+4, column=column_index+1) # openpyxl은 1부터 시작
    
    # 이미 이미지가 있는 경우 다음 행으로 넘어감
    if cell.value:
        print(f"이미지가 이미 있습니다. 행: {i+4}")
        continue
    else:
        print(f"이미지를 삽입합니다. 행: {i+4}")
        
        # 파일 탐색기 창 열기
        tkinter.Tk(useTk=True)
        root = tkinter.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename()
        
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
            cell.value = f" "
        else:
            # 파일을 선택하지 않았으면 작업 종료
            print("작업을 취소합니다.")
            break

# 수정된 엑셀 파일 저장  
output_file = 'output.xlsx'
workbook.save(output_file)