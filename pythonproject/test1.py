from openpyxl import load_workbook
from openpyxl.drawing.image import Image

workbook = load_workbook('output.xlsx')
sheet = workbook.active

img = Image('images/test.png')

# 이미지를 B3 셀의 오른쪽 아래 모서리에 삽입
sheet.add_image(img, anchor=(2.8, 3.8))

# 이미지를 C4 셀의 가운데에 삽입
sheet.add_image(img, anchor=(3.5, 4.5))

workbook.save('example_with_images.xlsx')