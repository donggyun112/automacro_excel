import xlwings as xw

# 새로운 Excel 워크북 생성
wb = xw.Book()

# 활성화된 시트 선택
sheet = wb.sheets.active

# 데이터 입력
data = [
    ['이름', '나이', '직업'],
    ['홍길동', 25, '학생'],
    ['김철수', 30, '회사원'],
    ['이영희', 28, '교사']
]

# 데이터를 Excel 시트에 쓰기
sheet.range('A1').value = data

# 파일 저장
wb.save('example.xlsx')

# 워크북 닫기
wb.close()