import xlwings as xw

def print_merged_cells_coordinates(file_path, sheet_name):
    # Excel 파일 열기
    wb = xw.Book(file_path)
    
    # 지정된 시트 선택
    sheet = wb.sheets[sheet_name]
    
    # 사용된 범위(Used Range) 가져오기
    used_range = sheet.used_range
    
    # 중복된 병합된 셀을 저장할 딕셔너리
    merged_cells_dict = {}
    
    # Used Range 내의 모든 셀에 대해 순회하면서 병합된 셀인지 확인
    for cell in used_range:
        if cell.merge_cells:
            start_cell_address = cell.address
            end_cell_address = cell.merge_area.address
            
            # 이미 시작 셀이 딕셔너리에 있을 경우, 무시
            if start_cell_address in merged_cells_dict:
                continue
            
            # 시작 셀을 키로 사용하여 종료 셀을 값으로 저장
            merged_cells_dict[start_cell_address] = end_cell_address
    
    # 저장된 병합된 셀 정보 출력
    print(f"{sheet_name} 시트의 병합된 셀 좌표:")
    for start_cell, end_cell in merged_cells_dict.items():
        print(f"시작 셀: {start_cell}")
        print(f"종료 셀: {end_cell}")
        print("-------------")
    
    # 엑셀 파일 닫기
    wb.close()

# 테스트용 Excel 파일 경로
excel_file_path = "qqqq.xlsx"

# 해당 시트 이름
sheet_name = "경비배분표"

# 함수 호출
print_merged_cells_coordinates(excel_file_path, sheet_name)
