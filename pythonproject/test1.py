def find_sheet_name(file_name):
	for i in range(0, len(df.columns)):
		df = pd.read_excel(file_name, header=i)

		column_names = ['위치', '원두종류', '사진', '합계']
		for column_name in column_names:
			if column_name in df.columns:
				print(f"'{column_name}' 열이 존재합니다.")
				return i
			else:
				print(f"'{column_name}' 열이 존재하지 않습니다.")
	return -1