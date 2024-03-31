import xlwings as xw
from tkinter import filedialog
import threading
import time
import sys
from openpyxl.drawing.image import Image

class Excel:
	user_cordinations = []
	def __init__(self):
		self.workbook = None
		self.path = None
		self.monitor_thread = None
		self.stop_event = threading.Event()


	def open_workbook(self):
		try:
			# 엑셀 라이센스 확인
			try:
				print("엑셀 라이센스 확인...")
				app = xw.App(visible=False)
				app.quit()
			except Exception as e:
				print("엑셀 라이센스가 없습니다. 프로그램을 종료합니다.")
				sys.exit(1)
			print("엑셀 라이센스 확인 완료")
			self.path = filedialog.askopenfilename(parent=None, filetypes=[("Excel files", "*.xlsx")])
		except Exception as e:
			print(f"파일 선택 오류: {str(e)}")
			sys.exit(1)

		if self.path:
			self.workbook = xw.Book(self.path)
			self.workbook.app.visible = True
			backup_path = self.path.replace(".xlsx", "_backup.xlsx")
			self.workbook.save(backup_path)
			print(f"백업 파일을 {backup_path}로 저장했습니다.")
		else:
			print("파일을 선택하지 않았습니다.")
			sys.exit(0)


	def insert_image(self):
		# 이미지 파일 선택
		image_paths = filedialog.askopenfilenames(parent=None, filetypes=[("Image files", "*.jpg *.jpeg *.png")])
		if not image_paths:
			print("이미지 파일을 선택하지 않았습니다.")
			return

		# 사용자가 선택한 셀 범위 확인
		if not self.user_cordinations:
			print("셀을 선택하지 않았습니다.")
			return

		sheet_name, start_column, start_row, end_column, end_row = self.user_cordinations

		# 선택한 셀 범위의 크기 계산
		num_rows = end_row - start_row + 1
		print("num_rows: ", num_rows)
		num_columns = end_column - start_column + 1
		print("num_columns: ", num_columns)

		# 선택한 셀 범위의 크기 계산
		range_width = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, end_column)).width
		range_height = self.workbook.sheets[sheet_name].range((start_row, start_column), (end_row, start_column)).height
		print("range_width: ", range_width)
		print("range_height: ", range_height)
		# 이미지 삽입
		for i, image_path in enumerate(image_paths):
			row_index = start_row + (i // num_columns)
			column_index = start_column + (i % num_columns)

			if row_index > end_row:
				print("선택한 셀 범위를 초과하여 이미지를 삽입할 수 없습니다.")
				break

			image = Image(image_path)
			image_width, image_height = image.width, image.height

			# 이미지 크기 조정
			ratio = min(range_width / image_width, range_height / image_height)
			new_width = range_width * 0.7
			new_height = range_height * 0.95
			# image = image.resize((new_width, new_height), resample=Image.LANCZOS)

			# 선택한 셀 범위의 중앙 위치 계산
			range_left = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, start_column)).left
			range_top = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, start_column)).top
			range_center_left = range_left + (range_width - new_width) / 2
			range_center_top = range_top + (range_height - new_height) / 2

			# 이미지 삽입
			self.workbook.sheets[sheet_name].pictures.add(image_path, top=range_center_top, left=range_center_left, width=new_width, height=new_height)

		print("이미지 삽입 완료")



	def _monitor(self):
		while not self.stop_event.is_set() and self.workbook.app.visible:
			try:
				selected_range = xw.apps.active.selection
				if selected_range:
					sheet = selected_range.sheet.name
					start_cell = selected_range[0]
					end_cell = selected_range[-1]
					start_column = start_cell.column
					start_row = start_cell.row
					end_column = end_cell.column
					end_row = end_cell.row
					self.user_cordinations = (sheet, start_column, start_row, end_column, end_row)
					print(f"선택한 범위: {selected_range.address}")
				else:
					print("선택한 셀이 없습니다.")
					self.user_cordinations = []
				time.sleep(1)  # 1초 간격으로 모니터링
			except KeyboardInterrupt:
				exit(0)
			except Exception as e:
				print(f"모니터링 오류: {str(e)}")
				print("파일이 닫혔거나 다른 이유로 모니터링을 종료합니다.")
				self.stop_event.set()
				

	def run(self):
		self.open_workbook()
		if self.workbook:
			self.monitor_thread = threading.Thread(target=self._monitor)
			self.monitor_thread.start()

			while self.monitor_thread.is_alive():
				try:
					time.sleep(1)  # 1초 간격으로 모니터링 스레드 상태 확인
					self.insert_image()
				except KeyboardInterrupt:
					print("키보드 인터럽트로 인해 프로그램을 종료합니다.")
					break

			
			if not self.monitor_thread.is_alive():
				print("모니터링 스레드가 종료되었습니다.")
				while True:
					try:
						cmd = input("엑셀 파일을 저장하시겠습니까? (y/n): ")
						if cmd.lower() in ["y", "yes", "네", "예", "ㅇ", "ㅇㅇ", "sp"]:
							self.save()
							sys.exit(0)
						elif cmd.lower() in ["n", "no", "아니요", "아니", "ㄴ", "ㄴㄴ", "s"]:
							print("엑셀 파일을 저장하지 않습니다.")
							sys.exit(0)
						else:
							raise ValueError
					except ValueError:
						print("y 또는 n을 입력해주세요.")
						continue
					except KeyboardInterrupt:
						sys.exit(0)
			self.stop_event.set()
			self.monitor_thread.join()
			self.workbook.close()
			sys.exit(0)


	def save(self):
		try:
			if self.workbook.app.visible:
				self.workbook.app.visible = False
			print("엑셀 파일을 저장합니다.")
			self.workbook.save(self.path)
			self.workbook.close()
		except Exception as e:
			print(f"오류 발생: {str(e)}")
			# 파일이 사용 중인 경우, 다른 이름으로 저장 시도
			try:
				self.workbook = xw.Book(self.path)
				new_path = self.path.replace(".xlsx", "_new.xlsx")
				self.workbook.save(new_path)
				self.workbook.close()
				print(f"파일을 {new_path}로 저장했습니다.")
			except Exception as e:
				print(f"파일 저장 실패: {str(e)}")

		def __del__(self):
			if self.workbook:
				self.workbook.close()
		def __exit__(self, exc_type, exc_value, traceback):
			if self.workbook:
				self.workbook.close()

			

if __name__ == "__main__":
	excel = Excel()
	excel.run()