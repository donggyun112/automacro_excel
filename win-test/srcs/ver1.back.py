import tkinter as tk
from tkinter import filedialog
import xlwings as xw
import threading
import time
import sys
import multiprocessing
from openpyxl.drawing.image import Image
from PIL import Image
import os

class Excel:
	user_cordinations = []
	old_path = None

	def __init__(self):
		self.workbook = None
		self.path = None
		self.monitor_thread = None
		self.stop_event = threading.Event()
		self.root = tk.Tk()
		self.root.title("Excel Image Inserter")
		self.workbook_idx = 0
		self.create_gui()

	def create_gui(self):
		self.path_label = tk.Label(self.root, text="파일 경로:")
		self.path_label.pack()

		self.path_entry = tk.Entry(self.root, width=50)
		self.path_entry.pack()

		self.browse_button = tk.Button(self.root, text="찾아보기", command=self.get_file_path)
		self.browse_button.pack()

		self.insert_button = tk.Button(self.root, text="실행", command=self.run)
		self.insert_button.pack()

		self.insert_button = tk.Button(self.root, text="사진 삽입", command=self.insert_image)
		self.insert_button.pack()

		self.command_label = tk.Label(self.root, text="명령어:")
		self.command_label.pack()

		self.command_entry = tk.Entry(self.root, width=50)
		self.command_entry.pack()

		self.execute_button = tk.Button(self.root, text="명령어 실행", command=self.execute_command)
		self.execute_button.pack()

	def browse_file(self):
		self.path = filedialog.askopenfilename(parent=self.root, filetypes=[("Excel files", "*.xlsx")])
		self.path_entry.delete(0, tk.END)
		self.path_entry.insert(0, self.path)
		self.open_workbook()


	def insert_image(self):
		if self.workbook:
			# if self.monitor_thread and self.monitor_thread.is_alive():
			# 	print("모니터링이 종료되어있습니다 다시 실행해주세요.")
			# 	return
			# 이미지 파일 선택
			image_paths = None
			image_paths = filedialog.askopenfilenames(parent=self.root, filetypes=[("Image files", "*.jpg *.jpeg *.png")])
			print(image_paths)
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
			num_columns = end_column - start_column + 1

			# 선택한 셀 범위의 크기 계산
			range_width = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, end_column)).width
			range_height = self.workbook.sheets[sheet_name].range((start_row, start_column), (end_row, start_column)).height

			for i, image_path in enumerate(image_paths):
				row_index = start_row + (i // num_columns)
				column_index = start_column + (i % num_columns)

				

				# 파일 경로 처리
				image_path = os.path.abspath(image_path)
				image_path = os.path.normpath(image_path)

				image = Image.open(image_path)
				image_width, image_height = image.size

				# 이미지 크기 조정
				ratio = min(range_width / image_width, range_height / image_height)
				new_width = int(image_width * ratio)
				new_height = int(image_height * ratio)
				image = image.resize((new_width, new_height), resample=Image.LANCZOS)

				# 선택한 셀 범위의 중앙 위치 계산
				range_left = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, start_column)).left
				range_top = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, start_column)).top
				range_center_left = range_left + (range_width - new_width) / 2
				range_center_top = range_top + (range_height - new_height) / 2

				# 이미지 삽입
				print(image_path)
				self.workbook.sheets[sheet_name].pictures.add(image_path, top=range_center_top, left=range_center_left, width=new_width, height=new_height)

			print("이미지 삽입 완료")
		else:
			print("Excel 파일을 먼저 선택해주세요.")

	def execute_command(self):
		command = self.command_entry.get()
		# 명령어 처리 로직 추가
		if command == "insert_image":
			self.insert_image()
		else:
			print(f"알 수 없는 명령어: {command}")

	def get_file_path(self):
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
			self.path_entry.delete(0, tk.END)
			self.path_entry.insert(0, self.path)

		except Exception as e:
			print(f"파일 선택 오류: {str(e)}")
			sys.exit(1)

	def open_workbook(self):
		if self.path:
			self.workbook = xw.Book(self.path)
			self.workbook.app.visible = True
			backup_path = self.path.replace(".xlsx", "_backup.xlsx")
			self.workbook.save(backup_path)
			print(f"백업 파일을 {backup_path}로 저장했습니다.")
		else:
			print("파일을 선택하지 않았습니다.")
			return

	def _monitor(self):
		while not self.stop_event.is_set():
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
					self.stop_event.clear()
					
				time.sleep(1)  # 1초 간격으로 모니터링
			except KeyboardInterrupt:
				self.stop_event.set()
				self.stop_event.clear()
				self.workbook.app.visible = False
				exit(0)
			except Exception as e:
				print(f"모니터링 오류: {str(e)}")
				print("파일이 닫혔거나 다른 이유로 모니터링을 종료합니다.")
				self.stop_event.set()
				self.stop_event.clear()
				exit(1)
		exit(0)

	def run(self):
		
		self.open_workbook()
		if self.workbook:
			self.stop_event.set()
			time.sleep(1)
			self.stop_event.clear()
			self.monitor_thread = threading.Thread(target=self._monitor)
			self.monitor_thread.name = "monitor_thread1"
			self.monitor_thread.start()
			print("모니터링을 시작합니다.")
		else:
			print("모니터링 중입니다. 모니터링을 종료합니다.")

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

	def stop(self):
		self.stop_event.set()
		self.workbook.close()
		self.stop_event.clear()

if __name__ == "__main__":
	excel = Excel()
	# p = multiprocessing.Process(target=excel.root.mainloop)
	excel.root.mainloop() 