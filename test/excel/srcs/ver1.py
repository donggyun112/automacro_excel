import tkinter as tk
from tkinter import filedialog
import xlwings as xw
import threading
import time
import sys
import multiprocessing
from PIL import Image

class Excel:
	user_cordinations = []
	old_path = None

	def __init__(self):
		self.workbooks = []
		self.paths = []
		self.paths.reverse()
		self.monitor_threads = []
		self.stop_events = []
		self.root = tk.Tk()
		self.root.title("Excel Image Inserter")
		self.create_gui()

	def create_gui(self):
		self.path_label = tk.Label(self.root, text="파일 경로:")
		self.path_label.pack()

		self.path_entries = []
		self.browse_buttons = []

		# for i in range(5):
		#     path_entry = tk.Entry(self.root, width=50)
		#     path_entry.pack()
		#     self.path_entries.append(path_entry)
		browse_button = tk.Button(self.root, text="찾아보기", command=lambda idx=0: self.get_file_path(idx))
		browse_button.pack()
		# self.browse_buttons.append(browse_button)

		self.insert_button = tk.Button(self.root, text="실행", command=self.run_all)
		self.insert_button.pack()

		self.insert_image_button = tk.Button(self.root, text="사진 삽입", command=self.insert_image_all)
		self.insert_image_button.pack()

		self.command_label = tk.Label(self.root, text="명령어:")
		self.command_label.pack()

		self.command_entry = tk.Entry(self.root, width=50)
		self.command_entry.pack()

		self.execute_button = tk.Button(self.root, text="명령어 실행", command=self.execute_command)
		self.execute_button.pack()

	def browse_file(self, idx):
		path = filedialog.askopenfilename(parent=self.root, filetypes=[("Excel files", "*.xlsx")])
		self.paths[idx] = path
		self.path_entries[idx].delete(0, tk.END)
		self.path_entries[idx].insert(0, path)
		self.open_workbook(idx)

	def insert_image_all(self):
		for idx, workbook in enumerate(self.workbooks):
			if workbook:
				self.insert_image(idx)

	def insert_image(self, idx):
		workbook = self.workbooks[idx]
		if workbook:
			image_paths = filedialog.askopenfilenames(parent=self.root, filetypes=[("Image files", "*.jpg *.jpeg *.png")])
			if not image_paths:
				print(f"이미지 파일을 선택하지 않았습니다. (파일 {idx+1})")
				return

			if idx >= len(self.user_cordinations) or not self.user_cordinations[idx]:
				print(f"셀을 선택하지 않았습니다. (파일 {idx+1})")
				return

			sheet_name, start_column, start_row, end_column, end_row = self.user_cordinations[idx]

			num_rows = end_row - start_row + 1
			num_columns = end_column - start_column + 1

			range_width = workbook.sheets[sheet_name].range((start_row, start_column), (start_row, end_column)).width
			range_height = workbook.sheets[sheet_name].range((start_row, start_column), (end_row, start_column)).height

			for i, image_path in enumerate(image_paths):
				row_index = start_row + (i // num_columns)
				column_index = start_column + (i % num_columns)

				if row_index > end_row:
					print(f"선택한 셀 범위를 초과하여 이미지를 삽입할 수 없습니다. (파일 {idx+1})")
					break

				image = Image.open(image_path)
				image_width, image_height = image.size

				ratio = min(range_width / image_width, range_height / image_height)
				new_width = int(image_width * ratio)
				new_height = int(image_height * ratio)
				image = image.resize((new_width, new_height), resample=Image.LANCZOS)

				range_left = workbook.sheets[sheet_name].range((start_row, start_column), (start_row, start_column)).left
				range_top = workbook.sheets[sheet_name].range((start_row, start_column), (start_row, start_column)).top
				range_center_left = range_left + (range_width - new_width) / 2
				range_center_top = range_top + (range_height - new_height) / 2

				workbook.sheets[sheet_name].pictures.add(image_path, top=range_center_top, left=range_center_left, width=new_width, height=new_height)

			print(f"이미지 삽입 완료 (파일 {idx+1})")
		else:
			print(f"Excel 파일을 먼저 선택해주세요. (파일 {idx+1})")

	def execute_command(self):
		command = self.command_entry.get()
		if command == "insert_image":
			self.insert_image_all()
		else:
			print(f"알 수 없는 명령어: {command}")

	def get_file_path(self, idx):
		try:
			try:
				print(f"엑셀 라이센스 확인... (파일 {idx+1})")
				app = xw.App(visible=False)
				app.quit()
			except Exception as e:
				print(f"엑셀 라이센스가 없습니다. 프로그램을 종료합니다. (파일 {idx+1})")
				sys.exit(1)
			print(f"엑셀 라이센스 확인 완료 (파일 {idx+1})")
			path = filedialog.askopenfilename(parent=None, filetypes=[("Excel files", "*.xlsx")])
			while len(self.paths) <= idx:
				self.paths.append(None)
			self.paths[idx] = path
			# self.path_entries[idx].delete(0, tk.END)
			# self.path_entries[idx].insert(0, path)

		except Exception as e:
			print(f"파일 선택 오류: {str(e)} (파일 {idx+1})")
			sys.exit(1)

	def open_workbook(self, idx):
		path = self.paths[idx]
		print(f"파일 경로: {path} (파일 {idx+1})")
		if path:
			workbook = xw.Book(path)
			while len(self.workbooks) <= idx:
				self.workbooks.append(None)
			self.workbooks[idx] = workbook
			workbook.app.visible = True
			backup_path = path.replace(".xlsx", "_backup.xlsx")
			workbook.save(backup_path)
			print(f"백업 파일을 {backup_path}로 저장했습니다. (파일 {idx+1})")
		else:
			print(f"파일을 선택하지 않았습니다. (파일 {idx+1})")
			return

	def _monitor(self, idx):
		while not self.stop_events[idx].is_set() and self.workbooks[idx].app.visible:
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
					if idx < len(self.user_cordinations):
						self.user_cordinations[idx] = (sheet, start_column, start_row, end_column, end_row)
					else:
						self.user_cordinations.append((sheet, start_column, start_row, end_column, end_row))
					print(f"선택한 범위: {selected_range.address} (파일 {idx+1})")
				else:
					print(f"선택한 셀이 없습니다. (파일 {idx+1})")
					if idx < len(self.user_cordinations):
						self.user_cordinations[idx] = None
					else:
						self.user_cordinations.append(None)

				time.sleep(1)
			except KeyboardInterrupt:
				self.stop_events[idx].set()
				self.stop_events[idx].clear()
				# self.workbooks[idx].app.visible = False
				exit(0)
			except Exception as e:
				print(f"모니터링 오류: {str(e)} (파일 {idx+1})")
				print(f"파일이 닫혔거나 다른 이유로 모니터링을 종료합니다. (파일 {idx+1})")
				self.stop_events[idx].set()
				self.stop_events[idx].clear()
				exit(1)

	def run(self, idx):
		self.open_workbook(idx)
		if self.workbooks[idx]:
				stop_event = threading.Event()
				self.stop_events.append(stop_event)
				monitor_thread = threading.Thread(target=self._monitor, args=(idx,))
				monitor_thread.name = f"monitor_thread{idx+1}"
				monitor_thread.start()
				self.monitor_threads.append(monitor_thread)
				print(f"모니터링을 시작합니다. (파일 {idx+1})")
		else:
			print(f"모니터링 중입니다. 모니터링을 종료합니다. (파일 {idx+1})")

	def run_all(self):
		print("모든 파일을 실행합니다.")
		print (len(self.paths))
		for idx in range(len(self.paths)):
			self.run(idx)

	def save(self, idx):
		try:
			workbook = self.workbooks[idx]
			if workbook.app.visible:
				workbook.app.visible = False
			print(f"엑셀 파일을 저장합니다. (파일 {idx+1})")
			workbook.save(self.paths[idx])
			workbook.close()
		except Exception as e:
			print(f"오류 발생: {str(e)} (파일 {idx+1})")
			try:
				workbook = xw.Book(self.paths[idx])
				new_path = self.paths[idx].replace(".xlsx", "_new.xlsx")
				workbook.save(new_path)
				workbook.close()
				print(f"파일을 {new_path}로 저장했습니다. (파일 {idx+1})")
			except Exception as e:
				print(f"파일 저장 실패: {str(e)} (파일 {idx+1})")

	def save_all(self):
		for idx in range(len(self.workbooks)):
			self.save(idx)

	def stop(self, idx):
		self.stop_events[idx].set()
		self.workbooks[idx].close()
		self.stop_events[idx].clear()

	def stop_all(self):
		for idx in range(len(self.stop_events)):
			self.stop(idx)

if __name__ == "__main__":
	excel = Excel()
	excel.root.mainloop()