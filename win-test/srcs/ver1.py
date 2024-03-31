import xlwings as xw
from tkinter import filedialog
import threading
import time
import os

class Excel:
    def __init__(self):
        self.workbook = None
        self.path = None
        self.monitor_thread = None
        self.stop_event = threading.Event()

    def open_workbook(self):
        self.path = filedialog.askopenfilename(parent=None, filetypes=[("Excel files", "*.xlsx")])
        if self.path:
            self.workbook = xw.Book(self.path)
            if not self.workbook.app.visible:
                    self.workbook.app.visible = True
        else:
            print("파일을 선택하지 않았습니다.")

    def monitor(self):
        #while not self.stop_event.is_set():
            #try:
                selected_range = xw.apps.active.selection
                if selected_range:
                    selected_cell = selected_range[0]
                    print(f"선택한 셀 좌표: {selected_cell.address}")
                else:
                    print("선택한 셀이 없습니다.")
                time.sleep(1)  # 1초 간격으로 모니터링
            #except:
             
            #    break

    def run(self):
        self.open_workbook()
        if self.workbook:
            self.monitor_thread = threading.Thread(target=self.monitor)
            self.monitor_thread.start()

            input("엔터를 누르면 프로그램이 종료됩니다...")
            self.stop_event.set()
            self.monitor_thread.join()

            try:
                
                # 파일 저장 경로 및 파일명 지정
                save_path = os.path.join(os.path.dirname(self.path), "test.xlsx")
                self.workbook.save(save_path)
                self.workbook.close()
            except Exception as e:
                print(f"파일 저장 실패: {str(e)}")
                # 파일이 사용 중인 경우, 다른 이름으로 저장 시도
                try:
                    new_path = self.path.replace(".xlsx", "_new.xlsx")
                    self.workbook.save(new_path)
                    self.workbook.close()
                    print(f"파일을 {new_path}로 저장했습니다.")
                except Exception as e:
                    print(f"파일 저장 실패: {str(e)}")

if __name__ == "__main__":
    excel = Excel()
    excel.run()