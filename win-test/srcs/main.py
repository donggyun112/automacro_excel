import win32com.client as win32
from tkinter import filedialog
import tkinter as tk

def get_selected_cell():
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True

    root = tk.Tk()
    root.withdraw()  # Tk root 창 숨기기

    path = filedialog.askopenfilename(parent=None, filetypes=[("Excel files", "*.xlsx")])
    if not path:
        print("파일을 선택하지 않았습니다.")
        return

    workbook = excel.Workbooks.Open(path)

    try:
        while True:
            print("엑셀에서 셀을 선택하세요...")
            selected_range = excel.Selection
            if selected_range:
                row = selected_range.Row
                column = selected_range.Column
                print(f"선택한 셀의 좌표: 행 {row}, 열 {column}")
                break
    finally:
        save_choice = input("변경 내용을 저장하시겠습니까? (y/n): ").lower()
        if save_choice == 'y':
            workbook.save(path)
        workbook.Close(SaveChanges=False)
        excel.Quit()

if __name__ == "__main__":
    get_selected_cell()
