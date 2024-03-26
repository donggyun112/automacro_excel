import tkinter as tk

def print_to_gui(text):
    output_text.insert(tk.END, text + "\n")
    output_text.see(tk.END)  # 자동으로 스크롤을 맨 아래로 내림

# Tkinter 창 생성
root = tk.Tk()
root.title("Output Viewer")

# 텍스트 위젯 생성
output_text = tk.Text(root, height=20, width=80)
output_text.pack()

# 출력문 예시
print_to_gui("Hello, World!")
print_to_gui("This is a sample output.")
print_to_gui("You can add multiple lines of text.")

# GUI 이벤트 루프 시작
root.mainloop()