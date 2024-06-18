
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import threading
import shutil
import atexit
import time
import sys
import os
from compare import CompareMerge
import ctypes
from imports import *
import re
import json

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("설정")
        self.setGeometry(100, 100, 400, 400)
        
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(50, 50, 50))
        self.setPalette(palette)

        layout = QVBoxLayout()

        # 이미지 가로폭 옵션
        image_width_layout = QHBoxLayout()
        image_width_label = QLabel("이미지 가로폭:")
        image_width_label.setStyleSheet("color: white; font-size: 14px;")
        self.image_width_spinbox = QDoubleSpinBox()
        self.image_width_spinbox.setRange(0.0, 1.0)  # 가로폭 범위 설정 (0.0 ~ 1.0)
        self.image_width_spinbox.setValue(self.parent.image_width)  # 초기값 설정
        self.image_width_spinbox.setDecimals(2)  # 소수점 이하 2자리까지 표시
        self.image_width_spinbox.setSingleStep(0.01)  # 스텝 크기 설정
        self.image_width_spinbox.setStyleSheet("""
            QDoubleSpinBox {
                background-color: #ecf0f1;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                font-size: 14px;
            }
        """)
        image_width_layout.addWidget(image_width_label)
        image_width_layout.addWidget(self.image_width_spinbox)
        layout.addLayout(image_width_layout)

        # 이미지 세로폭 옵션
        image_height_layout = QHBoxLayout()
        image_height_label = QLabel("이미지 세로폭:")
        image_height_label.setStyleSheet("color: white; font-size: 14px;")
        self.image_height_spinbox = QDoubleSpinBox()
        self.image_height_spinbox.setRange(0.0, 1.0)  # 세로폭 범위 설정 (0.0 ~ 1.0)
        self.image_height_spinbox.setValue(self.parent.image_height)  # 초기값 설정
        self.image_height_spinbox.setDecimals(2)  # 소수점 이하 2자리까지 표시
        self.image_height_spinbox.setSingleStep(0.01)  # 스텝 크기 설정
        self.image_height_spinbox.setStyleSheet("""
            QDoubleSpinBox {
                background-color: #ecf0f1;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                font-size: 14px;
            }
        """)
        image_height_layout.addWidget(image_height_label)
        image_height_layout.addWidget(self.image_height_spinbox)
        layout.addLayout(image_height_layout)

        # 이미지 품질 옵션 (기존 코드 유지)
        image_quality_layout = QHBoxLayout()
        image_quality_label = QLabel("이미지 품질:")
        image_quality_label.setStyleSheet("color: white; font-size: 14px;")
        self.image_quality_spinbox = QSpinBox()
        self.image_quality_spinbox.setRange(0, 100)
        self.image_quality_spinbox.setValue(self.parent.image_quality)
        self.image_quality_spinbox.setStyleSheet("""
            QSpinBox {
                background-color: #ecf0f1;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                font-size: 14px;
            }
        """)
        image_quality_layout.addWidget(image_quality_label)
        image_quality_layout.addWidget(self.image_quality_spinbox)
        layout.addLayout(image_quality_layout)

        # 저장한 파일 자동 열기 옵션 (기존 코드 유지)
        auto_open_layout = QHBoxLayout()
        auto_open_label = QLabel("저장한 파일 자동 열기:")
        auto_open_label.setStyleSheet("color: white; font-size: 14px;")
        self.auto_open_button = QPushButton()
        self.auto_open_button.setCheckable(True)
        self.auto_open_button.setChecked(self.parent.auto_open)  # 초기값 설정
        self.update_auto_open_button_text()
        self.auto_open_button.clicked.connect(self.toggle_auto_open)
        self.auto_open_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:checked {
                background-color: #2ecc71;
            }
        """)
        auto_open_layout.addWidget(auto_open_label)
        auto_open_layout.addWidget(self.auto_open_button)
        layout.addLayout(auto_open_layout)

        # 저장 경로 지정 옵션 (수정된 코드)
        save_path_layout = QHBoxLayout()
        save_path_label = QLabel("저장 경로:")
        save_path_label.setStyleSheet("color: white; font-size: 14px;")
        self.save_path_entry = QLineEdit(self.parent.save_path)
        self.save_path_entry.setReadOnly(True)  # 읽기 전용으로 설정
        self.save_path_entry.setStyleSheet("""
            QLineEdit {
                background-color: #ecf0f1;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                font-size: 14px;
            }
        """)
        save_path_button = QPushButton("찾아보기")
        save_path_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #3498db;
            }
        """)
        save_path_button.clicked.connect(self.choose_save_path)
        save_path_layout.addWidget(save_path_label)
        save_path_layout.addWidget(self.save_path_entry)
        save_path_layout.addWidget(save_path_button)
        layout.addLayout(save_path_layout)

        reset_button = QPushButton("초기화")
        reset_button.clicked.connect(self.reset_settings)
        reset_button.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
            QPushButton:pressed {
                background-color: #e67e22;
            }
        """)

        button_layout = QHBoxLayout()
        ok_button = QPushButton("확인")
        ok_button.clicked.connect(self.accept)
        ok_button.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
            QPushButton:pressed {
                background-color: #2ecc71;
            }
        """)
        cancel_button = QPushButton("취소")
        cancel_button.clicked.connect(self.reject)
        cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
            QPushButton:pressed {
                background-color: #95a5a6;
            }
        """)
        
        button_layout.addWidget(reset_button)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)


    def choose_save_path(self):
        print("choose_save_path")
        while True:
            save_path, _ = QFileDialog.getSaveFileName(self, "결과 파일 저장", "", "Excel Files (*.xlsx)")
            if not save_path:
                return
            if not save_path.endswith(".xlsx"):
                QMessageBox.critical(self, "오류", "파일 확장자는 .xlsx로 설정해주세요.")
                continue
            if os.path.exists(save_path):
                result = QMessageBox.question(self, "확인", "이미 존재하는 파일입니다. 덮어쓰시겠습니까?", QMessageBox.Yes | QMessageBox.No)
                if result == QMessageBox.Yes:
                    break
            base_name = os.path.basename(save_path)
            if not self.is_valid_filename(base_name):
                QMessageBox.critical(self, "오류", "파일 이름이 유효하지 않습니다.")
                continue
            break
        self.save_path_entry.setText(save_path)

    def is_valid_filename(self, filename):
        # 파일 이름 유효성 검사
        if filename.startswith(" "):
            return False
        if filename.strip() == "":
            return False
        pattern = r'^[a-zA-Z0-9가-힣_\-.()\s]+$'
        return re.match(pattern, filename) is not None

    def update_auto_open_button_text(self):
        self.auto_open_button.setText("On" if self.parent.auto_open else "Off")

    def toggle_auto_open(self):
        self.parent.auto_open = not self.parent.auto_open
        self.update_auto_open_button_text()

    def reset_settings(self):
        # 초기값으로 설정값 되돌리기
        self.image_width_spinbox.setValue(0.8)
        self.image_height_spinbox.setValue(0.97)
        self.image_quality_spinbox.setValue(85)
        self.auto_open_button.setChecked(True)
        self.parent.auto_open = True
        self.update_auto_open_button_text()
        self.save_path_entry.setText("")

    def accept(self):
        # 설정값 저장
        self.parent.image_width = self.image_width_spinbox.value()
        self.parent.image_height = self.image_height_spinbox.value()
        self.parent.image_quality = self.image_quality_spinbox.value()
        self.parent.auto_open = self.auto_open_button.isChecked()
        self.parent.save_path = self.save_path_entry.text()
        super().accept()

def optimize_image(image_path, output_path, quality=85):
    # 이미지 열기
    with PILImage.open(image_path) as img:
        # RGBA 모드인 경우 RGB 모드로 변환
        if img.mode == "RGBA":
            img = img.convert("RGB")
        # JPEG 포맷으로 저장
        img.save(output_path, "JPEG", optimize=True, quality=quality)
    original_size = os.path.getsize(image_path)
    # # 최적화된 이미지 크기
    optimized_size = os.path.getsize(output_path)

    # # 크기 비교
    print(f"Original Size: {original_size} bytes")
    print(f"Optimized Size: {optimized_size} bytes")
    print(f"Size Reduction: {(original_size - optimized_size) / original_size * 100:.2f}%")

class MonitorThread(QThread):
    selected_range_signal = pyqtSignal(str)
    monitoring_error_signal = pyqtSignal(str)

    def __init__(self, workbook, stop_event, parent=None):
        super().__init__()
        self.activate = True
        self.workbook = workbook
        self.stop_event = stop_event
        self.user_cordinations = []
        self.selected_range = None
        self.copy_selected_range = None
    def run(self):
        while not self.stop_event.is_set():
            try:
                self.selected_range = xw.apps.active.selection
                if self.selected_range:
                    sheet = self.selected_range.sheet.name
                    start_cell = self.selected_range[0]
                    end_cell = self.selected_range[-1]
                    start_column = start_cell.column
                    start_row = start_cell.row
                    end_column = end_cell.column
                    end_row = end_cell.row
                    self.user_cordinations = (sheet, start_column, start_row, end_column, end_row)
                    
                    # print(f"선택한 범위: {selected_range.address}")
                    self.selected_range_signal.emit(f"선택한 좌표: {self.selected_range.address}")
                    self.copy_selected_range = self.selected_range.address
                else:
                    self.user_cordinations = []
                    self.stop_event.clear()
                    self.selected_range_signal.emit("선택한 좌표: ")
                time.sleep(1)  # 1초 간격으로 모니터링
            except Exception as e:
                print(f"모니터링 오류: {str(e)}")
                print("파일이 닫혔거나 다른 이유로 모니터링을 종료합니다.")
                self.monitoring_error_signal.emit(str(e))
                break


class MonitorWorkbookThread(QThread):
    workbook_closed_signal = pyqtSignal()

    def __init__(self, workbook, name=None):
        super().__init__()
        self.workbook = workbook
        self.user_state = None
        self.name = name
        self.selected_range = None

    def run(self):
        while self.workbook and (not self.user_state or self.user_state != 2):
            try:
                self.selected_range = xw.apps.active.selection
                # active_workbook = xw.apps.active.books.active
                # if active_workbook.name == self.name:
                #     selected_range = xw.apps.active.selection
                #     pass
                # else:
                #     pass
            except Exception as e:
                print(e)
                self.workbook_closed_signal.emit()
                break
            time.sleep(0.5)  # 0.5초 간격으로 확인

    def stop(self):
        self.quit()
        self.user_state = 2
        self.quit()
        self.wait()


class ExcelHelper(QMainWindow):
    stop_monitoring_workbook = pyqtSignal()
    def __init__(self):
        self.activate = False
        error_path = os.path.join(os.path.dirname(__file__), "error.txt")
        self.f = open(error_path, "w")
        super().__init__()
        self.setWindowTitle("Excel Image Inserter")
        self.setGeometry(100, 100, 800, 600)
        
        # 감각적인 색상 팔레트 설정
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(50, 50, 50))
        palette.setColor(QPalette.WindowText, QColor(255, 255, 255))
        palette.setColor(QPalette.Base, QColor(80, 80, 80))
        palette.setColor(QPalette.AlternateBase, QColor(60, 60, 60))
        palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 220))
        palette.setColor(QPalette.ToolTipText, QColor(50, 50, 50))
        palette.setColor(QPalette.Text, QColor(255, 255, 255))
        palette.setColor(QPalette.Button, QColor(80, 80, 80))
        palette.setColor(QPalette.ButtonText, QColor(255, 255, 255))
        palette.setColor(QPalette.BrightText, QColor(255, 255, 255))
        palette.setColor(QPalette.Link, QColor(42, 130, 218))
        palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.HighlightedText, QColor(0, 0, 0))
        self.setPalette(palette)


        self.workbook = None
        self.path = None
        self.monitor_thread = None
        self.stop_event = None
        self.user_state = None
        self.monitor_workbook_thread = None
        self.image_width = 0.8
        self.image_height = 0.97
        self.image_quality = 85
        self.auto_open = True
        self.save_path = None
        self.settings_path = os.path.join(os.path.dirname(__file__), "setting", "settings.json")

        qmessagebox_style = """
        QMessageBox {
            background-color: #323232;
            color: white;
        }

        QMessageBox QLabel {
            color: white;
        }

        QMessageBox QPushButton {
            background-color: #34495e;
            color: white;
            border-radius: 5px;
            padding: 5px;
            min-width: 80px;
        }

        QMessageBox QPushButton:hover {
            background-color: #2980b9;
        }

        QMessageBox QPushButton:pressed {
            background-color: #2c3e50;
        }

        QMessageBox#information {
            background-color: #27ae60;
        }

        QMessageBox#critical {
            background-color: #c0392b;
        }

        QMessageBox#warning {
            background-color: #f39c12;
        }
        """
        QApplication.instance().setStyleSheet(qmessagebox_style)

        self.init_ui()
        self.check_excel_install()
        self.init_settings()
     

    def init_settings(self):
        try:
            if os.path.exists(self.settings_path):
                with open(self.settings_path, "r") as s:
                    settings = json.load(s)
                    self.image_width = settings.get("image_width", 0.8)
                    self.image_height = settings.get("image_height", 0.97)
                    self.image_quality = settings.get("image_quality", 85)
                    self.auto_open = settings.get("auto_open", True)
                # self.s = open(self.settings_path, "w")
        except Exception as e:
                with open(self.settings_path, "w") as s:
                    settings = {
                        "image_width": self.image_width,
                        "image_height": self.image_height,
                        "image_quality": self.image_quality,
                        "auto_open": self.auto_open
                    }
                    json.dump(settings, s, indent=4)

    def inset_settings(self):
        try:
            with open(self.settings_path, "w") as s:
                settings = {
                    "image_width": self.image_width,
                    "image_height": self.image_height,
                    "image_quality": self.image_quality,
                    "auto_open": self.auto_open
                }
                json.dump(settings, s, indent=4)
        except Exception as e:
            print("Failed to save settings")

    def init_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)		
        layout = QVBoxLayout()
        central_widget.setLayout(layout)		
        path_layout = QHBoxLayout()
        path_label = QLabel("파일 경로:")
        path_label.setFixedWidth(80)
        path_label.setStyleSheet("color: white;")
        self.path_display = QLineEdit()
        self.path_display.setReadOnly(True)
        self.path_display.setStyleSheet("background-color: #808080; color: white; border-radius: 10px; padding: 5px;")
        browse_button = QPushButton(QIcon("folder.png"), "찾아보기")
        browse_button.clicked.connect(self.browse_file)
        browse_button.setStyleSheet("""
            QPushButton {
                background-color: #808080;
                color: white;
                border-radius: 10px;
                padding: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #6d6d6d;
            }
            QPushButton:pressed {
                background-color: #5a5a5a;
            }
        """)
        path_layout.addWidget(path_label)
        path_layout.addWidget(self.path_display)
        path_layout.addWidget(browse_button)
        layout.addLayout(path_layout)		
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 10, 0, 10)
        button_widget = QWidget()
        button_widget.setLayout(button_layout)		
        # 실행 버튼
        self.run_button = QPushButton(QIcon("play.png"), "실행")
        self.run_button.clicked.connect(self.run)
        self.run_button.setEnabled(False)
        self.run_button.setFixedSize(100, 30)
        self.run_button.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border-radius: 10px;
                padding: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
            QPushButton:disabled {
                background-color: #808080;
                color: #a0a0a0;
            }
        """)		

        # 비교 병합 버튼
        self.compare_merge_button = QPushButton(QIcon("merge.png"), "비교 병합")
        self.compare_merge_button.clicked.connect(self.compare_merge)
        self.compare_merge_button.setFixedSize(100, 30)
        self.compare_merge_button.setEnabled(False)
        self.compare_merge_button.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                border-radius: 10px;
                padding: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
            QPushButton:disabled {
                background-color: #808080;
                color: #a0a0a0;
            }
        """)


        # 사진 삽입 버튼
        self.insert_image_button = QPushButton(QIcon("image.png"), "사진 삽입")
        self.insert_image_button.clicked.connect(self.insert_image)
        self.insert_image_button.setFixedSize(100, 30)
        self.insert_image_button.setEnabled(False)
        self.insert_image_button.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                border-radius: 10px;
                padding: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
            QPushButton:disabled {
                background-color: #808080;
                color: #a0a0a0;
            }
        """)		
        # 저장 버튼
        self.save_button = QPushButton(QIcon("save.png"), "저장")
        self.save_button.clicked.connect(self.save_workbook)
        self.save_button.setFixedSize(100, 30)
        self.save_button.setEnabled(False)
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 10px;
                padding: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #808080;
                color: #a0a0a0;
            }
        """)		
        button_layout.addStretch(1)
        button_layout.addWidget(self.run_button)
        button_layout.addWidget(self.compare_merge_button)
        button_layout.addWidget(self.insert_image_button)
        button_layout.addWidget(self.save_button)
        layout.addWidget(button_widget)		
        command_layout = QHBoxLayout()
        command_label = QLabel("명령어:")
        command_label.setFixedWidth(80)
        command_label.setStyleSheet("color: white;")
        self.command_entry = QLineEdit()
        self.command_entry.setStyleSheet("background-color: #808080; color: white; border-radius: 10px; padding: 5px;")
        execute_button = QPushButton(QIcon("execute.png"), "실행")
        execute_button.setFixedWidth(100)
        execute_button.clicked.connect(self.execute_command)
        execute_button.setStyleSheet("background-color: #f1c40f; color: white; border-radius: 10px; padding: 5px;")
        command_layout.addWidget(command_label)
        command_layout.addWidget(self.command_entry)
        command_layout.addWidget(execute_button)
        layout.addLayout(command_layout)		
        self.selected_range_label = QLabel("선택한 좌표: ")
        self.selected_range_label.setFont(QFont("Arial", 12))
        self.selected_range_label.setAlignment(Qt.AlignCenter)
        self.selected_range_label.setFixedHeight(40)
        self.selected_range_label.setStyleSheet("color: white;")
        layout.addWidget(self.selected_range_label)

        settings_button = QPushButton()
        _ = os.path.join(os.path.dirname(__file__), "setting", "settings_icon.png")
        settings_button.setIcon(QIcon(_))
        settings_button.setIconSize(QSize(24, 24))
        settings_button.setFixedSize(32, 32)
        settings_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
                border-radius: 5px;
            }
            QPushButton:pressed {
                background-color: #95a5a6;
                border-radius: 5px;
            }
        """)
        settings_button.clicked.connect(self.open_settings_dialog)
        layout.addWidget(settings_button, alignment=Qt.AlignLeft | Qt.AlignBottom)

        layout.addStretch(1)	
        # 윈도우 크기 조정
        self.setGeometry(100, 200, 500, 200)

    def check_excel_install(self):
        try:
            print("엑셀 설치 유무 확인...", file=self.f)
            app = xw.App(visible=False)
            app.quit()
        except Exception as e:
            print("엑셀이 설치되어있지 않습니다 프로그램을 종료합니다.", file=self.f)
            sys.exit(1)
        print("엑셀 설치 확인 완료", file=self.f)

    def browse_file(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Excel 파일 선택", "", "Excel Files (*.xlsx)")
        if self.path:
            try:
                self.stop_monitoring_workbook.emit()
                self.monitor_workbook_thread.quit()
                self.monitor_workbook_thread.wait()
                self.workbook.close()
                # self.stop_monitoring_workbook.disconnect()
                # self.monitor_workbook_thread.workbook_closed_signal.disconnect()
            except Exception as e:
                print("Failed to close")
            self.user_state = None
            try:
                self.open_workbook()
                self.path_display.setText(self.path)
                self.monitor_workbook_thread = MonitorWorkbookThread(self.workbook, name=self.workbook.name)
                self.stop_monitoring_workbook.connect(self.monitor_workbook_thread.stop)
                self.monitor_workbook_thread.workbook_closed_signal.connect(self.handle_workbook_closed)
                self.monitor_workbook_thread.start()
            except Exception as e:
                QMessageBox.critical(self, "오류", "파일을 열 수 없습니다.")
                self.path_display.setText("")
                self.run_button.setEnabled(False)
                self.save_button.setEnabled(False)
                self.compare_merge_button.setEnabled(False)
        else:
            print("엑셀 파일을 선택하지 않았습니다", file=self.f)
            self.path_display.setText("")
            self.run_button.setEnabled(False)
            self.save_button.setEnabled(False)

    def open_workbook(self):
        if self.path:
            try:
                self.workbook = xw.Book(self.path)
                self.workbook.app.visible = True
                self.compare_merge_button.setEnabled(True)
                self.run_button.setEnabled(True)
                self.save_button.setEnabled(True)
                self.compare_merge_button.setEnabled(True)
            except Exception as e:
                print(f"엑셀 파일 열기 중 오류 발생: {str(e)}", file=self.f)
                self.run_button.setEnabled(False)
                self.save_button.setEnabled(False)
        else:
            print("파일을 선택하지 않았습니다.", file=self.f)
            self.run_button.setEnabled(False)
            self.save_button.setEnabled(False)


    def run(self):
        if self.workbook:
            self.stop_monitoring_workbook.emit()
            self.user_state = 2
            self.activate = True
            self.stop_event = threading.Event()
            self.monitor_thread = MonitorThread(self.workbook, self.stop_event, parent=self)
            self.monitor_thread.selected_range_signal.connect(self.update_selected_range)
            self.monitor_thread.monitoring_error_signal.connect(self.handle_monitoring_error)
            self.monitor_thread.start()
            print("모니터링을 시작합니다.", file=self.f)
            self.run_button.setText("중단")
            self.run_button.clicked.disconnect()
            self.run_button.clicked.connect(self.stop)
            self.insert_image_button.setEnabled(True)
        else:
            print("엑셀이 열려있지 않습니다.")

    def open_settings_dialog(self):
        settings_dialog = SettingsDialog(self)
        settings_dialog.exec_()

    def stop(self):
        try:
            self.activate = False
            self.stop_event.set()
            self.run_button.setText("실행")
            self.selected_range_label.setText("선택한 좌표: ")
            self.run_button.clicked.disconnect()
            self.run_button.clicked.connect(self.run)
            self.insert_image_button.setEnabled(False)
            self.monitor_workbook_thread = MonitorWorkbookThread(self.workbook)
            self.monitor_workbook_thread.workbook_closed_signal.connect(self.handle_workbook_closed)
            self.stop_monitoring_workbook.connect(self.monitor_workbook_thread.stop)
            self.monitor_workbook_thread.start()
            self.compare_merge_dialog.close()
        except Exception as e:
            print("중단중 ...", file=self.f)

    def update_selected_range(self, selected_range):
        self.selected_range_label.setText(selected_range)

    def handle_monitoring_error(self, error_message):
        self.stop()
        return
        self.run_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.insert_image_button.setEnabled(False)
        self.compare_merge_button.setEnabled(False)
        self.path_display.setText("")
        self.run_button.setText("실행")
        self.selected_range_label.setText("선택한 좌표: ")
        self.stop_event.set()
        self.stop_event.clear()
        self.path_display.setText("")
        self.compare_merge_dialog.close()


    def insert_image(self):
        if self.workbook:
            if not self.monitor_thread or not self.monitor_thread.isRunning():
                print("모니터링이 종료되어있습니다 다시 실행해주세요.", file=self.f)
                return

            image_paths, _ = QFileDialog.getOpenFileNames(self, "이미지 파일 선택", "", "Image Files (*.jpg *.jpeg *.png)")
            if not image_paths:
                print("이미지 파일을 선택하지 않았습니다.", file=self.f)
                return

            if not self.monitor_thread.user_cordinations:
                print("셀을 선택하지 않았습니다.", file=self.f)
                return

            sheet_name, start_column, start_row, end_column, end_row = self.monitor_thread.user_cordinations

            first_image_width = None

            for i, image_path in enumerate(image_paths):
                # 이미지 최적화
                optimized_image_path = os.path.join(os.path.dirname(__file__), "images", f"optimized_image_{i}.jpg")
                optimize_image(image_path, optimized_image_path, self.image_quality)

                if i == 0:
                    # 첫 번째 사진은 현재 로직 그대로 삽입
                    range_width = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, end_column)).width
                    range_height = self.workbook.sheets[sheet_name].range((start_row, start_column), (end_row, start_column)).height
                    range_left = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, start_column)).left
                    range_top = self.workbook.sheets[sheet_name].range((start_row, start_column), (start_row, start_column)).top

                    image = Image(optimized_image_path)
                    image_width, image_height = image.width, image.height

                    first_image_width = range_width  # 첫 번째 이미지의 가로 폭 저장

                    new_width = range_width * self.image_width
                    new_height = range_height * self.image_height

                    range_center_left = range_left + (range_width - new_width) / 2
                    range_center_top = range_top + (range_height - new_height) / 2

                    self.workbook.sheets[sheet_name].pictures.add(optimized_image_path, top=range_center_top, left=range_center_left, width=new_width, height=new_height)
                else:
                    # 두 번째 사진부터는 이전에 삽입된 사진 아래의 셀에 삽입
                    next_row = end_row + 1
                    next_range = self.workbook.sheets[sheet_name].range((next_row, start_column), (next_row, end_column))

                    # 병합된 셀인지 확인
                    try:
                        if next_range.merge_area.size > 1:
                            next_range = next_range.merge_area
                    except:
                        pass

                    range_width = first_image_width  # 첫 번째 이미지의 가로 폭 사용
                    range_height = next_range.height # 다음 셀의 높이 사용
                    range_left = next_range.left
                    range_top = next_range.top

                    image = Image(optimized_image_path)
                    image_width, image_height = image.width, image.height

                    new_width = range_width * self.image_width
                    new_height = range_height * self.image_height

                    range_center_left = range_left + (range_width - new_width) / 2
                    range_center_top = range_top + (range_height - new_height) / 2

                    self.workbook.sheets[sheet_name].pictures.add(optimized_image_path, top=range_center_top, left=range_center_left, width=new_width, height=new_height)

                    end_row = next_range.last_cell.row

            print("이미지 삽입 완료")
        else:
            print("Excel 파일을 먼저 선택해주세요.")

    def execute_command(self):
        command = self.command_entry.text()
        command = command.split(":")
        if command[0] == "image":
            if len(command) != 2:
                print("이미지 명령어의 인수가 부족합니다.")
                return
            quality = int(command[1])
            self.quality = quality
            print(f"이미지 품질: {quality}")
        else:
            print(f"알 수 없는 명령어: {command}")


    def save_workbook(self):
        if self.workbook:
            try:
                self.workbook.save(self.path)
                print("엑셀 파일이 저장되었습니다.")
            except Exception as e:
                print(f"엑셀 파일 저장 중 오류 발생: {str(e)}")
        else:
            print("저장할 엑셀 파일이 없습니다.")


    def handle_workbook_closed(self):
        self.run_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.compare_merge_button.setEnabled(False)
        self.path_display.setText("")
        self.monitor_workbook_thread.exit()
        print("사용자가 엑셀 파일을 닫았습니다.")
        try:
            self.compare_merge_dialog.close()
            self.compare_merge_dialog = None
            with open(self.settings_path, "w") as f:
                settings = {
                	"image_width": self.image_width,
                	"image_height": self.image_height,
                	"image_quality": self.image_quality,
                	"auto_open": self.auto_open
                }
                json.dump(settings, f, indent=4)
        except Exception as e:
            pass


    def save_workbook(self):
        if self.workbook:
            try:
                self.workbook.save(self.path)
                print("엑셀 파일이 저장되었습니다.")
                QMessageBox.information(self, "저장 완료", "엑셀 파일이 성공적으로 저장되었습니다.")
            except Exception as e:
                print(f"엑셀 파일 저장 중 오류 발생: {str(e)}")
                QMessageBox.critical(self, "저장 실패", f"엑셀 파일 저장 중 오류가 발생했습니다: {str(e)}")
        else:
            print("저장할 엑셀 파일이 없습니다.")
            QMessageBox.warning(self, "저장 실패", "저장할 엑셀 파일이 없습니다.")

    def compare_merge(self):
        try:
            self.compare_merge_dialog = CompareMerge(self)
            self.compare_merge_dialog.exec_()
        except Exception as e:
            print(f"비교 병합 대화상자 실행 중 오류 발생: {str(e)}")
            try:
                self.compare_merge_dialog.close()
            except:
                pass


def create_hidden_folder():
    if sys.platform == "win32":
        hidden_folder_path = os.path.join(os.path.dirname(__file__), "images")
        if not os.path.exists(hidden_folder_path):
            os.makedirs(hidden_folder_path)
            ctypes.windll.kernel32.SetFileAttributesW(hidden_folder_path, 2)  # 폴더를 숨김 속성으로 설정
    else:
        hidden_folder_path = os.path.join(os.path.dirname(__file__), "images")
        if not os.path.exists(hidden_folder_path):
            os.makedirs(hidden_folder_path)
            if sys.platform == "darwin":
                os.system(f"chflags hidden {hidden_folder_path}")
            else:
                os.system(f"attrib +h {hidden_folder_path}")
    return hidden_folder_path

def delete_hidden_folder(hidden_folder_path):
    if os.path.exists(hidden_folder_path):
        if sys.platform == "win32":
            ctypes.windll.kernel32.SetFileAttributesW(hidden_folder_path, 128)  # 폴더의 숨김 속성 해제
        elif sys.platform == "darwin":
            os.system(f"chflags nohidden {hidden_folder_path}")
        shutil.rmtree(hidden_folder_path)
        print("숨김 폴더의 내용을 삭제했습니다.")

# 숨김 폴더 경로 생성
hidden_folder_path = create_hidden_folder()

# 프로그램 종료 시 실행될 함수 등록
atexit.register(delete_hidden_folder, hidden_folder_path)


def delete_images_folder():
    images_folder = hidden_folder_path
    if os.path.exists(images_folder):
        shutil.rmtree(images_folder)
        print("images 폴더의 내용을 삭제했습니다.")

# 프로그램 종료 시 실행될 함수 등록
atexit.register(delete_images_folder)


if __name__ == "__main__":
    settings_path = os.path.join(os.path.dirname(__file__), "setting", "settings.json")
    nested_folder_path = os.path.join(os.path.dirname(__file__), "images")
    try:
        os.makedirs(nested_folder_path)
        print(f"{nested_folder_path} 폴더 구조가 성공적으로 생성되었습니다.")
    except FileExistsError:
        print(f"{nested_folder_path} 폴더 구조가 이미 존재합니다.")
    except OSError:
        print(f"{nested_folder_path} 폴더 구조 생성에 실패했습니다.")
    try:
        if os.path.exists(settings_path):
            pass
        else:
            with open(settings_path, "w") as f:
                pass
    except Exception as e:
        print(f"설정 파일 생성 중 오류 발생: {str(e)}")
    app = QApplication(sys.argv)
    window = ExcelHelper()
    window.show()
    app.exec_()
    window.inset_settings()
    window.stop_monitoring_workbook.emit()
    window.close()
    window.destroy()
    window.clearFocus()
    sys.exit()