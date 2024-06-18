# -*- coding: utf-8 -*-

from imports import *
import re
from func import merge_sheets
import faulthandler
faulthandler.enable()



class SettingsDialog2(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("설정")
        self.setGeometry(100, 100, 400, 200)
        
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(50, 50, 50))
        self.setPalette(palette)

  
        layout = QVBoxLayout()

        # 키워드2 옵션
        keyword2_layout = QHBoxLayout()
        keyword2_label = QLabel("키워드1:")
        keyword2_label.setStyleSheet("color: white; font-size: 14px;")
        self.keyword2_entry = QLineEdit(self.parent.current_value)
        self.keyword2_entry.setPlaceholderText("예시: 당월, 이번달, '~20XX년. XX 누계' 등")
        self.keyword2_entry.setStyleSheet("""
            QLineEdit {
                background-color: #ecf0f1;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                font-size: 14px;
            }
        """)
        keyword2_layout.addWidget(keyword2_label)
        keyword2_layout.addWidget(self.keyword2_entry)
        layout.addLayout(keyword2_layout)

        # 키워드1 옵션
        keyword1_layout = QHBoxLayout()
        keyword1_label = QLabel("키워드2:")
        keyword1_label.setStyleSheet("color: white; font-size: 14px;")
        self.keyword1_entry = QLineEdit(self.parent.previous_value)
        self.keyword1_entry.setPlaceholderText("예시: 전월, 이전달, '~20XX년. XX 누계' 등")
        self.keyword1_entry.setStyleSheet("""
            QLineEdit {
                background-color: #ecf0f1;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                font-size: 14px;
            }
        """)
        keyword1_layout.addWidget(keyword1_label)
        keyword1_layout.addWidget(self.keyword1_entry)
        layout.addLayout(keyword1_layout)

        

        # 확인/취소 버튼
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
        button_layout.addStretch(1)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def accept(self):
        # 설정값 저장
        self.parent.previous_value = self.keyword1_entry.text()
        self.parent.current_value = self.keyword2_entry.text()
        super().accept()


class CompareMerge(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.path = None
        self.workbook = None
        self.selected_sheets = []
        self.sheet_ranges = {}
        self.app = None
        self.previous_value = None
        self.current_value = None

        self.setWindowTitle("비교 병합")
        self.resize(500, 400)

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

        

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        path_layout = QHBoxLayout()
        path_label = QLabel("파일 경로:")
        path_label.setFixedWidth(80)
        path_label.setStyleSheet("color: white;")
        self.path_display = QLineEdit()
        self.path_display.setReadOnly(True)
        self.path_display.setStyleSheet("background-color: #808080; color: white; border-radius: 10px; padding: 5px;")
        browse_button = QPushButton(QIcon("folder.png"), "찾아보기")
        browse_button.clicked.connect(self.browse_workbook)
        self.path_display.setMinimumWidth(250)
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

        settings_button = QPushButton(QIcon("settings.png"), "설정")
        settings_button.clicked.connect(self.open_settings)
        settings_button.setStyleSheet("""
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
        layout.addLayout(path_layout)
        path_layout.addStretch()
        path_layout.addWidget(browse_button)
        path_layout.addWidget(settings_button)

        label = QLabel("비교 병합할 시트를 선택하세요:")
        label.setStyleSheet("color: white;")
        layout.addWidget(label)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)

        # 스크롤 영역의 배경색 설정
        scroll_area.setStyleSheet("background-color: #323232;")

        self.sheet_buttons = []
        self.range_buttons = []

        if self.parent.workbook:
            for sheet in self.parent.workbook.sheets:
                sheet_layout = QHBoxLayout()


                button = QPushButton(sheet.name)
                button.setCheckable(True)
                button.setStyleSheet("""
                    QPushButton {
                        background-color: #bdc3c7;
                        color: #2c3e50;
                        border-radius: 10px;
                        padding: 10px;
                        font-size: 16px;
                    }
                    QPushButton:hover {
                        background-color: #a1a7ab;
                    }
                    QPushButton:pressed {
                        background-color: #8b9296;
                    }
                    QPushButton:checked {
                        background-color: #2ecc71;
                        color: white;
                    }
                    QPushButton:checked:hover {
                        background-color: #27ae60;
                    }
                    QPushButton:disabled {
                        background-color: #95a5a6;
                        color: #7f8c8d;
                    }
                """)
                button.clicked.connect(lambda _, s=sheet.name, b=button: self.toggle_sheet(s, b))
                sheet_layout.addWidget(button)
                self.sheet_buttons.append(button)

                range_button = QPushButton(f"범위 설정")
                range_button.setStyleSheet("""
                    QPushButton {
                        background-color: #95a5a6;
                        color: white;
                        border-radius: 10px;
                        padding: 5px;
                        font-size: 12px;
                    }
                    QPushButton:hover {
                        background-color: #7f8c8d;
                    }
                    QPushButton:pressed {
                        background-color: #6c7a7d;
                    }
                    QPushButton:disabled {
                        background-color: #bdc3c7;
                        color: #95a5a6;
                    }
                """)
                range_button.clicked.connect(lambda _, s=sheet.name, b=range_button: self.set_range(s, b))
                sheet_layout.addWidget(range_button)
                self.range_buttons.append(range_button)

                scroll_layout.addLayout(sheet_layout)

        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        self.execute_button = QPushButton(QIcon("execute.png"), "실행")
        self.execute_button.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                border-radius: 10px;
                padding: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
        """)
        self.execute_button.clicked.connect(self.execute_compare_merge)
        layout.addWidget(self.execute_button)

        close_button = QPushButton(QIcon("close.png"), "닫기")
        close_button.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border-radius: 10px;
                padding: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
            QPushButton:pressed {
                background-color: #a93226;
            }
        """)
        close_button.clicked.connect(self.__del__)
        layout.addWidget(close_button)

        self.setLayout(layout)

        # 파일 선택 여부에 따라 버튼 활성화/비활성화
        self.update_button_states()
        
    def update_button_states(self):
        file_selected = self.workbook is not None

        # 선택한 파일의 시트 이름 가져오기
        selected_file_sheet_names = [sheet.name for sheet in self.workbook.sheets] if file_selected else []

        for button, range_button in zip(self.sheet_buttons, self.range_buttons):
            sheet_name = button.text()
            button.setEnabled(sheet_name in selected_file_sheet_names)
            range_button.setEnabled(button.isChecked() and sheet_name in selected_file_sheet_names)

        self.execute_button.setEnabled(file_selected)
        if not file_selected:
            self.execute_button.setStyleSheet("""
                QPushButton {
                    background-color: #95a5a6;
                    color: #7f8c8d;
                    border-radius: 10px;
                    padding: 10px;
                    font-size: 16px;
                }
            """)
        else:
            self.execute_button.setStyleSheet("""
                QPushButton {
                    background-color: #2ecc71;
                    color: white;
                    border-radius: 10px;
                    padding: 10px;
                    font-size: 16px;
                }
                QPushButton:hover {
                    background-color: #27ae60;
                }
                QPushButton:pressed {
                    background-color: #1e8449;
                }
            """)

    def browse_workbook(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if path:
            if self.workbook:
                self.workbook.close()
                self.workbook = None
            self.path = path
            file_name = os.path.basename(self.path)
            self.path_display.setText(file_name)
            self.open_workbook()
            self.update_button_states()
            
    def closeEvent(self, event: QCloseEvent):
        try:
            if self.workbook:
                self.workbook.close()
                self.workbook = None
                self.app = None
        except Exception as e:
            event.accept()

    def open_settings(self):
       dialog = SettingsDialog2(self)
       dialog.exec_()

    # test
    def print_workbook_content(self):
       if self.workbook is None:
           print("워크북이 로드되지 않았습니다.")
           return
    
       for sheet_name in self.workbook.sheetnames:
           print(f"시트 이름: {sheet_name}")
           sheet = self.workbook[sheet_name]
           for row in sheet.iter_rows():
               row_data = [cell.value for cell in row]
               print(row_data)
           print()  # 시트 간 구분을 위한 빈 줄 출력
    
    def open_workbook(self):
       if self.path:
           try:
            #    self.app = xw.App(visible=False)
               self.workbook = xw.App(visible=False).books.open(self.path, update_links=False, read_only=True, ignore_read_only_recommended=True)
           except Exception as e:
               print(f"엑셀 파일 열기 중 오류 발생: {str(e)}")
  

    def toggle_sheet(self, sheet_name, button):
        if button.isChecked():
            self.selected_sheets.append(sheet_name)
        else:
            self.selected_sheets.remove(sheet_name)
        
        # 시트 버튼의 상태에 따라 범위 설정 버튼 활성화/비활성화
        for sheet_button, range_button in zip(self.sheet_buttons, self.range_buttons):
            if sheet_button == button:
                range_button.setEnabled(button.isChecked())
    
    def set_range(self, sheet_name, button):
        range_dialog = QDialog(self)
        range_dialog.setWindowTitle(f"{sheet_name} 범위 설정")

        layout = QVBoxLayout()

        range_label = QLabel("비교 병합할 범위를 입력하세요 (예: A1:B10, C1:D10):")
        layout.addWidget(range_label)

        range_input = QLineEdit()
        layout.addWidget(range_input)

        def update_range():
            # print(self.parent.monitor_thread.user_cordinations)
            # print(self.parent.monitor_thread.copy_selected_range)
            try:
                parent_sheet_name, start_column, start_row, end_column, end_row = self.parent.monitor_thread.user_cordinations
            except:
                pass
            try:
                if self.parent.monitor_thread and self.parent.monitor_thread.copy_selected_range:
                    if parent_sheet_name == sheet_name:
                        range_address = self.parent.monitor_thread.copy_selected_range
                        range_input.setText(range_address)
                range_input.setFocus()  # 입력란에 포커스 설정
            except Exception as e:
                print(f"범위 업데이트 중 오류 발생: {str(e)}")
                return

        def accept_range():
            range_value = range_input.text()
            self.sheet_ranges[sheet_name] = range_value if range_value else "All"
            if range_value:
                button.setText(range_value)  # 설정된 범위로 버튼 텍스트 변경
            else:
                button.setText("범위 설정")
            range_dialog.accept()

        ok_button = QPushButton("확인")
        ok_button.clicked.connect(accept_range)
        layout.addWidget(ok_button)

        cancel_button = QPushButton("취소")
        button.setText("범위 설정")
        cancel_button.clicked.connect(range_dialog.reject)
        layout.addWidget(cancel_button)

        range_dialog.setLayout(layout)

        # 타이머를 사용하여 주기적으로 범위 업데이트
        # if self.parent.monitor_thread:
        if self.parent.activate == True:
            timer = QTimer(range_dialog)
            timer.timeout.connect(update_range)
            timer.start(400)  # 400ms(0.4초) 간격으로 업데이트

        range_dialog.exec_()

    def is_valid_filename(self, filename):
        # 파일 이름 유효성 검사
        if filename.startswith(" "):
            return False
        if filename.strip() == "":
            return False
        pattern = r'^[a-zA-Z0-9가-힣_\-.()\s]+$'
        return re.match(pattern, filename) is not None

    def execute_compare_merge(self):
        cross_status = None
        selected_sheets = [button.text() for button in self.sheet_buttons if button.isChecked()]
        if not selected_sheets:
            QMessageBox.warning(self, "경고", "비교할 시트를 선택하세요.")
            return
        try:
            if self.parent.save_path and self.parent.save_path != "":
                output_path = self.parent.save_path
            else:
                while True:
                    output_path, _ = QFileDialog.getSaveFileName(self, "결과 파일 저장", "", "Excel Files (*.xlsx *.xls)")
                    if not output_path:
                        return 
                    if not output_path.endswith(".xlsx") and not output_path.endswith(".xls"):
                        QMessageBox.critical(self, "오류", "파일 확장자는 .xlsx로 설정해주세요.")
                        return
                    if os.path.exists(output_path):
                        result = QMessageBox.question(self, "확인", "이미 존재하는 파일입니다. 덮어쓰시겠습니까?", QMessageBox.Yes | QMessageBox.No)
                        if result == QMessageBox.Yes:
                            break
                    base_name = os.path.basename(output_path)
                    if not self.is_valid_filename(base_name):
                        QMessageBox.critical(self, "오류", "파일 이름이 유효하지 않습니다.")
                        continue
                    if output_path:
                        break
                    else:
                        QMessageBox.warning(self, "경고", "결과 파일을 저장할 경로를 선택하세요.")
                        return
            wb1 = self.parent.workbook
            wb2 = self.workbook
        except Exception as e:
            QMessageBox.warning(self, "오류", f"엑셀 파일 열기 중 오류 발생: {str(e)}")
            return

        print("선택한 시트:")
        merged_sheets = []  # 병합된 시트 이름을 저장할 리스트
        error_messages = []  # 오류 메시지를 저장할 리스트s
        if not self.workbook:
            print("비교 대상 워크북이 로드되지 않았습니다.")
            error_messages.append("비교 대상 워크북이 로드되지 않았습니다.")
        for sheet_name in self.selected_sheets:
            range_value = self.sheet_ranges.get(sheet_name, "All")
            if range_value == "All":
                # 전체 시트 범위 사용
                if self.parent.workbook:
                    sheet_range = self.parent.workbook.sheets[sheet_name].used_range
                    data = sheet_range.value
                    print(f"{sheet_name}: {range_value}")
                    print(f"선택한 범위: {sheet_range.address}")
                    # 비교 대상 워크북에서 동일한 시트 찾기
                    compare_sheet = self.find_compare_sheet(sheet_name)
                    if compare_sheet:
                        compare_data = compare_sheet.used_range.value
                        status, cross_status = self.compare_and_merge(data, compare_data, sheet_range, sheet_name, wb1, wb2, "All", error_messages)
                        if cross_status:
                            error_messages.append(f"{sheet_name}에서 병합 오류 발생: {cross_status}\n")
                        if status:
                            merged_sheets.append(sheet_name)
                        else:
                            error_messages.append(f"{sheet_name}'에서 병합 오류 발생")
                    else:
                        print(f"비교 대상 워크북에 '{sheet_name}' 시트가 존재하지 않습니다.")
                        error_messages.append(f"비교 대상 워크북에 시트 '{sheet_name}' 존재하지 않음")
                else:
                    print("원본 워크북이 선택되지 않았습니다.")
                    error_messages.append("원본 워크북이 선택되지 않았습니다.")
            else:
                # 사용자가 선택한 범위 사용
                if self.parent.workbook:
                    sheet = self.parent.workbook.sheets[sheet_name]
                    range_list = range_value.split(',')
                    # 비교 대상 워크북에서 동일한 시트 찾기
                    compare_sheet = self.find_compare_sheet(sheet_name)
                    if compare_sheet:
                        merge_success = True
                        for range_item in range_list:
                            sheet_range = sheet.range(range_item.strip())
                            data = sheet_range.value
                            print(data)
                            print("=" * 50)
                            print(f"{sheet_name}: {range_item}")
                            print(f"선택한 범위: {sheet_range.address}")
                            compare_range = compare_sheet.range(range_item.strip())
                            compare_data = compare_range.value
                            print(compare_data)
                            status, cross_status = self.compare_and_merge(data, compare_data, sheet_range, sheet_name, wb1, wb2, range_item.strip(), error_messages)
                            if not status:
                                merge_success = False
                                error_messages.append(f"시트 '{sheet_name}' 범위 '{range_item}'에서 병합 오류 발생")
                        if merge_success:
                            merged_sheets.append(sheet_name)
                    else:
                        print(f"비교 대상 워크북에 '{sheet_name}' 시트가 존재하지 않습니다.")
                        error_messages.append(f"비교 대상 워크북에 시트 '{sheet_name}' 존재하지 않음")
                else:
                    print("원본 워크북이 선택되지 않았습니다.")
                    error_messages.append("원본 워크북이 선택되지 않았습니다.")

        # 비교 병합 작업 완료 후 결과 메시지 출력
        try:

            if cross_status:
                QMessageBox.warning(self, "경고", f"병합 작업 중 오류 발생: {cross_status}")
            result_message = "비교 병합 작업 결과:\n\n"
            if merged_sheets:
                result_message += "성공적으로 병합된 시트:\n" + "\n".join(merged_sheets) + "\n\n"
            if error_messages:
                result_message += "병합 중 오류가 발생한 시트:\n" + "\n".join(error_messages)
            else:
                result_message += "모든 시트가 성공적으로 병합되었습니다."
            QMessageBox.information(self, "비교 병합 결과", result_message)
            wb1.save(output_path)
        except Exception as e:
            QMessageBox.critical(self, "오류", f"결과 파일 저장 중 오류 발생: {str(e)}")
            error_messages.append("결과 파일 저장 중 오류 발생")
            
            error_message = "비교 병합 작업 중 오류가 발생했습니다.\n\n오류 내용:\n" + "\n".join(error_messages)
            QMessageBox.warning(self, "비교 병합 실패", error_message)
        finally:
            if self.parent.auto_open:
                try:
                    xw.Book(output_path)
                except:
                    QMessageBox.warning(self, "경고", "결과 파일을 열 수 없습니다.")
            self.close()

    def find_compare_sheet(self, sheet_name):
        # 비교 대상 워크북에서 동일한 시트 이름 찾기
        sheet_name_without_number = re.sub(r'^\d+\.\s*', '', sheet_name)
        compare_workbook = self.workbook
        if compare_workbook:
            for sheet_name_with_number in compare_workbook.sheets:
                compare_sheet_name_without_number = re.sub(r'^\d+\.\s*', '', sheet_name_with_number.name)
                if compare_sheet_name_without_number == sheet_name_without_number:
                    return sheet_name_with_number
        return None

    def compare_and_merge(self, data, compare_data, sheet_range, sheet_name, wb1, wb2, _range_, error_messages):
        # 병합 수행
        if _range_ != "All":
            range__ = sheet_range.address
        else:
            range__ = None
        print("range__", range__)
        status, cross_status = merge_sheets(wb1, wb2, sheet_name, range__, _range_, self.previous_value, self.current_value)
        print(f"CROSS STATUS: {cross_status}")
        return status, cross_status

    def __del__(self, event: QCloseEvent):
        if self.workbook:
            self.workbook.close()
            self.app.quit()
            self.workbook = None
            self.app = None
        event.accept()
        
