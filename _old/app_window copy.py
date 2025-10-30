import sys
import os
import tempfile
import shutil
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout, 
    QLabel, QLineEdit, QPushButton, QFrame, QMessageBox, QFileDialog,
    QDialog, QHBoxLayout 
)
from PyQt5.QtCore import QThread, QObject, pyqtSignal, pyqtSlot, Qt
from pynput import keyboard 

import core_functions

# --- 백그라운드 스레드 워커 클래스 ---
class Worker(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(str)
    success = pyqtSignal(str)
    def __init__(self, function, *args, **kwargs):
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs
    @pyqtSlot()
    def run(self):
        try:
            result_msg = self.function(*self.args, **self.kwargs)
            self.success.emit(result_msg)
        except Exception as e:
            self.error.emit(f"오류 발생: {e}")
        finally:
            self.finished.emit()

# --- 캡처 세션 보조창 (변경 없음 - PPT 외 파일용) ---
class CaptureDialog(QDialog):
    session_finished = pyqtSignal(str, str, str)
    def __init__(self, output_dir, base_filename, parent=None):
        super().__init__(parent)
        self.output_dir = output_dir
        self.base_filename = base_filename
        self.parent = parent 
        self.page_count = 0
        self.temp_dir = tempfile.mkdtemp()
        self.listener = None
        self.init_ui()
        self.start_hotkey_listener()
    def init_ui(self):
        self.setWindowTitle("Capture Session")
        self.setWindowFlags(Qt.WindowStaysOnTopHint) 
        self.setFixedSize(300, 100)
        layout = QVBoxLayout()
        self.info_label = QLabel(f"Page {self.page_count} captured.\nPress [F9] to capture active window.")
        self.info_label.setAlignment(Qt.AlignCenter)
        btn_layout = QHBoxLayout()
        self.btn_capture = QPushButton("Capture (F9)")
        self.btn_finish = QPushButton("Finish & Zip")
        btn_layout.addWidget(self.btn_capture)
        btn_layout.addWidget(self.btn_finish)
        layout.addWidget(self.info_label)
        layout.addLayout(btn_layout)
        self.setLayout(layout)
        self.btn_capture.clicked.connect(self.do_capture)
        self.btn_finish.clicked.connect(self.finish_session)
    def start_hotkey_listener(self):
        self.listener = keyboard.Listener(on_press=self.on_key_press)
        self.listener.start()
    def on_key_press(self, key):
        if key == keyboard.Key.f9:
            self.do_capture()
    def do_capture(self):
        try:
            self.page_count += 1
            filename = f"{self.page_count}.png"
            core_functions.capture_active_window_to_clipboard(self.temp_dir, filename)
            self.info_label.setText(f"Page {self.page_count} captured.\nPress [F9] to capture active window.")
        except Exception as e:
            if self.parent:
                self.parent.show_error(f"캡처 실패: {e}\n클립보드에 이미지가 없거나, 문서 창이 활성화되지 않았습니다.")
            self.page_count -= 1 
    def finish_session(self):
        if self.page_count == 0:
            if self.parent:
                self.parent.show_error("최소 1페이지 이상 캡처해야 합니다.")
            return
        if self.listener:
            self.listener.stop()
        self.session_finished.emit(self.temp_dir, self.output_dir, self.base_filename)
        self.accept() 
    def reject(self):
        if self.listener:
            self.listener.stop()
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
        super().reject()

# --- 메인 윈도우 ---
class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('ScanDoc')
        self.thread = None
        self.worker = None
        self.capture_dialog = None
        self.init_ui()
        self.connect_signals()

    def init_ui(self):
        self.setFixedWidth(500)
        self.move(100, 100)
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        
        run_button_style = """
            QPushButton {
                background-color: dimgray; color: white; border: 1px solid #505050; padding: 4px 12px;
            }
            QPushButton:hover { background-color: #7A7A7A; }
            QPushButton:pressed { background-color: #5A5A5A; }
        """
        # --- 1. Scan Directory (변경 없음) ---
        main_layout.addWidget(QLabel("<b>1. Scan Directory</b>"))
        scan_layout = QGridLayout()
        self.line_scan_target = QLineEdit()
        self.btn_scan_find = QPushButton("FIND")
        self.line_scan_output = QLineEdit()
        self.btn_scan_set = QPushButton("SET")
        self.btn_scan_run = QPushButton("RUN")
        self.btn_scan_run.setStyleSheet(run_button_style) 
        scan_layout.addWidget(QLabel("- Target Dir :"), 0, 0)
        scan_layout.addWidget(self.line_scan_target, 0, 1)
        scan_layout.addWidget(self.btn_scan_find, 0, 2)
        scan_layout.addWidget(QLabel("- Output File :"), 1, 0)
        scan_layout.addWidget(self.line_scan_output, 1, 1)
        scan_layout.addWidget(self.btn_scan_set, 1, 2)
        run_layout_1 = QHBoxLayout()
        run_layout_1.addStretch() 
        run_layout_1.addWidget(self.btn_scan_run)
        run_layout_1.addStretch() 
        scan_layout.addLayout(run_layout_1, 2, 1, 1, 2) 
        main_layout.addLayout(scan_layout)
        main_layout.addSpacing(15) 
        line1 = QFrame()
        line1.setFrameShape(QFrame.HLine); line1.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line1)
        main_layout.addSpacing(15) 

        # --- 2. Convert To Image (변경 없음) ---
        main_layout.addWidget(QLabel("<b>2. Convert To Image (xlsx, pptx, docx, pdf, hwp, txt)</b>"))
        img_layout = QGridLayout()
        self.line_img_target = QLineEdit()
        self.btn_img_find = QPushButton("FIND")
        self.line_img_output = QLineEdit()
        self.btn_img_set = QPushButton("SET")
        self.btn_img_run = QPushButton("RUN")
        self.btn_img_run.setStyleSheet(run_button_style)
        img_layout.addWidget(QLabel("- Target File :"), 0, 0)
        img_layout.addWidget(self.line_img_target, 0, 1)
        img_layout.addWidget(self.btn_img_find, 0, 2)
        img_layout.addWidget(QLabel("- Output Dir :"), 1, 0)
        img_layout.addWidget(self.line_img_output, 1, 1)
        img_layout.addWidget(self.btn_img_set, 1, 2)
        run_layout_2 = QHBoxLayout()
        run_layout_2.addStretch()
        run_layout_2.addWidget(self.btn_img_run)
        run_layout_2.addStretch()
        img_layout.addLayout(run_layout_2, 2, 1, 1, 2)
        main_layout.addLayout(img_layout)
        main_layout.addSpacing(15) 
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine); line2.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line2)
        main_layout.addSpacing(15) 

        # --- 3. Convert To PDF (변경 없음) ---
        main_layout.addWidget(QLabel("<b>3. Convert To PDF</b>"))
        pdf_layout = QGridLayout()
        self.line_pdf_target = QLineEdit()
        self.btn_pdf_find = QPushButton("FIND")
        self.line_pdf_output = QLineEdit()
        self.btn_pdf_set = QPushButton("SET")
        self.btn_pdf_run = QPushButton("RUN")
        self.btn_pdf_run.setStyleSheet(run_button_style) 
        pdf_layout.addWidget(QLabel("- Target File :"), 0, 0)
        pdf_layout.addWidget(self.line_pdf_target, 0, 1)
        pdf_layout.addWidget(self.btn_pdf_find, 0, 2)
        pdf_layout.addWidget(QLabel("- Output Dir :"), 1, 0)
        pdf_layout.addWidget(self.line_pdf_output, 1, 1)
        pdf_layout.addWidget(self.btn_pdf_set, 1, 2)
        run_layout_3 = QHBoxLayout()
        run_layout_3.addStretch()
        run_layout_3.addWidget(self.btn_pdf_run)
        run_layout_3.addStretch()
        pdf_layout.addLayout(run_layout_3, 2, 1, 1, 2)
        main_layout.addLayout(pdf_layout)
        main_layout.addSpacing(20) 
        
        self.setCentralWidget(main_widget)
        self.adjustSize() 

    # --- (시그널 연결부 변경 없음) ---
    def connect_signals(self):
        self.btn_scan_find.clicked.connect(self.find_scan_dir)
        self.btn_scan_set.clicked.connect(self.set_scan_output)
        self.btn_img_find.clicked.connect(self.find_img_target) 
        self.btn_img_set.clicked.connect(self.set_img_output)
        self.btn_pdf_find.clicked.connect(self.find_pdf_target)
        self.btn_pdf_set.clicked.connect(self.set_pdf_output)

        self.btn_scan_run.clicked.connect(self.run_scan_directory)
        self.btn_img_run.clicked.connect(self.run_convert_to_image) 
        self.btn_pdf_run.clicked.connect(self.run_convert_to_pdf)

    # --- (파일/폴더 찾기 슬롯 변경 없음) ---
    def find_scan_dir(self):
        dir = QFileDialog.getExistingDirectory(self, "Select Target Directory")
        if dir: self.line_scan_target.setText(dir)
    def set_scan_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "Set Output Excel File", filter="Excel Files (*.xlsx)")
        if file: self.line_scan_output.setText(file)
    def find_img_target(self):
        # [수정] PPT 자동화를 위해 필터 수정
        filters = "Supported Files (*.pptx *.ppt *.docx *.pdf *.hwp *.txt);;All Files (*)"
        file, _ = QFileDialog.getOpenFileName(self, "Select Target File", filter=filters)
        if file: self.line_img_target.setText(file)
    def set_img_output(self):
        dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if dir: self.line_img_output.setText(dir)
    def find_pdf_target(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Target ZIP File", filter="ZIP Files (*.zip)")
        if file: self.line_pdf_target.setText(file)
    def set_pdf_output(self):
        dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if dir: self.line_pdf_output.setText(dir)

    # --- (RUN 로직 슬롯 변경됨) ---
    
    def run_scan_directory(self):
        # (변경 없음)
        target_dir = self.line_scan_target.text()
        output_file = self.line_scan_output.text()
        if not target_dir or not output_file:
            self.show_error("Target Directory와 Output File을 모두 지정해야 합니다.")
            return
        self.run_threaded_task(core_functions.scan_directory, target_dir, output_file)

    def run_convert_to_image(self):
        # [수정됨] 파일 확장자에 따라 분기
        target_file = self.line_img_target.text()
        output_dir = self.line_img_output.text()
        
        if not target_file or not output_dir:
            self.show_error("Target File과 Output Dir를 모두 지정해야 합니다.")
            return
            
        base_filename = os.path.splitext(os.path.basename(target_file))[0]
        file_ext = os.path.splitext(target_file)[1].lower()

        if file_ext in ['.pptx', '.ppt']:
            # PPT 파일: 자동화 스레드 실행
            QMessageBox.information(self, "자동 캡처 시작", 
                "PPT 자동화 캡처를 시작합니다.\n\n"
                "캡처가 완료될 때까지 PC 사용을 멈추고 기다려주세요.\n"
                "(파워포인트가 자동으로 실행되고 캡처됩니다.)")
            self.run_threaded_task(
                #core_functions.automate_ppt_capture,
                core_functions.capture_ppt_slides,
                target_file, output_dir, base_filename
            )
        else:
            # 그 외 파일: 기존 수동 캡처 다이얼로그 실행
            reply = QMessageBox.information(self, "수동 캡처 시작", 
                "DRM 캡처 모드를 시작합니다.\n\n"
                "1. 변환할 문서를 **수동으로** 여세요.\n"
                "2. 캡처할 창을 **활성화(클릭)**하세요.\n\n"
                "[확인]을 누르면 캡처 보조창이 뜹니다.",
                QMessageBox.Ok | QMessageBox.Cancel)
                
            if reply == QMessageBox.Ok:
                self.capture_dialog = CaptureDialog(output_dir, base_filename, self)
                self.capture_dialog.session_finished.connect(self.run_zip_task)
                self.capture_dialog.show()

    @pyqtSlot(str, str, str)
    def run_zip_task(self, temp_dir, output_dir, base_filename):
        # (변경 없음 - 수동 캡처의 Zip을 담당)
        self.run_threaded_task(
            core_functions.zip_images_and_cleanup, 
            temp_dir, output_dir, base_filename
        )

    def run_convert_to_pdf(self):
        # (변경 없음)
        target_zip = self.line_pdf_target.text()
        output_dir = self.line_pdf_output.text()
        if not target_zip or not output_dir:
            self.show_error("Target File (ZIP)과 Output Dir를 모두 지정해야 합니다.")
            return
        self.run_threaded_task(core_functions.convert_to_pdf, target_zip, output_dir)

    # --- (스레드 실행 및 메시지 박스 변경 없음) ---
    def run_threaded_task(self, function, *args):
        self.set_buttons_enabled(False)
        self.thread = QThread()
        self.worker = Worker(function, *args)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.success.connect(self.show_success)
        self.worker.error.connect(self.show_error)
        self.worker.finished.connect(lambda: self.set_buttons_enabled(True)) 
        self.thread.start()

    def set_buttons_enabled(self, enabled):
        self.btn_scan_run.setEnabled(enabled)
        self.btn_img_run.setEnabled(enabled)
        self.btn_pdf_run.setEnabled(enabled)

    def show_error(self, message):
        #QMessageBox.critical(self, "Error", message)
        QMessageBox.critical(self, "Error", "작업 실패!")

    def show_success(self, message):
        #QMessageBox.information(self, "Success", message)
        QMessageBox.information(self, "Success", "작업 성공!")
        
    def closeEvent(self, event):
        if self.capture_dialog:
            self.capture_dialog.reject()
        event.accept()