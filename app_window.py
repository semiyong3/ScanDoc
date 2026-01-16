import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QFrame, QMessageBox, QFileDialog, QHBoxLayout,
    QInputDialog
)
from PyQt5.QtCore import Qt
import core_functions
# import configparser 
# from datetime import datetime

class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('ScanDoc')
        self.init_ui()
        self.connect_signals()

    def init_ui(self):
        self.setFixedWidth(500)
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)

        # 공통 RUN 버튼 스타일
        run_button_style = """
            QPushButton {
                background-color: dimgray;
                color: white;
                border: 1px solid #505050;
                padding: 2px 8px;
                min-width: 80px;
                min-height: 26px;
            }
            QPushButton:hover { background-color: #7A7A7A; }
            QPushButton:pressed { background-color: #5A5A5A; }
        """
        # --- 1. Scan Directory ---
        main_layout.addWidget(QLabel("<b>1. Scan Directory</b>"))
        scan_layout = QGridLayout()
        self.line_scan_target = QLineEdit()
        self.btn_scan_find = QPushButton("FIND")
        self.line_scan_output = QLineEdit()
        self.btn_scan_set = QPushButton("SET")

        self.btn_scan_run = QPushButton("RUN")
        self.btn_scan_run.setStyleSheet(run_button_style)
        run_layout1 = QHBoxLayout()
        run_layout1.addStretch()
        run_layout1.addWidget(self.btn_scan_run)
        run_layout1.addStretch()

        scan_layout.addWidget(QLabel("- Target Dir :"), 0, 0)
        scan_layout.addWidget(self.line_scan_target, 0, 1)
        scan_layout.addWidget(self.btn_scan_find, 0, 2)
        scan_layout.addWidget(QLabel("- Output Dir :"), 1, 0)
        scan_layout.addWidget(self.line_scan_output, 1, 1)
        scan_layout.addWidget(self.btn_scan_set, 1, 2)
        scan_layout.addLayout(run_layout1, 2, 0, 1, 3)
        main_layout.addLayout(scan_layout)

        line1 = QFrame()
        line1.setFrameShape(QFrame.HLine)
        main_layout.addWidget(line1)

        # --- 2. Convert To Image  ---
        title_layout2 = QHBoxLayout()
        title_layout2.addWidget(QLabel("<b>2. Convert (ppt/xls/doc/pdf) To Image (Batch)  </b>"))
        title_layout2.addStretch()
        title_layout2.addWidget(QLabel("Interval (s):"))
        
        # 기본값을 "1.0"으로 변경
        self.line_img_interval = QLineEdit("1.0") 
        self.line_img_interval.setFixedWidth(40)  
        title_layout2.addWidget(self.line_img_interval)
        main_layout.addLayout(title_layout2) 

        img_layout = QGridLayout()
        self.line_img_target = QLineEdit()
        self.btn_img_find = QPushButton("FIND")
        self.line_img_output = QLineEdit()
        self.btn_img_set = QPushButton("SET")

        self.btn_img_run = QPushButton("RUN")
        self.btn_img_run.setStyleSheet(run_button_style)
        run_layout2 = QHBoxLayout()
        run_layout2.addStretch()
        run_layout2.addWidget(self.btn_img_run)
        run_layout2.addStretch()

        img_layout.addWidget(QLabel("- Target Dir :"), 0, 0)
        img_layout.addWidget(self.line_img_target, 0, 1)
        img_layout.addWidget(self.btn_img_find, 0, 2)
        img_layout.addWidget(QLabel("- Output Dir :"), 1, 0)
        img_layout.addWidget(self.line_img_output, 1, 1)
        img_layout.addWidget(self.btn_img_set, 1, 2)
        img_layout.addLayout(run_layout2, 2, 0, 1, 3)
        main_layout.addLayout(img_layout)

        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        main_layout.addWidget(line2)

        # --- 3. Convert To PDF (변경 없음) ---
        main_layout.addWidget(QLabel("<b>3. Convert To PDF</b>"))
        pdf_layout = QGridLayout()
        self.line_pdf_target = QLineEdit()
        self.btn_pdf_find = QPushButton("FIND")
        self.line_pdf_output = QLineEdit()
        self.btn_pdf_set = QPushButton("SET")

        self.btn_pdf_run = QPushButton("RUN")
        self.btn_pdf_run.setStyleSheet(run_button_style)
        run_layout3 = QHBoxLayout()
        run_layout3.addStretch()
        run_layout3.addWidget(self.btn_pdf_run)
        run_layout3.addStretch()

        pdf_layout.addWidget(QLabel("- Target Dir :"), 0, 0)
        pdf_layout.addWidget(self.line_pdf_target, 0, 1)
        pdf_layout.addWidget(self.btn_pdf_find, 0, 2)
        pdf_layout.addWidget(QLabel("- Output File :"), 1, 0)
        pdf_layout.addWidget(self.line_pdf_output, 1, 1)
        pdf_layout.addWidget(self.btn_pdf_set, 1, 2)
        pdf_layout.addLayout(run_layout3, 2, 0, 1, 3)
        main_layout.addLayout(pdf_layout)

        self.setCentralWidget(main_widget)

    def connect_signals(self):
        self.btn_scan_find.clicked.connect(self.find_scan_dir)
        self.btn_scan_set.clicked.connect(self.set_scan_output_dir) 
        self.btn_img_find.clicked.connect(self.find_img_target)
        self.btn_img_set.clicked.connect(self.set_img_output)
        self.btn_pdf_find.clicked.connect(self.find_pdf_target)
        self.btn_pdf_set.clicked.connect(self.set_pdf_output)
        self.btn_scan_run.clicked.connect(self.run_scan_directory)
        self.btn_img_run.clicked.connect(self.run_convert_to_image)
        self.btn_pdf_run.clicked.connect(self.run_convert_to_pdf)

    # --- 파일/디렉토리 선택 ---
    def find_scan_dir(self):
        dir = QFileDialog.getExistingDirectory(self, "Select Target Directory")
        if dir:
            self.line_scan_target.setText(dir)

    def set_scan_output_dir(self): 
        """Output Dir를 선택하는 다이얼로그를 엽니다."""
        dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if dir:
            self.line_scan_output.setText(dir)

    def find_img_target(self):
        # filters = "ppt, xls, doc, pdf (*.ppt *.pptx *.xls *.xlsx *.doc *.docx *.pdf)"
        # file, _ = QFileDialog.getOpenFileName(self, "Select Target File", filter=filters)
        # if file:
        #     self.line_img_target.setText(file)
        dir = QFileDialog.getExistingDirectory(self, "Select Target Directory (containing docs)")
        if dir:
            self.line_img_target.setText(dir)

    def set_img_output(self):
        dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if dir:
            self.line_img_output.setText(dir)

    def find_pdf_target(self):
        dir = QFileDialog.getExistingDirectory(self, "Select Target Directory")
        if dir:
            self.line_pdf_target.setText(dir)

    def set_pdf_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "Set Output PDF File", filter="PDF (*.pdf)")
        if file:
            self.line_pdf_output.setText(file)

    # --- 실행 기능들 ---
    def run_scan_directory(self):
        target_dir = self.line_scan_target.text()
        output_dir = self.line_scan_output.text()
        if not target_dir or not output_dir:
            self.show_error("Target, Output Directory 를 모두 지정해야 합니다.")
            return
        try:
            msg = core_functions.scan_directory(target_dir, output_dir)
            QMessageBox.information(self, "완료", msg)
        except Exception as e:
            self.show_error(str(e))

    def run_convert_to_image(self):
        target_dir = self.line_img_target.text()
        output_dir = self.line_img_output.text()
        
        if not target_dir or not output_dir:
            self.show_error("Target Dir과 Output Dir를 모두 지정해야 합니다.")
            return
        
        # Interval 값 체크
        try:
            interval_sec = float(self.line_img_interval.text())
            if interval_sec < 1.0:
                interval_sec = 1.0 
                self.line_img_interval.setText("1.0")
        except ValueError:
            self.show_error("Interval 값은 숫자(예: 1.0)여야 합니다.")
            return

        # 일괄 작업 경고 메시지
        reply = QMessageBox.warning(self, "일괄 변환 시작",
            "지정된 폴더 내의 **모든 문서 파일**을 변환합니다.\n\n"
            "**[주의사항]**\n"
            "1. 작업 중 **마우스/키보드 사용 금지** (창 포커스 유지 필요)\n"
            "2. Office/PDF 프로그램이 **주 모니터**에서 실행되어야 함\n"
            "3. 파일 개수에 따라 시간이 오래 걸릴 수 있음\n\n"
            "진행하시겠습니까?",
            QMessageBox.Ok | QMessageBox.Cancel)
        
        if reply == QMessageBox.Cancel:
            return

        try:
            # [변경] 폴더 일괄 처리 함수 호출
            msg = core_functions.process_directory_for_images(target_dir, output_dir, interval_sec)
            
            # 결과 메시지 표시
            QMessageBox.information(self, "작업 결과", msg)
            
        except Exception as e:
            self.show_error(f"작업 실행 중 오류 발생: {str(e)}")

    def run_convert_to_pdf(self):
        target_dir = self.line_pdf_target.text()
        output_file = self.line_pdf_output.text()
        if not target_dir or not output_file:
            self.show_error("Target Dir 과 Output File 을 모두 지정해야 합니다.")
            return
        try:
            msg = core_functions.convert_to_pdf(target_dir, output_file)
            QMessageBox.information(self, "완료", msg)
        except Exception as e:
            self.show_error(str(e))

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

# [수정] 라이선스 관련 헬퍼 함수 및 if __name__ == "__main__" 블록 *제거*