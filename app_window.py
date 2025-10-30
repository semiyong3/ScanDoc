import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QFrame, QMessageBox, QFileDialog, QHBoxLayout
)
from PyQt5.QtCore import Qt
import core_functions


class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('ScanDoc (PPT Auto Mode)')
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

        # RUN 버튼: 가운데 정렬, 크기 축소
        self.btn_scan_run = QPushButton("RUN")
        self.btn_scan_run.setStyleSheet(run_button_style)
        run_layout1 = QHBoxLayout()
        run_layout1.addStretch()
        run_layout1.addWidget(self.btn_scan_run)
        run_layout1.addStretch()

        scan_layout.addWidget(QLabel("- Target Dir :"), 0, 0)
        scan_layout.addWidget(self.line_scan_target, 0, 1)
        scan_layout.addWidget(self.btn_scan_find, 0, 2)
        scan_layout.addWidget(QLabel("- Output File :"), 1, 0)
        scan_layout.addWidget(self.line_scan_output, 1, 1)
        scan_layout.addWidget(self.btn_scan_set, 1, 2)
        scan_layout.addLayout(run_layout1, 2, 0, 1, 3)
        main_layout.addLayout(scan_layout)

        line1 = QFrame()
        line1.setFrameShape(QFrame.HLine)
        main_layout.addWidget(line1)

        # --- 2. Convert To Image ---
        main_layout.addWidget(QLabel("<b>2. Convert To Image</b>"))
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

        img_layout.addWidget(QLabel("- Target File :"), 0, 0)
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

        # --- 3. Convert To PDF ---
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
        self.btn_scan_set.clicked.connect(self.set_scan_output)
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

    def set_scan_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "Set Output Excel File", filter="Excel Files (*.xlsx)")
        if file:
            self.line_scan_output.setText(file)

    def find_img_target(self):
        filters = "PowerPoint Files (*.pptx *.ppt);;All Files (*)"
        file, _ = QFileDialog.getOpenFileName(self, "Select Target File", filter=filters)
        if file:
            self.line_img_target.setText(file)

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
        output_file = self.line_scan_output.text()
        if not target_dir or not output_file:
            self.show_error("Target Directory와 Output File을 모두 지정해야 합니다.")
            return
        try:
            msg = core_functions.scan_directory(target_dir, output_file)
            QMessageBox.information(self, "완료", msg)
        except Exception as e:
            self.show_error(str(e))

    def run_convert_to_image(self):
        target_file = self.line_img_target.text()
        output_dir = self.line_img_output.text()
        if not target_file or not output_dir:
            self.show_error("Target File과 Output Dir를 모두 지정해야 합니다.")
            return
        base_filename = os.path.splitext(os.path.basename(target_file))[0]
        ext = os.path.splitext(target_file)[1].lower()

        if ext not in [".ppt", ".pptx"]:
            self.show_error(f"현재는 PPT/PPTX 형식만 지원합니다. (입력 파일: {ext})")
            return

        try:
            msg = core_functions.capture_ppt_slides(target_file, output_dir, base_filename)
            QMessageBox.information(self, "완료", msg)
        except Exception as e:
            self.show_error(str(e))

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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = AppWindow()
    win.show()
    sys.exit(app.exec_())
