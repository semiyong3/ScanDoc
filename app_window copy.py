import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QFrame, QMessageBox, QFileDialog, QHBoxLayout,
    QInputDialog 
)
from PyQt5.QtCore import Qt
import core_functions
import configparser  
from datetime import datetime


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

        # --- 2. Convert To Image ---
        main_layout.addWidget(QLabel("<b>2. Convert To Image : ppt/xls/doc/pdf available</b>"))
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
        self.btn_scan_set.clicked.connect(self.set_scan_output_dir) # [수정] 이름 변경
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

    # [수정] 함수 이름 변경 (set_scan_output -> set_scan_output_dir)
    def set_scan_output_dir(self):
        """Output Dir를 선택하는 다이얼로그를 엽니다."""
        dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if dir:
            self.line_scan_output.setText(dir)

    def find_img_target(self):
        filters = "ppt, xls, doc, pdf (*.ppt *.pptx *.xls *.xlsx *.doc *.docx *.pdf)"
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
        target_file = self.line_img_target.text()
        output_dir = self.line_img_output.text()
        if not target_file or not output_dir:
            self.show_error("Target File과 Output Dir를 모두 지정해야 합니다.")
            return
        
        base_filename = os.path.splitext(os.path.basename(target_file))[0]
        ext = os.path.splitext(target_file)[1].lower()

        # --- 확장자별 변환 함수 매핑 및 호출 ---
        
        # 지원 확장자와 함수 매핑
        conversion_map = {
            # PPT
            ".ppt": core_functions.capture_ppt_slides,
            ".pptx": core_functions.capture_ppt_slides,
            # Excel
            ".xls": core_functions.capture_excel_sheets,
            ".xlsx": core_functions.capture_excel_sheets,
            # Word
            ".doc": core_functions.capture_word_document,
            ".docx": core_functions.capture_word_document,
            # PDF
            ".pdf": core_functions.capture_pdf_document 
        }

        if ext not in conversion_map:
            supported_exts = ", ".join(conversion_map.keys())
            self.show_error(f"현재 지원하지 않는 파일 형식입니다. (지원 형식: {supported_exts})")
            return
        
        reply = QMessageBox.warning(self, "자동 캡처 시작 - 중요!",
            "자동 캡처를 시작합니다.\n\n"
            "**1. [포커스 유지]**\n"
            "캡처가 완료될 때까지 **키보드나 마우스를 절대 조작하지 마세요!**\n"
            "(다른 창을 클릭하면 캡처가 실패합니다.)\n\n"
            "**2. [듀얼 모니터]**\n"
            "실행되는 프로그램(PPT, Word, Excel, PDF)이\n"
            "반드시 **'주 모니터(Main Monitor)'**에서 실행되도록 준비해주세요.\n"
            "(보조 모니터에서 실행되면 실패할 수 있습니다.)\n\n"
            "준비되었으면 [OK]를 누르세요.",
            QMessageBox.Ok | QMessageBox.Cancel)
        
        if reply == QMessageBox.Cancel:
            QMessageBox.information(self, "취소", "작업이 취소되었습니다.")
            return
            
        converter_func = conversion_map[ext]

        try:
            msg = converter_func(target_file, output_dir, base_filename)
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

def show_startup_error(message):
    """메인 윈도우 생성 전, 오류 메시지를 표시하고 종료합니다."""
    # (QApplication이 이미 생성되었다고 가정)
    msg_box = QMessageBox(QMessageBox.Critical, "Access Denied", message)
    msg_box.exec_()
    sys.exit(1) # 오류로 종료


if __name__ == "__main__":
    # 라이선스 체크 로직을 win.show() 이전에 추가
    
    # 1. PyQt5 앱은 QDialog를 띄우기 위해 Application 객체가 먼저 필요합니다.
    app = QApplication(sys.argv)
    
    config = configparser.ConfigParser()
    LICENSE_FILE = 'license.ini'

    try:
        # 2. license.ini 파일 읽기
        if not config.read(LICENSE_FILE):
            raise FileNotFoundError(f"{LICENSE_FILE} 파일을 찾을 수 없습니다.\n"
                                     "프로그램과 같은 폴더에 license.ini 파일을 생성해주세요.")

        # 3. [Security] 섹션에서 값 읽기
        stored_password = config.get('Security', 'Password')
        expire_date_str = config.get('Security', 'ExpireDate')

        # 4. 비밀번호 확인
        entered_pass, ok = QInputDialog.getText(None, 'Authentication', 
                                                'Enter Password:', QLineEdit.Password)
        
        if not ok:
            sys.exit(0) # 사용자가 'Cancel'을 누름
        
        if entered_pass != stored_password:
            show_startup_error("암호가 일치하지 않습니다.") # [수정] 정의된 함수 호출

        # 5. 만료일 확인 (ExpireDate가 "None"이 아닐 경우에만)
        if expire_date_str.strip().lower() != 'none':
            try:
                expire_date = datetime.strptime(expire_date_str, '%Y-%m-%d').date()
                today = datetime.now().date()
                
                if today > expire_date:
                    show_startup_error(f"라이선스가 만료되었습니다. (만료일: {expire_date_str})") # [수정] 정의된 함수 호출
                    
            except ValueError:
                # 날짜 형식이 잘못된 경우
                show_startup_error(f"license.ini 파일의 날짜 형식이 잘못되었습니다.\n" # [수정] 정의된 함수 호출
                                 "ExpireDate = YYYY-MM-DD 또는 None 으로 설정하세요.")
    
    except (FileNotFoundError, configparser.Error, KeyError) as e:
        # 파일이 없거나, [Security] 섹션이 없거나, 키가 없는 경우
        show_startup_error(f"라이선스 설정 오류:\n{e}") # [수정] 정의된 함수 호출
    
    except Exception as e:
        # [수정] 기타 예기치 못한 오류 (예: PyQt 모듈 누락 등)
        show_startup_error(f"알 수 없는 오류 발생: {e}")

    # --- 라이선스 체크 통과 ---
    
    win = AppWindow()
    win.show()
    sys.exit(app.exec_())