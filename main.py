import sys
import os
import configparser
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMessageBox, QInputDialog, QLineEdit
from app_window import AppWindow

def show_startup_error(message):
    """메인 윈도우 생성 전, 오류 메시지를 표시하고 종료합니다."""
    # (QApplication이 이미 생성되었다고 가정)
    msg_box = QMessageBox(QMessageBox.Critical, "Access Denied", message)
    msg_box.exec_()
    sys.exit(1) # 오류로 종료

if __name__ == "__main__":
    
    # 1. PyQt5 앱은 QDialog를 띄우기 위해 Application 객체가 먼저 필요합니다.
    app = QApplication(sys.argv)
    
    config = configparser.ConfigParser()
    LICENSE_FILE = 'license.ini' # main.py와 동일 폴더

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
            show_startup_error("암호가 일치하지 않습니다.")

        # 5. 만료일 확인 (ExpireDate가 "None"이 아닐 경우에만)
        if expire_date_str.strip().lower() != 'none':
            try:
                expire_date = datetime.strptime(expire_date_str, '%Y-%m-%d').date()
                today = datetime.now().date()
                
                if today > expire_date:
                    show_startup_error(f"라이선스 에러!!!")
                    
            except ValueError:
                # 날짜 형식이 잘못된 경우
                show_startup_error(f"license.ini 파일의 날짜 형식이 잘못되었습니다.\n"
                                 "ExpireDate = YYYY-MM-DD 또는 None 으로 설정하세요.")
    
    except (FileNotFoundError, configparser.Error, KeyError) as e:
        # 파일이 없거나, [Security] 섹션이 없거나, 키가 없는 경우
        show_startup_error(f"라이선스 설정 오류:\n{e}")
    
    except Exception as e:
        # 기타 예기치 못한 오류 (예: PyQt 모듈 누락 등)
        show_startup_error(f"알 수 없는 오류 발생: {e}")

    # --- 라이선스 체크 통과 ---
    # (이제 메인 윈도우를 띄웁니다)
    win = AppWindow()
    win.show()
    sys.exit(app.exec_())