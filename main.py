import sys
import os
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMessageBox, QInputDialog, QLineEdit
from app_window import AppWindow

# 1. 만료일을 지정 (YYYY-MM-DD)
# 2. 만료일 체크를 안 하려면 이 값을 None 으로 설정 (예: EXPIRE_DATE = None)
EXPIRE_DATE = "2026-06-30" 
# ---

# --- 헬퍼 함수 (변경 없음) ---
def show_startup_error(message):
    """메인 윈도우 생성 전, 오류 메시지를 표시하고 종료합니다."""
    msg_box = QMessageBox(QMessageBox.Critical, "Access Denied", message)
    msg_box.exec_()
    sys.exit(1) # 오류로 종료

# --- 메인 실행 로직 ---
if __name__ == "__main__":
    
    # 1. QApplication 먼저 생성
    app = QApplication(sys.argv)
    
    try:
        # 1. 현재 날짜 기준으로 '정답' 암호 생성 (예: "si2511")
        today = datetime.now()
        correct_password = f"si{today.strftime('%y%m')}"

        # 2. 사용자에게 암호 입력 받기
        entered_pass, ok = QInputDialog.getText(None, 'Authentication', 
                                                'Enter Password:', QLineEdit.Password)
        
        if not ok:
            sys.exit(0) # 사용자가 'Cancel'을 누름
        
        if entered_pass != correct_password:
            show_startup_error("암호가 일치하지 않습니다.")

        # 3. 만료일 확인 (변수값이 None이 아닐 경우에만)
        if EXPIRE_DATE is not None:
            try:
                expire_date = datetime.strptime(EXPIRE_DATE, '%Y-%m-%d').date()
                
                if today.date() > expire_date:
                    show_startup_error(f"라이선스가 만료되었습니다. (만료일: {EXPIRE_DATE})")
                    
            except ValueError:
                # 날짜 형식이 잘못된 경우
                show_startup_error(f"코드 내 만료일(EXPIRE_DATE) 형식이 잘못되었습니다.\n"
                                 "YYYY-MM-DD 또는 None 으로 설정하세요.")
    
    except Exception as e:
        show_startup_error(f"라이선스 확인 중 알 수 없는 오류 발생: {e}")

    win = AppWindow()
    win.show()
    sys.exit(app.exec_())