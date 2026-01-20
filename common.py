import hashlib
import win32gui
from PIL import ImageGrab
import ctypes

def _get_file_hash(filepath):
    """파일의 MD5 해시를 반환합니다."""
    hasher = hashlib.md5()
    try:
        with open(filepath, 'rb') as f:
            buf = f.read()
            hasher.update(buf)
        return hasher.hexdigest()
    except Exception as e:
        print(f"[WARN] 파일 해시 읽기 실패: {e}")
        return None


def capture_active_window(hwnd=None):
    """
    현재 활성화된 창(Foreground Window)만 캡처하여 Pillow 이미지 객체로 반환
    """
 
    # 1. 활성 창의 핸들(HWND) 가져오기
    if (hwnd == 0) or (hwnd is None):
        raise Exception("활성화된 창을 찾을 수 없습니다.")
        
    # 2. 핸들을 사용하여 창의 외곽 좌표(bbox) 가져오기
    #    bbox는 (left, top, right, bottom) 튜플입니다.
    rect = win32gui.GetWindowRect(hwnd)
    bbox = (rect[0], rect[1], rect[2], rect[3])

    # 3. bbox 좌표를 ImageGrab.grab()에 전달하여 해당 영역만 캡처
    screenshot = ImageGrab.grab(bbox=bbox)
    
    return screenshot

def _clear_system_clipboard():
    """
    Office 프로그램 종료 시 '복사한 데이터를 유지하시겠습니까?' 
    팝업이나 백그라운드 대기 현상을 막기 위해 클립보드를 비웁니다.
    """
    try:
        if ctypes.windll.user32.OpenClipboard(None):
            ctypes.windll.user32.EmptyClipboard()
            ctypes.windll.user32.CloseClipboard()
    except Exception:
        pass