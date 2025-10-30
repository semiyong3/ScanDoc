import win32gui     # (필수) 창 핸들 및 좌표 획득
import win32con     # (필수) 창 상태 확인
import os
import sys
import tempfile
import shutil
import zipfile
import glob
import time 
import pythoncom
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill  
from openpyxl.utils import get_column_letter 
from PIL import Image, ImageGrab
from pynput.keyboard import Key, Controller 
from win32com.client import Dispatch

try:
    import win32com.client
    import win32gui
    import win32con
except ImportError:
    print("pywin32 라이브러리가 필요합니다. pip install pywin32")
    sys.exit(1)

# --- 1. Scan Directory ---

def scan_directory(target_dir, output_file):
    """
    지정된 디렉터리를 스캔하여 엑셀 파일로 저장 
    """
    wb = Workbook()
    ws = wb.active
    
    sheet_name = os.path.splitext(os.path.basename(output_file))[0]
    ws.title = sheet_name
    
    base_depth = target_dir.count(os.sep)
    file_cells_coords = [] 
    
    for root, dirs, files in os.walk(target_dir, topdown=True):
        current_depth = root.count(os.sep) - base_depth
        
        folder_name = "📁 " + os.path.basename(root)
        row = [None] * current_depth + [folder_name]
        
        if files:
            files_str = "\n".join(["┣ " + f for f in files])
            row.append(files_str)
            
        ws.append(row)
        
        if files:
            current_row_index = ws.max_row
            ws.row_dimensions[current_row_index].height = 13 * len(files)
            file_col_letter = chr(ord('A') + current_depth + 1)
            file_cells_coords.append(f"{file_col_letter}{current_row_index}")

    # --- 열 너비 자동 조절 ---
    column_max_lengths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value:
                col_idx = cell.column - 1 
                cell_value_str = str(cell.value)
                length = 0
                if "\n" in cell_value_str:
                    lines = cell_value_str.split('\n')
                    length = max(len(line) for line in lines)
                else:
                    length = len(cell_value_str)
                current_max = column_max_lengths.get(col_idx, 0)
                column_max_lengths[col_idx] = max(current_max, length)

    for col_idx, max_length in column_max_lengths.items():
        col_letter = get_column_letter(col_idx + 1) 
        ws.column_dimensions[col_letter].width = max_length + 2

    # --- 전체 셀 서식 적용 ---
    
    font_9pt = Font(size=9)
    align_top_no_wrap = Alignment(vertical='top', wrap_text=False)
    align_top_wrap = Alignment(vertical='top', wrap_text=True)

    # 빈셀은 회색으로 채워서 가독성을 높임
    gray_fill = PatternFill(start_color='BFBFBF',
                            end_color='BFBFBF',
                            fill_type='solid')

    for row in ws.iter_rows():
        for cell in row:
            # 1. 기본 폰트 및 정렬 적용
            cell.font = font_9pt
            cell.alignment = align_top_no_wrap
            
            # 2. (신규) 값이 없는 셀(None)인 경우 회색으로 채우기
            if cell.value is None:
                cell.fill = gray_fill
            
    # 3. 파일 목록 셀에만 '줄바꿈 허용' 서식 덮어쓰기
    for cell_coord in file_cells_coords:
        ws[cell_coord].alignment = align_top_wrap

    wb.save(output_file)
    return f"디렉터리 스캔 완료!\n{output_file}"

# --- 2. Convert To Image ---

def _trigger_alt_printscreen_and_get_image():
    """
    pynput으로 Alt+PrintScreen을 시뮬레이션하고 클립보드에서 이미지를 읽어 반환합니다.
    """
    keyboard = Controller()
    with keyboard.pressed(Key.alt):
        keyboard.press(Key.print_screen)
        keyboard.release(Key.print_screen)
    
    # 클립보드가 업데이트될 때까지 잠시 대기
    time.sleep(0.1) 
    
    img = ImageGrab.grabclipboard()
    
    if img is None:
        raise Exception("클립보드에서 이미지를 찾을 수 없습니다. (DRM에 의해 차단되었거나, 활성 창이 없음)")
    
    return img

def capture_active_window(hwnd=None):
    """
    현재 활성화된 창(Foreground Window)만 캡처하여
    Pillow 이미지 객체로 반환합니다.
    """
    
    # 1. 활성 창의 핸들(HWND) 가져오기
    hwnd = win32gui.GetForegroundWindow()
    
    if hwnd == 0:
        raise Exception("활성화된 창을 찾을 수 없습니다.")
        
    # 2. 핸들을 사용하여 창의 외곽 좌표(bbox) 가져오기
    #    bbox는 (left, top, right, bottom) 튜플입니다.
    rect = win32gui.GetWindowRect(hwnd)
    
    # 3. GetWindowRect는 창의 그림자/테두리를 포함할 수 있습니다.
    #    정확한 클라이언트 영역을 원하면 다른 함수(GetClientRect, ClientToScreen)가
    #    필요하지만, 우선 GetWindowRect(외곽)를 사용합니다.
    bbox = (rect[0], rect[1], rect[2], rect[3])
    
    print(f"활성 창 캡처: {win32gui.GetWindowText(hwnd)} (좌표: {bbox})")

    # 4. bbox 좌표를 ImageGrab.grab()에 전달하여 해당 영역만 캡처
    screenshot = ImageGrab.grab(bbox=bbox)
    
    return screenshot

def capture_ppt_slides(target_file, output_dir, base_filename):
    """
    
    """
    #print("[DEBUG] capture_ppt_slides (슬라이드 쇼 + Alt+PrintScreen) 시작")
    #pythoncom.CoInitialize()
    
    output_path = os.path.join(os.path.abspath(output_dir), base_filename)
    os.makedirs(output_path, exist_ok=True)
    
    powerpoint = None
    presentation = None

    try:
        print("[DEBUG] 1. PowerPoint Dispatch 및 Open 시도...")
        powerpoint = Dispatch("PowerPoint.Application")
        # powerpoint.Visible = True (슬라이드 쇼가 어차피 보이게 함)
        file_path = os.path.abspath(target_file)
        # 문서는 백그라운드에서 열기
        presentation = powerpoint.Presentations.Open(file_path)
        slide_count = presentation.Slides.Count
        print(f"[DEBUG] 1. Open 성공. 총 슬라이드: {slide_count}개")

        """
        # [수정] 2. 슬라이드 쇼를 '전체 화면'으로 실행
        print("[DEBUG] 2. 슬라이드 쇼 전체 화면 실행 시도...")
        ss_settings = presentation.SlideShowSettings
        ss_window = ss_settings.Run() # 슬라이드 쇼 창 객체 반환
        print("[DEBUG] 2. 슬라이드 쇼 실행 성공.")
        
        # 슬라이드 쇼 창이 완전히 뜰 때까지 2초 대기
        
        """
        time.sleep(5.0) 

        for i in range(1, slide_count + 1):
            print(f"[DEBUG] 3-{i}. 슬라이드 {i} Select 시도...")
            slide = presentation.Slides(i)
            slide.Select()

            # [수정] 렌더링 대기 (매우 중요)
            time.sleep(1.0) 
            print(f"[DEBUG] 3-{i}. 렌더링 대기 완료.")

            # [수정] "활성 창" (즉, 슬라이드 쇼) 캡처
            print(f"[DEBUG] 4-{i}. Alt+PrintScreen 캡처 시도...")
            screenshot = capture_active_window()
            #screenshot = _trigger_alt_printscreen_and_get_image()
            print(f"[DEBUG] 4-{i}. 캡처 성공.")

            output_file_path = os.path.join(output_path, f"slide_{i:03}.png")
            screenshot.save(output_file_path, "PNG")
            print(f"[OK] {output_file_path} 저장 완료")

    except Exception as e:
        # (중요) GotoSlide가 DRM에 막히면 여기서 오류 발생
        print(f"\n[!!!] 자동화 작업 중 심각한 오류 발생: {e}\n")
        raise RuntimeError(f"PPT 변환 중 오류 발생: {e}")

    finally:
        print("[DEBUG] 6. finally 블록 실행 (정리 시작)")
        if presentation:
            presentation.Close()
            print("[DEBUG] 6-1. Presentation 닫기 완료.")
        if powerpoint:
            powerpoint.Quit()
            print("[DEBUG] 6-2. PowerPoint 종료 완료.")
        
        #pythoncom.CoUninitialize() # (CoInitialize가 아님)
        print("[DEBUG] 6-3. CoUninitialize 완료.")

    return f"PPT 슬라이드 {slide_count}개를 이미지로 저장 완료!\n{output_path}"

# --- 3. Convert To PDF (변경 없음) ---

def _numeric_sort_key(f):
    basename = os.path.splitext(os.path.basename(f))[0]
    try:
        return int(basename)
    except ValueError:
        return basename

def convert_to_pdf(target_zip, output_dir):
    temp_extract_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(target_zip, 'r') as zf:
            zf.extractall(temp_extract_dir)
            
        img_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
        image_files = [f for f in glob.glob(os.path.join(temp_extract_dir, "*")) 
                       if os.path.splitext(f)[1].lower() in img_extensions]
                       
        if not image_files:
            raise Exception("ZIP 파일 내에 변환할 수 있는 이미지 파일이 없습니다.")
            
        image_files.sort(key=_numeric_sort_key)
        images_pil = [Image.open(f).convert('RGB') for f in image_files]
        
        base_filename = os.path.splitext(os.path.basename(target_zip))[0]
        pdf_path = os.path.join(output_dir, f"{base_filename}.pdf")
        
        images_pil[0].save(
            pdf_path,
            save_all=True,
            append_images=images_pil[1:]
        )
    finally:
        if os.path.exists(temp_extract_dir):
            shutil.rmtree(temp_extract_dir)
    os.remove(target_zip)
    return f"PDF 변환 완료!\n{pdf_path}"