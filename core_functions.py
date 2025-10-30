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

def capture_ppt_slides(target_file, output_dir, base_filename):
    """
    ppt 파일을 열고 슬라이드를 한 페이지씩 이동하면서 화면을 캡처하고 파일로 저장
    """
    
    output_path = os.path.join(os.path.abspath(output_dir), base_filename)
    os.makedirs(output_path, exist_ok=True)
    
    powerpoint = None
    presentation = None

    try:
        print("[DEBUG] 1. PowerPoint Dispatch 및 Open 시도...")
        powerpoint = Dispatch("PowerPoint.Application")
        powerpoint.Visible = True
        file_path = os.path.abspath(target_file)

        presentation = powerpoint.Presentations.Open(file_path)
        slide_count = presentation.Slides.Count
        print(f"[DEBUG] 1. Open 성공. 총 슬라이드: {slide_count}개")

        #time.sleep(2.0) 
        #powerpoint.Activate()
        #time.sleep(2.0) 

        # 파워포인트 윈도우의 핸들을 찾아 최대화, 최상위로 설정
        hwnd = win32gui.FindWindow("PPTFrameClass", None)
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
        win32gui.SetForegroundWindow(hwnd)

        for i in range(1, slide_count + 1):
            print("[DEBUG] 2. Slide-{i} 캡처 시도...") 
            slide = presentation.Slides(i)
            slide.Select()
            time.sleep(0.5) 

            screenshot = capture_active_window(hwnd)
            output_file_path = os.path.join(output_path, f"slide_{i:03}.png")
            screenshot.save(output_file_path, "PNG")
            print("[DEBUG] 2. Slide-{i} 캡처 완료...") 

        print(f"[OK] {output_file_path} 저장 완료")

    except Exception as e:
        print(f"\n[!!!] 변환 작업 중 심각한 오류 발생: {e}\n")
        raise RuntimeError(f"PPT 변환 중 오류 발생: {e}")

    finally:
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()

    return f"PPT 슬라이드 {slide_count}개를 이미지로 저장 완료!\n{output_path}"

# --- 3. Convert To PDF (변경 없음) ---

def _numeric_sort_key(f):
    basename = os.path.splitext(os.path.basename(f))[0]
    try:
        # 파일명이 "slide_001.png" 같은 경우, "001"을 숫자로 변환하여 정렬
        # 숫자가 아닌 경우(예: "__MACOSX")는 basename으로 정렬
        return int(basename)
    except ValueError:
        return basename

def convert_to_pdf(target_dir, output_file):
    """
    지정된 디렉터리 내의 이미지 파일들을 모아 하나의 PDF 파일로 변환
    (이전 버전의 ZIP 파일 처리 로직 제거됨)
    """
    
    # 1. 이미지 파일 확장자 정의
    img_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')
    
    # 2. 지정된 디렉터리에서 이미지 파일 목록을 가져옵니다.
    # glob.glob을 사용하여 모든 파일을 검색하고, 확장자를 확인하여 필터링합니다.
    target_dir = os.path.abspath(target_dir)
    output_file = os.path.abspath(output_file)
    
    image_files = [f for f in glob.glob(os.path.join(target_dir, "*")) 
                   if os.path.splitext(f)[1].lower() in img_extensions]
                   
    if not image_files:
        raise Exception(f"'{target_dir}' 디렉터리 내에 변환할 수 있는 이미지 파일이 없습니다. (지원 확장자: {img_extensions})")
        
    # 3. 파일 목록을 순서대로 정렬 (slide_001.png, slide_002.png 순서 보장)
    image_files.sort(key=_numeric_sort_key)
    
    # 4. Pillow Image 객체로 로드 (PDF 변환을 위해 RGB로 변환)
    # PIL.Image.open() 시 파일이 잠기는 것을 방지하기 위해 .convert('RGB')까지 처리
    try:
        images_pil = [Image.open(f).convert('RGB') for f in image_files]
    except Exception as e:
        raise RuntimeError(f"이미지 파일을 로드하는 중 오류 발생: {e}")

    
    # 5. PDF 파일 저장 경로 설정 (output_file은 app_window.py에서 이미 전체 경로를 받음)
    pdf_path = output_file
    
    # 6. 첫 번째 이미지를 기준으로 PDF를 생성하고 나머지 이미지들을 추가합니다.
    if images_pil:
        images_pil[0].save(
            pdf_path,
            save_all=True,
            append_images=images_pil[1:]
        )
    else:
        # 이 else 블록은 2단계에서 이미 처리되었으나, 안전을 위해 남겨둡니다.
        raise Exception("변환할 이미지가 준비되지 않았습니다.")
    
    
    # 이전 버전에서 사용되던 shutil, tempfile 관련 로직은 제거되었습니다.
    
    return f"PDF 변환 완료!\n총 {len(image_files)}개의 이미지를 {pdf_path}로 병합했습니다."
