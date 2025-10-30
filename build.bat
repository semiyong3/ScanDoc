"""
pyinstaller --onefile --windowed --name ScanDoc main.py
--------------------------------------------------------

부분,역할,설명
pyinstaller,명령어,PyInstaller 도구를 호출합니다.

--onefile
단일 파일 옵션, (줄여서 -F) 하나의 실행 파일
(.exe on Windows, or a file with no extension on Linux/macOS)을 생성
실행 시 임시 폴더에 압축을 풀고 실행, 이 옵션을 사용하지 않으면 여러 파일과 폴더로 구성된 디렉터리가 생성됨

--windowed,GUI/창 모드 옵션,(줄여서 -w) 
콘솔 창(Command Prompt 또는 Terminal)을 표시하지 않고 애플리케이션을 실행
PyQt나 Tkinter 등 GUI 기반 애플리케이션에 필수적

--name ScanDoc,이름 지정 옵션,(줄여서 -n) 생성될 최종 실행 파일의 이름을 ScanDoc으로 지정
(예: ScanDoc.exe). 이 옵션을 생략하면 입력 파일 이름(main)을 따름
main.py,입력 파일,빌드할 파이썬 스크립트의 메인 파일
PyInstaller는 이 파일을 분석하여 필요한 모든 모듈을 찾음
"""

@echo OFF
echo [ScanDoc] .exe 빌드를 시작합니다.

echo [1/3] requirements.txt 라이브러리 설치 중...
pip install -r requirements.txt

echo [2/3] PyInstaller 실행 중...
pyinstaller --onefile --windowed --name ScanDoc main.py

echo [3/3] 빌드 완료. 'dist' 폴더를 확인하세요.
pause