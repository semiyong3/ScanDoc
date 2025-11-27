#### ScanDoc (PPT Auto Mode)

## 소개
ScanDoc은 Windows 환경에서 PowerPoint, Excel, Word, PDF 파일을 자동으로 이미지로 변환하거나, 폴더 내 파일을 스캔 및 PDF로 변환하는 데스크탑 도구입니다.  
PyQt5 기반의 직관적인 UI를 제공합니다.

## 주요 기능

### 1. Scan Directory
- **설명**: 지정한 폴더 내 파일을 스캔하여 결과를 출력 폴더에 저장합니다.
- **사용법**:
  - [FIND] 버튼으로 대상 폴더 선택
  - [SET] 버튼으로 출력 폴더 선택
  - [RUN] 버튼으로 스캔 실행

### 2. Convert To Image (ppt/xls/doc/pdf)
- **설명**: 선택한 파일을 자동으로 이미지로 변환(캡처)합니다. 슬라이드/페이지별로 Interval(초) 지정 가능.
- **사용법**:
  - [FIND] 버튼으로 대상 파일 선택
  - [SET] 버튼으로 출력 폴더 선택
  - Interval(초) 입력 (최소 1초)
  - [RUN] 버튼으로 변환 실행
- **지원 확장자**: ppt, xls, doc, pdf

### 3. Convert To PDF
- **설명**: 지정한 폴더 내 파일을 하나의 PDF로 변환합니다.
- **사용법**:
  - [FIND] 버튼으로 대상 폴더 선택
  - [SET] 버튼으로 출력 파일 지정
  - [RUN] 버튼으로 변환 실행

## 실행 방법

1. Python 3.x 설치
2. 필수 라이브러리 설치
   ```bash
   pip install -r requirements.txt
3. 프로그램 실행
    python main.py   

## 빌드 방법 (Windows)
PyInstaller로 실행 파일 생성:

생성된 exe 파일은 dist 폴더에서 확인할 수 있습니다.

## 사용 환경
Windows 10 이상
Python 3.7 이상
PyQt5

## 참고 및 주의사항
이미지 변환 시, 캡처 작업이 자동으로 진행되므로 작업 중 프로그램 창에 포커스를 유지하세요.
듀얼 모니터 환경에서는 캡처 위치가 달라질 수 있습니다.
Interval(초)는 슬라이드/페이지 전환 간격입니다.
