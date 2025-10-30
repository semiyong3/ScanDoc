# ScanDoc 프로그램 사용법

## 1. 프로그램 사용법 (중요!)

### 1-1. Scan Directory
(이전과 동일)

### 1-2. Convert To Image (DRM 캡처 방식)
이 기능은 `Alt+PrintScreen` (활성 창 캡처)을 이용한 수동 캡처 방식입니다.

1.  **수동으로** 변환할 DRM 문서(PPT, Word, PDF 뷰어 등)를 엽니다.
2.  ScanDoc 프로그램에서 `Target File`을 선택합니다. (출력될 `zip` 파일의 이름을 정하는 데 사용됩니다.)
3.  `Output Dir`에 `zip` 파일이 저장될 폴더를 지정합니다.
4.  [RUN] 버튼을 누르고, 안내 팝업창에서 [확인]을 누릅니다.
5.  **"Capture Session"**이라는 작은 보조창이 뜹니다.
6.  1번에서 열어둔 **문서 창을 마우스로 클릭**하여 활성화시킵니다.
7.  키보드의 **F9** 키를 누릅니다. (또는 보조창의 "Capture" 버튼 클릭)
    * (이때 프로그램이 내부적으로 `Alt+PrintScreen` 키를 눌러 클립보드에 복사 후 저장합니다.)
8.  캡처가 완료되면 수동으로 문서의 다음 페이지로 이동합니다.
9.  다시 **F9** 키를 눌러 다음 페이지를 캡처합니다.
10. 마지막 페이지까지 캡처를 반복합니다.
11. 모든 캡처가 끝나면, "Capture Session" 보조창의 **[Finish & Zip]** 버튼을 누릅니다.
12. `Output Dir`에 `파일명.zip` 파일이 생성됩니다.

### 1-3. Convert To PDF
(이전과 동일)

## 2. 프로그램 실행 (Python 스크립트로 실행 시)

1.  PC에 Python 3.8 이상을 설치합니다.
2.  프로젝트 폴더에서 다음 명령어를 실행하여 라이브러리를 설치합니다.
    ```bash
    pip install -r requirements.txt
    ```
3.  다음 명령어로 프로그램을 실행합니다.
    ```bash
    python main.py
    ```

## 3. 프로그램 실행 (.exe 빌드 및 실행)

1.  `build.bat` 파일을 실행합니다.
2.  빌드가 완료되면 `dist` 폴더 안에 `ScanDoc.exe` 파일이 생성됩니다.
3.  `ScanDoc.exe` 파일을 실행합니다.