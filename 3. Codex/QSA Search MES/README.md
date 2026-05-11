# QSA Audit Search MES

Windows 10/11에서 IATF16949 / QSA Audit 대응 문서를 빠르게 검색하기 위한 Python Tkinter 프로그램입니다.

## 설치 방법

1. Python 3.10 이상을 설치합니다.
2. 이 폴더에서 명령 프롬프트를 엽니다.
3. 아래 명령을 실행합니다.

```bat
python -m pip install -r requirements.txt
```

## 실행 방법

```bat
python main.py
```

프로그램 이름은 `QSA Audit Search MES`입니다.

## 인덱스 생성 방법

1. 기본 폴더는 `D:\QC\1. JIIN\1. IATF 16949`로 설정되어 있습니다.
2. 다른 폴더를 검색하려면 `폴더 선택` 버튼을 누릅니다.
3. `문서 인덱스 생성` 버튼을 누르면 하위 폴더까지 문서를 읽어 SQLite DB에 저장합니다.
4. DB 파일은 프로그램 폴더의 `qsa_index.db`입니다.
5. 원본 문서는 삭제, 이동, 복사, 수정하지 않습니다. DB에는 파일 경로와 추출 텍스트만 저장합니다.

## 검색 방법

1. 검색어를 입력합니다.
2. `검색` 버튼을 누르거나 Enter 키를 누릅니다.
3. 문서 내부 텍스트에서 대소문자 구분 없이 검색합니다.
4. 결과에는 파일명, 형식, 위치, 내용 일부, 전체 경로가 표시됩니다.
5. 검색 결과를 더블클릭하면 원본 문서가 열립니다.

## EXE 만드는 방법

`build_exe.bat` 파일을 더블클릭하거나 아래 명령을 실행합니다.

```bat
build_exe.bat
```

생성되는 실행 파일 이름은 다음과 같습니다.

```text
dist\QSA Audit Search MES.exe
```

## 지원 파일

- PowerPoint: `.pptx`, `.ppt`
- Word: `.docx`, `.doc`
- PDF: `.pdf`

`.ppt`, `.doc` 파일은 Microsoft Office가 설치된 Windows 환경에서 `pywin32`를 통해 읽습니다.

## 주의사항

- 원본 문서는 절대 삭제, 이동, 수정하지 않습니다.
- 읽기 실패한 파일은 `index_errors.log`에 기록됩니다.
- 이미 인덱스된 파일은 수정일이 동일하면 다시 처리하지 않습니다.
- 수정일이 변경된 파일만 재인덱싱합니다.
- 삭제된 원본 파일은 검색 결과에서 제외되도록 DB에서 정리합니다.
- PDF 특정 페이지 자동 이동은 사용자의 기본 PDF 프로그램에 따라 지원 여부가 달라질 수 있습니다.
- EXE 빌드 시 백신 또는 Windows 보안 정책 때문에 실행이 지연될 수 있습니다.
