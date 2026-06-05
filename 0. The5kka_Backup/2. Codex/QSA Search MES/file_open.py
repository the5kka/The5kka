import os
import re
import subprocess


def open_file(file_path, location=""):
    """검색 결과 더블클릭 시 원본 문서를 엽니다."""
    if not os.path.exists(file_path):
        raise FileNotFoundError("원본 파일을 찾을 수 없습니다.")

    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        open_pdf(file_path, location)
        return

    os.startfile(file_path)


def open_pdf(file_path, location=""):
    """PDF는 가능하면 기본 연결 프로그램으로 엽니다."""
    page = extract_number(location)
    try:
        if page:
            subprocess.Popen(["cmd", "/c", "start", "", file_path], shell=False)
        else:
            os.startfile(file_path)
    except Exception:
        os.startfile(file_path)


def extract_number(text):
    """Slide/Page/Paragraph 같은 위치 문자열에서 숫자를 추출합니다."""
    match = re.search(r"(\d+)", text or "")
    return int(match.group(1)) if match else None
