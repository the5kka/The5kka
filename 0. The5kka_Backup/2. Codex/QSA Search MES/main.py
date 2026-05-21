import os
import queue
import json
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

import db
import file_open
import indexer
import search


APP_TITLE = "QSA Audit Search MES"
DEFAULT_FOLDER = r"D:\QC\1. JIIN\1. IATF 16949"
SETTINGS_FILE = "settings.json"


class QsaSearchApp:
    """QSA Audit 문서 검색 프로그램 메인 GUI입니다."""

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1200x720")
        self.root.minsize(980, 600)
        self.default_font = tkfont.Font(family="Malgun Gothic", size=10)
        self.entry_font = tkfont.Font(family="Malgun Gothic", size=10)
        self.root.option_add("*Font", self.default_font)

        self.message_queue = queue.Queue()
        self.index_thread = None
        self.results = []

        db.init_db()
        self.create_styles()
        self.create_widgets()
        self.poll_queue()

    def get_settings_path(self):
        """설정 파일 경로를 반환합니다."""
        return os.path.join(db.get_app_dir(), SETTINGS_FILE)

    def load_last_folder(self):
        """마지막으로 선택한 문서 폴더를 불러옵니다."""
        try:
            settings_path = self.get_settings_path()
            if not os.path.exists(settings_path):
                return DEFAULT_FOLDER
            with open(settings_path, "r", encoding="utf-8") as file:
                data = json.load(file)
            folder = data.get("last_folder") or DEFAULT_FOLDER
            return folder if os.path.isdir(folder) else DEFAULT_FOLDER
        except Exception:
            return DEFAULT_FOLDER

    def save_last_folder(self, folder):
        """선택한 문서 폴더를 다음 실행 때도 사용할 수 있게 저장합니다."""
        try:
            with open(self.get_settings_path(), "w", encoding="utf-8") as file:
                json.dump({"last_folder": folder}, file, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def create_styles(self):
        """입력 중 화면 크기가 흔들리지 않도록 기본 스타일을 고정합니다."""
        style = ttk.Style()
        style.configure("TEntry", padding=(4, 2))
        style.configure("TButton", padding=(8, 3))
        style.configure("Treeview", rowheight=24)

    def create_widgets(self):
        """화면 구성 요소를 생성합니다."""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        top_frame = ttk.Frame(self.root, padding=(10, 10, 10, 4))
        top_frame.grid(row=0, column=0, sticky="ew")
        top_frame.columnconfigure(1, weight=1)

        ttk.Label(top_frame, text="문서 폴더", width=8).grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.folder_var = tk.StringVar(value=self.load_last_folder())
        folder_entry = ttk.Entry(top_frame, textvariable=self.folder_var, font=self.entry_font)
        folder_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        ttk.Button(top_frame, text="문서 폴더 선택", width=14, command=self.select_folder).grid(row=0, column=2, padx=(0, 8))
        self.index_button = ttk.Button(top_frame, text="선택 폴더 인덱스 생성", width=18, command=self.start_indexing)
        self.index_button.grid(row=0, column=3)

        search_frame = ttk.Frame(self.root, padding=(10, 4, 10, 4))
        search_frame.grid(row=1, column=0, sticky="ew")
        search_frame.columnconfigure(1, weight=1)

        ttk.Label(search_frame, text="검색어", width=8).grid(row=0, column=0, padx=(0, 8), sticky="w")
        self.keyword_var = tk.StringVar()
        keyword_entry = ttk.Entry(search_frame, textvariable=self.keyword_var, font=self.entry_font)
        keyword_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        keyword_entry.bind("<Return>", lambda _event: self.search_documents())

        ttk.Button(search_frame, text="검색", width=12, command=self.search_documents).grid(row=0, column=2, padx=(0, 8))
        self.result_count_var = tk.StringVar(value="검색 결과: 0건")
        ttk.Label(search_frame, textvariable=self.result_count_var, width=16, anchor="e").grid(row=0, column=3, sticky="e")

        table_frame = ttk.Frame(self.root, padding=(10, 4, 10, 4))
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        columns = ("no", "file_name", "file_type", "location", "preview", "file_path")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        self.tree.heading("no", text="No")
        self.tree.heading("file_name", text="파일명")
        self.tree.heading("file_type", text="형식")
        self.tree.heading("location", text="위치")
        self.tree.heading("preview", text="내용 일부")
        self.tree.heading("file_path", text="전체 경로")

        self.tree.column("no", width=55, anchor="center", stretch=False)
        self.tree.column("file_name", width=220, anchor="w")
        self.tree.column("file_type", width=70, anchor="center", stretch=False)
        self.tree.column("location", width=110, anchor="center", stretch=False)
        self.tree.column("preview", width=430, anchor="w")
        self.tree.column("file_path", width=360, anchor="w")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self.open_selected_file)

        bottom_frame = ttk.Frame(self.root, padding=(10, 4, 10, 10))
        bottom_frame.grid(row=3, column=0, sticky="ew")
        bottom_frame.columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="준비 완료")
        ttk.Label(bottom_frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w")

        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(bottom_frame, variable=self.progress_var, maximum=100)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(6, 0))

    def select_folder(self):
        """검색 대상 폴더를 선택합니다."""
        try:
            folder = filedialog.askdirectory(initialdir=self.folder_var.get() or DEFAULT_FOLDER)
            if folder:
                self.folder_var.set(folder)
                self.save_last_folder(folder)
        except Exception as exc:
            messagebox.showerror("폴더 선택 오류", str(exc))

    def start_indexing(self):
        """문서 인덱스 생성을 시작합니다."""
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("확인 필요", "올바른 폴더 경로를 입력하거나 선택하세요.")
            return

        self.save_last_folder(folder)
        if self.index_thread and self.index_thread.is_alive():
            messagebox.showinfo("진행 중", "이미 인덱스 생성이 진행 중입니다.")
            return

        self.index_button.configure(state="disabled")
        self.progress_var.set(0)
        self.status_var.set("인덱스 생성 준비 중...")
        self.index_thread = threading.Thread(target=self.index_worker, args=(folder,), daemon=True)
        self.index_thread.start()

    def index_worker(self, folder):
        """인덱스 작업 스레드입니다."""
        try:
            result = indexer.build_index(
                folder,
                progress_callback=self.on_index_progress,
                status_callback=self.on_index_status,
            )
            self.message_queue.put(("index_done", result))
        except Exception as exc:
            self.message_queue.put(("error", f"인덱스 생성 오류: {exc}"))

    def on_index_progress(self, current, total, file_name):
        """인덱스 진행률을 큐로 전달합니다."""
        percent = 0 if total == 0 else current / total * 100
        self.message_queue.put(("progress", percent, current, total, file_name))

    def on_index_status(self, message):
        """상태 메시지를 큐로 전달합니다."""
        self.message_queue.put(("status", message))

    def poll_queue(self):
        """작업 스레드에서 전달한 메시지를 GUI에 반영합니다."""
        try:
            while True:
                message = self.message_queue.get_nowait()
                kind = message[0]
                if kind == "progress":
                    _, percent, current, total, file_name = message
                    self.progress_var.set(percent)
                    if file_name:
                        self.status_var.set(f"{current}/{total} 처리 중: {file_name}")
                elif kind == "status":
                    self.status_var.set(message[1])
                elif kind == "index_done":
                    self.index_button.configure(state="normal")
                    result = message[1]
                    self.status_var.set(
                        "완료: 전체 {total}개, 신규/변경 {indexed}개, 건너뜀 {skipped}개, 실패 {failed}개, 삭제정리 {deleted}개".format(
                            **result
                        )
                    )
                    if result["failed"] > 0:
                        messagebox.showwarning(
                            "인덱스 완료",
                            "일부 파일 읽기에 실패했습니다.\nindex_errors.log 파일을 확인하세요.",
                        )
                elif kind == "error":
                    self.index_button.configure(state="normal")
                    self.status_var.set(message[1])
                    messagebox.showerror("오류", message[1])
        except queue.Empty:
            pass
        self.root.after(150, self.poll_queue)

    def search_documents(self):
        """DB에서 문서 내부 텍스트를 검색합니다."""
        keyword = self.keyword_var.get().strip()
        if not keyword:
            messagebox.showwarning("검색어 필요", "검색어를 입력하세요.")
            return

        try:
            self.clear_results()
            self.results = search.search_documents(keyword)
            for idx, result in enumerate(self.results, start=1):
                self.tree.insert(
                    "",
                    "end",
                    iid=str(idx - 1),
                    values=(
                        idx,
                        result["file_name"],
                        result["file_type"],
                        result["location"],
                        result["preview"],
                        result["file_path"],
                    ),
                )
            self.result_count_var.set(f"검색 결과: {len(self.results)}건")
            self.status_var.set("검색 완료")
        except Exception as exc:
            messagebox.showerror("검색 오류", str(exc))
            self.status_var.set(f"검색 오류: {exc}")

    def clear_results(self):
        """검색 결과 테이블을 비웁니다."""
        for item in self.tree.get_children():
            self.tree.delete(item)

    def open_selected_file(self, _event=None):
        """더블클릭한 검색 결과의 원본 문서를 엽니다."""
        selected = self.tree.selection()
        if not selected:
            return
        try:
            index = int(selected[0])
            result = self.results[index]
            file_open.open_file(result["file_path"], result["location"])
            self.status_var.set(f"원본 열기: {result['file_name']} / {result['location']}")
        except Exception as exc:
            messagebox.showerror("파일 열기 실패", str(exc))
            self.status_var.set(f"파일 열기 실패: {exc}")


def main():
    """프로그램 시작점입니다."""
    try:
        root = tk.Tk()
        app = QsaSearchApp(root)
        root.mainloop()
    except Exception as exc:
        messagebox.showerror("프로그램 오류", str(exc))


if __name__ == "__main__":
    main()
