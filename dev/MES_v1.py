# -*- coding: utf-8 -*-
import sqlite3
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk


def get_base_dir():
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).resolve().parent
    else:
        app_dir = Path(__file__).resolve().parent

    if app_dir.name.lower() in {"dev", "exe", "final"}:
        return app_dir.parent
    return app_dir


DB_PATH = get_base_dir() / "data" / "quality_history.db"


COLORS = {
    "bg": "#f6f7f9",
    "panel": "#ffffff",
    "border": "#d7dce2",
    "text": "#2f343a",
    "muted": "#5f6872",
    "header": "#eef2f6",
    "navy": "#1f4e79",
    "navy_hover": "#183d60",
    "button": "#f3f5f7",
    "button_hover": "#e7ebef",
    "selection": "#dbe9f6",
}


def get_connection():
    return sqlite3.connect(DB_PATH)


def init_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS quality_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_name TEXT NOT NULL,
                lot_no TEXT NOT NULL,
                inspection_date TEXT NOT NULL,
                inspector TEXT NOT NULL,
                result TEXT NOT NULL,
                defect_type TEXT,
                note TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )


class QualityHistoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("품질 이력 관리 프로그램")
        self.root.geometry("1200x740")
        self.root.minsize(1040, 640)
        self.root.configure(bg=COLORS["bg"])

        self.selected_id = None
        self.fields = {}

        self.setup_style()
        self.create_widgets()
        self.search_records()

    def setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        base_font = ("맑은 고딕", 10)
        heading_font = ("맑은 고딕", 10, "bold")

        style.configure(".", font=base_font, foreground=COLORS["text"])
        style.configure("App.TFrame", background=COLORS["bg"])
        style.configure("Panel.TFrame", background=COLORS["panel"])
        style.configure("Panel.TLabelframe", background=COLORS["panel"])
        style.configure(
            "Panel.TLabelframe.Label",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=("맑은 고딕", 10, "bold"),
        )
        style.configure("Form.TLabel", background=COLORS["panel"], foreground=COLORS["muted"])
        style.configure(
            "TEntry",
            fieldbackground="#ffffff",
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            padding=(8, 5),
        )
        style.configure(
            "TCombobox",
            fieldbackground="#ffffff",
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            arrowcolor=COLORS["muted"],
            padding=(8, 5),
        )
        style.configure(
            "Primary.TButton",
            background=COLORS["navy"],
            foreground="#ffffff",
            bordercolor=COLORS["navy"],
            focusthickness=1,
            focuscolor=COLORS["navy"],
            padding=(16, 7),
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLORS["navy_hover"]), ("pressed", COLORS["navy_hover"])],
            foreground=[("active", "#ffffff")],
        )
        style.configure(
            "Secondary.TButton",
            background=COLORS["button"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            padding=(16, 7),
        )
        style.map(
            "Secondary.TButton",
            background=[("active", COLORS["button_hover"]), ("pressed", COLORS["button_hover"])],
            foreground=[("active", COLORS["text"])],
        )
        style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            rowheight=30,
            font=base_font,
        )
        style.configure(
            "Treeview.Heading",
            background=COLORS["header"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            relief="flat",
            font=heading_font,
            padding=(8, 7),
        )
        style.map(
            "Treeview",
            background=[("selected", COLORS["selection"])],
            foreground=[("selected", COLORS["text"])],
        )

    def create_widgets(self):
        container = ttk.Frame(self.root, style="App.TFrame", padding=14)
        container.grid(row=0, column=0, sticky="nsew")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        self.create_input_area(container)
        self.create_search_area(container)
        self.create_table_area(container)

    def create_input_area(self, parent):
        input_frame = ttk.LabelFrame(
            parent, text="입력 영역", style="Panel.TLabelframe", padding=(16, 12)
        )
        input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        for col in (1, 3, 5):
            input_frame.columnconfigure(col, weight=1)

        items = [
            ("제품명", "product_name"),
            ("LOT 번호", "lot_no"),
            ("검사일자", "inspection_date"),
            ("검사자", "inspector"),
            ("결과", "result"),
            ("불량유형", "defect_type"),
            ("비고", "note"),
        ]

        for index, (label, key) in enumerate(items):
            row = index // 3
            col = (index % 3) * 2
            ttk.Label(input_frame, text=label, style="Form.TLabel").grid(
                row=row, column=col, sticky="w", padx=(0, 8), pady=7
            )

            if key == "result":
                widget = ttk.Combobox(
                    input_frame,
                    values=("합격", "불합격"),
                    state="readonly",
                    width=22,
                )
                widget.set("합격")
            elif key == "note":
                widget = ttk.Entry(input_frame)
                widget.grid(
                    row=row,
                    column=col + 1,
                    columnspan=3,
                    sticky="ew",
                    padx=(0, 16),
                    pady=7,
                )
                self.fields[key] = widget
                continue
            else:
                widget = ttk.Entry(input_frame, width=22)
                if key == "inspection_date":
                    widget.insert(0, datetime.now().strftime("%Y-%m-%d"))

            widget.grid(row=row, column=col + 1, sticky="ew", padx=(0, 16), pady=7)
            self.fields[key] = widget

        button_frame = ttk.Frame(input_frame, style="Panel.TFrame")
        button_frame.grid(row=3, column=0, columnspan=6, sticky="e", pady=(12, 0))

        ttk.Button(
            button_frame, text="입력", command=self.add_record, style="Primary.TButton"
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            button_frame, text="수정", command=self.update_record, style="Secondary.TButton"
        ).pack(side="left", padx=6)
        ttk.Button(
            button_frame, text="삭제", command=self.delete_record, style="Secondary.TButton"
        ).pack(side="left", padx=6)
        ttk.Button(
            button_frame, text="초기화", command=self.clear_form, style="Secondary.TButton"
        ).pack(side="left", padx=(6, 0))

    def create_search_area(self, parent):
        search_frame = ttk.LabelFrame(
            parent, text="조회/검색 영역", style="Panel.TLabelframe", padding=(16, 12)
        )
        search_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)

        ttk.Label(search_frame, text="검색어", style="Form.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        self.search_entry = ttk.Entry(search_frame)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12))
        self.search_entry.bind("<Return>", lambda _event: self.search_records())

        ttk.Button(
            search_frame, text="조회", command=self.search_records, style="Primary.TButton"
        ).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(
            search_frame,
            text="전체 조회",
            command=self.show_all_records,
            style="Secondary.TButton",
        ).grid(row=0, column=3)

    def create_table_area(self, parent):
        table_frame = ttk.LabelFrame(
            parent, text="결과 테이블", style="Panel.TLabelframe", padding=(16, 12)
        )
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        columns = (
            "id",
            "product_name",
            "lot_no",
            "inspection_date",
            "inspector",
            "result",
            "defect_type",
            "note",
            "created_at",
            "updated_at",
        )
        headings = {
            "id": "ID",
            "product_name": "제품명",
            "lot_no": "LOT 번호",
            "inspection_date": "검사일자",
            "inspector": "검사자",
            "result": "결과",
            "defect_type": "불량유형",
            "note": "비고",
            "created_at": "등록일",
            "updated_at": "수정일",
        }
        widths = {
            "id": 60,
            "product_name": 150,
            "lot_no": 130,
            "inspection_date": 115,
            "inspector": 105,
            "result": 80,
            "defect_type": 125,
            "note": 260,
            "created_at": 155,
            "updated_at": 155,
        }

        self.tree = ttk.Treeview(
            table_frame, columns=columns, show="headings", selectmode="browse"
        )
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], minwidth=widths[column])

        for column in ("id", "inspection_date", "result", "created_at", "updated_at"):
            self.tree.column(column, anchor="center")
        for column in ("product_name", "lot_no", "inspector", "defect_type", "note"):
            self.tree.column(column, anchor="w")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.tree.xview
        )
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

    def get_form_data(self):
        data = {key: widget.get().strip() for key, widget in self.fields.items()}

        required = {
            "product_name": "제품명",
            "lot_no": "LOT 번호",
            "inspection_date": "검사일자",
            "inspector": "검사자",
            "result": "결과",
        }
        for key, label in required.items():
            if not data[key]:
                messagebox.showwarning("입력 확인", f"{label}은(는) 필수 입력입니다.")
                self.fields[key].focus_set()
                return None

        try:
            datetime.strptime(data["inspection_date"], "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning(
                "날짜 형식 오류", "검사일자는 YYYY-MM-DD 형식으로 입력하세요."
            )
            self.fields["inspection_date"].focus_set()
            return None

        return data

    def add_record(self):
        data = self.get_form_data()
        if data is None:
            return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO quality_history (
                    product_name, lot_no, inspection_date, inspector,
                    result, defect_type, note, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    data["product_name"],
                    data["lot_no"],
                    data["inspection_date"],
                    data["inspector"],
                    data["result"],
                    data["defect_type"],
                    data["note"],
                    now,
                    now,
                ),
            )

        messagebox.showinfo("등록 완료", "품질 이력이 등록되었습니다.")
        self.clear_form()
        self.search_records()

    def search_records(self):
        keyword = self.search_entry.get().strip()
        like_keyword = f"%{keyword}%"

        with get_connection() as conn:
            rows = conn.execute(
                """
                SELECT id, product_name, lot_no, inspection_date, inspector,
                       result, defect_type, note, created_at, updated_at
                FROM quality_history
                WHERE product_name LIKE ?
                   OR lot_no LIKE ?
                   OR inspector LIKE ?
                   OR result LIKE ?
                   OR defect_type LIKE ?
                   OR note LIKE ?
                ORDER BY inspection_date DESC, id DESC
                """,
                (
                    like_keyword,
                    like_keyword,
                    like_keyword,
                    like_keyword,
                    like_keyword,
                    like_keyword,
                ),
            ).fetchall()

        self.load_table(rows)

    def show_all_records(self):
        self.search_entry.delete(0, tk.END)
        self.search_records()

    def load_table(self, rows):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for index, row in enumerate(rows):
            display_row = tuple("" if value is None else value for value in row)
            tag = "evenrow" if index % 2 == 0 else "oddrow"
            self.tree.insert("", tk.END, values=display_row, tags=(tag,))

        self.tree.tag_configure("evenrow", background="#ffffff")
        self.tree.tag_configure("oddrow", background="#fafbfc")

    def on_row_select(self, _event):
        selection = self.tree.selection()
        if not selection:
            return

        values = self.tree.item(selection[0], "values")
        self.selected_id = values[0]

        field_order = [
            "product_name",
            "lot_no",
            "inspection_date",
            "inspector",
            "result",
            "defect_type",
            "note",
        ]

        for index, key in enumerate(field_order, start=1):
            widget = self.fields[key]
            if isinstance(widget, ttk.Combobox):
                widget.set(values[index])
            else:
                widget.delete(0, tk.END)
                widget.insert(0, values[index])

    def update_record(self):
        if not self.selected_id:
            messagebox.showwarning("선택 필요", "수정할 행을 먼저 선택하세요.")
            return

        data = self.get_form_data()
        if data is None:
            return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with get_connection() as conn:
            conn.execute(
                """
                UPDATE quality_history
                SET product_name = ?,
                    lot_no = ?,
                    inspection_date = ?,
                    inspector = ?,
                    result = ?,
                    defect_type = ?,
                    note = ?,
                    updated_at = ?
                WHERE id = ?
                """,
                (
                    data["product_name"],
                    data["lot_no"],
                    data["inspection_date"],
                    data["inspector"],
                    data["result"],
                    data["defect_type"],
                    data["note"],
                    now,
                    self.selected_id,
                ),
            )

        messagebox.showinfo("수정 완료", "품질 이력이 수정되었습니다.")
        self.clear_form()
        self.search_records()

    def delete_record(self):
        if not self.selected_id:
            messagebox.showwarning("선택 필요", "삭제할 행을 먼저 선택하세요.")
            return

        if not messagebox.askyesno("삭제 확인", "선택한 품질 이력을 삭제하시겠습니까?"):
            return

        with get_connection() as conn:
            conn.execute("DELETE FROM quality_history WHERE id = ?", (self.selected_id,))

        messagebox.showinfo("삭제 완료", "품질 이력이 삭제되었습니다.")
        self.clear_form()
        self.search_records()

    def clear_form(self):
        self.selected_id = None
        for key, widget in self.fields.items():
            if isinstance(widget, ttk.Combobox):
                widget.set("합격")
            else:
                widget.delete(0, tk.END)
                if key == "inspection_date":
                    widget.insert(0, datetime.now().strftime("%Y-%m-%d"))

        for item in self.tree.selection():
            self.tree.selection_remove(item)


def main():
    init_db()
    root = tk.Tk()
    QualityHistoryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
