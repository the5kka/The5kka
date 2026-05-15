import os
import sqlite3
import sys
from datetime import datetime


DB_FILE_NAME = "qsa_index.db"


def get_app_dir():
    """프로그램 실행 위치를 반환합니다."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_db_path():
    """SQLite DB 파일 경로를 반환합니다."""
    return os.path.join(get_app_dir(), DB_FILE_NAME)


def normalize_path(path):
    """같은 파일이 다른 경로 표기로 중복 저장되지 않도록 정리합니다."""
    if not path:
        return ""
    return os.path.normcase(os.path.normpath(os.path.abspath(path)))


def connect_db():
    """DB에 연결하고 Row 형태로 결과를 받을 수 있게 설정합니다."""
    conn = sqlite3.connect(get_db_path())
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """문서 검색용 테이블과 인덱스를 생성합니다."""
    conn = connect_db()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                file_type TEXT NOT NULL,
                location TEXT NOT NULL,
                content TEXT NOT NULL,
                modified_time REAL NOT NULL,
                indexed_time TEXT NOT NULL
            )
            """
        )
        cur.execute("CREATE INDEX IF NOT EXISTS idx_documents_path ON documents(file_path)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_documents_modified ON documents(modified_time)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_documents_type ON documents(file_type)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_documents_content ON documents(content)")
        conn.commit()
    finally:
        conn.close()


def get_indexed_modified_time(file_path):
    """이미 인덱스된 파일의 수정일을 조회합니다."""
    file_path = normalize_path(file_path)
    conn = connect_db()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT modified_time FROM documents WHERE file_path = ? LIMIT 1",
            (file_path,),
        )
        row = cur.fetchone()
        return row["modified_time"] if row else None
    finally:
        conn.close()


def delete_file_records(file_path, conn=None):
    """특정 파일의 기존 인덱스 데이터를 삭제합니다."""
    file_path = normalize_path(file_path)
    own_conn = conn is None
    if own_conn:
        conn = connect_db()
    try:
        conn.execute("DELETE FROM documents WHERE file_path = ?", (file_path,))
        if own_conn:
            conn.commit()
    finally:
        if own_conn:
            conn.close()


def insert_records(records):
    """추출된 문서 텍스트 목록을 DB에 저장합니다."""
    if not records:
        return

    conn = connect_db()
    try:
        cur = conn.cursor()
        indexed_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = [
            (
                record["file_name"],
                record["file_path"],
                record["file_type"],
                record["location"],
                record["content"],
                record["modified_time"],
                indexed_time,
            )
            for record in records
            if record.get("content", "").strip()
        ]
        cur.executemany(
            """
            INSERT INTO documents
            (file_name, file_path, file_type, location, content, modified_time, indexed_time)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )
        conn.commit()
    finally:
        conn.close()


def remove_deleted_files():
    """원본 파일이 삭제된 경우 DB 검색 결과에서 제외되도록 인덱스를 정리합니다."""
    conn = connect_db()
    removed_count = 0
    try:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT file_path FROM documents")
        paths = [row["file_path"] for row in cur.fetchall()]
        for path in paths:
            if not os.path.exists(path):
                cur.execute("DELETE FROM documents WHERE file_path = ?", (path,))
                removed_count += 1
        conn.commit()
        return removed_count
    finally:
        conn.close()
