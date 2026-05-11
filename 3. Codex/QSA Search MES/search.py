import os
import sqlite3

import db


def make_preview(content, keyword, size=200):
    """검색어 주변 내용을 약 200자 미리보기로 만듭니다."""
    if not content:
        return ""

    lowered = content.lower()
    lowered_keyword = keyword.lower()
    found_at = lowered.find(lowered_keyword)
    if found_at < 0:
        return content[:size]

    half = size // 2
    start = max(found_at - half, 0)
    end = min(found_at + len(keyword) + half, len(content))
    preview = content[start:end].strip()
    if start > 0:
        preview = "..." + preview
    if end < len(content):
        preview += "..."
    return preview


def search_documents(keyword, limit=500):
    """문서 내부 텍스트에서 대소문자 구분 없이 검색합니다."""
    keyword = (keyword or "").strip()
    if not keyword:
        return []

    db.remove_deleted_files()
    conn = db.connect_db()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT id, file_name, file_path, file_type, location, content, modified_time
            FROM documents
            WHERE content LIKE ? COLLATE NOCASE
              AND file_path IS NOT NULL
            ORDER BY modified_time DESC, id DESC
            LIMIT ?
            """,
            (f"%{keyword}%", limit),
        )
        rows = cur.fetchall()
    except sqlite3.Error:
        return []
    finally:
        conn.close()

    results = []
    for row in rows:
        file_path = row["file_path"]
        file_path = db.normalize_path(file_path)
        if not os.path.exists(file_path):
            continue
        results.append(
            {
                "id": row["id"],
                "file_name": row["file_name"],
                "file_path": file_path,
                "file_type": row["file_type"],
                "location": row["location"],
                "preview": make_preview(row["content"], keyword),
                "modified_time": row["modified_time"],
            }
        )
    return results
