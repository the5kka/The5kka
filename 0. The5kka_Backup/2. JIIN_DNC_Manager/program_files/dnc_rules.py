import unicodedata
from pathlib import Path


def normalize_condition_filename(value: str) -> str:
    """조건 파일명 비교용 이름입니다. .txt만 조건 파일로 인정합니다."""
    text = unicodedata.normalize("NFC", str(value or "")).strip()
    if text.lower().endswith(".txt"):
        text = text[:-4]
    return text.casefold()


def search_condition_file_exact_txt(condition_name: str, source_folder: Path) -> list[Path]:
    """원본 폴더에서 .txt 조건 파일명 완전일치만 검색합니다."""
    if not source_folder.exists() or not source_folder.is_dir():
        return []
    normalized = normalize_condition_filename(condition_name)
    matches: list[Path] = []
    for file_path in source_folder.rglob("*.txt"):
        if not file_path.is_file():
            continue
        if normalize_condition_filename(file_path.name) == normalized:
            matches.append(file_path)
    return matches


def format_duplicate_condition_files(matches: list[Path]) -> str:
    """동일 조건 파일이 2개 이상일 때 표시할 현장용 문구입니다."""
    lines = ["동일한 DNC 파일이 2개 이상 발견되었습니다.", ""]
    for index, file_path in enumerate(matches, start=1):
        lines.append(f"{index}. {file_path}")
        lines.append("")
    lines.append(f"동일 조건 파일이 {len(matches)}개 입니다. 관리자 확인 바랍니다.")
    return "\n".join(lines)


def validate_process_paths(config: dict, process_name: str = "KCC PKG") -> tuple[bool, str]:
    """공정별 DNC 진행에 필요한 3개 경로를 확인합니다."""
    source_folders = config.get("source_dnc_folders", {}) if isinstance(config.get("source_dnc_folders"), dict) else {}
    source_path = source_folders.get(process_name) or config.get("source_dnc_folder", "")
    excel_path = config.get("excel_file", "")
    transfer_path = config.get("transfer_dnc_folder", "")

    checks = [
        ("DNC 조건 원본 폴더", source_path, "dir"),
        ("작업일보 파일", excel_path, "file"),
        ("DNC 전송 폴더", transfer_path, "dir"),
    ]
    missing: list[str] = []
    for label, raw_path, kind in checks:
        path_text = str(raw_path or "").strip()
        if not path_text:
            missing.append(label)
            continue
        path = Path(path_text)
        if kind == "file" and not path.exists():
            missing.append(label)
        elif kind == "dir":
            if not path.exists() or not path.is_dir():
                missing.append(label)
    if missing:
        return False, " / ".join(missing) + " 확인 필요"
    return True, ""
