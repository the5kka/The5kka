from __future__ import annotations

import json
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from simtek_sps_rules import clean_text


STATUS_AVAILABLE = "AVAILABLE"
STATUS_REVIEW = "REVIEW"
STATUS_DUPLICATE_DISABLED = "DUPLICATE_DISABLED"
STATUS_MISSING_FILE = "MISSING_FILE"
STATUS_BLOCKED = "BLOCKED"

STATUS_LABELS = {
    STATUS_AVAILABLE: "사용 가능",
    STATUS_REVIEW: "관리자 확인",
    STATUS_DUPLICATE_DISABLED: "중복_사용중지",
    STATUS_MISSING_FILE: "적용 불가",
    STATUS_BLOCKED: "적용 불가",
}


def _normalize_key(value: object) -> str:
    return clean_text(value).casefold()


def _parse_shift(value: object) -> float | None:
    if value in (None, ""):
        return None
    return float(value)


def _latest_date(left: str, right: str) -> str:
    return max(clean_text(left), clean_text(right))


@dataclass(frozen=True)
class ProgramPolicy:
    name: str
    size_key: str
    shift: float | None
    status: str
    history_count: int = 0
    recent_date: str = ""
    verified_in_archive: bool = False
    note: str = ""


@dataclass(frozen=True)
class SpsProgramOption:
    program: str
    size_key: str
    shift: float | None
    history_count: int
    recent_date: str = ""
    status: str = STATUS_REVIEW
    status_label: str = "관리자 확인"
    file_verified: bool = False
    file_status: str = "실제 파일 미확인"
    note: str = ""

    @property
    def selectable(self) -> bool:
        return self.file_verified and self.status in {STATUS_AVAILABLE, STATUS_REVIEW}


class SpsProgramCatalog:
    def __init__(
        self,
        policies: Iterable[ProgramPolicy],
        aliases: dict[str, str] | None = None,
    ):
        self._policies = {_normalize_key(policy.name): policy for policy in policies}
        self._aliases = {
            _normalize_key(alias): clean_text(representative)
            for alias, representative in (aliases or {}).items()
            if clean_text(alias) and clean_text(representative)
        }

    @classmethod
    def from_json(cls, path: Path) -> "SpsProgramCatalog":
        source = Path(path)
        with source.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        policies = [
            ProgramPolicy(
                name=clean_text(row.get("name")),
                size_key=clean_text(row.get("size_key")),
                shift=_parse_shift(row.get("shift")),
                status=clean_text(row.get("status")).upper() or STATUS_REVIEW,
                history_count=int(row.get("history_count") or 0),
                recent_date=clean_text(row.get("recent_date")),
                verified_in_archive=bool(row.get("verified_in_archive")),
                note=clean_text(row.get("note")),
            )
            for row in payload.get("programs", [])
            if clean_text(row.get("name"))
        ]
        aliases = {
            clean_text(row.get("alias")): clean_text(row.get("representative"))
            for row in payload.get("aliases", [])
            if clean_text(row.get("alias")) and clean_text(row.get("representative"))
        }
        catalog = cls(policies, aliases)
        catalog._validate()
        return catalog

    def _validate(self) -> None:
        for alias_key, representative in self._aliases.items():
            if _normalize_key(representative) not in self._policies:
                raise ValueError(f"SPS P/G 별칭 대표 파일이 목록에 없습니다: {alias_key} -> {representative}")
        valid_statuses = {
            STATUS_AVAILABLE,
            STATUS_REVIEW,
            STATUS_DUPLICATE_DISABLED,
            STATUS_MISSING_FILE,
            STATUS_BLOCKED,
        }
        for policy in self._policies.values():
            if policy.status not in valid_statuses:
                raise ValueError(f"SPS P/G 상태값 오류: {policy.name} / {policy.status}")

    @property
    def policies(self) -> tuple[ProgramPolicy, ...]:
        return tuple(self._policies.values())

    def resolve_name(self, value: object) -> str:
        name = clean_text(value)
        if not name:
            return ""
        alias_target = self._aliases.get(_normalize_key(name))
        if alias_target:
            return alias_target
        policy = self._policies.get(_normalize_key(name))
        return policy.name if policy else name

    def get_policy(self, value: object) -> ProgramPolicy | None:
        resolved = self.resolve_name(value)
        return self._policies.get(_normalize_key(resolved))

    def build_options(
        self,
        program_counts: Counter[str] | dict[str, int] | None = None,
        recent_dates: dict[str, str] | None = None,
        registered_programs: Iterable[str] = (),
        source_folder: Path | None = None,
    ) -> list[SpsProgramOption]:
        counts: Counter[str] = Counter()
        dates: dict[str, str] = {}
        for raw_name, count in (program_counts or {}).items():
            resolved = self.resolve_name(raw_name)
            counts[resolved] += int(count)
        for raw_name, recent_date in (recent_dates or {}).items():
            resolved = self.resolve_name(raw_name)
            dates[resolved] = _latest_date(dates.get(resolved, ""), clean_text(recent_date))

        registered = {self.resolve_name(name) for name in registered_programs if clean_text(name)}
        live_files: dict[str, list[Path]] | None = None
        root = Path(source_folder) if source_folder else None
        if root and root.is_dir():
            live_files = {}
            for path in root.rglob("*.txt"):
                if path.is_file():
                    live_files.setdefault(path.stem.casefold(), []).append(path)

        options: list[SpsProgramOption] = []
        for policy in self._policies.values():
            history_count = max(policy.history_count, counts.get(policy.name, 0))
            recent_date = _latest_date(policy.recent_date, dates.get(policy.name, ""))
            status = policy.status
            note = policy.note

            if live_files is None:
                file_verified = policy.verified_in_archive
                file_status = "기준 파일 확인" if file_verified else "실제 파일 없음"
            else:
                matches = live_files.get(policy.name.casefold(), [])
                file_verified = len(matches) == 1
                if len(matches) == 1:
                    file_status = "실제 파일 확인"
                elif not matches:
                    file_status = "실제 파일 없음"
                    status = STATUS_MISSING_FILE
                else:
                    file_status = f"동일 파일 {len(matches)}개"
                    status = STATUS_BLOCKED

            if status in {STATUS_MISSING_FILE, STATUS_BLOCKED}:
                file_verified = False
            if policy.status == STATUS_DUPLICATE_DISABLED:
                status = STATUS_DUPLICATE_DISABLED
            if policy.name in registered and status == STATUS_REVIEW:
                note = (note + " / Rule 등록 확인").strip(" /")

            options.append(
                SpsProgramOption(
                    program=policy.name,
                    size_key=policy.size_key,
                    shift=policy.shift,
                    history_count=history_count,
                    recent_date=recent_date,
                    status=status,
                    status_label=STATUS_LABELS[status],
                    file_verified=file_verified,
                    file_status=file_status,
                    note=note,
                )
            )

        return sorted(options, key=_option_sort_key)


def _option_sort_key(option: SpsProgramOption) -> tuple[tuple[int, int], int, int, int, str]:
    try:
        width, height = (int(value) for value in option.size_key.split("-", 1))
    except (TypeError, ValueError):
        width, height = 9999, 9999
    status_order = {
        STATUS_AVAILABLE: 0,
        STATUS_REVIEW: 1,
        STATUS_DUPLICATE_DISABLED: 2,
        STATUS_MISSING_FILE: 3,
        STATUS_BLOCKED: 4,
    }
    return (
        (width, height),
        status_order.get(option.status, 9),
        0 if option.shift is None else 1,
        -option.history_count,
        option.program.casefold(),
    )


def condition_mismatch_reason(
    option: SpsProgramOption,
    expected_size: str,
    expected_shift: float | None,
) -> str:
    if not option.selectable:
        if option.status == STATUS_DUPLICATE_DISABLED:
            return "중복 사용 중지 조건입니다. 대표 P/G를 선택하세요."
        if not option.file_verified:
            return "실제 P/G 파일을 확인할 수 없습니다."
        return option.note or "현재 상태에서는 적용할 수 없습니다."
    if option.size_key != expected_size:
        return f"사이즈 불일치: 필수 {expected_size or '확인 불가'} / P/G {option.size_key or '확인 불가'}"
    if expected_shift is None and option.shift is not None:
        return f"SHIFT 불일치: 필수 없음 / P/G {option.shift:+g}"
    if expected_shift is not None and option.shift is None:
        return f"SHIFT 누락: 필수 {expected_shift:+g} / P/G 없음"
    if expected_shift is not None and option.shift is not None and abs(expected_shift - option.shift) > 0.001:
        return f"SHIFT 불일치: 필수 {expected_shift:+g} / P/G {option.shift:+g}"
    return ""
