"""Сохранение и загрузка операционной выработки (JSON)."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any


@dataclass
class BarProgressRow:
    opening: int
    done: bool = False
    start_iso: str = ""
    end_iso: str = ""
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> BarProgressRow:
        return cls(
            opening=int(data.get("opening", 0)),
            done=bool(data.get("done", False)),
            start_iso=str(data.get("start_iso", "") or ""),
            end_iso=str(data.get("end_iso", "") or ""),
            note=str(data.get("note", "") or ""),
        )


@dataclass
class ShopFloorSession:
    version: int = 1
    project_name: str = ""
    project_cipher: str = ""
    spec_path: str = ""
    saved_at: str = ""
    rows: list[BarProgressRow] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "version": self.version,
            "project_name": self.project_name,
            "project_cipher": self.project_cipher,
            "spec_path": self.spec_path,
            "saved_at": self.saved_at,
            "rows": [r.to_dict() for r in self.rows],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ShopFloorSession:
        raw_rows = data.get("rows", [])
        rows: list[BarProgressRow] = []
        if isinstance(raw_rows, list):
            for item in raw_rows:
                if isinstance(item, dict):
                    rows.append(BarProgressRow.from_dict(item))
        return cls(
            version=int(data.get("version", 1)),
            project_name=str(data.get("project_name", "") or ""),
            project_cipher=str(data.get("project_cipher", "") or ""),
            spec_path=str(data.get("spec_path", "") or ""),
            saved_at=str(data.get("saved_at", "") or ""),
            rows=rows,
        )


def progress_by_opening(rows: list[BarProgressRow]) -> dict[int, BarProgressRow]:
    return {int(r.opening): r for r in rows if int(r.opening) > 0}


def merge_progress(
    layout_openings: list[int],
    stored: dict[int, BarProgressRow],
) -> list[BarProgressRow]:
    out: list[BarProgressRow] = []
    for oid in layout_openings:
        if oid in stored:
            out.append(stored[oid])
        else:
            out.append(BarProgressRow(opening=oid))
    return out


def duration_minutes(start_iso: str, end_iso: str) -> str:
    if not start_iso or not end_iso:
        return ""
    try:
        t0 = datetime.fromisoformat(start_iso)
        t1 = datetime.fromisoformat(end_iso)
        delta = t1 - t0
        mins = int(delta.total_seconds() // 60)
        if mins < 0:
            return ""
        h, m = divmod(mins, 60)
        if h:
            return f"{h} ч {m} мин"
        return f"{m} мин"
    except ValueError:
        return ""


def save_session(path: str | Path, session: ShopFloorSession) -> None:
    session.saved_at = datetime.now().isoformat(timespec="seconds")
    p = Path(path)
    p.write_text(
        json.dumps(session.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_session(path: str | Path) -> ShopFloorSession:
    p = Path(path)
    data = json.loads(p.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Некорректный формат JSON")
    return ShopFloorSession.from_dict(data)
