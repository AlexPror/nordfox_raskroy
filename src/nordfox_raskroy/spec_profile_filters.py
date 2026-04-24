from __future__ import annotations

from nordfox_raskroy.models import SpecRow
from nordfox_raskroy.profile_codes import profile_label_for_code


def profile_filter_key(profile_code: str) -> str:
    code = (profile_code or "").strip()
    if not code:
        return ""
    series = profile_label_for_code(code)
    if series != "—":
        return series
    return code


def filter_rows_by_selected_profiles(
    rows_all: list[SpecRow],
    allowed_profiles: set[str],
) -> tuple[list[SpecRow], list[str]]:
    kept: list[SpecRow] = []
    warns: list[str] = []
    for row in rows_all:
        key = profile_filter_key(row.profile_code)
        if key in allowed_profiles:
            kept.append(row)
        else:
            warns.append(
                f"Строка {row.row_index}: «{row.profile_code}» ({key}) не включён в раскрой"
            )
    return kept, warns
