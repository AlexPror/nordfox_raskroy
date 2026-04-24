"""
Шифр профиля в обозначении детали (NordFox, см. nordfox_module.profile_rules).

«СК-{d}-{длина}», «СС-{d}-{длина}», «Р-{d}-{длина}» — цифра d:
  0 → Н20, 1 → Н21, 2 → Н22, 3 → Н23
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from nordfox_raskroy.models import SpecRow

_PREFIX_RE = re.compile(
    r"^(?P<prefix>СК|СС|Р)-(?P<digit>\d+)-",
    re.IGNORECASE,
)

PROFILE_DIGIT_TO_NAME: dict[int, str] = {
    0: "Н20",
    1: "Н21",
    2: "Н22",
    3: "Н23",
}

NAME_TO_DIGIT: dict[str, int] = {v: k for k, v in PROFILE_DIGIT_TO_NAME.items()}


def parse_profile_series_digit(profile_code: str) -> int | None:
    """
    Цифра серии 0…3 для Н20–Н23. Иначе None (другой формат или цифра вне набора).
    """
    raw = (profile_code or "").strip()
    m = _PREFIX_RE.match(raw)
    if not m:
        return None
    d = int(m.group("digit"))
    if d in PROFILE_DIGIT_TO_NAME:
        return d
    return None


def profile_label_for_code(profile_code: str) -> str:
    d = parse_profile_series_digit(profile_code)
    if d is None:
        return "—"
    return PROFILE_DIGIT_TO_NAME[d]


def filter_spec_by_profiles(
    rows: list[SpecRow],
    allowed_profiles: set[int] | set[str],
) -> tuple[list[SpecRow], list[str]]:
    kept: list[SpecRow] = []
    warnings: list[str] = []
    as_text_mode = any(isinstance(x, str) for x in allowed_profiles)
    allowed_codes = {
        str(x).strip().upper() for x in allowed_profiles if isinstance(x, str) and str(x).strip()
    }
    for sr in rows:
        code_norm = (sr.profile_code or "").strip().upper()
        if as_text_mode:
            if code_norm not in allowed_codes:
                warnings.append(
                    f"Строка {sr.row_index}: «{sr.profile_code}» не включён в раскрой"
                )
                continue
            kept.append(sr)
            continue

        d = parse_profile_series_digit(sr.profile_code)
        if d is None:
            # Нестандартные профили (например L15, DT20 и т.п.) не отбрасываем:
            # фильтр Н20–Н23 применяется только к распознанным сериям 0..3.
            warnings.append(
                f"Строка {sr.row_index}: «{sr.profile_code}» оставлена в расчете "
                f"(нестандартный код, фильтр Н20–Н23 не применен)"
            )
            kept.append(sr)
            continue
        if d not in allowed_profiles:
            warnings.append(
                f"Строка {sr.row_index}: «{sr.profile_code}» — "
                f"{PROFILE_DIGIT_TO_NAME[d]} не включён в раскрой"
            )
            continue
        kept.append(sr)
    return kept, warnings
