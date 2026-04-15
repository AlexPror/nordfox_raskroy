"""
Справочник массы погонного метра (кг/м) для раскроя NordFox.

Источники:
- Профили каркаса Н20–Н23, DT, уголок L15 — из ``nordfox_specification`` (``materials_database.PROFILES_MASS_PER_METER``).
- Стандартный металлопрокат — типовые значения по ГОСТ (ориентиры для оценки; при необходимости уточнять по каталогу проката).

Цифра в коде СК-/СС-/Р- (0…3) соответствует сериям Н20…Н23 согласно ``profile_codes``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Final

from nordfox_raskroy.profile_codes import PROFILE_DIGIT_TO_NAME, parse_profile_series_digit

# --- NordFox: кг/м по цифре серии (как в nordfox_specification) ---
# digit 0 → Н20, 1 → Н21, 2 → Н22, 3 → Н23
SERIES_DIGIT_KG_PER_M: Final[dict[int, float]] = {
    0: 1.0315,  # H20 / Н20
    1: 1.206,  # H21 / Н21
    2: 1.2075,  # H22 / Н22 (экстраполяция в исходном проекте)
    3: 1.209,  # H23 / Н23
}

# Дополнительно из того же справочника (не привязаны к цифре СК/СС/Р)
EXTRA_PROFILES_KG_PER_M: Final[dict[str, float]] = {
    "dt20": 0.91,
    "dt21": 0.91,
    "l15": 0.49,
    "уголок": 0.49,
}

# --- Металлопрокат: подстрока в нижнем регистре → кг/м (первое совпадение по длине ключа) ---
# Ключи отсортированы при загрузке: длиннее раньше (чтобы «швеллер 16п» не перехватывал «швеллер»).
_ROLLED_RAW: list[tuple[str, float]] = [
    # Швеллер (ГОСТ 8240, ориентиры У/П)
    ("швеллер 30п", 36.2),
    ("швеллер 27п", 31.8),
    ("швеллер 24п", 26.7),
    ("швеллер 20п", 22.6),
    ("швеллер 18п", 18.1),
    ("швеллер 16п", 15.4),
    ("швеллер 14п", 12.3),
    ("швеллер 12п", 10.4),
    ("швеллер 10п", 8.59),
    ("швеллер 8п", 7.05),
    ("швеллер 6.5п", 5.9),
    ("швеллер 5п", 4.84),
    # Двутавр (ГОСТ 8239 — ориентиры)
    ("двутавр 40б1", 96.1),
    ("двутавр 30б1", 57.1),
    ("двутавр 24м", 27.9),
    ("двутавр 20б1", 25.4),
    ("двутавр 18м", 19.9),
    ("двутавр 16б1", 16.2),
    ("двутавр 14б1", 12.9),
    ("двутавр 12б1", 11.5),
    ("двутавр 10б1", 9.46),
    # Уголок равнополочный (ГОСТ 8509, ориентиры)
    ("уголок 125х125х10", 19.1),
    ("уголок 100х100х10", 15.1),
    ("уголок 90х90х8", 10.9),
    ("уголок 75х75х8", 8.99),
    ("уголок 63х63х6", 5.8),
    ("уголок 50х50х5", 3.77),
    # Труба профильная квадрат/прямоугольник (оценка, уточнять по толщине/ГОСТ)
    ("труба профильная 100х100х5", 14.7),
    ("труба профильная 80х80х4", 8.42),
    ("труба профильная 60х40х3", 4.07),
    ("труба профильная 40х40х3", 3.36),
    ("профильная труба 60х60х3", 5.19),
    # Полоса / квадрат (оценка)
    ("полоса 4х40", 1.26),
    ("полоса 5х50", 1.96),
    ("квадрат 10", 0.785),
]


def _rolled_sorted() -> list[tuple[str, float]]:
    return sorted(_ROLLED_RAW, key=lambda x: len(x[0]), reverse=True)


ROLLED_SUBSTRINGS_KG_M: Final[list[tuple[str, float]]] = _rolled_sorted()


@dataclass(frozen=True, slots=True)
class MassLookup:
    """Результат поиска кг/м."""

    kg_per_m: float
    source: str  # nordfox_series | extra_profile | rolled_steel | unknown


def kg_per_meter_nordfox_series(digit: int) -> float | None:
    """Масса погонного метра по цифре 0…3 (Н20…Н23)."""
    return SERIES_DIGIT_KG_PER_M.get(digit)


def kg_per_meter_from_profile_code(profile_code: str) -> MassLookup:
    """
    Для кодов вида СК-/СС-/Р- с цифрой 0–3 — берём таблицу NordFox.
    Иначе пробуем найти прокат по подстроке в названии.
    """
    raw = (profile_code or "").strip()
    if not raw:
        return MassLookup(0.0, "unknown")

    d = parse_profile_series_digit(raw)
    if d is not None:
        kg = SERIES_DIGIT_KG_PER_M[d]
        name = PROFILE_DIGIT_TO_NAME[d]
        return MassLookup(kg, f"nordfox_series:{name}")

    return kg_per_meter_from_description(raw)


def _norm_text(s: str) -> str:
    t = s.casefold().strip()
    t = re.sub(r"[\s_]+", " ", t)
    t = t.replace("х", "x").replace("×", "x")
    return t


def kg_per_meter_from_description(text: str) -> MassLookup:
    """
    Поиск кг/м по произвольной строке (наименование из спецификации).
    Сначала профили NordFox/DT/уголок по ключам, затем прокат.
    """
    n = _norm_text(text)
    if not n:
        return MassLookup(0.0, "unknown")

    for key, kg in EXTRA_PROFILES_KG_PER_M.items():
        if key in n:
            return MassLookup(kg, f"extra_profile:{key}")

    for key, kg in ROLLED_SUBSTRINGS_KG_M:
        if key in n:
            return MassLookup(kg, f"rolled_steel:{key}")

    # H/Н без кода СК — по номеру серии в тексте
    for m in re.finditer(r"[hн]\s*(\d{2})", n):
        num = int(m.group(1))
        if 20 <= num <= 23:
            digit = num - 20
            if digit in SERIES_DIGIT_KG_PER_M:
                kg = SERIES_DIGIT_KG_PER_M[digit]
                return MassLookup(kg, f"nordfox_series:H{num}")

    return MassLookup(0.0, "unknown")


def total_mass_kg(
    length_mm: float,
    kg_per_m: float,
    quantity: float = 1.0,
) -> float:
    """Масса, кг: L(м) × кг/м × количество."""
    return (length_mm / 1000.0) * kg_per_m * quantity


def row_mass_kg_display(
    profile_code: str,
    length_mm: float,
    quantity: float = 1.0,
) -> tuple[float | None, str]:
    """
    Масса строки раскроя для отображения и суммы.
    Возвращает (кг или None если справочник не сработал, текст ячейки).
    """
    m = kg_per_meter_from_profile_code(profile_code)
    if m.source == "unknown" or m.kg_per_m <= 0:
        return None, "—"
    kg = total_mass_kg(length_mm, m.kg_per_m, quantity)
    txt = f"{kg:.3f}".rstrip("0").rstrip(".")
    return kg, txt if txt else "0"


def materials_reference_rows() -> list[tuple[str, str, float, str]]:
    """
    Строки для таблицы: (категория, обозначение, кг/м, примечание).
    """
    rows: list[tuple[str, str, float, str]] = []
    for d in sorted(SERIES_DIGIT_KG_PER_M):
        name = PROFILE_DIGIT_TO_NAME[d]
        rows.append(
            (
                "Каркас NordFox (цифра в СК/СС/Р-)",
                f"{d} → {name}",
                SERIES_DIGIT_KG_PER_M[d],
                "nordfox_specification PROFILES_MASS_PER_METER",
            )
        )
    for k, v in sorted(EXTRA_PROFILES_KG_PER_M.items()):
        rows.append(("Профиль доп.", k, v, "nordfox_specification"))
    for k, v in _ROLLED_RAW:
        rows.append(("Металлопрокат (тип.)", k, v, "ГОСТ/каталог, оценка"))
    return rows
