"""
Справочник массы погонного метра (кг/м) для раскроя NordFox.

Источники:
- Профили каркаса Н20–Н23, DT, уголок L15 — из ``nordfox_specification`` (``materials_database.PROFILES_MASS_PER_METER``).
- Стандартный металлопрокат — типовые значения по ГОСТ (ориентиры для оценки; при необходимости уточнять по каталогу проката).

Цифра в коде СК-/СС-/Р- (0…3) соответствует сериям Н20…Н23 согласно ``profile_codes``.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Final

from nordfox_raskroy.profile_codes import PROFILE_DIGIT_TO_NAME, parse_profile_series_digit

_PROFILE_LIBRARY_PATH = Path(__file__).resolve().parents[2] / "profile_library.json"
_EDITABLE_PROFILE_CACHE: list[tuple[str, float]] | None = None

# --- NordFox: кг/м по цифре серии (как в nordfox_specification) ---
# digit 0 → Н20, 1 → Н21, 2 → Н22, 3 → Н23
SERIES_DIGIT_KG_PER_M: Final[dict[int, float]] = {
    0: 0.937,  # H20 / Н20 (актуализировано по согласованной таблице чертежей)
    1: 1.132,  # H21 / Н21
    2: 1.358,  # H22 / Н22
    3: 1.48,  # H23 / Н23
}

# Дополнительно по согласованной библиотеке профилей проекта (кг/м).
# Ключи должны быть в нормализованном виде (см. _norm_text): lower-case, "х" -> "x".
EXTRA_PROFILES_KG_PER_M: Final[dict[str, float]] = {
    "профиль t35": 0.669,
    "t35": 0.669,
    "бокс 80x40x3": 1.854,
    "профиль dt11": 1.016,
    "dt11": 1.016,
    "бокс120x60x3": 2.829,
    "бокс 120x60x3": 2.829,
    "профиль t21": 0.696,
    "t21": 0.696,
    "бокс 50x50x2": 1.041,
    "профиль h20": 0.937,
    "h20": 0.937,
    "профиль h50": 0.963,
    "h50": 0.963,
    "профиль ламели lz10": 0.588,
    "lz10": 0.588,
    "профиль h40": 1.138,
    "h40": 1.138,
    "профиль t20": 0.494,
    "t20": 0.494,
    "профиль t16": 0.686,
    "t16": 0.686,
    "профиль dt21": 1.017,
    "dt21": 1.017,
    "профиль dt11n": 1.019,
    "dt11n": 1.019,
    "профиль dt23": 1.274,
    "dt23": 1.274,
    "профиль l20": 0.324,
    "l20": 0.324,
    "профиль t11n": 0.428,
    "t11n": 0.428,
    "профиль h hat21": 1.295,
    "h hat21": 1.295,
    "профиль t22": 0.418,
    "t22": 0.418,
    "бокс100x50x2": 1.583,
    "бокс 100x50x2": 1.583,
    "профиль t30": 0.259,
    "t30": 0.259,
    "профиль dt22": 1.179,
    "dt22": 1.179,
    "бокс 200x50x4": 5.246,
    "профиль dt20": 0.781,
    "dt20": 0.781,
    "бокс 25x25x2": 0.499,
    "профиль h hat24": 1.717,
    "h hat24": 1.717,
    "профиль t15": 0.587,
    "t15": 0.587,
    "короб светильника lb10": 0.558,
    "lb10": 0.558,
    "бокс 80x50x2": 1.366,
    "профиль h hat23": 1.423,
    "h hat23": 1.423,
    "профиль h21": 1.132,
    "h21": 1.132,
    "профиль h hat22": 1.377,
    "h hat22": 1.377,
    "бокс 80x40x2": 1.257,
    "бокс150x50x4": 4.162,
    "бокс 150x50x4": 4.162,
    "профиль t hat20": 0.761,
    "t hat20": 0.761,
    "профиль h hat20": 1.08,
    "h hat20": 1.08,
    "профиль h22": 1.358,
    "h22": 1.358,
    "профиль h24": 1.691,
    "h24": 1.691,
    "профиль dt24": 1.49,
    "dt24": 1.49,
    "профиль h23 03.02.2026": 1.48,
    "профиль h23": 1.48,
    "h23": 1.48,
    "бокс 60x50x3": 1.691,
    "бокс 25x25x1,5": 0.382,
    "бокс 25x25x1.5": 0.382,
    "профиль t36": 0.783,
    "t36": 0.783,
    "профиль l11n": 0.288,
    "l11n": 0.288,
    "профиль l15": 0.42,
    "l15": 0.42,
    "профиль l16": 0.489,
    "l16": 0.489,
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


def _default_editable_profiles() -> list[tuple[str, float]]:
    out: list[tuple[str, float]] = []
    seen: set[str] = set()
    for key, kg in EXTRA_PROFILES_KG_PER_M.items():
        if not key.startswith("профиль "):
            continue
        display = " ".join(p.capitalize() for p in key.split())
        if display not in seen:
            seen.add(display)
            out.append((display, float(kg)))
    return sorted(out, key=lambda x: x[0].casefold())


def get_editable_profile_entries() -> list[tuple[str, float]]:
    """Список профилей для редактирования в UI (название, кг/м)."""
    global _EDITABLE_PROFILE_CACHE
    if _EDITABLE_PROFILE_CACHE is not None:
        return list(_EDITABLE_PROFILE_CACHE)
    if _PROFILE_LIBRARY_PATH.is_file():
        try:
            raw = json.loads(_PROFILE_LIBRARY_PATH.read_text(encoding="utf-8"))
            rows: list[tuple[str, float]] = []
            if isinstance(raw, list):
                for it in raw:
                    if not isinstance(it, dict):
                        continue
                    name = str(it.get("name", "")).strip()
                    try:
                        kg = float(it.get("kg_per_m", 0))
                    except Exception:
                        continue
                    if name and kg > 0:
                        rows.append((name, kg))
            if rows:
                _EDITABLE_PROFILE_CACHE = rows
                return list(rows)
        except Exception:
            pass
    _EDITABLE_PROFILE_CACHE = _default_editable_profiles()
    return list(_EDITABLE_PROFILE_CACHE)


def save_editable_profile_entries(entries: list[tuple[str, float]]) -> None:
    global _EDITABLE_PROFILE_CACHE
    rows = [(n.strip(), float(v)) for n, v in entries if n.strip() and float(v) > 0]
    _EDITABLE_PROFILE_CACHE = rows
    payload = [{"name": n, "kg_per_m": v} for n, v in rows]
    _PROFILE_LIBRARY_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _editable_profile_match_items() -> list[tuple[str, float, str]]:
    items: list[tuple[str, float, str]] = []
    for name, kg in get_editable_profile_entries():
        n = _norm_text(name)
        if not n:
            continue
        items.append((n, kg, name))
        if n.startswith("профиль "):
            items.append((n.removeprefix("профиль ").strip(), kg, name))
    uniq: dict[str, tuple[float, str]] = {}
    for key, kg, src in items:
        uniq[key] = (kg, src)
    return sorted([(k, v[0], v[1]) for k, v in uniq.items()], key=lambda x: len(x[0]), reverse=True)


def kg_per_meter_from_description(text: str) -> MassLookup:
    """
    Поиск кг/м по произвольной строке (наименование из спецификации).
    Сначала профили NordFox/DT/уголок по ключам, затем прокат.
    """
    n = _norm_text(text)
    if not n:
        return MassLookup(0.0, "unknown")

    for key, kg, src in _editable_profile_match_items():
        if key in n:
            return MassLookup(kg, f"editable_profile:{src}")

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
    for n, kg in get_editable_profile_entries():
        rows.append(("Профиль (редактируемый)", n, kg, "profile_library.json"))
    for k, v in _ROLLED_RAW:
        rows.append(("Металлопрокат (тип.)", k, v, "ГОСТ/каталог, оценка"))
    return rows
