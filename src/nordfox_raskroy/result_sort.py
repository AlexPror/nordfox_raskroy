"""Сортировка списка резов для таблицы и экспорта."""

from __future__ import annotations

import re
from nordfox_raskroy.models import CutEvent
from nordfox_raskroy.optimizer import sort_cuts_for_display
from nordfox_raskroy.profile_codes import parse_profile_series_digit

# id для QComboBox.currentData()
SORT_MODES: list[tuple[str, str]] = [
    ("opening", "Пруток № (1-й → 2-й → …)"),
    (
        "operator",
        "Оператор: профиль → обрезок → заготовка",
    ),
    ("module", "Модуль (M1 → …)"),
    ("module_length_profile", "Модуль → длина → тип профиля"),
    ("profile", "Тип профиля (код)"),
    ("length_desc", "Длина детали ↓"),
    ("length_asc", "Длина детали ↑"),
    ("series", "Серия (Н20 → Н23)"),
    ("source", "Источник (обрезок → новая)"),
    ("stock_desc", "Заготовка мм ↓"),
    ("remainder_desc", "Остаток мм ↓"),
    ("as_calculated", "Порядок расчёта"),
]


def _module_key(name: str) -> tuple[int, str]:
    m = re.search(r"M\s*(\d+)", name, flags=re.IGNORECASE)
    if m:
        return (int(m.group(1)), name)
    return (10**9, name)


def _series_key(profile_code: str) -> int:
    d = parse_profile_series_digit(profile_code)
    return d if d is not None else 99


def _source_key(source: str) -> int:
    return 0 if source == "scrap" else 1


def sort_cuts(cuts: list[CutEvent], mode: str) -> list[CutEvent]:
    """Вернуть новый список в выбранном порядке."""
    base = list(cuts)
    m = (mode or "opening").strip().lower()

    if m == "opening":
        return sorted(base, key=lambda c: c.stock_opening_id)
    if m == "operator":
        # Меньше переключений: один тип профиля подряд; сначала обрезки со стеллажа,
        # потом новые прутки; куски одной длины рядом.
        return sorted(
            base,
            key=lambda c: (
                c.demand.profile_code.upper(),
                _source_key(c.stock_source),
                -c.stock_length_mm,
                _module_key(c.demand.module_name),
                -c.demand.length_mm,
            ),
        )
    if m == "module":
        return sort_cuts_for_display(base, by_module=True)
    if m == "module_length_profile":
        return sorted(
            base,
            key=lambda c: (
                _module_key(c.demand.module_name),
                c.demand.length_mm,
                c.demand.profile_code.upper(),
            ),
        )
    if m == "as_calculated":
        return base

    if m == "profile":
        return sorted(
            base,
            key=lambda c: (
                c.demand.profile_code.upper(),
                _module_key(c.demand.module_name),
                -c.demand.length_mm,
            ),
        )
    if m == "length_desc":
        return sorted(
            base,
            key=lambda c: (
                -c.demand.length_mm,
                _module_key(c.demand.module_name),
                c.demand.profile_code,
            ),
        )
    if m == "length_asc":
        return sorted(
            base,
            key=lambda c: (
                c.demand.length_mm,
                _module_key(c.demand.module_name),
                c.demand.profile_code,
            ),
        )
    if m == "series":
        return sorted(
            base,
            key=lambda c: (
                _series_key(c.demand.profile_code),
                _module_key(c.demand.module_name),
                c.demand.profile_code,
                -c.demand.length_mm,
            ),
        )
    if m == "source":
        return sorted(
            base,
            key=lambda c: (
                _source_key(c.stock_source),
                _module_key(c.demand.module_name),
                c.demand.profile_code,
            ),
        )
    if m == "stock_desc":
        return sorted(
            base,
            key=lambda c: (
                -c.stock_length_mm,
                _module_key(c.demand.module_name),
                c.demand.profile_code,
            ),
        )
    if m == "remainder_desc":
        return sorted(
            base,
            key=lambda c: (
                -c.remainder_mm,
                _module_key(c.demand.module_name),
                c.demand.profile_code,
            ),
        )

    return sort_cuts_for_display(base, by_module=True)
