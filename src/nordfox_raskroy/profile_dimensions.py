"""
Габариты поперечного сечения профилей (высота × ширина, мм) для расчётов
(запас под скос пилы, визуализация и т.п.).

Источник — данные цеха; при необходимости дополняйте ``SECTION_BY_CANONICAL_NAME``.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Final

from nordfox_raskroy.profile_codes import PROFILE_DIGIT_TO_NAME, parse_profile_series_digit

# Ключи — каноническое имя в нижнем регистре: «н20», «dt20», «l15» …
SECTION_BY_CANONICAL_NAME: Final[dict[str, tuple[float, float]]] = {
    # Каркас NordFox (высота × ширина, мм)
    "н20": (60.0, 60.0),
    "н21": (82.9, 60.0),
    "н22": (90.0, 60.0),
    "н23": (100.0, 60.0),
    "н24": (120.0, 60.0),
    # Уголок
    "l15": (61.8, 40.0),
    "l16": (62.1, 40.0),
    "l20": (60.0, 30.9),
    # Тёплый / мостик (по обозначениям цеха)
    "dt20": (55.3, 60.0),
    "dt21": (83.8, 60.0),
}


@dataclass(frozen=True, slots=True)
class ProfileSectionMm:
    """Сечение: высота и ширина поперечника, мм."""

    height_mm: float
    width_mm: float


def _norm_text(s: str) -> str:
    t = (s or "").casefold().strip()
    t = re.sub(r"[\s_]+", " ", t)
    t = t.replace("х", "x").replace("×", "x")
    return t


def profile_section_mm(profile_code: str) -> ProfileSectionMm | None:
    """
    Внешние габариты сечения по коду строки (СК-/СС-/Р-, либо подстрока L15, DT20, Н22 …).

    Возвращает None, если профиль не занесён в справочник.
    """
    raw = (profile_code or "").strip()
    if not raw:
        return None

    d = parse_profile_series_digit(raw)
    if d is not None:
        name = PROFILE_DIGIT_TO_NAME[d].casefold()
        pair = SECTION_BY_CANONICAL_NAME.get(name)
        if pair is None:
            return None
        h, w = pair
        return ProfileSectionMm(h, w)

    n = _norm_text(raw)

    for key in sorted(SECTION_BY_CANONICAL_NAME, key=len, reverse=True):
        if key in n:
            h, w = SECTION_BY_CANONICAL_NAME[key]
            return ProfileSectionMm(h, w)

    for m in re.finditer(r"[hн]\s*(\d{2})", n):
        num = int(m.group(1))
        name = f"н{num}"
        pair = SECTION_BY_CANONICAL_NAME.get(name)
        if pair is not None:
            h, w = pair
            return ProfileSectionMm(h, w)

    return None


def profile_section_max_side_mm(profile_code: str) -> float | None:
    """max(высота, ширина) в мм или None, если профиль неизвестен."""
    sec = profile_section_mm(profile_code)
    if sec is None:
        return None
    return max(sec.height_mm, sec.width_mm)


def part_trailing_angle_deg(cut_angle: int, cut_angle_2: int | None) -> int:
    """Угол правого (заднего) подреза детали: второй угол, иначе как левый."""
    if cut_angle_2 is not None:
        return int(cut_angle_2)
    return int(cut_angle)


def extra_trailing_end_clearance_mm(profile_code: str, trailing_angle_deg: int) -> int:
    """
    Доп. осевая длина (мм), чтобы линия углового реза не «вылезла» за правый торец обрезка.

    Упрощённая модель: плоскость реза отклонена от перпендикуляра к оси прутка на
    φ = |90° − θ|, по габариту сечения h = max(высота, ширина) запас вдоль оси
    считаем как h·tan(φ). При θ≈90° запас 0. Для неизвестного профиля — 0.
    """
    h = profile_section_max_side_mm(profile_code)
    if h is None or h <= 0:
        return 0
    try:
        theta = float(trailing_angle_deg)
    except (TypeError, ValueError):
        return 0
    if abs(theta - 90.0) < 0.01:
        return 0
    phi_deg = abs(90.0 - theta)
    # Ограничиваем φ, чтобы tan не взрывался у 0°/180°.
    phi_deg = min(max(phi_deg, 0.5), 89.5)
    span = h * math.tan(math.radians(phi_deg))
    return max(0, int(math.ceil(span)))
