"""
Стабильные цвета по модулю (один модуль — один базовый оттенок).
Используется в Qt, Excel и PDF.
"""

from __future__ import annotations

import re

# Пастельная палитра (RGB), хорошо читается с чёрным текстом
_MODULE_PALETTE: list[tuple[int, int, int]] = [
    (219, 234, 254),  # blue-100
    (220, 252, 231),  # green-100
    (254, 249, 195),  # yellow-100
    (255, 228, 230),  # rose-100
    (237, 233, 254),  # violet-100
    (207, 250, 254),  # cyan-100
    (255, 237, 213),  # orange-100
    (226, 232, 240),  # slate-200
    (209, 250, 229),  # emerald-100
    (243, 232, 255),  # purple-100
    (254, 226, 226),  # red-100
    (224, 242, 254),  # sky-100
    (236, 253, 245),  # green-50
    (255, 247, 237),  # orange-50
    (238, 242, 255),  # indigo-50
    (240, 253, 250),  # teal-50
]


def module_palette_index(module_name: str) -> int:
    name = (module_name or "").strip()
    m = re.search(r"M\s*(\d+)", name, flags=re.IGNORECASE)
    if m:
        return (int(m.group(1)) - 1) % len(_MODULE_PALETTE)
    return abs(hash(name)) % len(_MODULE_PALETTE)


def module_base_rgb(module_name: str) -> tuple[int, int, int]:
    return _MODULE_PALETTE[module_palette_index(module_name)]


def module_row_rgb(module_name: str, *, is_scrap: bool) -> tuple[int, int, int]:
    """Строка результата: базовый цвет модуля; обрезок — чуть темнее."""
    r, g, b = module_base_rgb(module_name)
    if is_scrap:
        factor = 0.82
        return int(r * factor), int(g * factor), int(b * factor)
    return r, g, b


def rgb_to_openpyxl_argb(rgb: tuple[int, int, int]) -> str:
    """openpyxl: ARGB без решётки."""
    r, g, b = rgb
    return f"FF{r:02X}{g:02X}{b:02X}"


def rgb_to_pdf_hex(rgb: tuple[int, int, int]) -> str:
    """reportlab HexColor: #rrggbb"""
    r, g, b = rgb
    return f"#{r:02x}{g:02x}{b:02x}"


def rgb_to_qcolor_tuple(rgb: tuple[int, int, int]) -> tuple[int, int, int]:
    return rgb
