from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal


@dataclass(frozen=True, slots=True)
class SpecRow:
    """Одна строка спецификации после нормализации (после forward-fill)."""

    row_index: int
    item_no: int | None
    module_name: str
    profile_code: str
    length_mm: int
    cut_angle: int
    quantity: int
    qr: str | None = None
    cut_angle_2: int | None = None  # второй подрез (другая сторона), иначе один угол как раньше


@dataclass(frozen=True, slots=True)
class PartDemand:
    """Единичная потребность в отрезке (развёртка quantity)."""

    spec_row_index: int
    module_name: str
    profile_code: str
    length_mm: int
    cut_angle: int
    cut_angle_2: int | None = None


@dataclass(frozen=True, slots=True)
class CutEvent:
    """Один факт резки с заготовки или обрезка."""

    demand: PartDemand
    stock_length_mm: int
    stock_source: Literal["new_bar", "scrap"]
    remainder_mm: int
    waste_mm: int
    #: 0 — обрезок со склада до расчёта; 1+ — N-й открытый пруток и все резы с его хвостов.
    stock_opening_id: int = 0


@dataclass
class OptimizationResult:
    cuts: list[CutEvent] = field(default_factory=list)
    final_scraps_mm: list[int] = field(default_factory=list)
    bars_used: dict[int, int] = field(default_factory=dict)  # length -> count
