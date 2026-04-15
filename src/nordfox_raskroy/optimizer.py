from __future__ import annotations

from collections import Counter
from typing import Sequence

from nordfox_raskroy.models import CutEvent, OptimizationResult, PartDemand, SpecRow


def spec_rows_to_demands(rows: list[SpecRow]) -> list[PartDemand]:
    out: list[PartDemand] = []
    for sr in rows:
        for _ in range(sr.quantity):
            out.append(
                PartDemand(
                    spec_row_index=sr.row_index,
                    module_name=sr.module_name,
                    profile_code=sr.profile_code,
                    length_mm=sr.length_mm,
                    cut_angle=sr.cut_angle,
                )
            )
    return out


def _pick_bar_length(need_mm: int, bar_lengths: list[int]) -> int:
    candidates = [b for b in bar_lengths if b >= need_mm]
    if not candidates:
        raise ValueError(
            f"Нет заготовки длиной ≥ {need_mm} мм. Доступны: {sorted(bar_lengths)}"
        )
    return min(candidates)


def _remove_one(scraps: list[int], value: int) -> None:
    try:
        scraps.remove(value)
    except ValueError as e:
        raise RuntimeError(f"internal: scrap {value} not in bin") from e


def optimize_cutting(
    demands: list[PartDemand],
    *,
    bar_lengths_mm: list[int],
    kerf_mm: int = 0,
    min_scrap_mm: int = 50,
    initial_scraps_mm: Sequence[int] | None = None,
) -> OptimizationResult:
    """
    Жадный best-fit по убыванию длины детали: сначала обрезки, затем новая заготовка.
    Остаток после реза возвращается в пул обрезков, если >= min_scrap_mm.

    initial_scraps_mm — куски со склада (каждый элемент = длина одного куска, мм).
    """
    if kerf_mm < 0 or min_scrap_mm < 0:
        raise ValueError("kerf_mm и min_scrap_mm не могут быть отрицательными")
    bars = sorted(set(bar_lengths_mm))
    if not bars:
        raise ValueError("Не заданы длины заготовок")

    need = max((d.length_mm for d in demands), default=0)
    if need > max(bars):
        raise ValueError(
            f"Есть деталь длиной {need} мм — больше максимальной заготовки {max(bars)} мм"
        )

    sorted_demands = sorted(demands, key=lambda d: d.length_mm, reverse=True)
    scraps: list[int] = []
    if initial_scraps_mm:
        for x in initial_scraps_mm:
            xi = int(x)
            if xi > 0:
                scraps.append(xi)
    result = OptimizationResult()

    for d in sorted_demands:
        cut_len = d.length_mm + kerf_mm
        best_piece: int | None = None
        best_waste: int | None = None

        for piece in scraps:
            if piece >= cut_len:
                waste = piece - cut_len
                if best_waste is None or waste < best_waste:
                    best_waste = waste
                    best_piece = piece

        if best_piece is not None:
            assert best_waste is not None
            _remove_one(scraps, best_piece)
            rem = best_waste
            if rem >= min_scrap_mm:
                scraps.append(rem)
            result.cuts.append(
                CutEvent(
                    demand=d,
                    stock_length_mm=best_piece,
                    stock_source="scrap",
                    remainder_mm=rem,
                    waste_mm=0 if rem >= min_scrap_mm else rem,
                )
            )
            continue

        bar = _pick_bar_length(cut_len, bars)
        rem = bar - cut_len
        result.bars_used[bar] = result.bars_used.get(bar, 0) + 1
        if rem >= min_scrap_mm:
            scraps.append(rem)
        result.cuts.append(
            CutEvent(
                demand=d,
                stock_length_mm=bar,
                stock_source="new_bar",
                remainder_mm=rem,
                waste_mm=0 if rem >= min_scrap_mm else rem,
            )
        )

    result.final_scraps_mm = sorted(scraps)
    return result


def sort_cuts_for_display(
    cuts: list[CutEvent], *, by_module: bool = True
) -> list[CutEvent]:
    if not by_module:
        return list(cuts)

    import re

    def module_key(name: str) -> tuple[int, str]:
        m = re.search(r"M\s*(\d+)", name, flags=re.IGNORECASE)
        if m:
            return (int(m.group(1)), name)
        return (10**9, name)

    return sorted(
        cuts,
        key=lambda c: (
            module_key(c.demand.module_name),
            c.demand.profile_code,
            c.demand.length_mm,
        ),
    )


def summarize(result: OptimizationResult) -> str:
    bars = Counter(result.bars_used)
    total_bars = sum(bars.values())
    lines: list[str] = [f"Всего заготовок: {total_bars}"]
    for L in sorted(bars):
        lines.append(f"  {L} мм × {bars[L]} шт.")
    lines.append(f"Остатки на складе обрезков (мм): {result.final_scraps_mm}")
    return "\n".join(lines)
