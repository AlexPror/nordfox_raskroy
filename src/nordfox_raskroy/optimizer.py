from __future__ import annotations

from collections import Counter
import logging
from typing import Sequence

from nordfox_raskroy.models import CutEvent, OptimizationResult, PartDemand, SpecRow
from nordfox_raskroy.profile_codes import profile_label_for_code
logger = logging.getLogger("nordfox_raskroy.optimizer")


def angle_offset_mm(cut_angle: int) -> int:
    """Технологический отступ на одну сторону подреза: 90° -> 30 мм, иначе 50 мм."""
    return 30 if int(cut_angle) == 90 else 50


def total_angle_offset_mm(demand: PartDemand) -> int:
    """Технологический отступ только по первому (левому) углу."""
    return angle_offset_mm(demand.cut_angle)


def format_cut_angles(demand: PartDemand) -> str:
    """Текст для экспорта/таблицы: «90» или «45/87»."""
    if demand.cut_angle_2 is None:
        return str(int(demand.cut_angle))
    return f"{int(demand.cut_angle)}/{int(demand.cut_angle_2)}"


def demand_cut_length_mm(demand: PartDemand, kerf_mm: int = 0) -> int:
    """Полная длина съёма с заготовки для одной детали."""
    return int(demand.length_mm) + total_angle_offset_mm(demand) + int(kerf_mm) * 2


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
                    cut_angle_2=sr.cut_angle_2,
                )
            )
        ang_s = (
            f"{sr.cut_angle}/{sr.cut_angle_2}"
            if sr.cut_angle_2 is not None
            else str(sr.cut_angle)
        )
        logger.info(
            "spec_row expanded: row=%s module=%s profile=%s len=%d angle=%s qty=%d",
            sr.row_index,
            sr.module_name,
            sr.profile_code,
            sr.length_mm,
            ang_s,
            sr.quantity,
        )
    logger.info("spec_rows_to_demands done: demands=%d", len(out))
    return out


def _pick_bar_length(need_mm: int, bar_lengths: list[int]) -> int:
    candidates = [b for b in bar_lengths if b >= need_mm]
    if not candidates:
        raise ValueError(
            f"Нет заготовки длиной ≥ {need_mm} мм. Доступны: {sorted(bar_lengths)}"
        )
    return min(candidates)


def _remove_scrap(
    scraps: list[tuple[int, int, str]],
    piece: tuple[int, int, str],
) -> None:
    """Удалить один кусок (длина, партия прутка, профиль)."""
    try:
        scraps.remove(piece)
    except ValueError as e:
        raise RuntimeError(f"internal: scrap {piece} not in bin") from e


def _scrap_profile_key(profile_code: str) -> str:
    """
    Ключ совместимости обрезков.
    Для стандартных кодов используем серию (Н20..Н23), чтобы хвосты
    переиспользовались между СК/СС/Р одной и той же серии.
    """
    label = profile_label_for_code(profile_code)
    if label != "—":
        return label
    return profile_code.strip().upper()


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
    logger.info(
        "optimize_cutting start: demands=%d bars=%s kerf=%d min_scrap=%d initial_scraps=%d",
        len(demands),
        sorted(set(bar_lengths_mm)),
        kerf_mm,
        min_scrap_mm,
        len(initial_scraps_mm) if initial_scraps_mm else 0,
    )
    if kerf_mm < 0 or min_scrap_mm < 0:
        raise ValueError("kerf_mm и min_scrap_mm не могут быть отрицательными")
    bars = sorted(set(bar_lengths_mm))
    if not bars:
        raise ValueError("Не заданы длины заготовок")

    need = max((demand_cut_length_mm(d, kerf_mm) for d in demands), default=0)
    if need > max(bars):
        raise ValueError(
            "Есть деталь с учетом технологического отступа/пропила длиной "
            f"{need} мм — больше максимальной заготовки {max(bars)} мм"
        )

    sorted_demands = sorted(
        demands,
        key=lambda d: demand_cut_length_mm(d, kerf_mm),
        reverse=True,
    )
    # (длина мм, stock_opening_id, ключ профиля): 0 — склад до расчёта; новый пруток получает 1, 2, …
    scraps: list[tuple[int, int, str]] = []
    if initial_scraps_mm:
        for x in initial_scraps_mm:
            xi = int(x)
            if xi > 0:
                scraps.append((xi, 0, "*"))
    result = OptimizationResult()
    next_opening_id = 0

    for d in sorted_demands:
        offset_mm = total_angle_offset_mm(d)
        cut_len = demand_cut_length_mm(d, kerf_mm)
        prof_key = _scrap_profile_key(d.profile_code)
        logger.info(
            "demand processing: row=%s module=%s profile=%s key=%s len=%d angle=%s offset=%d cut_len=%d scraps=%d",
            d.spec_row_index,
            d.module_name,
            d.profile_code,
            prof_key,
            d.length_mm,
            format_cut_angles(d),
            offset_mm,
            cut_len,
            len(scraps),
        )
        best: tuple[int, int, str] | None = None
        best_waste: int | None = None

        for length_mm, oid, prof in scraps:
            if prof not in ("*", prof_key):
                continue
            if length_mm >= cut_len:
                waste = length_mm - cut_len
                if best_waste is None or waste < best_waste or (
                    waste == best_waste and best is not None and oid < best[1]
                ):
                    best_waste = waste
                    best = (length_mm, oid, prof)

        if best is not None:
            assert best_waste is not None
            piece_len, opening_id, source_profile = best
            _remove_scrap(scraps, best)
            rem = best_waste
            if rem >= min_scrap_mm:
                scraps.append((rem, opening_id, prof_key))
            logger.info(
                "cut from scrap: profile=%s key=%s len=%d source=%d source_key=%s rem=%d rem_saved=%s opening=%d",
                d.profile_code,
                prof_key,
                d.length_mm,
                piece_len,
                source_profile,
                rem,
                rem >= min_scrap_mm,
                opening_id,
            )
            result.cuts.append(
                CutEvent(
                    demand=d,
                    stock_length_mm=piece_len,
                    stock_source="scrap",
                    remainder_mm=rem,
                    waste_mm=0 if rem >= min_scrap_mm else rem,
                    stock_opening_id=opening_id,
                )
            )
            continue

        next_opening_id += 1
        opening_id = next_opening_id
        bar = _pick_bar_length(cut_len, bars)
        rem = bar - cut_len
        result.bars_used[bar] = result.bars_used.get(bar, 0) + 1
        if rem >= min_scrap_mm:
            scraps.append((rem, opening_id, prof_key))
        logger.info(
            "cut from new bar: profile=%s len=%d bar=%d rem=%d rem_saved=%s bar_count=%d opening=%d",
            d.profile_code,
            d.length_mm,
            bar,
            rem,
            rem >= min_scrap_mm,
            result.bars_used[bar],
            opening_id,
        )
        result.cuts.append(
            CutEvent(
                demand=d,
                stock_length_mm=bar,
                stock_source="new_bar",
                remainder_mm=rem,
                waste_mm=0 if rem >= min_scrap_mm else rem,
                stock_opening_id=opening_id,
            )
        )

    result.cuts.sort(key=lambda c: c.stock_opening_id)
    result.final_scraps_mm = sorted(length for length, _oid, _prof in scraps)
    logger.info(
        "optimize_cutting done: cuts=%d new_bars=%d final_scraps=%d",
        len(result.cuts),
        sum(result.bars_used.values()),
        len(result.final_scraps_mm),
    )
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
