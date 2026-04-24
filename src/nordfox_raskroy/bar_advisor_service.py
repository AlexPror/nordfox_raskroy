from __future__ import annotations

from dataclasses import dataclass

from nordfox_raskroy.bar_scenarios import (
    ScenarioOutcome,
    compare_bar_scenarios,
    format_scenario_report,
    pick_recommended,
)
from nordfox_raskroy.models import PartDemand
from nordfox_raskroy.optimizer import demand_cut_length_mm


@dataclass(frozen=True, slots=True)
class BarAdvisorResult:
    outcomes: list[ScenarioOutcome]
    report_text: str
    recommended_bars_mm: tuple[int, ...] | None
    recommended_length_mm: int | None
    recommended_name: str | None
    mode: str


def run_bar_advisor(
    demands: list[PartDemand],
    *,
    kerf_mm: int,
    offset_90_mm: int,
    offset_other_mm: int,
    base_len_mm: int,
    standard_mode: bool,
    initial_scraps_mm: list[int] | None,
    min_scrap_mm: int = 0,
    mode: str = "waste_first",
) -> BarAdvisorResult:
    if base_len_mm <= 0:
        raise ValueError("Базовая длина должна быть > 0")
    if base_len_mm > 12000:
        raise ValueError("Базовая длина не должна превышать 12000 мм")

    required_max = max(
        (
            demand_cut_length_mm(
                d,
                kerf_mm,
                offset_90_mm=offset_90_mm,
                offset_other_mm=offset_other_mm,
            )
            for d in demands
        ),
        default=0,
    )
    if required_max > 12000:
        raise ValueError(
            "Есть детали, требующие заготовку больше 12000 мм "
            f"(максимум требуется: {required_max} мм)"
        )

    initial_scraps = initial_scraps_mm if initial_scraps_mm else None
    if standard_mode:
        scan_from = ((required_max + 49) // 50) * 50
        candidates = set(range(scan_from, 12001, 50))
        if base_len_mm % 50 == 0:
            candidates.add(base_len_mm)
        candidates = {x for x in candidates if scan_from <= x <= 12000}
        scenarios = [(f"Только {b} мм", (b,)) for b in sorted(candidates)]
        outcomes = compare_bar_scenarios(
            demands,
            kerf_mm=kerf_mm,
            offset_90_mm=offset_90_mm,
            offset_other_mm=offset_other_mm,
            min_scrap_mm=min_scrap_mm,
            initial_scraps_mm=initial_scraps,
            scenarios=scenarios,
        )
    else:
        coarse_from = ((required_max + 49) // 50) * 50
        coarse_candidates = set(range(coarse_from, 12001, 50))
        if base_len_mm % 50 == 0 and base_len_mm >= coarse_from:
            coarse_candidates.add(base_len_mm)
        coarse_scenarios = [(f"Только {b} мм", (b,)) for b in sorted(coarse_candidates)]
        coarse_outcomes = compare_bar_scenarios(
            demands,
            kerf_mm=kerf_mm,
            offset_90_mm=offset_90_mm,
            offset_other_mm=offset_other_mm,
            min_scrap_mm=min_scrap_mm,
            initial_scraps_mm=initial_scraps,
            scenarios=coarse_scenarios,
        )
        coarse_ok = [o for o in coarse_outcomes if o.ok and o.result is not None]
        if coarse_ok:
            top = sorted(
                coarse_ok,
                key=lambda o: (o.waste_pct, o.total_bars, o.material_mm),
            )[:5]
            refine_candidates: set[int] = {required_max, min(12000, base_len_mm)}
            for o in top:
                center = o.bars_mm[0]
                lo = max(required_max, center - 60)
                hi = min(12000, center + 60)
                refine_candidates.update(range(lo, hi + 1))
            refine_scenarios = [(f"Только {b} мм", (b,)) for b in sorted(refine_candidates)]
            outcomes = compare_bar_scenarios(
                demands,
                kerf_mm=kerf_mm,
                offset_90_mm=offset_90_mm,
                offset_other_mm=offset_other_mm,
                min_scrap_mm=min_scrap_mm,
                initial_scraps_mm=initial_scraps,
                scenarios=refine_scenarios,
            )
        else:
            outcomes = coarse_outcomes

    rec = pick_recommended(outcomes, mode=mode)
    return BarAdvisorResult(
        outcomes=outcomes,
        report_text=format_scenario_report(outcomes, mode=mode),
        recommended_bars_mm=rec.bars_mm if rec else None,
        recommended_length_mm=rec.bars_mm[0] if rec and rec.bars_mm else None,
        recommended_name=rec.name if rec else None,
        mode=mode,
    )
