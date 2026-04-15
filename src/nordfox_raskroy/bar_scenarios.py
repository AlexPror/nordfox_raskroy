"""Сравнение наборов длин заготовок для одной и той же спецификации."""

from __future__ import annotations

from dataclasses import dataclass

from nordfox_raskroy.models import OptimizationResult, PartDemand
from nordfox_raskroy.optimizer import optimize_cutting


@dataclass
class ScenarioOutcome:
    name: str
    bars_mm: tuple[int, ...]
    ok: bool
    error: str
    total_bars: int
    material_mm: int
    result: OptimizationResult | None


# Имя сценария → разрешённые длины новых заготовок (мм)
STANDARD_BAR_SCENARIOS: list[tuple[str, tuple[int, ...]]] = [
    ("Только 6000 мм", (6000,)),
    ("Только 7500 мм", (7500,)),
    ("Только 12000 мм", (12000,)),
    ("6000 + 7500 мм", (6000, 7500)),
    ("6000 + 12000 мм", (6000, 12000)),
    ("7500 + 12000 мм", (7500, 12000)),
    ("6000 + 7500 + 12000 мм", (6000, 7500, 12000)),
]


def evaluate_scenario(
    demands: list[PartDemand],
    name: str,
    bars_mm: tuple[int, ...],
    *,
    kerf_mm: int,
    min_scrap_mm: int,
    initial_scraps_mm: list[int] | None,
) -> ScenarioOutcome:
    try:
        r = optimize_cutting(
            demands,
            bar_lengths_mm=list(bars_mm),
            kerf_mm=kerf_mm,
            min_scrap_mm=min_scrap_mm,
            initial_scraps_mm=initial_scraps_mm,
        )
        total = sum(r.bars_used.values())
        mat = sum(L * c for L, c in r.bars_used.items())
        return ScenarioOutcome(name, bars_mm, True, "", total, mat, r)
    except Exception as e:  # noqa: BLE001
        return ScenarioOutcome(name, bars_mm, False, str(e), 0, 0, None)


def compare_bar_scenarios(
    demands: list[PartDemand],
    *,
    kerf_mm: int,
    min_scrap_mm: int,
    initial_scraps_mm: list[int] | None = None,
    scenarios: list[tuple[str, tuple[int, ...]]] | None = None,
) -> list[ScenarioOutcome]:
    sc = scenarios or STANDARD_BAR_SCENARIOS
    return [
        evaluate_scenario(
            demands,
            name,
            bars,
            kerf_mm=kerf_mm,
            min_scrap_mm=min_scrap_mm,
            initial_scraps_mm=initial_scraps_mm,
        )
        for name, bars in sc
    ]


def pick_recommended(outcomes: list[ScenarioOutcome]) -> ScenarioOutcome | None:
    """Минимум числа новых прутков, затем минимум суммарной длины новых прутков."""
    ok = [o for o in outcomes if o.ok and o.result is not None]
    if not ok:
        return None
    return min(ok, key=lambda o: (o.total_bars, o.material_mm))


def format_scenario_report(outcomes: list[ScenarioOutcome]) -> str:
    lines: list[str] = []
    for o in outcomes:
        if o.ok and o.result is not None:
            parts = [f"{L}×{o.result.bars_used.get(L, 0)}" for L in sorted(o.bars_mm)]
            use = ", ".join(parts)
            lines.append(
                f"• {o.name}: новых прутков {o.total_bars}, "
                f"сумма длин новых {o.material_mm} мм ({use})"
            )
        else:
            lines.append(f"• {o.name}: невозможно — {o.error}")
    rec = pick_recommended(outcomes)
    if rec:
        lines.append("")
        lines.append(
            f"Рекомендация: «{rec.name}» — меньше всего новых прутков "
            f"({rec.total_bars} шт., {rec.material_mm} мм суммарно)."
        )
    else:
        lines.append("")
        lines.append("Ни один вариант не сошёлся; проверьте длины деталей и склад.")
    return "\n".join(lines)
