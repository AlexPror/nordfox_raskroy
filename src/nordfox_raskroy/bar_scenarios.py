"""Сравнение наборов длин заготовок для одной и той же спецификации."""

from __future__ import annotations

from dataclasses import dataclass
import logging
from typing import Literal

from nordfox_raskroy.models import OptimizationResult, PartDemand
from nordfox_raskroy.optimizer import optimize_cutting

OptimizationMode = Literal["bars_first", "waste_first", "material_first", "balanced"]
logger = logging.getLogger("nordfox_raskroy.bar_scenarios")


@dataclass
class ScenarioOutcome:
    name: str
    bars_mm: tuple[int, ...]
    ok: bool
    error: str
    total_bars: int
    material_mm: int
    used_mm: int
    waste_mm: int
    waste_pct: float
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
    offset_90_mm: int,
    offset_other_mm: int,
    min_scrap_mm: int,
    initial_scraps_mm: list[int] | None,
) -> ScenarioOutcome:
    try:
        logger.info("scenario start: %s bars=%s", name, list(bars_mm))
        r = optimize_cutting(
            demands,
            bar_lengths_mm=list(bars_mm),
            kerf_mm=kerf_mm,
            offset_90_mm=offset_90_mm,
            offset_other_mm=offset_other_mm,
            min_scrap_mm=min_scrap_mm,
            initial_scraps_mm=initial_scraps_mm,
        )
        total = sum(r.bars_used.values())
        mat = sum(L * c for L, c in r.bars_used.items())
        used = sum(c.stock_length_mm - c.remainder_mm for c in r.cuts)
        waste = max(mat - used, 0)
        waste_pct = (waste / mat * 100.0) if mat > 0 else 0.0
        logger.info(
            "scenario done: %s bars=%d material=%d used=%d waste=%d waste_pct=%.3f",
            name,
            total,
            mat,
            used,
            waste,
            waste_pct,
        )
        return ScenarioOutcome(name, bars_mm, True, "", total, mat, used, waste, waste_pct, r)
    except Exception as e:  # noqa: BLE001
        logger.exception("scenario failed: %s", name)
        return ScenarioOutcome(name, bars_mm, False, str(e), 0, 0, 0, 0, 0.0, None)


def compare_bar_scenarios(
    demands: list[PartDemand],
    *,
    kerf_mm: int,
    offset_90_mm: int = 30,
    offset_other_mm: int = 50,
    min_scrap_mm: int,
    initial_scraps_mm: list[int] | None = None,
    scenarios: list[tuple[str, tuple[int, ...]]] | None = None,
) -> list[ScenarioOutcome]:
    logger.info(
        "compare_bar_scenarios: demands=%d kerf=%d min_scrap=%d with_initial_scrap=%s",
        len(demands),
        kerf_mm,
        min_scrap_mm,
        bool(initial_scraps_mm),
    )
    sc = scenarios or STANDARD_BAR_SCENARIOS
    return [
        evaluate_scenario(
            demands,
            name,
            bars,
            kerf_mm=kerf_mm,
            offset_90_mm=offset_90_mm,
            offset_other_mm=offset_other_mm,
            min_scrap_mm=min_scrap_mm,
            initial_scraps_mm=initial_scraps_mm,
        )
        for name, bars in sc
    ]


def _norm(val: float, lo: float, hi: float) -> float:
    if hi <= lo:
        return 0.0
    return (val - lo) / (hi - lo)


def pick_recommended(
    outcomes: list[ScenarioOutcome],
    mode: OptimizationMode = "bars_first",
) -> ScenarioOutcome | None:
    """Выбор лучшего сценария в выбранном режиме оптимизации."""
    ok = [o for o in outcomes if o.ok and o.result is not None]
    if not ok:
        logger.warning("pick_recommended: no successful scenarios")
        return None

    m = (mode or "bars_first").strip().lower()
    logger.info("pick_recommended: mode=%s candidates=%d", m, len(ok))
    if m == "waste_first":
        best = min(ok, key=lambda o: (o.waste_pct, o.total_bars, o.material_mm))
        logger.info("pick_recommended result: %s", best.name)
        return best
    if m == "material_first":
        best = min(ok, key=lambda o: (o.material_mm, o.total_bars, o.waste_pct))
        logger.info("pick_recommended result: %s", best.name)
        return best
    if m == "balanced":
        bars_min = min(o.total_bars for o in ok)
        bars_max = max(o.total_bars for o in ok)
        waste_min = min(o.waste_pct for o in ok)
        waste_max = max(o.waste_pct for o in ok)
        best = min(
            ok,
            key=lambda o: (
                0.6 * _norm(o.waste_pct, waste_min, waste_max)
                + 0.4 * _norm(float(o.total_bars), float(bars_min), float(bars_max)),
                o.material_mm,
            ),
        )
        logger.info("pick_recommended result: %s", best.name)
        return best
    # bars_first
    best = min(ok, key=lambda o: (o.total_bars, o.material_mm, o.waste_pct))
    logger.info("pick_recommended result: %s", best.name)
    return best


def _mode_caption(mode: OptimizationMode) -> str:
    if mode == "waste_first":
        return "минимум % отходов"
    if mode == "material_first":
        return "минимум суммарной длины новых заготовок"
    if mode == "balanced":
        return "сбалансированный (60% отходы, 40% число прутков)"
    return "минимум числа новых прутков"


def format_scenario_report(
    outcomes: list[ScenarioOutcome],
    mode: OptimizationMode = "bars_first",
) -> str:
    lines: list[str] = []
    lines.append(f"Режим оптимизации: {_mode_caption(mode)}")
    lines.append("")
    for o in outcomes:
        if o.ok and o.result is not None:
            parts = [f"{L}×{o.result.bars_used.get(L, 0)}" for L in sorted(o.bars_mm)]
            use = ", ".join(parts)
            lines.append(
                f"• {o.name}: новых прутков {o.total_bars}, "
                f"сумма длин новых {o.material_mm} мм ({use}), "
                f"отходы {o.waste_pct:.1f}%"
            )
        else:
            lines.append(f"• {o.name}: невозможно — {o.error}")
    rec = pick_recommended(outcomes, mode=mode)
    if rec:
        lines.append("")
        reason = _mode_caption(mode)
        lines.append(f"Рекомендация: «{rec.name}» — {reason}.")
        lines.append(
            f"Показатели: {rec.total_bars} шт., {rec.material_mm} мм, отходы {rec.waste_pct:.1f}%."
        )
    else:
        lines.append("")
        lines.append("Ни один вариант не сошёлся; проверьте длины деталей и склад.")
    return "\n".join(lines)
