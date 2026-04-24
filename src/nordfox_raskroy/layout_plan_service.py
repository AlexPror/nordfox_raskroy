from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Callable

from nordfox_raskroy.models import CutEvent
from nordfox_raskroy.module_names import module_short_name
from nordfox_raskroy.optimizer import format_cut_angles
from nordfox_raskroy.profile_names import display_profile_name


@dataclass(frozen=True, slots=True)
class LayoutPlanResult:
    rows: list[dict[str, object]]
    profile_names: set[str]


def build_layout_plan_rows(
    cuts: list[CutEvent],
    *,
    kerf_mm: int,
    offset_90_mm: int,
    offset_other_mm: int,
    opening_color_rgb: Callable[[int], tuple[int, int, int]],
) -> LayoutPlanResult:
    by_opening: dict[int, list[CutEvent]] = defaultdict(list)
    for c in cuts:
        if c.stock_opening_id <= 0:
            continue
        by_opening[c.stock_opening_id].append(c)
    if not by_opening:
        return LayoutPlanResult(rows=[], profile_names=set())

    rows: list[dict[str, object]] = []
    profile_names: set[str] = set()
    for opening in sorted(by_opening):
        group = by_opening[opening]
        first_new = next((c for c in group if c.stock_source == "new_bar"), None)
        if first_new is None:
            continue
        bar_len = int(first_new.stock_length_mm)
        segs: list[dict[str, object]] = []
        consumed = 0
        for c in group:
            d = c.demand
            profile_name = display_profile_name(d.profile_code)
            profile_names.add(profile_name)
            cut_angle = int(d.cut_angle)
            right_angle = int(d.cut_angle_2) if d.cut_angle_2 is not None else int(d.cut_angle)
            tech = int(offset_90_mm) if cut_angle == 90 else int(offset_other_mm)
            left_kerf = int(kerf_mm)
            segs.append(
                {
                    "kind": "tech",
                    "length_mm": float(tech),
                    "label": f"тех. {tech} мм",
                    "tech_mm": tech,
                }
            )
            consumed += tech
            if left_kerf > 0:
                segs.append(
                    {
                        "kind": "kerf",
                        "length_mm": float(left_kerf),
                        "label": "пропил",
                        "side": "leading",
                        "angle": cut_angle,
                        "kerf_mm": int(kerf_mm),
                    }
                )
                consumed += left_kerf
            segs.append(
                {
                    "kind": "profile",
                    "length_mm": int(d.length_mm),
                    "label": f"{module_short_name(d.module_name)} {d.profile_code}",
                    "module_name": d.module_name,
                    "profile_code": d.profile_code,
                    "part_length_mm": int(d.length_mm),
                    "cut_angles": format_cut_angles(d),
                    "left_angle": cut_angle,
                    "right_angle": right_angle,
                    "source_label": "Обрезок" if c.stock_source == "scrap" else "Новая",
                    "opening_color": opening_color_rgb(opening),
                    "profile_name": profile_name,
                    "is_scrap": c.stock_source == "scrap",
                    "origin": (
                        "из склада"
                        if c.stock_source == "scrap" and c.stock_opening_id == 0
                        else f"из прутка {c.stock_opening_id}"
                        if c.stock_source == "scrap"
                        else f"новый пруток {opening}"
                    ),
                }
            )
            consumed += int(d.length_mm)
            right_kerf = int(kerf_mm)
            if right_kerf > 0:
                segs.append(
                    {
                        "kind": "kerf",
                        "length_mm": right_kerf,
                        "label": "пропил",
                        "side": "trailing",
                        "angle": right_angle,
                        "kerf_mm": int(kerf_mm),
                    }
                )
                consumed += right_kerf
        remainder = max(bar_len - consumed, 0)
        rows.append(
            {
                "opening": opening,
                "bar_len": bar_len,
                "segments": segs,
                "remainder": remainder,
            }
        )
    return LayoutPlanResult(rows=rows, profile_names=profile_names)
