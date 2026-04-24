from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Literal

from nordfox_raskroy.models import CutEvent
from nordfox_raskroy.module_names import module_short_name
from nordfox_raskroy.profile_names import display_profile_name

AlbumMode = Literal["details", "joints"]


@dataclass(frozen=True, slots=True)
class AlbumPlanResult:
    rows: list[dict[str, object]]
    profile_names: set[str]


def build_album_plan_rows(
    cuts: list[CutEvent],
    *,
    mode: AlbumMode,
    offset_90_mm: int,
    offset_other_mm: int,
    kerf_mm: int,
) -> AlbumPlanResult:
    by_opening: dict[int, list[CutEvent]] = defaultdict(list)
    for c in cuts:
        if c.stock_opening_id <= 0:
            continue
        by_opening[c.stock_opening_id].append(c)
    rows: list[dict[str, object]] = []
    profile_names: set[str] = set()
    for opening in sorted(by_opening):
        g = by_opening[opening]
        if mode == "details":
            for c in g:
                d = c.demand
                prof = display_profile_name(d.profile_code)
                profile_names.add(prof)
                la = int(d.cut_angle)
                ra = int(d.cut_angle_2) if d.cut_angle_2 is not None else int(d.cut_angle)
                rows.append(
                    {
                        "kind": "detail",
                        "opening": opening,
                        "left_title": f"{module_short_name(d.module_name)} {d.profile_code}",
                        "right_title": "",
                        "left_profile_name": prof,
                        "right_profile_name": prof,
                        "left_right_angle": la,
                        "right_left_angle": ra,
                        "left_left_angle": la,
                        "right_right_angle": ra,
                        "left_tech_mm": offset_90_mm if la == 90 else offset_other_mm,
                        "right_tech_mm": offset_90_mm if ra == 90 else offset_other_mm,
                        "kerf_mm": kerf_mm,
                    }
                )
        else:
            for i in range(len(g) - 1):
                left = g[i].demand
                right = g[i + 1].demand
                left_prof = display_profile_name(left.profile_code)
                right_prof = display_profile_name(right.profile_code)
                profile_names.add(left_prof)
                profile_names.add(right_prof)
                rows.append(
                    {
                        "kind": "joint",
                        "opening": opening,
                        "left_title": f"{module_short_name(left.module_name)} {left.profile_code}",
                        "right_title": f"{module_short_name(right.module_name)} {right.profile_code}",
                        "left_profile_name": left_prof,
                        "right_profile_name": right_prof,
                        "left_right_angle": int(left.cut_angle_2) if left.cut_angle_2 is not None else int(left.cut_angle),
                        "right_left_angle": int(right.cut_angle),
                        "left_left_angle": int(left.cut_angle),
                        "right_right_angle": int(right.cut_angle_2) if right.cut_angle_2 is not None else int(right.cut_angle),
                        "left_tech_mm": offset_90_mm if int(left.cut_angle) == 90 else offset_other_mm,
                        "right_tech_mm": offset_90_mm if int(right.cut_angle) == 90 else offset_other_mm,
                        "kerf_mm": kerf_mm,
                    }
                )
    return AlbumPlanResult(rows=rows, profile_names=profile_names)
