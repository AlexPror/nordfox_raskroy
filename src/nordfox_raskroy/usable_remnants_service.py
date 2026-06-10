"""Деловые остатки после раскроя (куски ≥ min_length_mm) для схемы-«колонны»."""

from __future__ import annotations

from nordfox_raskroy.models import CutEvent, StockScrapPiece
from nordfox_raskroy.profile_names import display_profile_name

REMNANT_SORT_MODES: list[tuple[str, str]] = [
    ("length_desc", "Длина, мм ↓"),
    ("length_asc", "Длина, мм ↑"),
    ("profile", "Тип профиля"),
    ("opening", "Пруток №"),
    ("as_stock", "Порядок расчёта"),
]

MIN_USABLE_REMNANT_MM = 100


def _profile_from_layout_row(row: dict[str, object]) -> tuple[str, str]:
    segs = row.get("segments", [])
    if not isinstance(segs, list):
        return "", ""
    for seg in segs:
        if not isinstance(seg, dict):
            continue
        if str(seg.get("kind", "")) != "profile":
            continue
        code = str(seg.get("profile_code", ""))
        name = str(seg.get("profile_name", "")) or display_profile_name(code)
        return code, name
    return "", ""


def _opening_profile_map(cuts: list[CutEvent]) -> dict[int, str]:
    out: dict[int, str] = {}
    for c in cuts:
        oid = int(c.stock_opening_id)
        if oid <= 0:
            continue
        if oid not in out:
            out[oid] = c.demand.profile_code
    return out


def build_usable_remnant_rows(
    *,
    layout_rows: list[dict[str, object]],
    cuts: list[CutEvent],
    scrap_pieces: list[StockScrapPiece],
    min_length_mm: int = MIN_USABLE_REMNANT_MM,
    sort_mode: str = "length_desc",
) -> list[dict[str, object]]:
    """
    Собирает деловые остатки: хвосты прутков из схемы и нетронутые куски со склада (opening_id=0).
    """
    lo = int(min_length_mm)
    seen_keys: set[tuple[int, int, str]] = set()
    items: list[dict[str, object]] = []
    opening_codes = _opening_profile_map(cuts)

    for row in layout_rows:
        rem = int(row.get("remainder", 0) or 0)
        if rem < lo:
            continue
        opening = int(row.get("opening", 0) or 0)
        code, name = _profile_from_layout_row(row)
        if not code and opening in opening_codes:
            code = opening_codes[opening]
            name = display_profile_name(code)
        key = (opening, rem, code or name)
        if key in seen_keys:
            continue
        seen_keys.add(key)
        items.append(
            {
                "kind": "remnant",
                "opening": opening,
                "length_mm": rem,
                "profile_code": code,
                "profile_name": name or code,
                "source": "хвост прутка",
            }
        )

    for piece in scrap_pieces:
        ln = int(piece.length_mm)
        if ln < lo:
            continue
        oid = int(piece.opening_id)
        if oid > 0 and any(
            int(it.get("opening", -1)) == oid and int(it.get("length_mm", 0)) == ln
            for it in items
        ):
            continue
        pk = piece.profile_key
        code = opening_codes.get(oid, "") if oid > 0 else ""
        if not code:
            code = pk if pk != "*" else "—"
        name = display_profile_name(code) if code not in ("—", "*", "") else pk
        key = (oid, ln, str(code))
        if key in seen_keys:
            continue
        seen_keys.add(key)
        items.append(
            {
                "kind": "remnant",
                "opening": oid,
                "length_mm": ln,
                "profile_code": code,
                "profile_name": name,
                "source": "склад" if oid == 0 else "хвост прутка",
            }
        )

    m = (sort_mode or "length_desc").strip().lower()

    def _sort_key(it: dict[str, object]) -> tuple:
        code = str(it.get("profile_code", ""))
        return (
            -int(it.get("length_mm", 0)),
            code,
            int(it.get("opening", 0)),
        )

    if m == "length_asc":
        items.sort(key=lambda it: (int(it["length_mm"]), str(it.get("profile_code", "")), int(it.get("opening", 0))))
    elif m == "profile":
        items.sort(
            key=lambda it: (
                str(it.get("profile_code", "")),
                -int(it.get("length_mm", 0)),
                int(it.get("opening", 0)),
            )
        )
    elif m == "opening":
        items.sort(
            key=lambda it: (
                int(it.get("opening", 0)),
                -int(it.get("length_mm", 0)),
            )
        )
    elif m == "as_stock":
        pass
    else:
        items.sort(key=_sort_key)

    max_len = max((int(it["length_mm"]) for it in items), default=1)
    for i, it in enumerate(items, start=1):
        it["remnant_no"] = i
        it["max_length_mm"] = max_len
    return items
