"""Сбор PartDemand из строк таблицы результата (после ручного редактирования)."""

from __future__ import annotations

from nordfox_raskroy.models import PartDemand
from nordfox_raskroy.profile_codes import PROFILE_DIGIT_TO_NAME, parse_profile_series_digit


def demands_from_cut_table_rows(
    rows: list[list[str]],
    allowed_digits: set[int],
) -> tuple[list[PartDemand] | None, str]:
    """
    rows: каждая строка — 8 ячеек данных (допускается 9-я «Масса» — игнорируется)
    [Модуль, Тип профиля, Серия, Длина, Угол, Источник, Заготовка, Остаток]
    Для расчёта: модуль, профиль, длина, угол (серия и пр. игнорируются).
    Два подреза в ячейке угла: «45/87».
    """
    if not rows:
        return None, "Таблица пуста"

    out: list[PartDemand] = []
    logical_row = 0
    for i, cells in enumerate(rows):
        rnum = i + 1
        if len(cells) < 5:
            return None, f"Строка {rnum}: недостаточно данных"
        if not any((c or "").strip() for c in (cells + [""] * 8)[:8]):
            continue
        module = (cells[0] or "").strip()
        profile = (cells[1] or "").strip()
        len_raw = (cells[3] or "").strip().replace(" ", "").replace(",", ".")
        ang_raw = (cells[4] or "").strip().replace(" ", "").replace(",", ".")

        if not module:
            return None, f"Строка {rnum}: пустой модуль"
        if not profile:
            return None, f"Строка {rnum}: пустой тип профиля"
        if not len_raw:
            return None, f"Строка {rnum}: пустая длина"

        digit = parse_profile_series_digit(profile)
        if digit is None:
            return (
                None,
                f"Строка {rnum}: «{profile}» — ожидается код СК-/СС-/Р- с цифрой 0–3",
            )
        if digit not in allowed_digits:
            return (
                None,
                f"Строка {rnum}: профиль {PROFILE_DIGIT_TO_NAME[digit]} не отмечен в выборе раскроя",
            )

        try:
            length_mm = int(round(float(len_raw)))
        except ValueError:
            return None, f"Строка {rnum}: не число в колонке «Длина»: {len_raw!r}"

        cut_angle_2: int | None = None
        if not ang_raw:
            angle = 90
        elif "/" in ang_raw:
            parts = [p.strip() for p in ang_raw.split("/", 1)]
            if len(parts) != 2:
                return None, f"Строка {rnum}: ожидались два угла через «/»: {ang_raw!r}"
            try:
                angle = int(round(float(parts[0])))
                cut_angle_2 = int(round(float(parts[1])))
            except ValueError:
                return None, f"Строка {rnum}: не число в колонке «Угол»: {ang_raw!r}"
        else:
            try:
                angle = int(round(float(ang_raw)))
            except ValueError:
                return None, f"Строка {rnum}: не число в колонке «Угол»: {ang_raw!r}"

        if length_mm <= 0:
            return None, f"Строка {rnum}: длина должна быть > 0"

        logical_row += 1
        out.append(
            PartDemand(
                spec_row_index=1000 + logical_row,
                module_name=module,
                profile_code=profile,
                length_mm=length_mm,
                cut_angle=angle,
                cut_angle_2=cut_angle_2,
            )
        )
    if not out:
        return None, "Нет ни одной непустой строки с данными"
    return out, ""
