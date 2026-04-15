from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook, load_workbook

from nordfox_raskroy.models import SpecRow

HEADER_ALIASES = {
    "№ п/п": "item_no",
    "Наименование": "module_name",
    "Тип профиля": "profile_code",
    "Длина": "length_mm",
    "Угол запила": "cut_angle",
    "Количество": "quantity",
    "QR-код": "qr",
}


def _find_header_row(ws, max_scan: int = 60) -> int:
    for r in range(1, max_scan + 1):
        for c in range(1, 15):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip() == "Наименование":
                return r
    raise ValueError("Не найдена строка заголовка с колонкой «Наименование»")


def _map_headers(ws, header_row: int) -> dict[str, int]:
    col_by_key: dict[str, int] = {}
    for c in range(1, 32):
        raw = ws.cell(header_row, c).value
        if raw is None:
            continue
        label = str(raw).strip()
        if label in HEADER_ALIASES:
            col_by_key[HEADER_ALIASES[label]] = c
    required = {"module_name", "profile_code", "length_mm", "cut_angle", "quantity"}
    missing = required - set(col_by_key)
    if missing:
        raise ValueError(f"Не хватает колонок: {missing}")
    return col_by_key


def _cell_str(ws, r: int, c: int | None) -> str | None:
    if c is None:
        return None
    v = ws.cell(r, c).value
    if v is None:
        return None
    return str(v).strip()


def _cell_int(ws, r: int, c: int) -> int:
    v = ws.cell(r, c).value
    if v is None:
        raise ValueError(f"Пустая ячейка (строка {r}, колонка {c})")
    if isinstance(v, bool):
        raise ValueError("Некорректное число")
    if isinstance(v, int):
        return int(v)
    if isinstance(v, float):
        return int(round(v))
    s = str(v).strip().replace(" ", "").replace(",", ".")
    if not s:
        raise ValueError("Пустое число")
    return int(round(float(s)))


def parse_specification(path: str | Path) -> list[SpecRow]:
    path = Path(path)
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hr = _find_header_row(ws)
    cols = _map_headers(ws, hr)

    item_col = cols.get("item_no")
    rows_out: list[SpecRow] = []
    last_item: int | None = None
    last_module = ""

    r = hr + 1
    while True:
        profile = _cell_str(ws, r, cols["profile_code"])
        length_cell = ws.cell(r, cols["length_mm"]).value
        if profile is None and length_cell is None:
            break
        if not profile or length_cell is None:
            raise ValueError(f"Строка {r}: задайте тип профиля и длину")

        raw_item = _cell_str(ws, r, item_col) if item_col else None
        if raw_item:
            m = re.match(r"^(\d+)$", raw_item)
            if not m:
                raise ValueError(f"Строка {r}: некорректный «№ п/п»: {raw_item!r}")
            last_item = int(m.group(1))
        elif last_item is None:
            raise ValueError(f"Строка {r}: отсутствует «№ п/п» в первой строке блока")

        mod = _cell_str(ws, r, cols["module_name"])
        if mod:
            last_module = mod
        if not last_module:
            raise ValueError(f"Строка {r}: не указано наименование модуля")

        length_mm = _cell_int(ws, r, cols["length_mm"])
        cut_angle = _cell_int(ws, r, cols["cut_angle"])
        qty = _cell_int(ws, r, cols["quantity"])
        if length_mm <= 0 or qty <= 0:
            raise ValueError(f"Строка {r}: длина и количество должны быть > 0")

        qr = _cell_str(ws, r, cols.get("qr")) if "qr" in cols else None

        rows_out.append(
            SpecRow(
                row_index=r,
                item_no=last_item,
                module_name=last_module.strip(),
                profile_code=profile.strip(),
                length_mm=length_mm,
                cut_angle=cut_angle,
                quantity=qty,
                qr=qr,
            )
        )
        r += 1

    if not rows_out:
        raise ValueError("Нет данных ниже заголовка")
    return rows_out


def write_spec_workbook(path: str | Path, rows: list[tuple]) -> None:
    """Служебная запись тестовой спецификации. rows: список кортежей по колонкам."""
    path = Path(path)
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    headers = list(HEADER_ALIASES.keys())
    for j, h in enumerate(headers, start=1):
        ws.cell(1, j, h)
    for i, row in enumerate(rows, start=2):
        for j, val in enumerate(row, start=1):
            ws.cell(i, j, val if val is not None else "")
    wb.save(path)
