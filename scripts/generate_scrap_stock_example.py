"""Пример файла склада обрезков для тестов."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    out = root / "test" / "scrap_stock_example.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Склад"
    ws.append(["Длина", "Количество"])
    ws.append([2800, 2])
    ws.append([1500, 3])
    ws.append([450, 8])
    wb.save(out)
    print(out)


if __name__ == "__main__":
    main()
