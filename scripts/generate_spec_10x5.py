"""Генерация тестовой спецификации: 10 модулей × 5 деталей."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.excel_io import write_spec_workbook  # noqa: E402


def build_rows() -> list[tuple]:
    rows: list[tuple] = []
    for i in range(1, 11):
        a = 640 + i * 33
        b = 2410 + i * 27
        d = 1520 + i * 21
        e = 890 + i * 31
        name = f"Модуль M{i}"
        block = [
            (i, name, f"СК-0-{a}Л", a, 90, 1, None),
            ("", "", f"СК-0-{a}П", a, 90, 1, None),
            ("", "", f"Р-1-{b}В", b, 90, 1, None),
            ("", "", f"Р-1-{b}Н", b, 90, 1, None),
            ("", "", f"СС-1-{e}", e, 90, 1, None),
        ]
        rows.extend(block)
    return rows


def main() -> None:
    out = Path(__file__).resolve().parents[1] / "test" / "spec_10x5_modules.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    write_spec_workbook(out, build_rows())
    print(out)


if __name__ == "__main__":
    main()
