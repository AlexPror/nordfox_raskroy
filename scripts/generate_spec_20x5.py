"""Генерация спецификации: 20 модулей × 5 деталей, цифры профиля 0–3 по кругу."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.excel_io import write_spec_workbook  # noqa: E402


def build_rows() -> list[tuple]:
    rows: list[tuple] = []
    for i in range(1, 21):
        d0 = (i - 1) % 4
        d1 = i % 4
        d2 = (i + 1) % 4
        d3 = (i + 2) % 4
        d4 = (i + 3) % 4
        a = 620 + (i % 9) * 35 + (i // 4) * 12
        b = 2100 + (i % 6) * 45 + (i // 5) * 30
        e = 500 + (i % 8) * 28
        name = f"Модуль M{i}"
        block = [
            (i, name, f"СК-{d0}-{a}Л", a, 90, 1, None),
            ("", "", f"СК-{d1}-{a}П", a, 90, 1, None),
            ("", "", f"Р-{d2}-{b}В", b, 90, 1, None),
            ("", "", f"Р-{d3}-{b}Н", b, 90, 1, None),
            ("", "", f"СС-{d4}-{e}", e, 90, 1, None),
        ]
        rows.extend(block)
    return rows


def main() -> None:
    out = Path(__file__).resolve().parents[1] / "test" / "spec_20x5_modules.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    write_spec_workbook(out, build_rows())
    print(out)


if __name__ == "__main__":
    main()
