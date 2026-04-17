from __future__ import annotations

from pathlib import Path

from nordfox_raskroy.excel_io import write_spec_workbook


def main() -> None:
    out = Path(__file__).resolve().parents[1] / "test" / "spec_angles_0_180_step10.xlsx"
    rows: list[tuple[object, ...]] = []
    item_no = 1
    for angle in range(0, 181, 10):
        rows.append(
            (
                item_no,
                f"Модуль М{item_no}",
                "СК-1-1200",
                1200,
                angle,
                1,
                f"ANG-{angle:03d}",
            )
        )
        item_no += 1
    write_spec_workbook(out, rows)
    print(out)


if __name__ == "__main__":
    main()
