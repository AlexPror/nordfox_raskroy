import sys
import tempfile
import unittest
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.excel_io import parse_specification, write_spec_workbook  # noqa: E402


class ExcelIoTests(unittest.TestCase):
    def test_parse_specification_skips_blank_rows(self):
        rows = [
            (1, "Модуль M1", "L15", 1000, 90, 1, None),
            (None, None, None, None, None, None, None),
            (2, "Модуль M2", "СК-0-700", 700, 45, 2, None),
        ]
        fd, temp_name = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        p = Path(temp_name)
        try:
            write_spec_workbook(p, rows)
            parsed = parse_specification(p)
        finally:
            try:
                os.remove(p)
            except OSError:
                pass
        self.assertEqual(len(parsed), 2)
        self.assertEqual(parsed[0].profile_code, "L15")
        self.assertEqual(parsed[1].profile_code, "СК-0-700")

    def test_parse_block_template_two_angles_per_piece(self):
        path = ROOT / "test" / "Блок спецификации.xlsx"
        if not path.is_file():
            self.skipTest("нет файла образца")
        rows = parse_specification(path)
        first = next(x for x in rows if x.profile_code == "СК-0-1045")
        self.assertEqual(first.cut_angle, 45)
        self.assertEqual(first.cut_angle_2, 87)

    def test_parse_generated_20_modules_block(self):
        path = ROOT / "test" / "spec_20_modules_block.xlsx"
        if not path.is_file():
            self.skipTest("запустите scripts/generate_spec_20_modules_block.py")
        rows = parse_specification(path)
        self.assertEqual(len({r.module_name for r in rows}), 20)
        self.assertGreater(len(rows), 200)
        self.assertTrue(any(r.cut_angle_2 is not None for r in rows))


if __name__ == "__main__":
    unittest.main()
