import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.table_demand_import import demands_from_cut_table_rows  # noqa: E402


def _row(
    mod="Модуль M1",
    prof="СК-0-100Л",
    ser="Н20",
    ln="1000",
    ang="90",
    src="Новая",
    st="6000",
    rem="5000",
) -> list[str]:
    return [mod, prof, ser, ln, ang, src, st, rem]


class TableDemandImportTests(unittest.TestCase):
    def test_ok(self):
        d, err = demands_from_cut_table_rows([_row()], {0, 1, 2, 3})
        self.assertEqual(err, "")
        assert d is not None
        self.assertEqual(d[0].length_mm, 1000)
        self.assertEqual(d[0].profile_code, "СК-0-100Л")

    def test_bad_profile(self):
        d, err = demands_from_cut_table_rows([_row(prof="XX-0-1")], {0, 1, 2, 3})
        self.assertIsNone(d)
        self.assertIn("ожидается", err)

    def test_filter_digit(self):
        d, err = demands_from_cut_table_rows([_row(prof="СК-2-1")], {0, 1})
        self.assertIsNone(d)
        self.assertIn("Н22", err)

    def test_ok_with_extra_mass_column(self):
        r = _row()
        r.append("1.234")
        d, err = demands_from_cut_table_rows([r], {0, 1, 2, 3})
        self.assertEqual(err, "")
        assert d is not None
        self.assertEqual(len(d), 1)


if __name__ == "__main__":
    unittest.main()
