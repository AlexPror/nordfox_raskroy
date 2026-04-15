import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.scrap_stock_io import parse_scrap_inventory  # noqa: E402


class ScrapStockIoTests(unittest.TestCase):
    def test_parse_example(self):
        p = ROOT / "test" / "scrap_stock_example.xlsx"
        if not p.is_file():
            self.skipTest("Запустите scripts/generate_scrap_stock_example.py")
        pieces, warns = parse_scrap_inventory(p)
        self.assertEqual(len(pieces), 2 + 3 + 8)
        self.assertEqual(pieces.count(2800), 2)
        self.assertFalse(warns)


if __name__ == "__main__":
    unittest.main()
