import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.excel_io import parse_specification  # noqa: E402
from nordfox_raskroy.optimizer import optimize_cutting, spec_rows_to_demands  # noqa: E402
from nordfox_raskroy.profile_codes import filter_spec_by_profiles  # noqa: E402
from nordfox_raskroy.result_sort import sort_cuts  # noqa: E402


class Spec20PipelineTests(unittest.TestCase):
    spec_path = ROOT / "test" / "spec_20x5_modules.xlsx"

    def test_parse_and_optimize_20_modules(self):
        if not self.spec_path.is_file():
            self.skipTest("Запустите: python scripts/generate_spec_20x5.py")
        rows = parse_specification(self.spec_path)
        self.assertEqual(len(rows), 100)
        rows_f, _ = filter_spec_by_profiles(rows, {0, 1, 2, 3})
        self.assertEqual(len(rows_f), 100)
        demands = spec_rows_to_demands(rows_f)
        result = optimize_cutting(
            demands,
            bar_lengths_mm=[6000, 7500, 12000],
            kerf_mm=0,
            min_scrap_mm=50,
        )
        self.assertEqual(len(result.cuts), 100)
        by_len = sort_cuts(result.cuts, "length_desc")
        self.assertGreaterEqual(by_len[0].demand.length_mm, by_len[-1].demand.length_mm)


if __name__ == "__main__":
    unittest.main()
