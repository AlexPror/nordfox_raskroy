import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.models import CutEvent, PartDemand  # noqa: E402
from nordfox_raskroy.result_sort import sort_cuts  # noqa: E402


def _cut(
    module: str,
    profile: str,
    length: int,
    src: str = "new_bar",
    stock: int = 6000,
    opening: int = 0,
) -> CutEvent:
    return CutEvent(
        PartDemand(1, module, profile, length, 90),
        stock,
        src,  # type: ignore[arg-type]
        100,
        0,
        stock_opening_id=opening,
    )


class ResultSortTests(unittest.TestCase):
    def test_length_desc(self):
        cuts = [_cut("Модуль M2", "Р-0-100", 500), _cut("Модуль M1", "СК-0-200", 2000)]
        out = sort_cuts(cuts, "length_desc")
        self.assertEqual(out[0].demand.length_mm, 2000)

    def test_series_groups_digits(self):
        cuts = [
            _cut("Модуль M1", "СК-3-100", 100),
            _cut("Модуль M1", "СК-0-200", 200),
        ]
        out = sort_cuts(cuts, "series")
        self.assertEqual(out[0].demand.profile_code, "СК-0-200")

    def test_as_calculated_preserves_order(self):
        cuts = [_cut("Модуль M1", "A", 1), _cut("Модуль M2", "B", 2)]
        out = sort_cuts(cuts, "as_calculated")
        self.assertEqual([c.demand.profile_code for c in out], ["A", "B"])

    def test_source_scrap_first(self):
        cuts = [
            _cut("Модуль M1", "A", 100, "new_bar"),
            _cut("Модуль M1", "B", 100, "scrap", 500),
        ]
        out = sort_cuts(cuts, "source")
        self.assertEqual(out[0].stock_source, "scrap")

    def test_operator_profile_then_scrap_then_stock_length(self):
        cuts = [
            _cut("Модуль M2", "ZZ", 100, "new_bar", 6000, opening=2),
            _cut("Модуль M1", "AA", 200, "new_bar", 6000, opening=1),
            _cut("Модуль M1", "AA", 150, "scrap", 800, opening=1),
        ]
        out = sort_cuts(cuts, "operator")
        profiles = [c.demand.profile_code for c in out]
        self.assertEqual(profiles, ["AA", "AA", "ZZ"])
        self.assertEqual(out[0].stock_source, "scrap")
        self.assertEqual(out[1].stock_source, "new_bar")

    def test_opening_sort_by_stock_opening_id(self):
        cuts = [
            _cut("M", "P", 100, opening=2),
            _cut("M", "P", 200, opening=1),
        ]
        out = sort_cuts(cuts, "opening")
        self.assertEqual([c.stock_opening_id for c in out], [1, 2])


if __name__ == "__main__":
    unittest.main()
