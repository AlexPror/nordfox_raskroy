import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.models import CutEvent, PartDemand, StockScrapPiece  # noqa: E402
from nordfox_raskroy.usable_remnants_service import build_usable_remnant_rows  # noqa: E402


class UsableRemnantsServiceTests(unittest.TestCase):
    def test_from_layout_remainder(self):
        layout = [
            {
                "opening": 1,
                "bar_len": 6000,
                "remainder": 450,
                "segments": [
                    {
                        "kind": "profile",
                        "profile_code": "СК-0-1000",
                        "profile_name": "Н20",
                    }
                ],
            }
        ]
        rows = build_usable_remnant_rows(
            layout_rows=layout,
            cuts=[],
            scrap_pieces=[],
            min_length_mm=100,
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["length_mm"], 450)
        self.assertEqual(rows[0]["profile_code"], "СК-0-1000")

    def test_sort_length_asc(self):
        layout = [
            {"opening": 1, "bar_len": 6000, "remainder": 200, "segments": [{"kind": "profile", "profile_code": "A", "profile_name": "Н20"}]},
            {"opening": 2, "bar_len": 6000, "remainder": 500, "segments": [{"kind": "profile", "profile_code": "B", "profile_name": "Н21"}]},
        ]
        rows = build_usable_remnant_rows(
            layout_rows=layout,
            cuts=[],
            scrap_pieces=[],
            sort_mode="length_asc",
        )
        lengths = [int(r["length_mm"]) for r in rows]
        self.assertEqual(lengths, [200, 500])

    def test_warehouse_scrap_piece(self):
        rows = build_usable_remnant_rows(
            layout_rows=[],
            cuts=[],
            scrap_pieces=[StockScrapPiece(length_mm=150, opening_id=0, profile_key="Н20")],
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["source"], "склад")


if __name__ == "__main__":
    unittest.main()
