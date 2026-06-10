import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.shop_floor_progress import (  # noqa: E402
    BarProgressRow,
    ShopFloorSession,
    duration_minutes,
    load_session,
    merge_progress,
    progress_by_opening,
    save_session,
)


class ShopFloorProgressTests(unittest.TestCase):
    def test_merge_progress_keeps_stored(self):
        stored = progress_by_opening(
            [
                BarProgressRow(opening=1, done=True, start_iso="2026-06-10T08:00:00", end_iso="2026-06-10T08:30:00"),
            ]
        )
        merged = merge_progress([1, 2], stored)
        self.assertTrue(merged[0].done)
        self.assertEqual(merged[1].opening, 2)
        self.assertFalse(merged[1].done)

    def test_duration_minutes(self):
        txt = duration_minutes("2026-06-10T08:00:00", "2026-06-10T09:15:00")
        self.assertEqual(txt, "1 ч 15 мин")

    def test_json_roundtrip(self):
        session = ShopFloorSession(
            project_name="Тест",
            rows=[BarProgressRow(opening=3, start_iso="2026-06-10T10:00:00")],
        )
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "progress.json"
            save_session(path, session)
            loaded = load_session(path)
        self.assertEqual(loaded.project_name, "Тест")
        self.assertEqual(len(loaded.rows), 1)
        self.assertEqual(loaded.rows[0].opening, 3)


if __name__ == "__main__":
    unittest.main()
