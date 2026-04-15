import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))


class VersionTests(unittest.TestCase):
    def test_version_file_exists(self):
        vf = ROOT / "src" / "nordfox_raskroy" / "VERSION"
        self.assertTrue(vf.is_file())
        raw = vf.read_text(encoding="utf-8").strip()
        self.assertTrue(re.match(r"^\d+\.\d+\.\d+", raw), raw)

    def test_package_version(self):
        from nordfox_raskroy import __version__

        self.assertTrue(__version__)
        self.assertRegex(__version__, r"^\d+\.\d+\.\d+")


if __name__ == "__main__":
    unittest.main()
