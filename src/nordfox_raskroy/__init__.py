"""NordFox — линейный раскрой (прототип)."""

from __future__ import annotations

from pathlib import Path


def _read_package_version() -> str:
    vf = Path(__file__).with_name("VERSION")
    if vf.is_file():
        return vf.read_text(encoding="utf-8").strip()
    try:
        from importlib.metadata import version

        return version("nordfox-raskroy")
    except Exception:  # noqa: BLE001
        return "0.0.0"


__version__ = _read_package_version()
