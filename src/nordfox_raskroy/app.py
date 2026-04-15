"""Точка входа GUI: Qt (PySide6), при отсутствии пакета — tkinter."""

from __future__ import annotations

import importlib.util
import sys

__all__ = ["run_app"]


def run_app() -> None:
    if importlib.util.find_spec("PySide6") is not None:
        from nordfox_raskroy.qt_app import run_app as _qt_run

        _qt_run()
        return
    print(
        "PySide6 не найден. Установите: pip install PySide6\n"
        "Временно открывается интерфейс tkinter (тот же фильтр профилей).\n",
        file=sys.stderr,
    )
    from nordfox_raskroy.app_tk import run_tk_app

    run_tk_app()
