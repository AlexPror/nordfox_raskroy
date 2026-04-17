"""Точка входа GUI: Qt (PySide6), при отсутствии пакета — tkinter."""

from __future__ import annotations

import importlib.util
import sys

from nordfox_raskroy.logging_utils import setup_logging

__all__ = ["run_app"]


def run_app() -> None:
    log_file = setup_logging()
    import logging

    logger = logging.getLogger("nordfox_raskroy.app")
    logger.info("Application start requested. Log file: %s", log_file)
    if importlib.util.find_spec("PySide6") is not None:
        from nordfox_raskroy.qt_app import run_app as _qt_run

        logger.info("Launching Qt UI")
        _qt_run()
        return
    logger.warning("PySide6 not found. Falling back to tkinter UI")
    print(
        "PySide6 не найден. Установите: pip install PySide6\n"
        "Временно открывается интерфейс tkinter (тот же фильтр профилей).\n",
        file=sys.stderr,
    )
    from nordfox_raskroy.app_tk import run_tk_app

    run_tk_app()
