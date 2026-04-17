from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_logging() -> Path:
    """
    Настройка общего логирования приложения.
    Логи пишутся в %LOCALAPPDATA%/nordfox_raskroy/logs/app.log (Windows)
    либо в ~/.nordfox_raskroy/logs/app.log на других ОС.
    """
    root = logging.getLogger("nordfox_raskroy")
    if root.handlers:
        # Уже настроено (например, повторный вход из теста/перезапуска окна).
        for h in root.handlers:
            if isinstance(h, RotatingFileHandler):
                return Path(h.baseFilename)
        return Path("app.log")

    base_dir = Path.home() / ".nordfox_raskroy" / "logs"
    local_app_data = Path.home()
    try:
        import os

        lad = os.getenv("LOCALAPPDATA")
        if lad:
            local_app_data = Path(lad)
            base_dir = local_app_data / "nordfox_raskroy" / "logs"
    except Exception:
        pass

    base_dir.mkdir(parents=True, exist_ok=True)
    log_file = base_dir / "app.log"

    root.setLevel(logging.INFO)
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    fh = RotatingFileHandler(log_file, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    fh.setLevel(logging.INFO)
    fh.setFormatter(fmt)

    sh = logging.StreamHandler()
    sh.setLevel(logging.WARNING)
    sh.setFormatter(fmt)

    root.addHandler(fh)
    root.addHandler(sh)
    root.propagate = False
    root.info("Logging initialized: %s", log_file)
    return log_file
