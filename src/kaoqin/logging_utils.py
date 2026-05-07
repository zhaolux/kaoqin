from __future__ import annotations

import logging
from pathlib import Path


DEFAULT_LOG_DIRNAME = "logs"
DEFAULT_LOG_FILENAME = "kaoqin.log"


def setup_logging(base_dir: Path | None = None, level: str = "INFO", log_file: str | None = None) -> Path:
    root_dir = base_dir or Path.cwd()
    log_path = Path(log_file) if log_file else root_dir / DEFAULT_LOG_DIRNAME / DEFAULT_LOG_FILENAME
    if not log_path.is_absolute():
        log_path = root_dir / log_path

    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger()
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return log_path
