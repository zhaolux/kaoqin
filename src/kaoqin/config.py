from __future__ import annotations

import json
import logging
from importlib.resources import files
from pathlib import Path
from typing import Any

from .exceptions import ConfigError


RESOURCE_PACKAGE = "kaoqin.resources"
CONFIG_FILES = {
    "holiday_calendars": "holiday_calendars.json",
    "attendance_layout": "attendance_layout.json",
    "attendance_config": "attendance_config.json",
    "attendance_rules": "attendance_rules.json",
}


def _resource_path(filename: str):
    return files(RESOURCE_PACKAGE).joinpath(filename)


def _project_override_path(filename: str, base_dir: Path | None = None) -> Path:
    root_dir = base_dir or Path.cwd()
    return root_dir / filename


def load_json_config(name: str, base_dir: Path | None = None) -> Any:
    logger = logging.getLogger(__name__)
    filename = CONFIG_FILES[name]
    override_path = _project_override_path(filename, base_dir)
    # Allow a project-local override without changing the packaged defaults.
    if override_path.exists():
        try:
            with override_path.open("r", encoding="utf-8") as file:
                return json.load(file)
        except json.JSONDecodeError as exc:
            raise ConfigError(f"配置文件格式错误：{override_path}") from exc

    resource = _resource_path(filename)
    try:
        with resource.open("r", encoding="utf-8") as file:
            logger.debug("Use packaged config: %s", filename)
            return json.load(file)
    except FileNotFoundError as exc:
        raise ConfigError(f"缺少默认配置文件：{filename}") from exc
