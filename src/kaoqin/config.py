from __future__ import annotations

import json
from importlib.resources import files
from pathlib import Path
from typing import Any


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
    filename = CONFIG_FILES[name]
    override_path = _project_override_path(filename, base_dir)
    if override_path.exists():
        with override_path.open("r", encoding="utf-8") as file:
            return json.load(file)

    resource = _resource_path(filename)
    with resource.open("r", encoding="utf-8") as file:
        return json.load(file)
