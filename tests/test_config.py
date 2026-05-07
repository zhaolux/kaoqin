import json

import pytest

from kaoqin.config import load_json_config
from kaoqin.exceptions import ConfigError


def test_load_json_config_uses_packaged_default(tmp_path):
    config = load_json_config("attendance_config", base_dir=tmp_path)
    assert "auto_workday_people" in config
    assert "name_alias" in config


def test_load_json_config_prefers_project_override(tmp_path):
    override = {
        "auto_workday_people": ["测试人员"],
        "name_alias": {"甲": "乙"},
    }
    (tmp_path / "attendance_config.json").write_text(
        json.dumps(override, ensure_ascii=False),
        encoding="utf-8",
    )

    config = load_json_config("attendance_config", base_dir=tmp_path)
    assert config == override


def test_load_json_config_raises_for_invalid_json(tmp_path):
    (tmp_path / "attendance_config.json").write_text("{bad json", encoding="utf-8")

    with pytest.raises(ConfigError):
        load_json_config("attendance_config", base_dir=tmp_path)
