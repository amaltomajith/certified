"""
config_utils.py
Safe, atomic helpers for reading and writing config.yaml.
config.yaml always lives at the project root (one level above backend/).
"""

from pathlib import Path
import yaml

CONFIG_PATH = Path(__file__).parent.parent / "config.yaml"


def read_config() -> dict:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    
    # Ensure nested types structure exists
    if "active_type" not in cfg:
        cfg["active_type"] = "participation"
    if "types" not in cfg:
        cfg["types"] = {
            "participation": {"template_mode": "image", "excel_path": "", "columns": {}, "image_text_fields": {}},
            "winner": {"template_mode": "image", "excel_path": "", "columns": {}, "image_text_fields": {}},
            "school": {"template_mode": "image", "excel_path": "", "columns": {}, "image_text_fields": {}},
        }
    return cfg


def write_config(cfg: dict) -> None:
    """Write the entire config dict back to config.yaml."""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(cfg, f, allow_unicode=True, default_flow_style=False, sort_keys=False)


def patch_config(**kwargs) -> dict:
    """Read config, apply top-level key patches, write back, return updated config."""
    cfg = read_config()
    for key, value in kwargs.items():
        cfg[key] = value
    write_config(cfg)
    return cfg
