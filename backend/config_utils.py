"""
config_utils.py
Safe, atomic helpers for reading and writing config.yaml.
config.yaml always lives at the project root (one level above backend/).
"""

from pathlib import Path
import shutil
import yaml

CONFIG_PATH = Path(__file__).parent.parent / "config.yaml"
DEFAULT_CONFIG_PATH = Path(__file__).parent.parent / "config.default.yaml"


def _ensure_config_exists() -> None:
    """Create config.yaml from config.default.yaml if it doesn't exist (e.g. first run in prod)."""
    if not CONFIG_PATH.exists():
        if DEFAULT_CONFIG_PATH.exists():
            shutil.copy(DEFAULT_CONFIG_PATH, CONFIG_PATH)
        else:
            # Minimal fallback if even the default is missing
            default_cfg = {
                "active_type": "participation",
                "types": {
                    "participation": {"template_mode": "image", "excel_path": "", "columns": {}, "image_text_fields": {}},
                    "winner": {"template_mode": "image", "excel_path": "", "columns": {}, "image_text_fields": {}},
                    "school": {"template_mode": "image", "excel_path": "", "columns": {}, "image_text_fields": {}},
                    "volunteer": {"template_mode": "image", "excel_path": "", "columns": {}, "image_text_fields": {}},
                },
                "output_dir": "output",
                "email": {"dry_run": False, "sender_email": "", "sender_app_password": "", "smtp_host": "smtp.gmail.com", "smtp_port": 465},
            }
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                yaml.dump(default_cfg, f, allow_unicode=True, default_flow_style=False, sort_keys=False)


def _merge_defaults(cfg: dict) -> dict:
    """If user config has empty template paths or text fields, fall back to config.default.yaml."""
    if not DEFAULT_CONFIG_PATH.exists():
        return cfg
    try:
        with open(DEFAULT_CONFIG_PATH, "r", encoding="utf-8") as f:
            d_cfg = yaml.safe_load(f) or {}
    except Exception:
        return cfg

    # Ensure types exist
    types = cfg.get("types", {})
    d_types = d_cfg.get("types", {})

    for cert_type, d_type_cfg in d_types.items():
        if cert_type not in types:
            types[cert_type] = d_type_cfg
            continue
        type_cfg = types[cert_type]
        
        # Fallback template path
        if not type_cfg.get("image_template_path") and d_type_cfg.get("image_template_path"):
            type_cfg["image_template_path"] = d_type_cfg["image_template_path"]
        
        for k in ("image_template_path_1st", "image_template_path_2nd", "image_template_path_3rd"):
            if not type_cfg.get(k) and d_type_cfg.get(k):
                type_cfg[k] = d_type_cfg[k]

        # Fallback text fields if empty
        if not type_cfg.get("image_text_fields") and d_type_cfg.get("image_text_fields"):
            type_cfg["image_text_fields"] = d_type_cfg["image_text_fields"]

        for k in ("image_text_fields_1st", "image_text_fields_2nd", "image_text_fields_3rd"):
            if not type_cfg.get(k) and d_type_cfg.get(k):
                type_cfg[k] = d_type_cfg[k]

    cfg["types"] = types
    return cfg


def read_config() -> dict:
    import os
    _ensure_config_exists()
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

    # Fallback to defaults for missing template paths or fields
    cfg = _merge_defaults(cfg)

    # Fallback to Environment Variables if set (for cloud deployments like Render)
    if "email" not in cfg:
        cfg["email"] = {}
    
    env_email = os.environ.get("SENDER_EMAIL") or os.environ.get("GMAIL_EMAIL")
    env_pass = os.environ.get("SENDER_APP_PASSWORD") or os.environ.get("GMAIL_APP_PASSWORD")

    if env_email and not cfg["email"].get("sender_email"):
        cfg["email"]["sender_email"] = env_email
    if env_pass and not cfg["email"].get("sender_app_password"):
        cfg["email"]["sender_app_password"] = env_pass

    return cfg


def write_config(cfg: dict) -> None:
    """Write the entire config dict back to config.yaml, always preserving output_dir."""
    cfg.setdefault("output_dir", "output")
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(cfg, f, allow_unicode=True, default_flow_style=False, sort_keys=False)


def patch_config(**kwargs) -> dict:
    """Read config, apply top-level key patches, write back, return updated config."""
    cfg = read_config()
    for key, value in kwargs.items():
        cfg[key] = value
    write_config(cfg)
    return cfg
