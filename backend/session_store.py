"""
session_store.py
Stores Gmail credentials both in-memory (for the current request) and
optionally persisted to a local .auth file so you don't have to re-enter
on every restart.

Security note: this is a local-only tool.  The credentials file is
base64-encoded (not encrypted) — acceptable for a developer tool that
never leaves the developer's machine.
"""

import base64
import json
import os
from pathlib import Path

_store: dict = {}

# Where the persisted credentials live — same directory as config.yaml
_CREDS_FILE = Path(__file__).parent.parent / ".auth_credentials"


# ── In-memory helpers ──────────────────────────────────────────────────────

def set_credentials(sender_email: str, app_password: str, persist: bool = True) -> None:
    """Store credentials in memory and (optionally) on disk."""
    _store["sender_email"] = sender_email
    _store["app_password"] = app_password.replace(" ", "")
    if persist:
        _save_to_disk(sender_email, _store["app_password"])


def get_app_password() -> str | None:
    return _store.get("app_password")


def get_sender_email() -> str | None:
    return _store.get("sender_email")


def has_app_password() -> bool:
    return bool(_store.get("app_password"))


# Keep old API name so other code doesn't break
def set_app_password(raw: str) -> None:
    _store["app_password"] = raw.replace(" ", "")


def clear_credentials() -> None:
    """Sign out: wipe memory AND delete the credentials file."""
    _store.pop("app_password", None)
    _store.pop("sender_email", None)
    if _CREDS_FILE.exists():
        _CREDS_FILE.unlink()


# Legacy alias
clear_app_password = clear_credentials


# ── Disk persistence ───────────────────────────────────────────────────────

def _save_to_disk(sender_email: str, app_password: str) -> None:
    payload = json.dumps({"sender_email": sender_email, "app_password": app_password})
    encoded = base64.b64encode(payload.encode()).decode()
    _CREDS_FILE.write_text(encoded)
    # Restrict permissions on non-Windows; Windows doesn't support chmod the same way
    try:
        os.chmod(_CREDS_FILE, 0o600)
    except Exception:
        pass


def load_from_disk() -> bool:
    """Try to restore credentials from disk. Returns True if successful."""
    if not _CREDS_FILE.exists():
        return False
    try:
        encoded = _CREDS_FILE.read_text().strip()
        payload = json.loads(base64.b64decode(encoded).decode())
        _store["sender_email"] = payload.get("sender_email", "")
        _store["app_password"] = payload.get("app_password", "")
        return bool(_store["app_password"])
    except Exception:
        return False
