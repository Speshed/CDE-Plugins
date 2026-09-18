from __future__ import annotations

import json
import os
import sys
from pathlib import Path

__all__ = [
    "DEFAULT_SSO_BASE_URL", "DEFAULT_REALM", "DEFAULT_BROKER_ALIAS", "DEFAULT_ADFS_BASE_URL",
    "DEFAULT_CLIENT_ID", "DEFAULT_BASE_URL", "AUTH_TIMEOUT_SECONDS", "CONNECTION_PRESETS",
    "PROFILES_FILE", "load_connection_profile", "save_connection_profile",
]


DEFAULT_SSO_BASE_URL = "https://sso.oz-mine.com:8444"


DEFAULT_REALM = "highlandgold"


DEFAULT_BROKER_ALIAS = "highlandgoldsaml"


DEFAULT_ADFS_BASE_URL = "https://adfs.hgml.ru"


DEFAULT_CLIENT_ID = "projectpoint"


DEFAULT_BASE_URL = "https://projectpoint.areal.ru"


AUTH_TIMEOUT_SECONDS = 60


CONNECTION_PRESETS = {
    "Production": {
        "auth_mode": "keycloak_broker",
        "base_url": "https://projectpoint.areal.ru",
        "client_id": "projectpoint",
        "sso_base_url": "https://sso.oz-mine.com:8444",
        "realm": "highlandgold",
        "broker_alias": "highlandgoldsaml",
        "adfs_base_url": "https://adfs.hgml.ru",
    },
    "Test": {
        "auth_mode": "adfs_direct",
        "base_url": "https://ibim-test.cloud.projectpoint.ru",
        "client_id": "ebd1641a-a126-45ef-8a8e-f7f0de6fd3fa",
        "sso_base_url": "",
        "realm": "",
        "broker_alias": "",
        "adfs_base_url": "https://adfs.tnsv.projectpoint.ru",
    },
    "Eyurevich": {
        "auth_mode": "adfs_direct",
        "base_url": "https://eyurevich.cloud.projectpoint.ru",
        "client_id": "4e41e442-caf5-452f-94fb-b937a4b97942",
        "sso_base_url": "",
        "realm": "",
        "broker_alias": "",
        "adfs_base_url": "https://adfs.tnsv.projectpoint.ru",
    },
    "Custom": {
        "auth_mode": "keycloak_broker",
        "base_url": "",
        "client_id": "",
        "sso_base_url": "",
        "realm": "",
        "broker_alias": "",
        "adfs_base_url": "",
    },
}


PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PACKAGE_DIR.parent


def _profiles_file_path():
    """Return the editable connection-profile cache path.

    A frozen portable build must not create files beside Larix_CDE.exe.
    Store user state in the Windows user profile instead.  Source runs keep
    the historical project-local file for developer convenience.
    """
    if getattr(sys, "frozen", False):
        base = (
            os.environ.get("LOCALAPPDATA")
            or os.environ.get("APPDATA")
            or str(Path.home() / ".larix_cde")
        )
        return Path(base).expanduser() / "Larix CDE" / "ProjectPoint" / "connection_profiles.json"
    return PROJECT_ROOT / "connection_profiles.json"


PROFILES_FILE = _profiles_file_path()


def load_connection_profile(base_url):
    """Load cached connection profile for a given base_url."""
    if not os.path.exists(PROFILES_FILE):
        return None
    try:
        with open(PROFILES_FILE, "r", encoding="utf-8") as f:
            profiles = json.load(f)
        return profiles.get(base_url)
    except Exception:
        return None


def save_connection_profile(base_url, profile):
    """Save connection profile to cache."""
    profiles = {}
    if os.path.exists(PROFILES_FILE):
        try:
            with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                profiles = json.load(f)
        except Exception:
            pass
    
    profiles[base_url] = profile
    try:
        PROFILES_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(profiles, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Warning: Failed to save connection profile: {e}")
