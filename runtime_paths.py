# -*- coding: utf-8 -*-
import os
from pathlib import Path


def _default_app_home() -> Path:
    env_home = os.environ.get("ECOUNT_SYSTEM_HOME") or os.environ.get("YIKAN_APP_HOME")
    if env_home:
        return Path(env_home).expanduser()

    if os.name == "nt":
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / "ecount-system"

    return Path.home() / ".ecount-system"


APP_HOME = _default_app_home()
CONFIG_FILE = APP_HOME / "config.json"
BASE_DATA_DB_FILE = APP_HOME / "base_data.db"


def runtime_file(name: str) -> Path:
    return APP_HOME / name


def ensure_app_home() -> Path:
    APP_HOME.mkdir(parents=True, exist_ok=True)
    return APP_HOME


__all__ = [
    "APP_HOME",
    "CONFIG_FILE",
    "BASE_DATA_DB_FILE",
    "runtime_file",
    "ensure_app_home",
]
