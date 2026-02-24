import os
import sys
from pathlib import Path


def resource_path(relative: str) -> Path:
    if getattr(sys, "_MEIPASS", None):
        return Path(sys._MEIPASS) / relative
    return Path(__file__).resolve().parent / relative


def get_version() -> str:
    try:
        path = resource_path("VERSION")
        return path.read_text(encoding="utf-8").strip() or "0.0.0"
    except Exception:
        return "0.0.0"


def appdata_dirname() -> str:
    return "RECA Inclusion Laboral"


def appdata_logs_dir() -> Path:
    base = os.getenv("APPDATA")
    if base:
        root = Path(base) / appdata_dirname() / "logs"
    else:
        root = Path.home() / "AppData" / "Roaming" / appdata_dirname() / "logs"
    root.mkdir(parents=True, exist_ok=True)
    return root
