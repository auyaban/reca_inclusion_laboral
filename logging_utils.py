import os
import tempfile
import threading
from datetime import datetime


LOGS_FOLDER_NAME = "Inclusion laboral logs"
LOG_MAX_BYTES = 10 * 1024 * 1024
LOG_TRIM_BYTES = 5 * 1024 * 1024
LOG_FILE_MAP = {
    "app": "app.log",
    "excel": "excel.log",
    "drive": "drive.log",
    "supabase": "supabase.log",
}


_LOG_WRITE_LOCK = threading.Lock()
_LOG_ROOT_CACHE = None
_LOG_ROOT_SOURCE = None
_FALLBACK_NOTICE_WRITTEN = False


def _resolve_desktop_dir():
    def _validate(path):
        candidate = str(path or "").strip()
        if not candidate:
            return None
        if os.path.isdir(candidate):
            return candidate
        return None

    if os.name == "nt":
        try:
            import winreg

            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            ) as key:
                desktop_raw, _ = winreg.QueryValueEx(key, "Desktop")
            resolved = _validate(os.path.expandvars(str(desktop_raw or "").strip()))
            if resolved:
                return resolved
        except Exception:
            pass

    for env_key in ("OneDrive", "OneDriveConsumer"):
        base = os.getenv(env_key)
        if base:
            resolved = _validate(os.path.join(base, "Desktop"))
            if resolved:
                return resolved

    userprofile = os.getenv("USERPROFILE")
    if userprofile:
        for name in ("Desktop", "Escritorio"):
            resolved = _validate(os.path.join(userprofile, name))
            if resolved:
                return resolved

    home = os.path.expanduser("~")
    for name in ("Desktop", "Escritorio"):
        resolved = _validate(os.path.join(home, name))
        if resolved:
            return resolved
    return None


def _ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def get_logs_root():
    global _LOG_ROOT_CACHE, _LOG_ROOT_SOURCE
    if _LOG_ROOT_CACHE is not None:
        return _LOG_ROOT_CACHE

    desktop = _resolve_desktop_dir()
    if desktop:
        try:
            _LOG_ROOT_CACHE = _ensure_dir(os.path.join(desktop, LOGS_FOLDER_NAME))
            _LOG_ROOT_SOURCE = "desktop"
            return _LOG_ROOT_CACHE
        except Exception:
            pass

    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        try:
            _LOG_ROOT_CACHE = _ensure_dir(os.path.join(local_app_data, "RECA", LOGS_FOLDER_NAME))
            _LOG_ROOT_SOURCE = "localappdata"
            return _LOG_ROOT_CACHE
        except Exception:
            pass

    try:
        _LOG_ROOT_CACHE = _ensure_dir(os.path.join(os.getcwd(), LOGS_FOLDER_NAME))
        _LOG_ROOT_SOURCE = "cwd"
        return _LOG_ROOT_CACHE
    except Exception:
        pass

    try:
        _LOG_ROOT_CACHE = _ensure_dir(os.path.join(tempfile.gettempdir(), "RECA", LOGS_FOLDER_NAME))
        _LOG_ROOT_SOURCE = "temp"
        return _LOG_ROOT_CACHE
    except Exception:
        pass

    _LOG_ROOT_CACHE = ""
    _LOG_ROOT_SOURCE = "none"
    return _LOG_ROOT_CACHE


def get_logs_root_source():
    get_logs_root()
    return _LOG_ROOT_SOURCE or "unknown"


def get_log_file_path(channel):
    key = str(channel or "").strip().lower()
    file_name = LOG_FILE_MAP.get(key, f"{key or 'app'}.log")
    root = get_logs_root()
    if not root:
        return ""
    return os.path.join(root, file_name)


def _trim_log_file(path, incoming_bytes=0):
    if not os.path.exists(path):
        return
    try:
        size = os.path.getsize(path)
    except OSError:
        return

    if size + int(incoming_bytes or 0) <= LOG_MAX_BYTES or size <= 0:
        return

    try:
        with open(path, "rb") as handle:
            handle.seek(LOG_TRIM_BYTES)
            newline_pos = -1
            chunk = handle.read(512)
            if chunk:
                idx = chunk.find(b"\n")
                if idx != -1:
                    newline_pos = LOG_TRIM_BYTES + idx + 1
            if newline_pos == -1:
                newline_pos = LOG_TRIM_BYTES
            handle.seek(newline_pos)
            kept = handle.read()
    except OSError:
        return
    if not kept:
        return

    tmp_path = f"{path}.trim"
    try:
        with open(tmp_path, "wb") as handle:
            handle.write(kept)
        os.replace(tmp_path, path)
    except OSError:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass


def _write_log_line(path, line):
    if not path:
        return
    encoded = line.encode("utf-8", errors="replace")
    _trim_log_file(path, incoming_bytes=len(encoded))
    with open(path, "ab") as handle:
        handle.write(encoded)


def log_event(channel, message, level="INFO"):
    global _FALLBACK_NOTICE_WRITTEN
    channel_key = str(channel or "app").strip().lower() or "app"
    level_key = str(level or "INFO").strip().upper() or "INFO"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] [{channel_key.upper()}] [{level_key}] {str(message).rstrip()}\n"
    path = get_log_file_path(channel_key)
    with _LOG_WRITE_LOCK:
        _write_log_line(path, line)
        source = get_logs_root_source()
        if channel_key == "app" and source != "desktop" and not _FALLBACK_NOTICE_WRITTEN:
            fallback_line = (
                f"[{timestamp}] [APP] [WARN] logs_root_fallback source={source} path={get_logs_root()}\n"
            )
            _write_log_line(path, fallback_line)
            _FALLBACK_NOTICE_WRITTEN = True


def log_app_event(message, level="INFO"):
    log_event("app", message, level=level)


def log_excel_event(message, level="INFO"):
    log_event("excel", message, level=level)


def log_drive_event(message, level="INFO"):
    log_event("drive", message, level=level)


def log_supabase_event(message, level="INFO"):
    log_event("supabase", message, level=level)
