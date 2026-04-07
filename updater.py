from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
import uuid
from datetime import datetime
from pathlib import Path

from version_info import appdata_logs_dir


DEFAULT_REPO_OWNER = "auyaban"
DEFAULT_REPO_NAME = "reca_inclusion_laboral"
DEFAULT_INSTALLER_ASSET = "RECA_INCLUSION_LABORAL_Setup.exe"
DEFAULT_HASH_ASSET = f"{DEFAULT_INSTALLER_ASSET}.sha256"
DEFAULT_UPDATE_HELPER_NAME = "RECA_UPDATER_HELPER.exe"
DEFAULT_UPDATE_SESSION_ROOT = "RECAUpdater"
UPDATE_POINTER_FILE_NAME = "last_update_session.json"
UPDATE_MANIFEST_FILE_NAME = "update_manifest.json"
UPDATE_STATUS_FILE_NAME = "update_status.json"
UPDATE_READY_ACK_FILE_NAME = "update_ready.ack"
FINAL_UPDATE_STATES = {"succeeded", "failed"}
_LOG_CURRENT_DAY = None


def _update_log_path() -> Path:
    try:
        desktop = Path(os.path.expanduser("~")) / "Desktop"
        desktop.mkdir(parents=True, exist_ok=True)
        return desktop / "updater.log"
    except Exception:
        return appdata_logs_dir() / "updater.log"


def _ensure_daily_log(path: Path) -> None:
    global _LOG_CURRENT_DAY
    today = datetime.now().strftime("%Y-%m-%d")
    rotate = False
    if _LOG_CURRENT_DAY != today:
        rotate = True
    elif path.exists():
        try:
            file_day = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d")
            if file_day != today:
                rotate = True
        except Exception:
            rotate = True
    else:
        rotate = True

    if not rotate:
        return

    try:
        with path.open("w", encoding="utf-8") as handle:
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            handle.write(f"[{stamp}] [SYSTEM] Inicio de log diario ({today})\n")
    except Exception:
        pass
    _LOG_CURRENT_DAY = today


def _log_update(message: str) -> None:
    try:
        log_path = _update_log_path()
        _ensure_daily_log(log_path)
        with log_path.open("a", encoding="utf-8") as handle:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            handle.write(f"[{timestamp}] [UPDATER] {message.rstrip()}\n")
    except Exception:
        pass


def _resolve_env_candidates(env_name: str = ".env") -> list[Path]:
    candidates: list[Path] = []
    appdata = str(os.getenv("APPDATA") or "").strip()
    if appdata:
        candidates.append(Path(appdata) / "RECA Inclusion Laboral" / env_name)
    if getattr(sys, "frozen", False):
        try:
            candidates.append(Path(sys.executable).resolve().parent / env_name)
        except Exception:
            pass
    candidates.append(Path.cwd() / env_name)
    return candidates


def _load_env_file(env_name: str = ".env") -> dict:
    for candidate in _resolve_env_candidates(env_name):
        if not candidate.exists():
            continue
        env: dict[str, str] = {}
        try:
            with candidate.open("r", encoding="utf-8") as handle:
                for raw in handle:
                    line = str(raw or "").strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    env[key.strip().lstrip("\ufeff")] = value.strip().strip('"').strip("'")
        except OSError:
            return {}
        return env
    return {}


def _repo_config() -> tuple[str, str, str, str, str]:
    env = _load_env_file(".env")
    owner = DEFAULT_REPO_OWNER
    name = DEFAULT_REPO_NAME
    token = (env.get("GITHUB_TOKEN") or "").strip()
    installer_asset = (
        env.get("GITHUB_INSTALLER_ASSET")
        or env.get("INSTALLER_ASSET_NAME")
        or DEFAULT_INSTALLER_ASSET
    ).strip()
    hash_asset = (env.get("GITHUB_HASH_ASSET") or f"{installer_asset}.sha256").strip()
    _log_update(
        "Repo config loaded: "
        f"owner={owner}, repo={name}, token={'yes' if token else 'no'}, "
        f"installer_asset={installer_asset}, hash_asset={hash_asset}"
    )
    return owner, name, token, installer_asset, hash_asset


def get_release_page_url(version: str | None = None) -> str:
    owner, repo, _token, _installer_asset, _hash_asset = _repo_config()
    clean_version = str(version or "").strip().lstrip("v")
    if clean_version:
        return f"https://github.com/{owner}/{repo}/releases/tag/v{clean_version}"
    return f"https://github.com/{owner}/{repo}/releases/latest"


def _json_safe(value):
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    return value


def _atomic_write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_name(f"{path.name}.{uuid.uuid4().hex}.tmp")
    tmp_path.write_text(content, encoding="utf-8")
    os.replace(tmp_path, path)


def _atomic_write_json(path: Path, payload: dict) -> None:
    serialized = json.dumps(_json_safe(payload), ensure_ascii=False, indent=2, sort_keys=True)
    _atomic_write_text(path, serialized + "\n")


def _read_json(path: Path | str | None) -> dict:
    target = Path(path) if path else None
    if target is None or not target.exists():
        return {}
    try:
        return json.loads(target.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _update_state_root() -> Path:
    base = str(os.getenv("LOCALAPPDATA") or "").strip()
    if base:
        root = Path(base) / "RECA Inclusion Laboral" / "updates"
    else:
        root = Path(tempfile.gettempdir()) / "RECA" / "updates"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _update_pointer_path() -> Path:
    return _update_state_root() / UPDATE_POINTER_FILE_NAME


def _update_session_root() -> Path:
    root = Path(tempfile.gettempdir()) / DEFAULT_UPDATE_SESSION_ROOT
    root.mkdir(parents=True, exist_ok=True)
    return root


def create_update_session_dir() -> Path:
    session_dir = _update_session_root() / uuid.uuid4().hex
    session_dir.mkdir(parents=True, exist_ok=True)
    return session_dir


def get_installed_updater_helper_path(app_executable: str | os.PathLike | None = None) -> Path:
    executable = str(app_executable or sys.executable or "").strip()
    base_dir = Path(executable).resolve().parent if executable else Path.cwd()
    return base_dir / DEFAULT_UPDATE_HELPER_NAME


def get_installed_version_file_candidates(
    app_executable: str | os.PathLike | None = None,
) -> list[Path]:
    executable = str(app_executable or sys.executable or "").strip()
    base_dir = Path(executable).resolve().parent if executable else Path.cwd()
    return [
        base_dir / "_internal" / "VERSION",
        base_dir / "VERSION",
    ]


def is_self_update_supported(
    app_executable: str | os.PathLike | None = None,
    frozen: bool | None = None,
) -> bool:
    runtime_frozen = getattr(sys, "frozen", False) if frozen is None else bool(frozen)
    if not runtime_frozen:
        return False
    return get_installed_updater_helper_path(app_executable=app_executable).exists()


def _http_get_json(url: str, timeout: int = 20, token: str = "") -> dict:
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "reca-inclusion-laboral-updater",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, headers=headers)
    _log_update(
        f"HTTP GET JSON start: url={url}, timeout={timeout}, auth={'yes' if token else 'no'}"
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        status = getattr(response, "status", "?")
        _log_update(f"HTTP GET JSON response: status={status}, url={response.geturl()}")
        payload = response.read().decode("utf-8", errors="replace")
    _log_update(f"HTTP GET JSON payload length: {len(payload)}")
    return json.loads(payload)


def _latest_release_via_redirect(owner: str, repo: str, timeout: int = 20) -> str | None:
    url = f"https://github.com/{owner}/{repo}/releases/latest"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "reca-inclusion-laboral-updater"},
    )
    _log_update(f"Fallback redirect start: {url}")
    with urllib.request.urlopen(req, timeout=timeout) as response:
        final_url = str(response.geturl() or "")
    _log_update(f"Fallback redirect final url: {final_url}")
    match = re.search(r"/releases/tag/([^/?#]+)", final_url)
    if not match:
        return None
    return str(match.group(1)).lstrip("v")


def _latest_version_via_github_raw(owner: str, repo: str, timeout: int = 20) -> str | None:
    url = f"https://github.com/{owner}/{repo}/raw/main/VERSION"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "reca-inclusion-laboral-updater"},
    )
    _log_update(f"Fallback github.com raw VERSION start: {url}")
    with urllib.request.urlopen(req, timeout=timeout) as response:
        status = getattr(response, "status", "?")
        _log_update(f"Fallback github.com raw VERSION response: status={status}")
        payload = response.read().decode("utf-8", errors="replace")
    version = (payload or "").strip().splitlines()[0].strip().lstrip("v")
    if not re.match(r"^\d+\.\d+\.\d+$", version):
        return None
    return version


def _latest_version_via_raw(owner: str, repo: str, timeout: int = 20) -> str | None:
    url = f"https://raw.githubusercontent.com/{owner}/{repo}/main/VERSION"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "reca-inclusion-laboral-updater"},
    )
    _log_update(f"Fallback raw VERSION start: {url}")
    with urllib.request.urlopen(req, timeout=timeout) as response:
        status = getattr(response, "status", "?")
        _log_update(f"Fallback raw VERSION response: status={status}")
        payload = response.read().decode("utf-8", errors="replace")
    version = (payload or "").strip().splitlines()[0].strip().lstrip("v")
    if not re.match(r"^\d+\.\d+\.\d+$", version):
        return None
    return version


def _get_latest_release() -> tuple[str | None, dict]:
    owner, repo, token, installer_asset, hash_asset = _repo_config()
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    _log_update(f"Resolve latest release start: {api_url}")

    def _assets_for_version(version: str) -> dict:
        tag = f"v{version}"
        return {
            installer_asset: f"https://github.com/{owner}/{repo}/releases/download/{tag}/{installer_asset}",
            hash_asset: f"https://github.com/{owner}/{repo}/releases/download/{tag}/{hash_asset}",
        }

    def _fallback_version_lookup() -> tuple[str | None, dict]:
        try:
            version = _latest_release_via_redirect(owner, repo, timeout=20)
            if version:
                _log_update(f"FALLBACK releases/latest redirect OK: v{version}")
                return version, _assets_for_version(version)
        except Exception as fallback_exc:
            _log_update(f"ERROR fallback release/latest redirect: {fallback_exc}")
        try:
            version = _latest_version_via_github_raw(owner, repo, timeout=20)
            if version:
                _log_update(f"FALLBACK github.com raw VERSION OK: v{version}")
                return version, _assets_for_version(version)
        except Exception as fallback_exc:
            _log_update(f"ERROR fallback github.com raw VERSION: {fallback_exc}")
        try:
            version = _latest_version_via_raw(owner, repo, timeout=20)
            if version:
                _log_update(f"FALLBACK raw VERSION OK: v{version}")
                return version, _assets_for_version(version)
        except Exception as fallback_exc:
            _log_update(f"ERROR fallback raw VERSION: {fallback_exc}")
        return None, {}

    try:
        data = _http_get_json(api_url, timeout=20, token=token)
    except urllib.error.HTTPError as exc:
        code = getattr(exc, "code", "?")
        detail = ""
        try:
            detail = exc.read().decode("utf-8", errors="replace").strip()
        except Exception:
            pass
        _log_update(f"ERROR release/latest HTTP {code}: {exc} {detail}")
        return _fallback_version_lookup()
    except Exception as exc:
        _log_update(f"ERROR release/latest: {exc}")
        return _fallback_version_lookup()

    remote_version = str(data.get("tag_name", "")).lstrip("v")
    assets = {}
    for asset in data.get("assets", []) or []:
        name = asset.get("name")
        url = asset.get("browser_download_url")
        if name and url:
            assets[str(name)] = str(url)
    _log_update(
        f"Resolve latest release success: remote_version={remote_version or 'None'}, "
        f"assets={list(assets.keys())}"
    )
    return remote_version or None, assets


def get_latest_version() -> str | None:
    remote_version, _ = _get_latest_release()
    return remote_version


def get_latest_release_assets() -> tuple[str | None, dict]:
    return _get_latest_release()


def _parse_version(value: str | None) -> tuple[int, ...]:
    if not value:
        return ()
    cleaned = str(value).strip().lstrip("v")
    parts = []
    for chunk in cleaned.split("."):
        try:
            parts.append(int(chunk))
        except ValueError:
            parts.append(0)
    return tuple(parts)


def is_update_available(local_version: str | None, remote_version: str | None) -> bool:
    if not local_version or not remote_version:
        return False
    return _parse_version(remote_version) > _parse_version(local_version)


def _download_file(url: str, destination: Path, progress_callback=None) -> None:
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "reca-inclusion-laboral-updater"},
    )
    _log_update(f"Download start: url={url}, destination={destination}")
    try:
        with urllib.request.urlopen(req, timeout=40) as response:
            total = int(response.headers.get("Content-Length") or 0)
            status = getattr(response, "status", "?")
            _log_update(f"Download response: status={status}, total_bytes={total}")
            downloaded = 0
            destination.parent.mkdir(parents=True, exist_ok=True)
            with destination.open("wb") as handle:
                while True:
                    chunk = response.read(1024 * 256)
                    if not chunk:
                        break
                    handle.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total:
                        percent = int((downloaded / total) * 100)
                        progress_callback("Descargando instalador...", max(1, min(99, percent)))
        _log_update(f"Download completed: bytes={downloaded}, destination={destination}")
        return
    except Exception as exc:
        _log_update(f"ERROR urllib download: {exc}")
    raise RuntimeError(
        "No se pudo descargar el instalador. "
        "Revisa la conexión o descarga manualmente el instalador desde el release."
    )


def _verify_hash(installer_path: Path, assets: dict) -> None:
    _owner, _repo, _token, _installer_asset, hash_asset = _repo_config()
    url = assets.get(hash_asset)
    if not url:
        _log_update("Hash asset not present in release assets; skipping hash verification.")
        return
    hash_path = installer_path.with_suffix(".sha256")
    _download_file(url, hash_path)
    expected_line = hash_path.read_text(encoding="utf-8", errors="replace").strip()
    expected = expected_line.split()[0] if expected_line else ""
    if not expected:
        _log_update("Hash file downloaded but expected digest is empty; skipping hash verification.")
        return
    digest = hashlib.sha256(installer_path.read_bytes()).hexdigest()
    _log_update(f"Hash verification: expected={expected.lower()} actual={digest.lower()}")
    if expected.lower() != digest.lower():
        raise RuntimeError("Hash del instalador no coincide.")


def calculate_file_sha256(path: Path | str) -> str:
    target = Path(path)
    return hashlib.sha256(target.read_bytes()).hexdigest()


def download_installer(assets: dict, progress_callback=None, target_dir: Path | str | None = None) -> Path:
    _owner, _repo, _token, installer_asset, _hash_asset = _repo_config()
    url = assets.get(installer_asset)
    if not url:
        raise RuntimeError(f"No se encontró el instalador '{installer_asset}' en el release.")
    if target_dir is None:
        installer_path = Path(tempfile.gettempdir()) / installer_asset
    else:
        installer_path = Path(target_dir) / installer_asset
    _log_update(
        f"Download installer selected: asset={installer_asset}, url={url}, target={installer_path}"
    )
    _download_file(url, installer_path, progress_callback=progress_callback)
    _verify_hash(installer_path, assets)
    _log_update(f"Download installer completed: {installer_path}")
    return installer_path


def _installer_args(installer_path: Path, installer_log_path: Path | None = None) -> list[str]:
    args = [
        str(installer_path),
        "/VERYSILENT",
        "/CURRENTUSER",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
    ]
    if installer_log_path is not None:
        args.append(f"/LOG={installer_log_path}")
    return args


def build_update_manifest(
    *,
    session_dir: Path | str,
    installer_path: Path | str,
    expected_version: str,
    current_pid: int,
    relaunch_command: list[str] | tuple[str, ...] | None,
    installed_app_path: Path | str,
    installed_version_paths: list[Path | str] | tuple[Path | str, ...],
    release_url: str = "",
) -> dict:
    session_path = Path(session_dir)
    installer_path = Path(installer_path)
    installer_log_path = session_path / "installer.log"
    status_path = session_path / UPDATE_STATUS_FILE_NAME
    ack_path = session_path / UPDATE_READY_ACK_FILE_NAME
    manifest_path = session_path / UPDATE_MANIFEST_FILE_NAME
    manifest = {
        "session_dir": str(session_path),
        "manifest_path": str(manifest_path),
        "target_pid": int(current_pid),
        "expected_version": str(expected_version or "").strip().lstrip("v"),
        "installer_path": str(installer_path),
        "installer_sha256": calculate_file_sha256(installer_path),
        "installer_log_path": str(installer_log_path),
        "installed_app_path": str(Path(installed_app_path)),
        "installed_version_paths": [str(Path(item)) for item in installed_version_paths],
        "status_path": str(status_path),
        "ack_path": str(ack_path),
        "relaunch_args": [str(item) for item in (relaunch_command or [])],
        "release_url": str(release_url or "").strip(),
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    return manifest


def write_update_status(status_path: Path | str, state: str, message: str, **extra) -> dict:
    payload = {
        "state": str(state or "").strip(),
        "message": str(message or "").strip(),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    if extra:
        payload.update(_json_safe(extra))
    target = Path(status_path)
    _atomic_write_json(target, payload)
    _log_update(
        f"Update status written: state={payload['state']} message={payload['message']} path={target}"
    )
    return payload


def _write_pending_update_pointer(manifest: dict) -> Path:
    pointer_path = _update_pointer_path()
    pointer_payload = {
        "session_dir": manifest.get("session_dir"),
        "manifest_path": manifest.get("manifest_path"),
        "status_path": manifest.get("status_path"),
        "expected_version": manifest.get("expected_version"),
        "release_url": manifest.get("release_url"),
        "created_at": manifest.get("created_at"),
    }
    _atomic_write_json(pointer_path, pointer_payload)
    _log_update(f"Pending update pointer stored: {pointer_path}")
    return pointer_path


def clear_pending_update_session(
    session: dict | None = None,
    *,
    remove_session_dir: bool = False,
) -> None:
    pointer_path = _update_pointer_path()
    pointer_payload = dict(session or {})
    if not pointer_payload and pointer_path.exists():
        pointer_payload = _read_json(pointer_path)
    try:
        pointer_path.unlink(missing_ok=True)
    except Exception:
        pass
    if remove_session_dir:
        session_dir = Path(str(pointer_payload.get("session_dir") or "").strip()) if pointer_payload else None
        if session_dir and session_dir.exists():
            try:
                shutil.rmtree(session_dir, ignore_errors=True)
            except Exception:
                pass


def _creation_flags() -> int:
    return (
        getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        | getattr(subprocess, "DETACHED_PROCESS", 0)
        | getattr(subprocess, "CREATE_NO_WINDOW", 0)
    )


def _launch_detached_process(command: list[str]) -> subprocess.Popen:
    return subprocess.Popen(
        command,
        close_fds=True,
        creationflags=_creation_flags(),
    )


def _wait_for_helper_ready(
    proc: subprocess.Popen,
    ack_path: Path,
    status_path: Path,
    timeout_seconds: float,
) -> None:
    deadline = time.monotonic() + max(1.0, float(timeout_seconds))
    while time.monotonic() < deadline:
        if ack_path.exists():
            return
        if proc.poll() is not None:
            raise RuntimeError("El helper de actualización terminó antes de confirmar arranque.")
        status = _read_json(status_path)
        state = str(status.get("state") or "").strip().lower()
        if state == "failed":
            raise RuntimeError(str(status.get("message") or "El helper reportó un error de arranque."))
        time.sleep(0.1)
    raise RuntimeError("El helper de actualización no confirmó arranque a tiempo.")


def start_update_handoff(
    *,
    installer_path: Path | str,
    expected_version: str,
    current_pid: int,
    relaunch_command: list[str] | tuple[str, ...] | None,
    app_executable: str | os.PathLike | None = None,
    release_url: str = "",
    ack_timeout: float = 15.0,
) -> dict:
    app_path = Path(str(app_executable or sys.executable or "")).resolve()
    helper_source_path = get_installed_updater_helper_path(app_executable=app_path)
    if not helper_source_path.exists():
        raise RuntimeError(
            f"No se encontró el helper de actualización en la instalación ({helper_source_path})."
        )
    installer_path = Path(installer_path).resolve()
    session_dir = installer_path.parent
    session_dir.mkdir(parents=True, exist_ok=True)
    helper_runtime_path = session_dir / DEFAULT_UPDATE_HELPER_NAME
    shutil.copy2(helper_source_path, helper_runtime_path)
    manifest = build_update_manifest(
        session_dir=session_dir,
        installer_path=installer_path,
        expected_version=expected_version,
        current_pid=current_pid,
        relaunch_command=relaunch_command,
        installed_app_path=app_path,
        installed_version_paths=get_installed_version_file_candidates(app_executable=app_path),
        release_url=release_url or get_release_page_url(expected_version),
    )
    manifest_path = Path(manifest["manifest_path"])
    status_path = Path(manifest["status_path"])
    ack_path = Path(manifest["ack_path"])
    _atomic_write_json(manifest_path, manifest)
    _write_pending_update_pointer(manifest)
    command = [
        str(helper_runtime_path),
        "--manifest",
        str(manifest_path),
    ]
    _log_update(
        "Launching update helper: "
        f"helper={helper_runtime_path}, manifest={manifest_path}, target_pid={current_pid}"
    )
    proc = _launch_detached_process(command)
    try:
        _wait_for_helper_ready(proc, ack_path, status_path, timeout_seconds=ack_timeout)
    except Exception:
        clear_pending_update_session(remove_session_dir=False)
        raise
    return {
        "session_dir": str(session_dir),
        "manifest_path": str(manifest_path),
        "status_path": str(status_path),
        "ack_path": str(ack_path),
        "helper_path": str(helper_runtime_path),
        "helper_pid": getattr(proc, "pid", None),
        "installer_log_path": manifest.get("installer_log_path"),
        "expected_version": manifest.get("expected_version"),
        "release_url": manifest.get("release_url"),
    }


def _read_version_from_paths(paths: list[Path | str] | tuple[Path | str, ...]) -> str:
    for candidate in paths or []:
        try:
            payload = Path(candidate).read_text(encoding="utf-8-sig").strip()
        except Exception:
            continue
        if payload:
            return payload.lstrip("v")
    return ""


def _wait_for_process_exit_windows(pid: int, timeout_seconds: float | None = None) -> None:
    try:
        import ctypes

        synchronize = 0x00100000
        query_limited = 0x1000
        wait_object_0 = 0x00000000
        wait_timeout = 0x00000102
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(synchronize | query_limited, False, int(pid))
        if not handle:
            return
        try:
            if timeout_seconds is None:
                timeout_ms = 0xFFFFFFFF
            else:
                timeout_ms = max(0, int(float(timeout_seconds) * 1000))
            result = kernel32.WaitForSingleObject(handle, timeout_ms)
            if result in {wait_object_0, wait_timeout}:
                return
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        pass


def _wait_for_process_exit(pid: int, timeout_seconds: float | None = None) -> None:
    if os.name == "nt":
        _wait_for_process_exit_windows(pid, timeout_seconds=timeout_seconds)
        return
    deadline = None if timeout_seconds is None else time.monotonic() + float(timeout_seconds)
    while True:
        try:
            os.kill(int(pid), 0)
        except OSError:
            return
        if deadline is not None and time.monotonic() >= deadline:
            return
        time.sleep(0.2)


def _coerce_returncode(result) -> int:
    if hasattr(result, "returncode"):
        try:
            return int(getattr(result, "returncode"))
        except Exception:
            return 1
    try:
        return int(result)
    except Exception:
        return 1


def _run_installer_command(args: list[str]):
    return subprocess.run(args, check=False, close_fds=True)


def _relaunch_application(command: list[str]) -> subprocess.Popen:
    return _launch_detached_process(command)


def execute_update_manifest(
    manifest_path: Path | str,
    *,
    wait_for_target=None,
    installer_runner=None,
    relaunch_runner=None,
) -> dict:
    manifest = _read_json(manifest_path)
    if not manifest:
        raise RuntimeError("No se pudo leer el manifest de actualización.")
    status_path = Path(manifest["status_path"])
    ack_path = Path(
        str(manifest.get("ack_path") or (Path(manifest["session_dir"]) / UPDATE_READY_ACK_FILE_NAME))
    )
    installer_path = Path(manifest["installer_path"])
    installer_log_path = Path(manifest["installer_log_path"])
    expected_version = str(manifest.get("expected_version") or "").strip().lstrip("v")
    target_pid = int(manifest.get("target_pid") or 0)
    relaunch_args = [str(item) for item in (manifest.get("relaunch_args") or [])]
    installed_version_paths = [Path(item) for item in (manifest.get("installed_version_paths") or [])]

    write_update_status(
        status_path,
        "starting",
        "El helper de actualización inició.",
        helper_pid=os.getpid(),
        manifest_path=str(manifest_path),
    )
    ack_path.parent.mkdir(parents=True, exist_ok=True)
    _atomic_write_text(ack_path, f"{os.getpid()}\n")
    write_update_status(
        status_path,
        "ready",
        "El helper confirmó arranque y esperará el cierre de la app.",
        helper_pid=os.getpid(),
        ack_path=str(ack_path),
    )

    wait_fn = wait_for_target or _wait_for_process_exit
    write_update_status(
        status_path,
        "waiting_for_exit",
        "Esperando a que la aplicación se cierre por completo.",
        target_pid=target_pid,
    )
    wait_fn(target_pid)

    installer_args = _installer_args(installer_path, installer_log_path=installer_log_path)
    write_update_status(
        status_path,
        "installing",
        "Ejecutando el instalador.",
        installer_log_path=str(installer_log_path),
        installer_command=installer_args,
    )
    runner = installer_runner or _run_installer_command
    completed = runner(installer_args)
    installer_exit_code = _coerce_returncode(completed)
    if installer_exit_code != 0:
        return write_update_status(
            status_path,
            "failed",
            f"La instalación no se completó correctamente (código {installer_exit_code}).",
            installer_exit_code=installer_exit_code,
            installer_log_path=str(installer_log_path),
        )
    if not installer_log_path.exists():
        return write_update_status(
            status_path,
            "failed",
            "La instalación terminó sin generar el log esperado.",
            installer_exit_code=installer_exit_code,
            installer_log_path=str(installer_log_path),
        )

    write_update_status(
        status_path,
        "verifying_install",
        "Verificando la versión instalada.",
        installer_exit_code=installer_exit_code,
        installer_log_path=str(installer_log_path),
    )
    detected_version = _read_version_from_paths(installed_version_paths)
    if not detected_version or detected_version != expected_version:
        return write_update_status(
            status_path,
            "failed",
            "La instalación terminó pero la versión instalada no coincide con la esperada.",
            installer_exit_code=installer_exit_code,
            detected_installed_version=detected_version or None,
            expected_version=expected_version or None,
            installer_log_path=str(installer_log_path),
        )

    if relaunch_args:
        relaunch = relaunch_runner or _relaunch_application
        relaunch(relaunch_args)

    return write_update_status(
        status_path,
        "succeeded",
        "La actualización se aplicó correctamente.",
        installer_exit_code=installer_exit_code,
        detected_installed_version=detected_version,
        expected_version=expected_version,
        installer_log_path=str(installer_log_path),
        relaunch_args=relaunch_args,
    )


def mark_update_failed_from_manifest(manifest_path: Path | str, message: str) -> dict | None:
    manifest = _read_json(manifest_path)
    if not manifest:
        return None
    status_path = manifest.get("status_path")
    if not status_path:
        return None
    return write_update_status(
        status_path,
        "failed",
        str(message or "El helper de actualización terminó con error."),
        helper_pid=os.getpid(),
        installer_log_path=manifest.get("installer_log_path"),
    )


def inspect_pending_update(current_version: str | None = None) -> dict | None:
    pointer = _read_json(_update_pointer_path())
    if not pointer:
        return None
    manifest_path = Path(str(pointer.get("manifest_path") or "").strip()) if pointer.get("manifest_path") else None
    status_path = Path(str(pointer.get("status_path") or "").strip()) if pointer.get("status_path") else None
    manifest = _read_json(manifest_path)
    status = _read_json(status_path)
    expected_version = str((manifest or pointer).get("expected_version") or "").strip().lstrip("v")
    release_url = str((manifest or pointer).get("release_url") or "").strip()
    installer_log_path = str((status or manifest or {}).get("installer_log_path") or "").strip()
    state = str(status.get("state") or "").strip().lower()
    message = str(status.get("message") or "").strip()
    session = {
        "session_dir": str((manifest or pointer).get("session_dir") or ""),
        "manifest_path": str(manifest_path or ""),
        "status_path": str(status_path or ""),
        "expected_version": expected_version or None,
        "release_url": release_url or None,
        "installer_log_path": installer_log_path or None,
        "status": status,
        "manifest": manifest,
    }
    if state == "succeeded":
        local = str(current_version or "").strip().lstrip("v")
        detected_version = str(status.get("detected_installed_version") or "").strip().lstrip("v")
        if expected_version and local and local != expected_version:
            session.update(
                {
                    "outcome": "failed",
                    "message": (
                        "La actualización reportó éxito, pero la versión actual sigue siendo "
                        f"{local} en vez de {expected_version}."
                    ),
                    "detected_installed_version": detected_version or None,
                }
            )
            return session
        session.update(
            {
                "outcome": "succeeded",
                "message": message or "La actualización terminó correctamente.",
                "detected_installed_version": detected_version or None,
            }
        )
        return session
    if state == "failed":
        session.update(
            {
                "outcome": "failed",
                "message": message or "La actualización anterior falló.",
                "detected_installed_version": status.get("detected_installed_version"),
                "installer_exit_code": status.get("installer_exit_code"),
            }
        )
        return session
    session.update(
        {
            "outcome": "interrupted",
            "message": (
                message
                or "La actualización anterior quedó interrumpida antes de completarse."
            ),
            "detected_installed_version": status.get("detected_installed_version"),
            "installer_exit_code": status.get("installer_exit_code"),
        }
    )
    return session


def run_installer(installer_path: Path, wait: bool = True) -> None:
    args = _installer_args(installer_path)
    _log_update(f"Run installer: path={installer_path}, wait={wait}, args={args}")
    if wait:
        completed = subprocess.run(args, check=False)
        _log_update(f"Installer process finished: returncode={completed.returncode}")
        if completed.returncode != 0:
            raise RuntimeError(
                "La instalación no se completó correctamente "
                f"(código {completed.returncode})."
            )
    else:
        proc = subprocess.Popen(args, close_fds=True)
        _log_update(f"Installer process started: pid={getattr(proc, 'pid', '?')}")
