import hashlib
import json
import os
import re
import subprocess
import tempfile
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

from formularios.common import _load_env_file
from version_info import appdata_logs_dir


DEFAULT_REPO_OWNER = "auyaban"
DEFAULT_REPO_NAME = "reca_inclusion_laboral"
DEFAULT_INSTALLER_ASSET = "RECA_INCLUSION_LABORAL_Setup.exe"
DEFAULT_HASH_ASSET = f"{DEFAULT_INSTALLER_ASSET}.sha256"
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


def _repo_config():
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


def _http_get_json(url: str, timeout: int = 20, token: str = "") -> dict:
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "reca-inclusion-laboral-updater",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(
        url,
        headers=headers,
    )
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
        # 1) Redirect publico de releases/latest.
        try:
            version = _latest_release_via_redirect(owner, repo, timeout=20)
            if version:
                _log_update(f"FALLBACK releases/latest redirect OK: v{version}")
                return version, _assets_for_version(version)
        except Exception as fallback_exc:
            _log_update(f"ERROR fallback release/latest redirect: {fallback_exc}")
        # 2) VERSION servido desde github.com.
        try:
            version = _latest_version_via_github_raw(owner, repo, timeout=20)
            if version:
                _log_update(f"FALLBACK github.com raw VERSION OK: v{version}")
                return version, _assets_for_version(version)
        except Exception as fallback_exc:
            _log_update(f"ERROR fallback github.com raw VERSION: {fallback_exc}")
        # 3) VERSION en raw.githubusercontent.com.
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


def download_installer(assets: dict, progress_callback=None) -> Path:
    _owner, _repo, _token, installer_asset, _hash_asset = _repo_config()
    url = assets.get(installer_asset)
    if not url:
        raise RuntimeError(f"No se encontró el instalador '{installer_asset}' en el release.")
    installer_path = Path(tempfile.gettempdir()) / installer_asset
    _log_update(
        f"Download installer selected: asset={installer_asset}, url={url}, target={installer_path}"
    )
    _download_file(url, installer_path, progress_callback=progress_callback)
    _verify_hash(installer_path, assets)
    _log_update(f"Download installer completed: {installer_path}")
    return installer_path


def _installer_args(installer_path: Path) -> list[str]:
    return [
        str(installer_path),
        "/VERYSILENT",
        "/CURRENTUSER",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
    ]


def _build_post_exit_installer_cmd(
    installer_path: Path,
    current_pid: int,
    relaunch_command: list[str] | tuple[str, ...] | None = None,
) -> Path:
    installer_command = subprocess.list2cmdline(_installer_args(installer_path))
    relaunch_command_line = subprocess.list2cmdline(list(relaunch_command or []))
    script_path = Path(tempfile.gettempdir()) / f"reca_updater_{int(current_pid)}.cmd"
    lines = [
        "@echo off",
        "setlocal",
        ":wait_for_app",
        f'tasklist /FI "PID eq {int(current_pid)}" /FO CSV /NH 2>nul | find /I "\\"{int(current_pid)}\\"" >nul',
        "if not errorlevel 1 (",
        "  timeout /t 1 /nobreak >nul",
        "  goto wait_for_app",
        ")",
        "timeout /t 1 /nobreak >nul",
        f"start \"\" /wait {installer_command}",
        "set \"RECA_INSTALL_EXIT=%ERRORLEVEL%\"",
    ]
    if relaunch_command_line:
        lines.extend(
            [
                "if not \"%RECA_INSTALL_EXIT%\"==\"0\" goto cleanup",
                "timeout /t 1 /nobreak >nul",
                f"start \"\" {relaunch_command_line}",
            ]
        )
    lines.extend(
        [
            ":cleanup",
            "del \"%~f0\" >nul 2>&1",
            "exit /b",
        ]
    )
    script_path.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
    return script_path


def schedule_installer_after_exit(
    installer_path: Path,
    current_pid: int,
    relaunch_command: list[str] | tuple[str, ...] | None = None,
) -> None:
    script_path = _build_post_exit_installer_cmd(
        installer_path,
        current_pid=current_pid,
        relaunch_command=relaunch_command,
    )
    command = [
        "cmd.exe",
        "/c",
        str(script_path),
    ]
    creationflags = (
        getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        | getattr(subprocess, "DETACHED_PROCESS", 0)
        | getattr(subprocess, "CREATE_NO_WINDOW", 0)
    )
    _log_update(
        "Schedule installer after exit: "
        f"path={installer_path}, current_pid={current_pid}, relaunch={list(relaunch_command or [])}, "
        f"script={script_path}"
    )
    proc = subprocess.Popen(
        command,
        close_fds=True,
        creationflags=creationflags,
    )
    _log_update(f"Post-exit installer watcher started: pid={getattr(proc, 'pid', '?')}")


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

