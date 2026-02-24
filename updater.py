import hashlib
import json
import os
import subprocess
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

from formularios.common import _load_env_file
from version_info import appdata_logs_dir


DEFAULT_REPO_OWNER = "auyaban"
DEFAULT_REPO_NAME = "reca_inclusion_laboral"
DEFAULT_INSTALLER_ASSET = "RECA_INCLUSION_LABORAL_Setup.exe"
DEFAULT_HASH_ASSET = f"{DEFAULT_INSTALLER_ASSET}.sha256"


def _update_log_path() -> Path:
    return appdata_logs_dir() / "updater.log"


def _log_update(message: str) -> None:
    try:
        with _update_log_path().open("a", encoding="utf-8") as handle:
            handle.write(message.rstrip() + "\n")
    except Exception:
        pass


def _repo_config():
    env = _load_env_file(".env")
    owner = (env.get("GITHUB_REPO_OWNER") or DEFAULT_REPO_OWNER).strip()
    name = (env.get("GITHUB_REPO_NAME") or DEFAULT_REPO_NAME).strip()
    installer_asset = (
        env.get("GITHUB_INSTALLER_ASSET")
        or env.get("INSTALLER_ASSET_NAME")
        or DEFAULT_INSTALLER_ASSET
    ).strip()
    hash_asset = (env.get("GITHUB_HASH_ASSET") or f"{installer_asset}.sha256").strip()
    return owner, name, installer_asset, hash_asset


def _http_get_json(url: str, timeout: int = 20) -> dict:
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "reca-inclusion-laboral-updater",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        payload = response.read().decode("utf-8", errors="replace")
    return json.loads(payload)


def _get_latest_release() -> tuple[str | None, dict]:
    owner, repo, _installer_asset, _hash_asset = _repo_config()
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    try:
        data = _http_get_json(api_url, timeout=20)
    except urllib.error.HTTPError as exc:
        _log_update(f"ERROR release/latest HTTP {getattr(exc, 'code', '?')}: {exc}")
        return None, {}
    except Exception as exc:
        _log_update(f"ERROR release/latest: {exc}")
        return None, {}

    remote_version = str(data.get("tag_name", "")).lstrip("v")
    assets = {}
    for asset in data.get("assets", []) or []:
        name = asset.get("name")
        url = asset.get("browser_download_url")
        if name and url:
            assets[str(name)] = str(url)
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
    with urllib.request.urlopen(req, timeout=40) as response:
        total = int(response.headers.get("Content-Length") or 0)
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


def _verify_hash(installer_path: Path, assets: dict) -> None:
    _owner, _repo, _installer_asset, hash_asset = _repo_config()
    url = assets.get(hash_asset)
    if not url:
        return
    hash_path = installer_path.with_suffix(".sha256")
    _download_file(url, hash_path)
    expected_line = hash_path.read_text(encoding="utf-8", errors="replace").strip()
    expected = expected_line.split()[0] if expected_line else ""
    if not expected:
        return
    digest = hashlib.sha256(installer_path.read_bytes()).hexdigest()
    if expected.lower() != digest.lower():
        raise RuntimeError("Hash del instalador no coincide.")


def download_installer(assets: dict, progress_callback=None) -> Path:
    _owner, _repo, installer_asset, _hash_asset = _repo_config()
    url = assets.get(installer_asset)
    if not url:
        raise RuntimeError(f"No se encontró el instalador '{installer_asset}' en el release.")
    installer_path = Path(tempfile.gettempdir()) / installer_asset
    _download_file(url, installer_path, progress_callback=progress_callback)
    _verify_hash(installer_path, assets)
    return installer_path


def run_installer(installer_path: Path, wait: bool = True) -> None:
    args = [
        str(installer_path),
        "/VERYSILENT",
        "/CURRENTUSER",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
    ]
    if wait:
        subprocess.run(args, check=False)
    else:
        subprocess.Popen(args, close_fds=True)
