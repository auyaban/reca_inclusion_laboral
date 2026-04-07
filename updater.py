from __future__ import annotations

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

from version_info import appdata_logs_dir


DEFAULT_REPO_OWNER = "auyaban"
DEFAULT_REPO_NAME = "reca_inclusion_laboral"
DEFAULT_INSTALLER_ASSET = "RECA_INCLUSION_LABORAL_Setup.exe"
DEFAULT_HASH_ASSET = f"{DEFAULT_INSTALLER_ASSET}.sha256"
_LOG_CURRENT_DAY = None


def _update_log_path() -> Path:
    return appdata_logs_dir() / "updater.log"


def _ensure_daily_log(path: Path) -> None:
    global _LOG_CURRENT_DAY
    today = datetime.now().strftime("%Y-%m-%d")
    if _LOG_CURRENT_DAY == today:
        return
    _LOG_CURRENT_DAY = today
    try:
        with path.open("a", encoding="utf-8") as handle:
            handle.write(f"\n{'='*60}\n{today}\n{'='*60}\n")
    except Exception:
        pass


def _log_update(message: str) -> None:
    try:
        path = _update_log_path()
        _ensure_daily_log(path)
        with path.open("a", encoding="utf-8") as handle:
            ts = datetime.now().strftime("%H:%M:%S")
            handle.write(f"[{ts}] {message.rstrip()}\n")
    except Exception:
        pass


def _resolve_env_candidates(env_name: str = ".env") -> list[Path]:
    candidates = []
    roaming = os.getenv("APPDATA") or ""
    if roaming:
        candidates.append(Path(roaming) / "RECA Inclusion Laboral" / env_name)
    exe = getattr(__import__("sys"), "executable", None)
    if exe:
        candidates.append(Path(exe).parent / env_name)
    candidates.append(Path(env_name))
    return candidates


def _load_env_file(env_name: str = ".env") -> dict:
    for path in _resolve_env_candidates(env_name):
        try:
            lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
            result = {}
            for line in lines:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, _, value = line.partition("=")
                    result[key.strip()] = value.strip()
            if result:
                return result
        except Exception:
            continue
    return {}


def _repo_config() -> tuple[str, str, str, str, str]:
    env = _load_env_file(".env")
    owner = (env.get("GITHUB_REPO_OWNER") or DEFAULT_REPO_OWNER).strip()
    name = (env.get("GITHUB_REPO_NAME") or DEFAULT_REPO_NAME).strip()
    token = (env.get("GITHUB_TOKEN") or "").strip()
    installer_asset = (
        env.get("GITHUB_INSTALLER_ASSET")
        or env.get("INSTALLER_ASSET_NAME")
        or DEFAULT_INSTALLER_ASSET
    ).strip()
    hash_asset = (env.get("GITHUB_HASH_ASSET") or f"{installer_asset}.sha256").strip()
    return owner, name, token, installer_asset, hash_asset


def get_release_page_url(version: str | None = None) -> str:
    owner, repo, _token, _installer_asset, _hash_asset = _repo_config()
    base = f"https://github.com/{owner}/{repo}/releases"
    if version:
        tag = str(version).strip().lstrip("v")
        return f"{base}/tag/v{tag}"
    return f"{base}/latest"


def _http_get_json(url: str, timeout: int = 20, token: str = "") -> dict:
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": "reca-inclusion-laboral-updater",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as response:
        payload = response.read().decode("utf-8", errors="replace")
    return json.loads(payload)


def _latest_release_via_redirect(owner: str, repo: str, timeout: int = 20) -> str | None:
    url = f"https://github.com/{owner}/{repo}/releases/latest"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "reca-inclusion-laboral-updater"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        final_url = str(response.geturl() or "")
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
    with urllib.request.urlopen(req, timeout=timeout) as response:
        content = response.read().decode("utf-8", errors="replace").strip()
    version = content.lstrip("v").strip()
    return version if version else None


def _latest_version_via_raw(owner: str, repo: str, timeout: int = 20) -> str | None:
    url = f"https://raw.githubusercontent.com/{owner}/{repo}/main/VERSION"
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "reca-inclusion-laboral-updater"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as response:
        content = response.read().decode("utf-8", errors="replace").strip()
    version = content.lstrip("v").strip()
    return version if version else None


def _assets_for_version(version: str) -> dict:
    owner, repo, _token, installer_asset, hash_asset = _repo_config()
    tag = f"v{version}"
    base = f"https://github.com/{owner}/{repo}/releases/download/{tag}"
    return {
        installer_asset: f"{base}/{installer_asset}",
        hash_asset: f"{base}/{hash_asset}",
    }


def _get_latest_release() -> tuple[str | None, dict]:
    owner, repo, token, installer_asset, hash_asset = _repo_config()
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    try:
        data = _http_get_json(api_url, timeout=20, token=token)
        remote_version = str(data.get("tag_name", "")).lstrip("v")
        assets = {}
        for asset in data.get("assets", []) or []:
            name = asset.get("name")
            url = asset.get("browser_download_url")
            if name and url:
                assets[str(name)] = str(url)
        if remote_version:
            _log_update(f"release/latest API OK: v{remote_version}")
            return remote_version, assets
    except urllib.error.HTTPError as exc:
        _log_update(f"ERROR release/latest HTTP {getattr(exc, 'code', '?')}: {exc}")
    except Exception as exc:
        _log_update(f"ERROR release/latest: {exc}")

    # Fallback 1: redirect público desde github.com
    try:
        version = _latest_release_via_redirect(owner, repo, timeout=20)
        if version:
            _log_update(f"FALLBACK releases/latest redirect OK: v{version}")
            return version, _assets_for_version(version)
    except Exception as exc:
        _log_update(f"ERROR fallback release/latest redirect: {exc}")

    # Fallback 2: VERSION desde github.com/raw
    try:
        version = _latest_version_via_github_raw(owner, repo, timeout=20)
        if version:
            _log_update(f"FALLBACK github.com raw VERSION OK: v{version}")
            return version, _assets_for_version(version)
    except Exception as exc:
        _log_update(f"ERROR fallback github.com raw VERSION: {exc}")

    # Fallback 3: VERSION desde raw.githubusercontent.com
    try:
        version = _latest_version_via_raw(owner, repo, timeout=20)
        if version:
            _log_update(f"FALLBACK raw.githubusercontent VERSION OK: v{version}")
            return version, _assets_for_version(version)
    except Exception as exc:
        _log_update(f"ERROR fallback raw VERSION: {exc}")

    return None, {}


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
