import json
import os
import re
import sys
import time


SCOPE = "https://www.googleapis.com/auth/drive.file"
DEFAULT_CONFIG_PATH = "config.json"


def _get_bundle_dir():
    """Return the directory that contains bundled/sibling files.

    When running as a PyInstaller frozen executable the extracted files live
    in ``sys._MEIPASS`` (the ``_internal`` folder next to the .exe).  When
    running from source the files live next to this module.
    """
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def _sanitize_filename(value):
    safe = re.sub(r"[\\/:*?\"<>|]+", " ", str(value or ""))
    safe = re.sub(r"\s+", " ", safe).strip()
    return safe or "archivo"


def _get_credentials_path():
    path = os.getenv("GOOGLE_DRIVE_SA_JSON")
    if not path:
        config = _load_config()
        path = config.get("google_drive_sa_json")
    if not path:
        raise RuntimeError(
            "Falta GOOGLE_DRIVE_SA_JSON o config.json con google_drive_sa_json."
        )
    # Resolve relative paths against the bundle / script directory so the app
    # works correctly whether running from source or as a PyInstaller bundle.
    if not os.path.isabs(path):
        path = os.path.join(_get_bundle_dir(), path)
    if not os.path.exists(path):
        raise RuntimeError(f"No existe el JSON de credenciales: {path}")
    return path


def _get_folder_id():
    env_folder = str(os.getenv("GOOGLE_DRIVE_FOLDER_ID") or "").strip()
    if env_folder:
        return env_folder
    config = _load_config()
    cfg_folder = str(config.get("google_drive_folder_id") or "").strip()
    if cfg_folder:
        return cfg_folder
    raise RuntimeError(
        "Falta GOOGLE_DRIVE_FOLDER_ID o config.json con google_drive_folder_id."
    )


def _get_excel_folder_id():
    if os.getenv("GOOGLE_DRIVE_EXCEL_FOLDER_ID"):
        return os.getenv("GOOGLE_DRIVE_EXCEL_FOLDER_ID")
    config = _load_config()
    return config.get("google_drive_excel_folder_id") or _get_folder_id()


def _resolve_target_root_id(service, configured_id, log_base_path=None):
    target_id = str(configured_id or "").strip()
    if not target_id:
        raise RuntimeError("No se pudo resolver la carpeta raíz de Google Drive.")

    # Shared drive roots are commonly referenced directly by their drive ID.
    if target_id.startswith("0A"):
        _log_drive(f"ROOT_SHARED_DRIVE drive_id={target_id}", log_base_path)
        return target_id

    try:
        metadata = service.files().get(
            fileId=target_id,
            fields="id,name,driveId,parents,mimeType,trashed",
            supportsAllDrives=True,
        ).execute()
    except Exception as exc:
        _log_drive(f"WARN root_metadata_unavailable id={target_id} error={exc}", log_base_path)
        return target_id

    if metadata.get("trashed") and metadata.get("driveId"):
        drive_root_id = str(metadata.get("driveId") or "").strip()
        if drive_root_id:
            _log_drive(
                f"WARN root_folder_trashed id={target_id} name={metadata.get('name')!r} "
                f"fallback_drive_root={drive_root_id}",
                log_base_path,
            )
            return drive_root_id

    _log_drive(
        f"ROOT_FOLDER id={metadata.get('id')} name={metadata.get('name')!r} "
        f"drive_id={metadata.get('driveId')} trashed={metadata.get('trashed')}",
        log_base_path,
    )
    return target_id


def _get_or_create_folder(service, parent_folder_id, folder_name):
    safe_name = _sanitize_filename(folder_name)
    safe_query_name = safe_name.replace("'", "\\'")
    query = (
        "mimeType='application/vnd.google-apps.folder' "
        f"and name='{safe_query_name}' "
        f"and '{parent_folder_id}' in parents and trashed=false"
    )
    result = service.files().list(
        q=query,
        fields="files(id,name)",
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
        pageSize=1,
    ).execute()
    files = result.get("files", [])
    if files:
        return files[0]["id"]

    metadata = {
        "name": safe_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }
    created = service.files().create(
        body=metadata,
        fields="id,name",
        supportsAllDrives=True,
    ).execute()
    return created["id"]


def _load_config():
    # Look for config.json next to the bundle / script first, then fall back to
    # the current working directory so the dev workflow still works.
    bundle_path = os.path.join(_get_bundle_dir(), DEFAULT_CONFIG_PATH)
    cwd_path = DEFAULT_CONFIG_PATH

    for candidate in (bundle_path, cwd_path):
        if os.path.exists(candidate):
            try:
                with open(candidate, "r", encoding="utf-8") as handle:
                    return json.load(handle) or {}
            except (OSError, json.JSONDecodeError):
                return {}
    return {}


def _get_log_dir(base_path=None):
    if base_path and os.path.exists(base_path):
        base_dir = os.path.dirname(base_path)
    else:
        base_dir = os.getcwd()
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    return log_dir


def _get_desktop_log_path():
    candidates = []
    one_drive = str(os.getenv("OneDrive") or "").strip()
    if one_drive:
        candidates.append(os.path.join(one_drive, "Desktop", "log"))
    one_drive_consumer = str(os.getenv("OneDriveConsumer") or "").strip()
    if one_drive_consumer:
        candidates.append(os.path.join(one_drive_consumer, "Desktop", "log"))
    userprofile = str(os.getenv("USERPROFILE") or "").strip()
    if userprofile:
        candidates.append(os.path.join(userprofile, "Desktop", "log"))
    local_app_data = str(os.getenv("LOCALAPPDATA") or "").strip()
    if local_app_data:
        candidates.append(os.path.join(local_app_data, "RECA", "logs", "log"))
    candidates.append(os.path.join(os.getcwd(), "log"))

    for path in candidates:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            return path
        except OSError:
            continue
    return os.path.join(os.getcwd(), "log")


def _log_drive_desktop(message):
    try:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        log_path = _get_desktop_log_path()
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] [DRIVE] {message}\n")
    except OSError:
        return


def _log_drive(message, base_path=None):
    _log_drive_desktop(message)
    try:
        log_dir = _get_log_dir(base_path)
        log_path = os.path.join(log_dir, "drive_log.txt")
        if os.path.exists(log_path):
            try:
                if os.path.getsize(log_path) >= 5 * 1024 * 1024:
                    with open(log_path, "w", encoding="utf-8") as log_file:
                        log_file.write("")
            except OSError:
                with open(log_path, "w", encoding="utf-8") as log_file:
                    log_file.write("")
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] {message}\n")
    except OSError:
        return


def upload_excel_to_drive(
    excel_path,
    base_name=None,
    folder_name=None,
    professional_name=None,
):
    if not excel_path:
        raise RuntimeError("Falta excel_path para subir a Drive.")
    if not os.path.exists(excel_path):
        raise RuntimeError(f"No existe el archivo de Excel: {excel_path}")

    # Log early so every failure path is captured in the drive log.
    _log_drive(
        f"START_EXCEL path={excel_path} base_name={base_name!r} folder_name={folder_name!r} "
        f"bundle_dir={_get_bundle_dir()}",
        excel_path,
    )

    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
    except ImportError as exc:
        _log_drive("ERROR missing_dependencies", excel_path)
        raise RuntimeError(
            "Faltan dependencias para Google Drive. Instala google-api-python-client y google-auth."
        ) from exc

    try:
        creds_path = _get_credentials_path()
    except RuntimeError as exc:
        _log_drive(f"ERROR credentials_path {exc}", excel_path)
        raise

    try:
        configured_root_folder_id = _get_excel_folder_id()
    except RuntimeError as exc:
        _log_drive(f"ERROR folder_id {exc}", excel_path)
        raise

    filename = _sanitize_filename(base_name or os.path.basename(excel_path))
    resolved_folder_name = folder_name if folder_name is not None else professional_name

    _log_drive(
        f"RESOLVED creds={creds_path} folder_id={configured_root_folder_id} "
        f"target_folder={resolved_folder_name!r}",
        excel_path,
    )

    credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
    service = build("drive", "v3", credentials=credentials)
    root_folder_id = _resolve_target_root_id(service, configured_root_folder_id, excel_path)
    target_folder_id = root_folder_id
    if resolved_folder_name:
        try:
            target_folder_id = _get_or_create_folder(
                service,
                root_folder_id,
                resolved_folder_name,
            )
        except Exception as exc:
            _log_drive(
                f"WARN folder_fallback folder_name={resolved_folder_name!r} error={exc}",
                excel_path,
            )
            target_folder_id = root_folder_id

    metadata = {"name": filename, "parents": [target_folder_id]}
    media = MediaFileUpload(
        excel_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False,
    )
    try:
        result = service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink",
            supportsAllDrives=True,
        ).execute()
    except Exception as exc:
        _log_drive(f"ERROR upload_excel {exc}", excel_path)
        raise
    file_id = result.get("id")
    file_name = result.get("name")
    web_link = result.get("webViewLink")
    _log_drive(
        f"SUCCESS_EXCEL id={file_id} name={file_name} folder={target_folder_id} folder_name={resolved_folder_name!r} link={web_link}",
        excel_path,
    )
    return {
        "file_id": file_id,
        "file_name": file_name,
        "webViewLink": web_link,
    }
