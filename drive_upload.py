import json
import os
import re
import time


DEFAULT_FOLDER_ID = "1zbuUkpHPEDfNuLN-tPQax-3ua26oxjjA"
SCOPE = "https://www.googleapis.com/auth/drive.file"
DEFAULT_CONFIG_PATH = "config.json"


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
    if not os.path.exists(path):
        raise RuntimeError(f"No existe el JSON de credenciales: {path}")
    return path


def _get_folder_id():
    if os.getenv("GOOGLE_DRIVE_FOLDER_ID"):
        return os.getenv("GOOGLE_DRIVE_FOLDER_ID")
    config = _load_config()
    return config.get("google_drive_folder_id") or DEFAULT_FOLDER_ID


def _get_excel_folder_id():
    if os.getenv("GOOGLE_DRIVE_EXCEL_FOLDER_ID"):
        return os.getenv("GOOGLE_DRIVE_EXCEL_FOLDER_ID")
    config = _load_config()
    return config.get("google_drive_excel_folder_id") or _get_folder_id()


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
    if not os.path.exists(DEFAULT_CONFIG_PATH):
        return {}
    try:
        with open(DEFAULT_CONFIG_PATH, "r", encoding="utf-8") as handle:
            return json.load(handle) or {}
    except (OSError, json.JSONDecodeError):
        return {}


def _get_log_dir(base_path=None):
    if base_path and os.path.exists(base_path):
        base_dir = os.path.dirname(base_path)
    else:
        base_dir = os.getcwd()
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    return log_dir


def _log_drive(message, base_path=None):
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


def upload_excel_to_drive(excel_path, base_name=None, professional_name=None):
    if not excel_path:
        raise RuntimeError("Falta excel_path para subir a Drive.")
    if not os.path.exists(excel_path):
        raise RuntimeError(f"No existe el archivo de Excel: {excel_path}")
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
    except ImportError as exc:
        _log_drive("ERROR missing_dependencies", excel_path)
        raise RuntimeError(
            "Faltan dependencias para Google Drive. Instala google-api-python-client y google-auth."
        ) from exc

    creds_path = _get_credentials_path()
    root_folder_id = _get_excel_folder_id()
    filename = _sanitize_filename(base_name or os.path.basename(excel_path))

    _log_drive(
        f"START_EXCEL path={excel_path} base_name={base_name!r} professional={professional_name!r}",
        excel_path,
    )

    credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
    service = build("drive", "v3", credentials=credentials)
    target_folder_id = root_folder_id
    if professional_name:
        try:
            target_folder_id = _get_or_create_folder(
                service,
                root_folder_id,
                professional_name,
            )
        except Exception as exc:
            _log_drive(
                f"WARN folder_profesional_fallback professional={professional_name!r} error={exc}",
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
        f"SUCCESS_EXCEL id={file_id} name={file_name} folder={target_folder_id} professional={professional_name!r} link={web_link}",
        excel_path,
    )
    return file_id, file_name
