import json
import os
import sys
import time
import uuid
from logging_utils import log_drive_event
from formularios.common import _load_env_file, _sanitize_filename as _shared_sanitize_filename


SCOPE = "https://www.googleapis.com/auth/drive"
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


def _load_runtime_env():
    try:
        return _load_env_file(".env") or {}
    except Exception:
        return {}


def _sanitize_filename(value):
    return _shared_sanitize_filename(value, default="archivo", max_length=200)


def _split_filename(filename):
    name = str(filename or "").strip()
    stem, ext = os.path.splitext(name)
    if not stem:
        stem = name or "archivo"
    return stem, ext


def _get_credentials_path():
    runtime_env = _load_runtime_env()
    path = (
        os.getenv("GOOGLE_DRIVE_SA_JSON")
        or os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
        or runtime_env.get("GOOGLE_DRIVE_SA_JSON")
        or runtime_env.get("GOOGLE_SERVICE_ACCOUNT_FILE")
    )
    if not path:
        config = _load_config()
        path = config.get("google_drive_sa_json")
    if not path:
        raise RuntimeError(
            "Falta GOOGLE_DRIVE_SA_JSON/GOOGLE_SERVICE_ACCOUNT_FILE o config.json con google_drive_sa_json."
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


def _list_drive_items(service, *, parent_id=None, drive_id=None, query=None, fields="files(id,name)"):
    params = {
        "fields": fields,
        "includeItemsFromAllDrives": True,
        "supportsAllDrives": True,
        "pageSize": 100,
    }
    if drive_id:
        params["corpora"] = "drive"
        params["driveId"] = drive_id
    if query:
        params["q"] = query
    elif parent_id:
        params["q"] = f"'{parent_id}' in parents and trashed=false"
    return service.files().list(**params).execute()


def _drive_item_exists(service, parent_folder_id, filename):
    safe_name = _sanitize_filename(filename)
    safe_query_name = safe_name.replace("'", "\\'")
    query = (
        "mimeType!='application/vnd.google-apps.folder' "
        f"and name='{safe_query_name}' "
        f"and '{parent_folder_id}' in parents and trashed=false"
    )
    result = _list_drive_items(
        service,
        parent_id=parent_folder_id,
        query=query,
        fields="files(id,name)",
    )
    return bool(result.get("files", []))


def _get_available_filename(service, parent_folder_id, filename, log_base_path=None):
    base_name = _sanitize_filename(filename)
    stem, ext = _split_filename(base_name)
    if not _drive_item_exists(service, parent_folder_id, base_name):
        return base_name

    suffix = 1
    while True:
        candidate = f"{stem} ({suffix}){ext}"
        if not _drive_item_exists(service, parent_folder_id, candidate):
            _log_drive(
                f"NAME_CONFLICT original={base_name!r} resolved={candidate!r} parent={parent_folder_id}",
                log_base_path,
            )
            return candidate
        suffix += 1


def _probe_parent_read_access(service, target_id, log_base_path=None):
    if str(target_id).startswith("0A"):
        result = _list_drive_items(
            service,
            drive_id=target_id,
            query="mimeType!='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id,name)",
        )
        files = result.get("files", [])
        sample = files[0] if files else {}
        _log_drive(
            f"PROBE_READ_OK shared_drive={target_id} sample_id={sample.get('id')} sample_name={sample.get('name')!r}",
            log_base_path,
        )
        return {"sample_id": sample.get("id"), "sample_name": sample.get("name")}

    metadata = service.files().get(
        fileId=target_id,
        fields="id,name,driveId,mimeType,trashed",
        supportsAllDrives=True,
    ).execute()
    _log_drive(
        f"PROBE_READ_OK folder_id={metadata.get('id')} name={metadata.get('name')!r} drive_id={metadata.get('driveId')}",
        log_base_path,
    )
    return {"sample_id": metadata.get("id"), "sample_name": metadata.get("name")}


def _probe_parent_write_access(service, parent_folder_id, log_base_path=None):
    probe_name = f".reca_drive_probe_{uuid.uuid4().hex[:8]}"
    metadata = {
        "name": probe_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }
    created = service.files().create(
        body=metadata,
        fields="id,name",
        supportsAllDrives=True,
    ).execute()
    probe_id = created.get("id")
    _log_drive(
        f"PROBE_WRITE_CREATE_OK parent={parent_folder_id} probe_id={probe_id} probe_name={created.get('name')!r}",
        log_base_path,
    )
    try:
        service.files().update(
            fileId=probe_id,
            body={"trashed": True},
            fields="id,trashed",
            supportsAllDrives=True,
        ).execute()
        _log_drive(
            f"PROBE_WRITE_TRASH_OK parent={parent_folder_id} probe_id={probe_id}",
            log_base_path,
        )
    except Exception as exc:
        _log_drive(
            f"WARN probe_cleanup_failed parent={parent_folder_id} probe_id={probe_id} error={exc}",
            log_base_path,
        )
    return {"probe_id": probe_id, "probe_name": created.get("name")}


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


def _log_drive(message, base_path=None):
    try:
        _ = base_path
        log_drive_event(message)
    except Exception:
        return


def probe_drive_service(timeout=6, log_enabled=False, require_write=True):
    started_at = time.perf_counter()

    def _result(ok, status_text, error_code="", detail=""):
        latency_ms = int((time.perf_counter() - started_at) * 1000)
        payload = {
            "ok": bool(ok),
            "status_text": str(status_text or "").strip(),
            "error_code": str(error_code or "").strip(),
            "detail": str(detail or "").strip(),
            "latency_ms": latency_ms,
        }
        if log_enabled:
            _log_drive(
                f"PROBE ok={payload['ok']} status={payload['status_text']!r} "
                f"code={payload['error_code']!r} detail={payload['detail']!r} "
                f"latency_ms={latency_ms}"
            )
        return payload

    _ = timeout

    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError as exc:
        return _result(False, "Dependencias faltantes", "missing_dependencies", exc)

    try:
        creds_path = _get_credentials_path()
    except Exception as exc:
        return _result(False, "Credenciales no disponibles", "credentials", exc)

    try:
        configured_root_folder_id = _get_excel_folder_id()
    except Exception as exc:
        return _result(False, "Carpeta de Drive no configurada", "folder_config", exc)

    try:
        credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
        service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        target_id = _resolve_target_root_id(service, configured_root_folder_id)
        read_meta = _probe_parent_read_access(service, target_id)
        detail = f"target_id={target_id} sample_id={read_meta.get('sample_id')}"
        if require_write:
            write_meta = _probe_parent_write_access(service, target_id)
            detail = f"{detail} probe_id={write_meta.get('probe_id')}"
            return _result(
                True,
                "Configurado y con escritura",
                "",
                detail,
            )
        return _result(
            True,
            "Configurado y autenticado",
            "",
            detail,
        )
    except Exception as exc:
        status = int(getattr(getattr(exc, "resp", None), "status", 0) or 0)
        if status in {401, 403}:
            return _result(False, "Permisos insuficientes", "auth", exc)
        if status == 404:
            return _result(False, "Carpeta de Drive no accesible", "folder_not_found", exc)
        return _result(False, "No se pudo conectar a Drive", "connectivity", exc)


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

    requested_filename = _sanitize_filename(base_name or os.path.basename(excel_path))
    resolved_folder_name = folder_name if folder_name is not None else professional_name

    _log_drive(
        f"RESOLVED creds={creds_path} folder_id={configured_root_folder_id} "
        f"target_folder={resolved_folder_name!r}",
        excel_path,
    )

    credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
    service = build("drive", "v3", credentials=credentials, cache_discovery=False)
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

    filename = _get_available_filename(service, target_folder_id, requested_filename, excel_path)
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


def publish_evaluacion_accesibilidad_sheet(
    *,
    sheet_writes,
    base_name=None,
    folder_name=None,
    professional_name=None,
    clear_ranges=None,
    template_id=None,
):
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError as exc:
        _log_drive("ERROR missing_dependencies publish_sheet")
        raise RuntimeError(
            "Faltan dependencias para Google Drive. Instala google-api-python-client y google-auth."
        ) from exc

    try:
        from google_sheets_client import (
            batch_write_sheet_updates,
            clear_sheet_ranges,
            extract_spreadsheet_id,
            get_evaluacion_accesibilidad_template_id,
        )
    except ImportError as exc:
        _log_drive("ERROR missing_google_sheets_client")
        raise RuntimeError("No se pudo cargar el cliente de Google Sheets.") from exc

    try:
        creds_path = _get_credentials_path()
    except RuntimeError as exc:
        _log_drive(f"ERROR credentials_path publish_sheet {exc}")
        raise

    try:
        configured_root_folder_id = _get_excel_folder_id()
    except RuntimeError as exc:
        _log_drive(f"ERROR folder_id publish_sheet {exc}")
        raise

    resolved_template_id = str(template_id or "").strip()
    if resolved_template_id:
        resolved_template_id = extract_spreadsheet_id(resolved_template_id)
    else:
        resolved_template_id = get_evaluacion_accesibilidad_template_id()

    requested_filename = _sanitize_filename(base_name or "Evaluacion de Accesibilidad")
    requested_filename = _split_filename(requested_filename)[0]
    resolved_folder_name = folder_name if folder_name is not None else professional_name

    _log_drive(
        f"START_SHEET template_id={resolved_template_id} base_name={requested_filename!r} "
        f"folder_name={resolved_folder_name!r} bundle_dir={_get_bundle_dir()}"
    )

    credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
    service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    root_folder_id = _resolve_target_root_id(service, configured_root_folder_id)
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
                f"WARN sheet_folder_fallback folder_name={resolved_folder_name!r} error={exc}"
            )
            target_folder_id = root_folder_id

    filename = _get_available_filename(service, target_folder_id, requested_filename)
    metadata = {"name": filename, "parents": [target_folder_id]}
    spreadsheet_id = ""
    try:
        copied = service.files().copy(
            fileId=resolved_template_id,
            body=metadata,
            fields="id,name",
            supportsAllDrives=True,
        ).execute()
        spreadsheet_id = str(copied.get("id") or "").strip()
        if not spreadsheet_id:
            raise RuntimeError("Google Drive no devolvió el ID de la copia creada.")
        if clear_ranges:
            clear_sheet_ranges(spreadsheet_id, clear_ranges)
        batch_write_sheet_updates(spreadsheet_id, sheet_writes)
    except Exception as exc:
        _log_drive(f"ERROR publish_sheet {exc}")
        if spreadsheet_id:
            try:
                service.files().update(
                    fileId=spreadsheet_id,
                    body={"trashed": True},
                    fields="id,trashed",
                    supportsAllDrives=True,
                ).execute()
                _log_drive(f"CLEANUP_SHEET_TRASH_OK id={spreadsheet_id}")
            except Exception as cleanup_exc:
                _log_drive(
                    f"WARN publish_sheet_cleanup_failed id={spreadsheet_id} error={cleanup_exc}"
                )
        raise

    file_id = copied.get("id")
    file_name = copied.get("name")
    web_link = f"https://docs.google.com/spreadsheets/d/{file_id}/edit" if file_id else ""
    _log_drive(
        f"SUCCESS_SHEET id={file_id} name={file_name} folder={target_folder_id} "
        f"folder_name={resolved_folder_name!r} link={web_link}"
    )
    return {
        "file_id": file_id,
        "file_name": file_name,
        "webViewLink": web_link,
    }
