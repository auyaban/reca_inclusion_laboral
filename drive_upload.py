import json
import os
import re
import sys
import time
import uuid
from datetime import datetime

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover
    ZoneInfo = None

from logging_utils import log_drive_event
from formularios.common import _load_env_file, _sanitize_filename as _shared_sanitize_filename
from google_api_requests import (
    execute_google_create_with_confirmation,
    execute_google_request_with_retry,
)


SCOPE = "https://www.googleapis.com/auth/drive"
DEFAULT_CONFIG_PATH = "config.json"
GOOGLE_SHEETS_MIME = "application/vnd.google-apps.spreadsheet"
GOOGLE_FOLDER_MIME = "application/vnd.google-apps.folder"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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
    return _shared_sanitize_filename(value, default="archivo", max_length=200)


def _google_request_logger(log_base_path=None):
    return lambda message: _log_drive(message, log_base_path)


def _load_runtime_env():
    try:
        return _load_env_file(".env") or {}
    except Exception:
        return {}


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
        path = (
            config.get("google_drive_sa_json")
            or config.get("google_service_account_file")
            or config.get("google_sheets_sa_json")
        )
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
        metadata = execute_google_request_with_retry(
            service.files().get(
                fileId=target_id,
                fields="id,name,driveId,parents,mimeType,trashed",
                supportsAllDrives=True,
            ),
            operation_name="drive.get_root_metadata",
            logger=_google_request_logger(log_base_path),
        )
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


def _get_or_create_folder(service, parent_folder_id, folder_name, log_base_path=None):
    safe_name = _sanitize_filename(folder_name)
    safe_query_name = safe_name.replace("'", "\\'")
    query = (
        f"mimeType='{GOOGLE_FOLDER_MIME}' "
        f"and name='{safe_query_name}' "
        f"and '{parent_folder_id}' in parents and trashed=false"
    )
    result = execute_google_request_with_retry(
        service.files().list(
            q=query,
            fields="files(id,name)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            pageSize=1,
        ),
        operation_name="drive.find_folder",
        logger=_google_request_logger(log_base_path),
    )
    files = result.get("files", [])
    if files:
        return files[0]["id"]

    metadata = {
        "name": safe_name,
        "mimeType": GOOGLE_FOLDER_MIME,
        "parents": [parent_folder_id],
    }
    created = execute_google_create_with_confirmation(
        lambda: service.files().create(
            body=metadata,
            fields="id,name",
            supportsAllDrives=True,
        ),
        lambda: _find_named_folder(service, parent_folder_id, safe_name),
        operation_name="drive.create_folder",
        logger=_google_request_logger(log_base_path),
    )
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
    return execute_google_request_with_retry(
        service.files().list(**params),
        operation_name="drive.list_items",
        logger=_google_request_logger(),
    )


def _build_request_app_properties(kind, request_id, extra_properties=None):
    app_properties = {}
    for key, value in (extra_properties or {}).items():
        if value is None:
            continue
        app_properties[str(key)] = str(value)
    if kind:
        app_properties["kind"] = str(kind)
    if request_id:
        app_properties["request_id"] = str(request_id)
    return app_properties


def _find_drive_file_by_request_id(
    service,
    parent_folder_id,
    filename,
    request_id,
    *,
    mime_type=None,
):
    safe_name = _sanitize_filename(filename)
    safe_query_name = safe_name.replace("'", "\\'")
    query_parts = [f"name='{safe_query_name}'", f"'{parent_folder_id}' in parents", "trashed=false"]
    if mime_type:
        query_parts.insert(0, f"mimeType='{mime_type}'")
    result = _list_drive_items(
        service,
        parent_id=parent_folder_id,
        query=" and ".join(query_parts),
        fields="files(id,name,mimeType,webViewLink,appProperties)",
    )
    for item in result.get("files", []):
        app_properties = item.get("appProperties") or {}
        if str(app_properties.get("request_id") or "").strip() == str(request_id or "").strip():
            return item
    return None


def _find_named_folder(service, parent_folder_id, folder_name):
    safe_name = _sanitize_filename(folder_name)
    safe_query_name = safe_name.replace("'", "\\'")
    query = (
        f"mimeType='{GOOGLE_FOLDER_MIME}' "
        f"and name='{safe_query_name}' "
        f"and '{parent_folder_id}' in parents and trashed=false"
    )
    result = _list_drive_items(
        service,
        parent_id=parent_folder_id,
        query=query,
        fields="files(id,name,webViewLink)",
    )
    files = result.get("files", [])
    return files[0] if files else None


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


def _find_existing_spreadsheet(service, parent_folder_id, filename):
    """Return the file ID of an existing Google Sheets file in the folder, or None."""
    safe_name = _sanitize_filename(filename)
    safe_query_name = safe_name.replace("'", "\\'")
    query = (
        f"mimeType='{GOOGLE_SHEETS_MIME}' "
        f"and name='{safe_query_name}' "
        f"and '{parent_folder_id}' in parents and trashed=false"
    )
    try:
        result = _list_drive_items(
            service,
            parent_id=parent_folder_id,
            query=query,
            fields="files(id,name)",
        )
    except AttributeError:
        return None
    files = result.get("files", [])
    if files:
        return files[0]["id"]
    return None


def _extract_sheet_name_from_a1(range_name):
    match = re.match(r"^'([^']+)'!", str(range_name or "").strip())
    if match:
        return match.group(1)
    return ""


def _replace_sheet_name_in_a1(range_name, replacements):
    text = str(range_name or "").strip()
    match = re.match(r"^'([^']+)'!(.+)$", text)
    if not match:
        return text
    current_sheet = match.group(1)
    next_sheet = str((replacements or {}).get(current_sheet) or current_sheet).strip()
    return f"'{next_sheet}'!{match.group(2)}"


def _has_nonempty_sheet_values(value_ranges):
    for rows in (value_ranges or {}).values():
        if _range_has_sheet_values(rows):
            return True
    return False


def _cell_has_sheet_value(value):
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    return True


def _range_has_sheet_values(rows):
    for row in rows or []:
        for value in row or []:
            if _cell_has_sheet_value(value):
                return True
    return False


def _count_populated_target_ranges(value_ranges, expected_ranges):
    populated = 0
    for range_name in expected_ranges or []:
        if _range_has_sheet_values((value_ranges or {}).get(range_name)):
            populated += 1
    return populated


def _all_target_ranges_populated(value_ranges, expected_ranges):
    expected = [str(range_name or "").strip() for range_name in (expected_ranges or []) if str(range_name or "").strip()]
    if not expected:
        return False
    return _count_populated_target_ranges(value_ranges, expected) == len(expected)


def _get_bogota_today():
    if ZoneInfo is not None:
        try:
            return datetime.now(ZoneInfo("America/Bogota")).date()
        except Exception:
            pass
    return datetime.now().date()


def _build_dated_sheet_title(base_title, existing_titles, *, current_date=None):
    safe_base = str(base_title or "").strip() or "Hoja"
    date_text = str(current_date or _get_bogota_today().isoformat()).strip()
    suffix = f" - {date_text}" if date_text else ""
    max_title_len = 100
    trimmed_base = safe_base[: max_title_len - len(suffix)].rstrip() or "Hoja"
    candidate = f"{trimmed_base}{suffix}"
    if candidate not in existing_titles:
        return candidate

    counter = 2
    while True:
        numbered_suffix = f"{suffix} ({counter})"
        trimmed_base = safe_base[: max_title_len - len(numbered_suffix)].rstrip() or "Hoja"
        candidate = f"{trimmed_base}{numbered_suffix}"
        if candidate not in existing_titles:
            return candidate
        counter += 1


def _rewrite_sheet_payloads(
    sheet_writes,
    clear_ranges,
    checkbox_cells,
    unmerge_areas,
    row_insertions,
    replacements,
):
    rewritten_writes = []
    for item in sheet_writes or []:
        if not isinstance(item, dict):
            continue
        next_item = dict(item)
        next_item["range"] = _replace_sheet_name_in_a1(next_item.get("range"), replacements)
        rewritten_writes.append(next_item)

    rewritten_clear_ranges = [
        _replace_sheet_name_in_a1(range_name, replacements)
        for range_name in (clear_ranges or [])
    ]

    rewritten_checkboxes = []
    for item in checkbox_cells or []:
        if not isinstance(item, dict):
            continue
        next_item = dict(item)
        next_item["range"] = _replace_sheet_name_in_a1(next_item.get("range"), replacements)
        rewritten_checkboxes.append(next_item)

    rewritten_unmerge_areas = []
    for item in unmerge_areas or []:
        if not isinstance(item, dict):
            continue
        next_item = dict(item)
        sheet_name = str(next_item.get("sheet_name") or "").strip()
        if sheet_name in replacements:
            next_item["sheet_name"] = replacements[sheet_name]
        rewritten_unmerge_areas.append(next_item)

    rewritten_row_insertions = []
    for item in row_insertions or []:
        if not isinstance(item, dict):
            continue
        next_item = dict(item)
        sheet_name = str(next_item.get("sheet_name") or "").strip()
        if sheet_name in replacements:
            next_item["sheet_name"] = replacements[sheet_name]
        rewritten_row_insertions.append(next_item)

    return (
        rewritten_writes,
        rewritten_clear_ranges,
        rewritten_checkboxes,
        rewritten_unmerge_areas,
        rewritten_row_insertions,
    )


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

    metadata = execute_google_request_with_retry(
        service.files().get(
            fileId=target_id,
            fields="id,name,driveId,mimeType,trashed",
            supportsAllDrives=True,
        ),
        operation_name="drive.probe_read_parent",
        logger=_google_request_logger(log_base_path),
    )
    _log_drive(
        f"PROBE_READ_OK folder_id={metadata.get('id')} name={metadata.get('name')!r} drive_id={metadata.get('driveId')}",
        log_base_path,
    )
    return {"sample_id": metadata.get("id"), "sample_name": metadata.get("name")}


def _probe_parent_write_access(service, parent_folder_id, log_base_path=None):
    probe_name = f".reca_drive_probe_{uuid.uuid4().hex[:8]}"
    metadata = {
        "name": probe_name,
        "mimeType": GOOGLE_FOLDER_MIME,
        "parents": [parent_folder_id],
    }
    created = execute_google_create_with_confirmation(
        lambda: service.files().create(
            body=metadata,
            fields="id,name",
            supportsAllDrives=True,
        ),
        lambda: _find_named_folder(service, parent_folder_id, probe_name),
        operation_name="drive.probe_create_parent",
        logger=_google_request_logger(log_base_path),
    )
    probe_id = created.get("id")
    _log_drive(
        f"PROBE_WRITE_CREATE_OK parent={parent_folder_id} probe_id={probe_id} probe_name={created.get('name')!r}",
        log_base_path,
    )
    try:
        execute_google_request_with_retry(
            service.files().update(
                fileId=probe_id,
                body={"trashed": True},
                fields="id,trashed",
                supportsAllDrives=True,
            ),
            operation_name="drive.probe_trash_parent",
            logger=_google_request_logger(log_base_path),
        )
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
            return _result(True, "Configurado y con escritura", "", detail)
        return _result(True, "Configurado y autenticado", "", detail)
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
                log_base_path=excel_path,
            )
        except Exception as exc:
            _log_drive(
                f"WARN folder_fallback folder_name={resolved_folder_name!r} error={exc}",
                excel_path,
            )
            target_folder_id = root_folder_id

    filename = _get_available_filename(service, target_folder_id, requested_filename, excel_path)
    request_id = uuid.uuid4().hex
    metadata = {
        "name": filename,
        "parents": [target_folder_id],
        "appProperties": _build_request_app_properties("excel_upload", request_id),
    }
    try:
        result = execute_google_create_with_confirmation(
            lambda: service.files().create(
                body=metadata,
                media_body=MediaFileUpload(excel_path, mimetype=XLSX_MIME, resumable=False),
                fields="id,name,webViewLink,mimeType,appProperties",
                supportsAllDrives=True,
            ),
            lambda: _find_drive_file_by_request_id(
                service,
                target_folder_id,
                filename,
                request_id,
                mime_type=XLSX_MIME,
            ),
            operation_name="drive.upload_excel",
            logger=_google_request_logger(excel_path),
        )
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


def publish_sheet_from_template(
    *,
    template_id,
    sheet_writes,
    base_name=None,
    folder_name=None,
    clear_ranges=None,
    unmerge_areas=None,
    row_insertions=None,
    extra_visible_sheets=None,
    checkbox_cells=None,
):
    """Copy a Google Sheets template and write data into the copy.

    Parameters
    ----------
    template_id : str
        Spreadsheet ID (or URL) of the template to copy.
    sheet_writes : list[dict]
        List of ``{"range": "'Tab'!Cell", "value": ...}`` dicts.
    base_name : str | None
        Desired filename for the new spreadsheet.
    folder_name : str | None
        Subfolder inside the configured root Drive folder.
    clear_ranges : list[str] | None
        Ranges to clear before writing (optional).

    Returns
    -------
    dict  with keys ``file_id``, ``file_name``, ``webViewLink``.
    """
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
            batch_read_sheet_values,
            batch_write_sheet_updates,
            clear_protected_ranges,
            clear_sheet_ranges,
            copy_sheet_to_spreadsheet,
            extract_spreadsheet_id,
            get_sheet_titles,
            hide_sheets,
            insert_template_block_rows,
            insert_template_rows,
            set_native_checkboxes,
            unmerge_cells_in_area,
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

    resolved_template_id = extract_spreadsheet_id(template_id)

    requested_filename = _sanitize_filename(base_name or "Acta")
    requested_filename = _split_filename(requested_filename)[0]

    _log_drive(
        f"START_SHEET template_id={resolved_template_id} base_name={requested_filename!r} "
        f"folder_name={folder_name!r} bundle_dir={_get_bundle_dir()}"
    )

    credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
    service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    root_folder_id = _resolve_target_root_id(service, configured_root_folder_id)
    target_folder_id = root_folder_id
    if folder_name:
        try:
            target_folder_id = _get_or_create_folder(
                service,
                root_folder_id,
                folder_name,
                log_base_path=requested_filename,
            )
        except Exception as exc:
            _log_drive(
                f"WARN sheet_folder_fallback folder_name={folder_name!r} error={exc}"
            )
            target_folder_id = root_folder_id

    # --- Reuse existing spreadsheet or copy template -------------------
    existing_id = _find_existing_spreadsheet(service, target_folder_id, requested_filename)
    is_reuse = bool(existing_id)
    spreadsheet_id = existing_id or ""
    file_name = requested_filename
    preferred_sheet_gid = ""

    try:
        if is_reuse:
            _log_drive(f"REUSE_SHEET id={spreadsheet_id} name={requested_filename!r}")
        else:
            filename = _get_available_filename(service, target_folder_id, requested_filename)
            request_id = uuid.uuid4().hex
            metadata = {
                "name": filename,
                "parents": [target_folder_id],
                "appProperties": _build_request_app_properties("google_sheet_publish", request_id),
            }
            copied = execute_google_create_with_confirmation(
                lambda: service.files().copy(
                    fileId=resolved_template_id,
                    body=metadata,
                    fields="id,name,webViewLink,mimeType,appProperties",
                    supportsAllDrives=True,
                ),
                lambda: _find_drive_file_by_request_id(
                    service,
                    target_folder_id,
                    filename,
                    request_id,
                    mime_type=GOOGLE_SHEETS_MIME,
                ),
                operation_name="drive.copy_template_spreadsheet",
                logger=_google_request_logger(requested_filename),
            )
            spreadsheet_id = str(copied.get("id") or "").strip()
            file_name = copied.get("name", filename)
            if not spreadsheet_id:
                raise RuntimeError("Google Drive no devolvió el ID de la copia creada.")

        target_sheet_ranges = {}
        for item in sheet_writes or []:
            if not isinstance(item, dict):
                continue
            range_name = str(item.get("range") or "").strip()
            sheet_name = _extract_sheet_name_from_a1(range_name)
            if not sheet_name or not range_name:
                continue
            target_sheet_ranges.setdefault(sheet_name, set()).add(range_name)
        for item in checkbox_cells or []:
            if not isinstance(item, dict):
                continue
            range_name = str(item.get("range") or "").strip()
            sheet_name = _extract_sheet_name_from_a1(range_name)
            if not sheet_name or not range_name:
                continue
            target_sheet_ranges.setdefault(sheet_name, set()).add(range_name)

        copied_template_sheet = False
        if is_reuse and target_sheet_ranges:
            existing_titles = set(get_sheet_titles(spreadsheet_id))
            replacements = {}
            for sheet_name in sorted(target_sheet_ranges):
                ranges = sorted(target_sheet_ranges.get(sheet_name) or [])
                if not ranges:
                    continue
                if sheet_name not in existing_titles:
                    copied_sheet = copy_sheet_to_spreadsheet(
                        resolved_template_id,
                        sheet_name,
                        spreadsheet_id,
                        new_sheet_name=sheet_name,
                    )
                    preferred_sheet_gid = str(copied_sheet.get("sheetId") or "").strip()
                    existing_titles.add(sheet_name)
                    copied_template_sheet = True
                    _log_drive(
                        f"COPY_TEMPLATE_SHEET missing_sheet={sheet_name!r} destination_id={spreadsheet_id}"
                    )
                    continue

                current_values = batch_read_sheet_values(spreadsheet_id, ranges)
                populated_ranges = _count_populated_target_ranges(current_values, ranges)
                if populated_ranges < len(ranges):
                    _log_drive(
                        f"REUSE_PARTIAL_SHEET sheet={sheet_name!r} id={spreadsheet_id} "
                        f"populated_ranges={populated_ranges}/{len(ranges)}"
                    )
                    continue

                new_sheet_name = _build_dated_sheet_title(sheet_name, existing_titles)
                copied_sheet = copy_sheet_to_spreadsheet(
                    resolved_template_id,
                    sheet_name,
                    spreadsheet_id,
                    new_sheet_name=new_sheet_name,
                )
                preferred_sheet_gid = str(copied_sheet.get("sheetId") or "").strip()
                existing_titles.add(new_sheet_name)
                replacements[sheet_name] = new_sheet_name
                copied_template_sheet = True
                _log_drive(
                    f"COPY_TEMPLATE_SHEET occupied_sheet={sheet_name!r} "
                    f"new_sheet={new_sheet_name!r} destination_id={spreadsheet_id}"
                )

            if replacements:
                (
                    sheet_writes,
                    clear_ranges,
                    checkbox_cells,
                    unmerge_areas,
                    row_insertions,
                ) = _rewrite_sheet_payloads(
                    sheet_writes,
                    clear_ranges,
                    checkbox_cells,
                    unmerge_areas,
                    row_insertions,
                    replacements,
                )

        try:
            clear_protected_ranges(spreadsheet_id)
        except Exception as rp_exc:
            _log_drive(f"WARN remove_protection_failed id={spreadsheet_id} error={rp_exc}")

        if row_insertions:
            for row_spec in row_insertions:
                if not isinstance(row_spec, dict):
                    continue
                sheet_name = str(row_spec.get("sheet_name") or "").strip()
                start_row = int(row_spec.get("start_row") or 0)
                base_rows = int(row_spec.get("base_rows") or 0)
                total_rows = int(row_spec.get("total_rows") or 0)
                template_start_row = int(row_spec.get("template_start_row") or 0)
                template_end_row = int(row_spec.get("template_end_row") or 0)
                repeat_count = int(row_spec.get("repeat_count") or 0)
                paste_type = str(row_spec.get("paste_type") or "PASTE_NORMAL").strip() or "PASTE_NORMAL"

                if template_start_row > 0 and template_end_row >= template_start_row:
                    insert_at_row = int(row_spec.get("insert_at_row") or start_row or 0)
                    if not sheet_name or insert_at_row <= 0:
                        continue
                    inserted = max(0, repeat_count)
                    if inserted <= 0 and base_rows > 0 and total_rows > base_rows:
                        block_height = (template_end_row - template_start_row) + 1
                        inserted = max(0, (total_rows - base_rows) // max(1, block_height))
                    if inserted <= 0:
                        continue
                    insert_template_block_rows(
                        spreadsheet_id,
                        sheet_name,
                        insert_at_row=insert_at_row,
                        template_start_row=template_start_row,
                        template_end_row=template_end_row,
                        repeat_count=inserted,
                        paste_type=paste_type,
                    )
                    _log_drive(
                        f"INSERT_TEMPLATE_BLOCK_ROWS id={spreadsheet_id} sheet={sheet_name!r} "
                        f"insert_at_row={insert_at_row} template_start_row={template_start_row} "
                        f"template_end_row={template_end_row} repeat_count={inserted}"
                    )
                    continue

                if not sheet_name or start_row <= 0 or base_rows <= 0 or total_rows <= base_rows:
                    continue
                insert_at_row = int(row_spec.get("insert_at_row") or (start_row + base_rows))
                template_row = int(row_spec.get("template_row") or (start_row + base_rows - 1))
                insert_template_rows(
                    spreadsheet_id,
                    sheet_name,
                    insert_at_row=insert_at_row,
                    template_row=template_row,
                    count=total_rows - base_rows,
                    paste_type=paste_type,
                )
                _log_drive(
                    f"INSERT_TEMPLATE_ROWS id={spreadsheet_id} sheet={sheet_name!r} "
                    f"insert_at_row={insert_at_row} template_row={template_row} "
                    f"count={total_rows - base_rows}"
                )

        if clear_ranges:
            clear_sheet_ranges(spreadsheet_id, clear_ranges)
        if unmerge_areas:
            for area in unmerge_areas:
                unmerge_cells_in_area(
                    spreadsheet_id,
                    area["sheet_name"],
                    area["start_row"],
                    area["end_row"],
                    area.get("start_col", 0),
                    area.get("end_col", 21),
                )
        batch_write_sheet_updates(spreadsheet_id, sheet_writes)

        if checkbox_cells:
            try:
                set_native_checkboxes(spreadsheet_id, checkbox_cells)
            except Exception as cb_exc:
                _log_drive(f"WARN set_checkboxes_failed id={spreadsheet_id} error={cb_exc}")

        # --- Hide unused sheets -------------------------------------------
        used_sheet_names = set()
        for w in (sheet_writes or []):
            r = str(w.get("range") or "")
            sheet_name = _extract_sheet_name_from_a1(r)
            if sheet_name:
                used_sheet_names.add(sheet_name)
        for w in (checkbox_cells or []):
            r = str(w.get("range") or "")
            sheet_name = _extract_sheet_name_from_a1(r)
            if sheet_name:
                used_sheet_names.add(sheet_name)
        if extra_visible_sheets:
            used_sheet_names.update(extra_visible_sheets)
        if used_sheet_names:
            try:
                hide_sheets(spreadsheet_id, list(used_sheet_names))
            except Exception as hide_exc:
                _log_drive(
                    f"WARN hide_sheets_failed id={spreadsheet_id} error={hide_exc}"
                )
        # ------------------------------------------------------------------

    except Exception as exc:
        _log_drive(f"ERROR publish_sheet {exc}")
        if spreadsheet_id and not is_reuse:
            try:
                execute_google_request_with_retry(
                    service.files().update(
                        fileId=spreadsheet_id,
                        body={"trashed": True},
                        fields="id,trashed",
                        supportsAllDrives=True,
                    ),
                    operation_name="drive.trash_failed_spreadsheet",
                    logger=_google_request_logger(requested_filename),
                )
                _log_drive(f"CLEANUP_SHEET_TRASH_OK id={spreadsheet_id}")
            except Exception as cleanup_exc:
                _log_drive(
                    f"WARN publish_sheet_cleanup_failed id={spreadsheet_id} error={cleanup_exc}"
                )
        raise

    web_link = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit" if spreadsheet_id else ""
    if web_link and preferred_sheet_gid:
        web_link = f"{web_link}#gid={preferred_sheet_gid}"
    _log_drive(
        f"SUCCESS_SHEET id={spreadsheet_id} name={file_name} folder={target_folder_id} "
        f"folder_name={folder_name!r} reused={is_reuse} link={web_link}"
    )
    return {
        "file_id": spreadsheet_id,
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
    format_ranges=None,
    row_insertions=None,
    extra_visible_sheets=None,
    template_id=None,
):
    """Legacy wrapper – delegates to :func:`publish_sheet_from_template`."""
    from google_sheets_client import (
        get_evaluacion_accesibilidad_template_id,
        set_sheet_ranges_bold,
    )

    resolved_template_id = str(template_id or "").strip()
    if not resolved_template_id:
        resolved_template_id = get_evaluacion_accesibilidad_template_id()

    resolved_folder_name = folder_name if folder_name is not None else professional_name

    result = publish_sheet_from_template(
        template_id=resolved_template_id,
        sheet_writes=sheet_writes,
        base_name=base_name or "Evaluacion de Accesibilidad",
        folder_name=resolved_folder_name,
        clear_ranges=clear_ranges,
        row_insertions=row_insertions,
        extra_visible_sheets=extra_visible_sheets,
    )
    spreadsheet_id = str(result.get("file_id") or "").strip()
    if spreadsheet_id and format_ranges:
        set_sheet_ranges_bold(spreadsheet_id, format_ranges, bold=False)
    return result
