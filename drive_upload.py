import io
import json
import os
import re
import sys
import time
import uuid
from datetime import date, datetime

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover
    ZoneInfo = None

from logging_utils import log_drive_event
from formularios.common import (
    DEFAULT_SERVICE_ACCOUNT_FILE_NAME,
    FORBIDDEN_CONFIG_KEYS,
    _get_roaming_app_dir,
    _load_env_file,
    _load_json_config,
    _resolve_existing_path,
    _sanitize_filename as _shared_sanitize_filename,
)
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
    runtime_env, env_source = _load_env_file(".env", include_source=True)
    path = (
        os.getenv("GOOGLE_DRIVE_SA_JSON")
        or os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
        or runtime_env.get("GOOGLE_DRIVE_SA_JSON")
        or runtime_env.get("GOOGLE_SERVICE_ACCOUNT_FILE")
    )
    search_dirs = []
    if env_source:
        search_dirs.append(os.path.dirname(os.path.abspath(env_source)))
    roaming_dir = _get_roaming_app_dir(create=False)
    if roaming_dir:
        search_dirs.append(roaming_dir)
    search_dirs.extend([_get_bundle_dir(), os.getcwd()])
    if path:
        resolved_path = _resolve_existing_path(path, base_dirs=search_dirs)
        if not resolved_path:
            raise RuntimeError("El archivo de credenciales de Google Drive configurado no es válido.")
        return resolved_path
    resolved_path = _resolve_existing_path(DEFAULT_SERVICE_ACCOUNT_FILE_NAME, base_dirs=search_dirs)
    if not resolved_path:
        raise RuntimeError(
            "Faltan credenciales de Google Drive. Configure GOOGLE_SERVICE_ACCOUNT_FILE "
            "en el .env de la aplicación o coloque service-account.json en %APPDATA%\\RECA Inclusion Laboral."
        )
    return resolved_path


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


def _rewrite_excluded_auto_resize_rows(excluded_rows_by_sheet, replacements):
    rewritten = {}
    for sheet_name, rows in (excluded_rows_by_sheet or {}).items():
        source_name = str(sheet_name or "").strip()
        if not source_name:
            continue
        target_name = replacements.get(source_name, source_name)
        target_rows = rewritten.setdefault(target_name, set())
        for row in rows or []:
            try:
                row_number = int(row or 0)
            except Exception:
                continue
            if row_number > 0:
                target_rows.add(row_number)
    return rewritten


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
    return _load_json_config(
        DEFAULT_CONFIG_PATH,
        forbidden_keys=FORBIDDEN_CONFIG_KEYS,
    )


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
    auto_resize_excluded_rows=None,
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
                if populated_ranges <= 0:
                    _log_drive(
                        f"REUSE_EMPTY_TARGET_SHEET sheet={sheet_name!r} id={spreadsheet_id} "
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
                auto_resize_excluded_rows = _rewrite_excluded_auto_resize_rows(
                    auto_resize_excluded_rows,
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
        batch_write_sheet_updates(
            spreadsheet_id,
            sheet_writes,
            auto_resize_excluded_rows=auto_resize_excluded_rows,
        )

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


# ---------------------------------------------------------------------------
# PDF export con metadata RECA
# ---------------------------------------------------------------------------

_MESES_ES = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}

_PDF_EXPORT_FOLDER_KEY = "pdf_export_folder_id"
_PDF_EXPORT_FOLDER_ENV = "PDF_EXPORT_FOLDER_ID"


def _get_pdf_folder_id():
    """Retorna el ID de la carpeta Drive donde se guardan los PDFs de actas."""
    env_val = str(os.getenv(_PDF_EXPORT_FOLDER_ENV) or "").strip()
    if env_val:
        return env_val
    config = _load_config()
    cfg_val = str(config.get(_PDF_EXPORT_FOLDER_KEY) or "").strip()
    if cfg_val:
        return cfg_val
    raise RuntimeError(
        f"Falta {_PDF_EXPORT_FOLDER_ENV} o config.json con {_PDF_EXPORT_FOLDER_KEY}."
    )


def get_acta_pdf_name(tipo_acta: str, fecha_servicio: date, extra: str | None = None) -> str:
    """Retorna el nombre del archivo PDF según los CRITERIOS DE ROTULACIÓN de RECA.

    Args:
        tipo_acta: Identificador del tipo de acta (ej. "presentacion_programa").
        fecha_servicio: Fecha del servicio como objeto ``date``.
        extra: Campo variable según el tipo (nombre de vacante, oferente, empresa, etc.).

    Returns:
        Nombre del archivo sin extensión.
    """
    d = fecha_servicio.day
    m = _MESES_ES.get(fecha_servicio.month, str(fecha_servicio.month))
    y = fecha_servicio.year
    fecha_str = f"{d:02d}_{m}_{y}"

    extra_clean = str(extra or "").strip()

    nombres = {
        "presentacion_programa": f"PRESENTACIÓN DEL PROGRAMA DE INCLUSIÓN LABORAL- {fecha_str}",
        "reactivacion_programa": f"REACTIVACIÓN DEL PROGRAMA DE INCLUSIÓN LABORAL- {fecha_str}",
        "evaluacion_accesibilidad": f"EVALUACIÓN DE ACCESIBILIDAD- {fecha_str}",
        "condiciones_vacante": (
            f"REVISIÓN DE LAS CONDICIONES DE LA VACANTE- {extra_clean}- {fecha_str}"
            if extra_clean
            else f"REVISIÓN DE LAS CONDICIONES DE LA VACANTE- {fecha_str}"
        ),
        "seleccion_individual": (
            f"PROCESO DE SELECCIÓN INCLUYENTE INDIVIDUAL- {extra_clean}- {fecha_str}"
            if extra_clean
            else f"PROCESO DE SELECCIÓN INCLUYENTE INDIVIDUAL- {fecha_str}"
        ),
        "seleccion_grupal": (
            f"PROCESO DE SELECCIÓN INCLUYENTE GRUPAL \u2013 ({extra_clean}) OFERENTES- {fecha_str}"
            if extra_clean
            else f"PROCESO DE SELECCIÓN INCLUYENTE GRUPAL- {fecha_str}"
        ),
        "contratacion_individual": (
            f"PROCESO DE CONTRATACIÓN INCLUYENTE INDIVIDUAL- {extra_clean}- {fecha_str}"
            if extra_clean
            else f"PROCESO DE CONTRATACIÓN INCLUYENTE INDIVIDUAL- {fecha_str}"
        ),
        "contratacion_grupal": (
            f"PROCESO CONTRATACION INCLUYENTE GRUPAL \u2013 ({extra_clean}) VINCULADOS- {fecha_str}"
            if extra_clean
            else f"PROCESO CONTRATACION INCLUYENTE GRUPAL- {fecha_str}"
        ),
        "induccion_operativa": (
            f"INDUCCIÓN OPERATIVA- {extra_clean}- {fecha_str}"
            if extra_clean
            else f"INDUCCIÓN OPERATIVA- {fecha_str}"
        ),
        "induccion_organizacional": (
            f"INDUCCIÓN ORGANIZACIONAL- {extra_clean}- {fecha_str}"
            if extra_clean
            else f"INDUCCIÓN ORGANIZACIONAL- {fecha_str}"
        ),
        "capacitacion_sensibilizacion": f"CAPACITACIÓN Y SENSIBILIZACIÓN -{fecha_str}",
        "seguimiento": (
            f"SEGUIMIENTO AL PROCESO DE INCLUSIÓN LABORAL- {extra_clean}- {fecha_str}"
            if extra_clean
            else f"SEGUIMIENTO AL PROCESO DE INCLUSIÓN LABORAL- {fecha_str}"
        ),
        "control_asistencia": (
            f"CONTROL ASISTENCIA INCLUSIÓN LABORAL- {extra_clean}- {fecha_str}"
            if extra_clean
            else f"CONTROL ASISTENCIA INCLUSIÓN LABORAL- {fecha_str}"
        ),
        "interprete_individual": (
            f"SERVICIO INTÉRPRETE LSC \u2013 {extra_clean} -{fecha_str}"
            if extra_clean
            else f"SERVICIO INTÉRPRETE LSC- {fecha_str}"
        ),
        "interprete_grupal": (
            f"SERVICIO INTÉRPRETE LSC - ({extra_clean}) OFERENTES -{fecha_str}"
            if extra_clean
            else f"SERVICIO INTÉRPRETE LSC- {fecha_str}"
        ),
    }

    return nombres.get(tipo_acta, f"ACTA- {fecha_str}")


def _build_drive_service_for_pdf():
    """Construye y retorna un cliente autenticado de Google Drive."""
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    creds_path = _get_credentials_path()
    credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def export_google_sheet_as_pdf(service, file_id: str) -> bytes:
    """Exporta un Google Sheets a PDF usando la Drive API y retorna los bytes del PDF.

    Args:
        service: Cliente autenticado de Google Drive.
        file_id: ID del archivo Google Sheets a exportar.

    Returns:
        Contenido del PDF como ``bytes``.
    """
    from googleapiclient.http import MediaIoBaseDownload

    request = service.files().export_media(
        fileId=file_id,
        mimeType="application/pdf",
    )
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def inject_reca_metadata(pdf_bytes: bytes, metadata_dict: dict) -> bytes:
    """Inyecta el JSON del acta en el campo /RECA_Data de la metadata del PDF.

    El campo es invisible para el lector humano pero permite a RECA ODS
    extraer la información estructurada sin depender del parser de regex.

    Args:
        pdf_bytes: Contenido original del PDF.
        metadata_dict: Diccionario con los datos del acta a embeber.

    Returns:
        Nuevo contenido del PDF con la metadata inyectada.
    """
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)
    writer.add_metadata({"/RECA_Data": json.dumps(metadata_dict, ensure_ascii=False)})
    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


def upload_pdf_to_folder(
    service,
    pdf_bytes: bytes,
    file_name: str,
    folder_id: str,
) -> dict:
    """Sube el PDF (en memoria) a la carpeta de Drive especificada.

    Args:
        service: Cliente autenticado de Google Drive.
        pdf_bytes: Contenido del PDF.
        file_name: Nombre del archivo (sin extensión; se añade .pdf automáticamente).
        folder_id: ID de la carpeta destino en Drive.

    Returns:
        Dict con ``file_id``, ``file_name`` y ``webViewLink``.
    """
    from googleapiclient.http import MediaIoBaseUpload

    safe_name = _sanitize_filename(file_name) + ".pdf"
    media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf", resumable=False)
    file_metadata = {
        "name": safe_name,
        "mimeType": "application/pdf",
        "parents": [folder_id],
    }
    created = execute_google_create_with_confirmation(
        lambda: service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id,name,webViewLink",
            supportsAllDrives=True,
        ),
        lambda: _find_existing_pdf(service, folder_id, safe_name),
        operation_name="drive.upload_pdf",
        logger=_google_request_logger(),
    )
    _log_drive(
        f"PDF_UPLOADED file_id={created.get('id')} name={created.get('name')!r} "
        f"folder_id={folder_id}"
    )
    return {
        "file_id": str(created.get("id") or ""),
        "file_name": str(created.get("name") or safe_name),
        "webViewLink": str(created.get("webViewLink") or ""),
    }


def _find_existing_pdf(service, folder_id: str, file_name: str) -> dict | None:
    """Busca un PDF con el mismo nombre en la carpeta para evitar duplicados."""
    safe_query_name = file_name.replace("'", "\\'")
    query = (
        f"mimeType='application/pdf' "
        f"and name='{safe_query_name}' "
        f"and '{folder_id}' in parents "
        f"and trashed=false"
    )
    result = execute_google_request_with_retry(
        service.files().list(
            q=query,
            fields="files(id,name,webViewLink)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            pageSize=1,
        ),
        operation_name="drive.find_pdf",
        logger=_google_request_logger(),
    )
    files = result.get("files", [])
    return files[0] if files else None


def create_and_upload_acta_pdf(
    *,
    service,
    sheet_file_id: str,
    tipo_acta: str,
    acta_metadata: dict,
    fecha_servicio: date,
    folder_id: str,
    folder_name: str | None = None,
    extra: str | None = None,
) -> dict:
    """Orquesta la exportación, inyección de metadata y subida del PDF del acta.

    1. Exporta el Google Sheet como PDF via Drive API.
    2. Inyecta el JSON del acta en /RECA_Data (invisible para el cliente).
    3. Si ``folder_name`` está presente, crea o reutiliza una subcarpeta con ese
       nombre dentro de ``folder_id`` (ej. nombre de empresa) y sube el PDF ahí.
       Si no, sube directamente a ``folder_id``.

    Args:
        service: Cliente autenticado de Google Drive.
        sheet_file_id: ID del Google Sheet a exportar.
        tipo_acta: Tipo de acta (ej. "presentacion_programa").
        acta_metadata: Diccionario de datos del acta para embeber en la metadata.
        fecha_servicio: Fecha del servicio como ``date``.
        folder_id: ID de la carpeta raíz de PDFs en Drive.
        folder_name: Nombre de la subcarpeta por empresa (opcional).
        extra: Campo extra para el nombre del archivo (vacante, oferente, etc.).

    Returns:
        Dict con ``file_id``, ``file_name`` y ``webViewLink`` del PDF subido.
    """
    pdf_name = get_acta_pdf_name(tipo_acta, fecha_servicio, extra=extra)
    _log_drive(f"PDF_EXPORT_START sheet_id={sheet_file_id} tipo={tipo_acta!r} name={pdf_name!r}")

    # Resolver carpeta destino: subcarpeta por empresa si se indica
    target_folder_id = folder_id
    clean_folder_name = str(folder_name or "").strip()
    if clean_folder_name:
        try:
            target_folder_id = _get_or_create_folder(service, folder_id, clean_folder_name)
            _log_drive(f"PDF_EXPORT_SUBFOLDER folder={clean_folder_name!r} id={target_folder_id}")
        except Exception as exc:
            _log_drive(f"WARN PDF_EXPORT_SUBFOLDER_FAILED folder={clean_folder_name!r} err={exc} — usando raíz")
            target_folder_id = folder_id

    pdf_bytes = export_google_sheet_as_pdf(service, sheet_file_id)
    pdf_bytes = inject_reca_metadata(pdf_bytes, acta_metadata)
    result = upload_pdf_to_folder(service, pdf_bytes, pdf_name, target_folder_id)

    _log_drive(
        f"PDF_EXPORT_OK file_id={result['file_id']} name={result['file_name']!r} "
        f"folder={clean_folder_name or 'raiz'!r}"
    )
    return result
