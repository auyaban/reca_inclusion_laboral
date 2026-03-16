import threading
import re
import os
import time
import sys
import subprocess
import ctypes
import unicodedata
import shutil
import uuid
import base64
import hashlib
import hmac
import secrets
import json
import copy
import urllib.error
import urllib.request
import webbrowser
from zoneinfo import ZoneInfo
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry

from formularios.presentacion_programa import presentacion_programa
from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.condiciones_vacante import condiciones_vacante
from formularios.seleccion_incluyente import seleccion_incluyente
from formularios.contratacion_incluyente import contratacion_incluyente
from formularios.induccion_organizacional import induccion_organizacional
from formularios.induccion_operativa import induccion_operativa
from formularios.sensibilizacion import sensibilizacion
from formularios.seguimientos import seguimientos
import drive_upload
from dictation import attach_dictation, cleanup_stale_audio
import text_review
from formularios.common import (
    _supabase_upsert,
    _supabase_enqueue_upsert,
    _supabase_upsert_with_queue,
    _supabase_patch_with_queue,
    _supabase_ping,
    _supabase_get_paged,
    _get_supabase_write_queue_stats,
    _get_supabase_write_queue_snapshot,
    _get_supabase_failed_writes_snapshot,
    _supabase_retry_all_queued_writes,
    _supabase_rpc,
    _supabase_get_access_token,
    _supabase_auth_password_login,
    _supabase_auth_update_password,
    _clear_supabase_session,
    _get_desktop_dir,
    _load_env_file,
    _normalize_decimal_value,
    _next_available_file_path,
    probe_supabase_service,
)
from version_info import get_version
from updater import (
    get_latest_release_assets,
    is_update_available,
    download_installer,
    run_installer,
)
from logging_utils import get_log_file_path, get_logs_root, log_app_event


APP_NAME = "RECA Inclusion Laboral"
COLOR_PURPLE = "#7C3D96"
COLOR_TEAL = "#07B499"
COLOR_LIGHT_BG = "#F7F5FA"
COLOR_GROUP_EMPRESA = "#E6F4EA"
COLOR_GROUP_COMPENSAR = "#FFF3E0"
COLOR_GROUP_RECA = "#F3E5F5"
FONT_TITLE = ("Arial", 18, "bold")
FONT_SUBTITLE = ("Arial", 11)
FONT_SECTION = ("Arial", 12, "bold")
FONT_LABEL = ("Arial", 10, "bold")
FORM_PADX = 24
FORM_PADY = 12
ROW_PADY = 4
ENTRY_W_SHORT = 12
ENTRY_W_NARROW = 14
ENTRY_W_MED = 18
ENTRY_W_LONG = 28
ENTRY_W_WIDE = 42
ENTRY_W_XL = 60
TEXT_WIDE = 120
SCROLLBAR_WIDTH = 18
PASSWORD_HASH_ALGO = "pbkdf2_sha256"
PASSWORD_HASH_ITERATIONS = 260000
DEFAULT_EMPRESA_ESTADOS = [
    "Activa",
    "Inactiva",
    "En pausa",
    "Cerrada",
    "No viable",
]
_MOJIBAKE_PATTERNS = ("Ã", "Â", "â€", "ï¿½", "\ufffd", "Ð", "Ñ")
_ENCODING_CHECK_DONE = False
DRAFTS_FILE_NAME = "form_drafts_il.json"
OFFLINE_AUTH_FILE_NAME = "offline_auth_users.json"
LOGIN_CREDENTIALS_FILE_NAME = "login_credentials.json"
DRIVE_UPLOAD_QUEUE_FILE_NAME = "drive_upload_queue.json"
DRIVE_UPLOAD_FAILED_FILE_NAME = "drive_upload_failed.json"
_USAGE_EXEMPT_LOGINS_CACHE = None
_DRIVE_UPLOAD_LOCK = threading.RLock()
_DRIVE_UPLOAD_QUEUE = []
_DRIVE_UPLOAD_FAILED_QUEUE = []
_DRIVE_UPLOAD_WORKER_STARTED = False
_DRIVE_UPLOAD_QUEUE_LOADED = False
_DRIVE_UPLOAD_FAILED_LOADED = False
_SINGLE_INSTANCE_MUTEX_HANDLE = None
SEGUIMIENTO_NORMATIVA_TEMPLATE_TEXT = """
Se reitera la importancia de contar con la retroalimentación vía correo electrónico a la Agencia con copia a RECA, ante procesos de entrevista en el caso de no pasar candidatos filtros de selección y solicitar nuevos candidatos; y firma de contrato.

Se dialoga de la nueva ley 2466 del 2025, en donde se orienta ante totalidad de colaboradores la vinculación de 2 personas con discapacidad, se informa beneficios tangibles y no tangibles bajo la ley 361 art. 31 deducción en la renta por vinculación de personas con discapacidad y el apoyo que está entregando la secretaria de desarrollo.

El Decreto 0223 de 2026 es explícito al indicar en el numeral 1 de su artículo 2.2.6.3.3.33. que:
“Los aprendices no integran la base de trabajadores de carácter permanente de la empresa, para efectos del cálculo de la cuota de empleo para personas en situación de discapacidad, prevista en el numeral 17 del artículo 57 del Código Sustantivo del Trabajo.”

En consecuencia y a la luz de esta nueva norma, contratar aprendices con discapacidad no sirve para aumentar el número de personas con discapacidad computables dentro de la cuota de empleo exigida, por lo cual la cuota se calculará ahora sobre la base de trabajadores permanentes, y el decreto 0223 excluye a los aprendices de esa base.

Sin embargo, el Decreto genera un incentivo distinto en el numeral 2 del mismo artículo, donde establece que:
“La cuota de aprendices se reducirá en un 50% si las personas contratadas tienen una discapacidad comprobada no inferior al 25%”, en cumplimiento del parágrafo del artículo 31 de la Ley 361 de 1997.

Es decir, que sí es posible contratar aprendices con discapacidad, pero el efecto jurídico directo es sobre la cuota de aprendices (Ley 789 de 2002), no sobre la cuota de empleo para personas con discapacidad (Ley 2466 de 2025 art. 57 num. 17 CST).
""".strip()
FORM_MODULE_MAP = {
    "presentacion_programa": presentacion_programa,
    "evaluacion_accesibilidad": evaluacion_accesibilidad,
    "condiciones_vacante": condiciones_vacante,
    "seleccion_incluyente": seleccion_incluyente,
    "contratacion_incluyente": contratacion_incluyente,
    "induccion_organizacional": induccion_organizacional,
    "induccion_operativa": induccion_operativa,
    "sensibilizacion": sensibilizacion,
}
WINDOW_CLASS_FORM_ID_MAP = {
    "Section1Window": "presentacion_programa",
    "EvaluacionAccesibilidadWindow": "evaluacion_accesibilidad",
    "CondicionesVacanteWindow": "condiciones_vacante",
    "SeleccionIncluyenteWindow": "seleccion_incluyente",
    "ContratacionIncluyenteWindow": "contratacion_incluyente",
    "InduccionOrganizacionalWindow": "induccion_organizacional",
    "InduccionOperativaWindow": "induccion_operativa",
    "SensibilizacionWindow": "sensibilizacion",
    "SeguimientosWindow": "seguimientos",
}


def _acquire_single_instance_mutex():
    if os.name != "nt":
        return True
    global _SINGLE_INSTANCE_MUTEX_HANDLE
    if _SINGLE_INSTANCE_MUTEX_HANDLE:
        return True
    try:
        kernel32 = ctypes.windll.kernel32
        mutex_name = "Local\\RECA_INCLUSION_LABORAL_SINGLE_INSTANCE"
        handle = kernel32.CreateMutexW(None, False, mutex_name)
        if not handle:
            return True
        error_code = kernel32.GetLastError()
        _SINGLE_INSTANCE_MUTEX_HANDLE = handle
        return int(error_code or 0) != 183
    except Exception:
        return True


def _release_single_instance_mutex():
    global _SINGLE_INSTANCE_MUTEX_HANDLE
    handle = _SINGLE_INSTANCE_MUTEX_HANDLE
    _SINGLE_INSTANCE_MUTEX_HANDLE = None
    if not handle or os.name != "nt":
        return
    try:
        ctypes.windll.kernel32.ReleaseMutex(handle)
    except Exception:
        pass
    try:
        ctypes.windll.kernel32.CloseHandle(handle)
    except Exception:
        pass


def _show_single_instance_warning():
    message = (
        "La aplicación ya está abierta.\n\n"
        "Cierra la instancia actual antes de abrir otra."
    )
    if os.name == "nt":
        try:
            ctypes.windll.user32.MessageBoxW(
                None,
                message,
                APP_NAME,
                0x00000030,
            )
            return
        except Exception:
            pass
    print(message)

def _log_capture(message):
    try:
        log_app_event(message)
    except Exception:
        pass


def _set_module_cache_snapshot(module, cache_snapshot):
    if not module or not isinstance(cache_snapshot, dict):
        return
    form_cache = getattr(module, "FORM_CACHE", None)
    if isinstance(form_cache, dict):
        form_cache.clear()
        form_cache.update(copy.deepcopy(cache_snapshot))
    section_1_cache = getattr(module, "SECTION_1_CACHE", None)
    if isinstance(section_1_cache, dict):
        section_1_cache.clear()
        section_1 = cache_snapshot.get("section_1")
        if isinstance(section_1, dict):
            section_1_cache.update(copy.deepcopy(section_1))


def _desktop_log_path():
    return get_log_file_path("app")


def _get_local_cache_dir():
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        base = os.path.join(local_app_data, "RECA", "cache")
    else:
        base = os.path.join(os.getcwd(), ".cache")
    os.makedirs(base, exist_ok=True)
    return base


def _get_drafts_path():
    return os.path.join(_get_local_cache_dir(), DRAFTS_FILE_NAME)


def _get_offline_auth_path():
    return os.path.join(_get_local_cache_dir(), OFFLINE_AUTH_FILE_NAME)


def _get_login_credentials_path():
    return os.path.join(_get_local_cache_dir(), LOGIN_CREDENTIALS_FILE_NAME)


def _get_drive_upload_queue_path():
    return os.path.join(_get_local_cache_dir(), DRIVE_UPLOAD_QUEUE_FILE_NAME)


def _get_drive_upload_failed_queue_path():
    return os.path.join(_get_local_cache_dir(), DRIVE_UPLOAD_FAILED_FILE_NAME)


def _atomic_write_json_file(path, payload):
    tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _get_colombia_now():
    try:
        return datetime.now(ZoneInfo("America/Bogota"))
    except Exception:
        return datetime.now()


def _drive_job_identity(job):
    registro_id = str((job or {}).get("registro_id") or "").strip()
    if registro_id:
        return f"registro:{registro_id}"
    local_excel_path = str((job or {}).get("local_excel_path") or "").strip().lower()
    return f"path:{local_excel_path}"


def _persist_drive_upload_queue_locked():
    _atomic_write_json_file(_get_drive_upload_queue_path(), _DRIVE_UPLOAD_QUEUE)


def _persist_drive_failed_queue_locked():
    _atomic_write_json_file(_get_drive_upload_failed_queue_path(), _DRIVE_UPLOAD_FAILED_QUEUE)


def _load_drive_upload_queue_once():
    global _DRIVE_UPLOAD_QUEUE_LOADED
    if _DRIVE_UPLOAD_QUEUE_LOADED:
        return
    path = _get_drive_upload_queue_path()
    if not os.path.exists(path):
        _DRIVE_UPLOAD_QUEUE_LOADED = True
        return
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except Exception:
        _DRIVE_UPLOAD_QUEUE_LOADED = True
        return
    if isinstance(data, list):
        with _DRIVE_UPLOAD_LOCK:
            _DRIVE_UPLOAD_QUEUE[:] = [item for item in data if isinstance(item, dict)]
    _DRIVE_UPLOAD_QUEUE_LOADED = True


def _load_drive_failed_queue_once():
    global _DRIVE_UPLOAD_FAILED_LOADED
    if _DRIVE_UPLOAD_FAILED_LOADED:
        return
    path = _get_drive_upload_failed_queue_path()
    if not os.path.exists(path):
        _DRIVE_UPLOAD_FAILED_LOADED = True
        return
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except Exception:
        _DRIVE_UPLOAD_FAILED_LOADED = True
        return
    if isinstance(data, list):
        with _DRIVE_UPLOAD_LOCK:
            _DRIVE_UPLOAD_FAILED_QUEUE[:] = [item for item in data if isinstance(item, dict)]
    _DRIVE_UPLOAD_FAILED_LOADED = True


def _get_drive_upload_queue_snapshot(limit=200):
    path = _get_drive_upload_queue_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    rows = [item for item in data if isinstance(item, dict)]
    rows.sort(key=lambda r: float(r.get("next_try_at") or 0))
    if limit and limit > 0:
        rows = rows[: int(limit)]
    return rows


def _get_drive_failed_uploads_snapshot(limit=200):
    path = _get_drive_upload_failed_queue_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    rows = [item for item in data if isinstance(item, dict)]
    rows.sort(key=lambda r: float(r.get("failed_at") or 0), reverse=True)
    if limit and limit > 0:
        rows = rows[: int(limit)]
    return rows


def _get_drive_upload_queue_stats():
    pending_rows = _get_drive_upload_queue_snapshot(limit=0)
    failed_rows = _get_drive_failed_uploads_snapshot(limit=0)
    pending = len(pending_rows)
    if not pending_rows:
        return {
            "pending": 0,
            "failed": len(failed_rows),
            "max_attempts": 0,
            "oldest_next_try_at": None,
        }
    max_attempts = max(int(r.get("attempts") or 0) for r in pending_rows)
    oldest_next_try_at = min(float(r.get("next_try_at") or 0) for r in pending_rows)
    return {
        "pending": pending,
        "failed": len(failed_rows),
        "max_attempts": max_attempts,
        "oldest_next_try_at": oldest_next_try_at,
    }


def _remove_drive_job_locked(job):
    identity = _drive_job_identity(job)
    _DRIVE_UPLOAD_QUEUE[:] = [
        item for item in _DRIVE_UPLOAD_QUEUE if _drive_job_identity(item) != identity
    ]
    _DRIVE_UPLOAD_FAILED_QUEUE[:] = [
        item for item in _DRIVE_UPLOAD_FAILED_QUEUE if _drive_job_identity(item) != identity
    ]


def _store_failed_drive_job_locked(job, error):
    record = dict(job or {})
    record["error"] = str(error or "").strip()
    record["failed_at"] = time.time()
    record["attempts"] = int(record.get("attempts") or 0)
    identity = _drive_job_identity(record)
    _DRIVE_UPLOAD_FAILED_QUEUE[:] = [
        item for item in _DRIVE_UPLOAD_FAILED_QUEUE if _drive_job_identity(item) != identity
    ]
    _DRIVE_UPLOAD_FAILED_QUEUE.append(record)
    if len(_DRIVE_UPLOAD_FAILED_QUEUE) > 2000:
        _DRIVE_UPLOAD_FAILED_QUEUE[:] = _DRIVE_UPLOAD_FAILED_QUEUE[-2000:]


def _next_drive_retry_delay_seconds(attempts):
    tries = max(1, int(attempts))
    return min(300, 2 ** min(tries, 8))


def _update_form_completion_upload_status(
    registro_id,
    *,
    upload_status,
    upload_error=None,
    upload_attempted_at=None,
    uploaded_at=None,
    path_formato=None,
    drive_file_id=None,
):
    registro_key = str(registro_id or "").strip()
    if not registro_key:
        return False
    values = {"upload_status": str(upload_status or "").strip()}
    if upload_error is not None:
        values["upload_error"] = str(upload_error or "").strip()
    if upload_attempted_at is not None:
        values["upload_attempted_at"] = str(upload_attempted_at or "").strip()
    if uploaded_at is not None:
        values["uploaded_at"] = str(uploaded_at or "").strip()
    if path_formato is not None:
        values["path_formato"] = str(path_formato or "").strip()
    if drive_file_id is not None:
        values["drive_file_id"] = str(drive_file_id or "").strip()
    try:
        result = _supabase_patch_with_queue(
            "formatos_finalizados_il",
            {"registro_id": registro_key},
            values,
        )
        return (result or {}).get("status") in {"synced", "queued"}
    except Exception as exc:
        _log_capture(
            f"formatos_finalizados_il patch failed registro_id={registro_key} status={upload_status} err={exc}"
        )
        return False


def _is_transient_drive_exception(exc):
    if exc is None:
        return False
    root = exc
    if isinstance(root, RuntimeError) and getattr(root, "__cause__", None) is not None:
        root = root.__cause__

    if isinstance(root, urllib.error.HTTPError):
        code = int(getattr(root, "code", 0) or 0)
        return code >= 500 or code == 429
    if isinstance(root, urllib.error.URLError):
        return True
    if isinstance(root, TimeoutError):
        return True
    if isinstance(root, OSError):
        return True

    resp = getattr(root, "resp", None)
    status = int(getattr(resp, "status", 0) or getattr(root, "status_code", 0) or 0)
    if status >= 500 or status == 429:
        return True
    if status in {400, 401, 403, 404}:
        return False

    text = str(root).lower()
    transient_markers = (
        "timeout",
        "timed out",
        "temporarily unavailable",
        "temporary failure",
        "connection aborted",
        "connection reset",
        "network is unreachable",
        "name resolution",
        "failed to establish a new connection",
        "connection refused",
        "remote end closed connection",
        "ssl",
    )
    if any(marker in text for marker in transient_markers):
        return True
    permanent_markers = (
        "permission",
        "forbidden",
        "insufficient permissions",
        "no existe el json",
        "falta google_drive",
        "invalid",
        "not found",
        "no existe el archivo",
    )
    if any(marker in text for marker in permanent_markers):
        return False
    return False


def _drive_upload_operation_label(job):
    upload_kind = str((job or {}).get("upload_kind") or "excel_upload").strip().lower()
    if upload_kind == "evaluacion_sheet":
        return "sheet_copy"
    return "upload"


def _perform_drive_upload_attempt(job):
    attempted_at = _get_colombia_now().isoformat()
    excel_path = str((job or {}).get("local_excel_path") or "").strip()
    company_name = str((job or {}).get("company_name") or "").strip()
    upload_kind = str((job or {}).get("upload_kind") or "excel_upload").strip().lower()
    try:
        if upload_kind == "evaluacion_sheet":
            sheet_export = (job or {}).get("sheet_export") or {}
            if not isinstance(sheet_export, dict):
                raise RuntimeError("El payload de Google Sheets es inválido.")
            remote_file_name = str((job or {}).get("remote_file_name") or "").strip()
            if not remote_file_name:
                remote_file_name = os.path.splitext(os.path.basename(excel_path))[0]
            upload_result = drive_upload.publish_evaluacion_accesibilidad_sheet(
                sheet_writes=sheet_export.get("writes") or [],
                clear_ranges=sheet_export.get("clear_ranges") or [],
                base_name=remote_file_name,
                folder_name=company_name,
            )
        else:
            upload_result = drive_upload.upload_excel_to_drive(
                excel_path,
                base_name=os.path.basename(excel_path),
                folder_name=company_name,
            )
    except Exception as exc:
        error = str(exc).strip() or repr(exc)
        status = "pending" if _is_transient_drive_exception(exc) else "failed"
        _update_form_completion_upload_status(
            job.get("registro_id"),
            upload_status=status,
            upload_error=error,
            upload_attempted_at=attempted_at,
        )
        return {
            "status": status,
            "error": error,
            "attempted_at": attempted_at,
            "output_path": excel_path,
            "remote_url": "",
            "drive_file_id": "",
        }

    remote_url = str(upload_result.get("webViewLink") or "").strip()
    drive_file_id = str(upload_result.get("file_id") or "").strip()
    uploaded_at = _get_colombia_now().isoformat()
    _update_form_completion_upload_status(
        job.get("registro_id"),
        upload_status="synced",
        upload_error="",
        upload_attempted_at=attempted_at,
        uploaded_at=uploaded_at,
        path_formato=remote_url,
        drive_file_id=drive_file_id,
    )
    return {
        "status": "synced",
        "error": "",
        "attempted_at": attempted_at,
        "uploaded_at": uploaded_at,
        "output_path": excel_path,
        "remote_url": remote_url,
        "drive_file_id": drive_file_id,
    }


def _drive_upload_worker_loop():
    while True:
        job = None
        with _DRIVE_UPLOAD_LOCK:
            now = time.time()
            for item in _DRIVE_UPLOAD_QUEUE:
                if float(item.get("next_try_at") or 0) <= now:
                    job = dict(item)
                    break

        if not job:
            time.sleep(0.8)
            continue

        result = _perform_drive_upload_attempt(job)
        _do_persist_upload = False
        _do_persist_failed = False
        _skip_sleep = False
        with _DRIVE_UPLOAD_LOCK:
            current = None
            for item in _DRIVE_UPLOAD_QUEUE:
                if item.get("id") == job.get("id"):
                    current = item
                    break
            if current is None:
                _skip_sleep = True
            elif result.get("status") == "synced":
                _remove_drive_job_locked(job)
                _do_persist_upload = True
                _do_persist_failed = True
                _skip_sleep = True
            elif result.get("status") == "failed":
                _remove_drive_job_locked(job)
                failed_job = dict(job)
                failed_job["attempts"] = int(job.get("attempts") or 0) + 1
                _store_failed_drive_job_locked(failed_job, result.get("error"))
                _do_persist_upload = True
                _do_persist_failed = True
                _skip_sleep = True
            else:
                for idx, item in enumerate(_DRIVE_UPLOAD_QUEUE):
                    if item.get("id") != job.get("id"):
                        continue
                    item["attempts"] = int(item.get("attempts") or 0) + 1
                    item["last_error"] = str(result.get("error") or "")
                    item["next_try_at"] = time.time() + _next_drive_retry_delay_seconds(item["attempts"])
                    _DRIVE_UPLOAD_QUEUE[idx] = item
                    break
                _do_persist_upload = True
        if _do_persist_upload:
            _persist_drive_upload_queue_locked()
        if _do_persist_failed:
            _persist_drive_failed_queue_locked()
        if _skip_sleep:
            time.sleep(0.2)
            continue
        time.sleep(0.4)


def _ensure_drive_upload_worker():
    global _DRIVE_UPLOAD_WORKER_STARTED
    with _DRIVE_UPLOAD_LOCK:
        if _DRIVE_UPLOAD_WORKER_STARTED:
            return
        _load_drive_upload_queue_once()
        _load_drive_failed_queue_once()
        worker = threading.Thread(target=_drive_upload_worker_loop, daemon=True)
        worker.start()
        _DRIVE_UPLOAD_WORKER_STARTED = True


def _enqueue_drive_upload_job(job):
    _ensure_drive_upload_worker()
    record = {
        "id": str(uuid.uuid4()),
        "registro_id": str((job or {}).get("registro_id") or "").strip(),
        "form_name": str((job or {}).get("form_name") or "").strip(),
        "company_name": str((job or {}).get("company_name") or "").strip(),
        "local_excel_path": str((job or {}).get("local_excel_path") or "").strip(),
        "upload_kind": str((job or {}).get("upload_kind") or "excel_upload").strip(),
        "remote_file_name": str((job or {}).get("remote_file_name") or "").strip(),
        "sheet_export": (job or {}).get("sheet_export") or {},
        "attempts": int((job or {}).get("attempts") or 0),
        "last_error": str((job or {}).get("last_error") or "").strip(),
        "next_try_at": float((job or {}).get("next_try_at") or time.time()),
    }
    with _DRIVE_UPLOAD_LOCK:
        _remove_drive_job_locked(record)
        _DRIVE_UPLOAD_QUEUE.append(record)
        _persist_drive_upload_queue_locked()
        _persist_drive_failed_queue_locked()
    return record["id"]


def _drive_retry_all_queued_uploads():
    _ensure_drive_upload_worker()
    with _DRIVE_UPLOAD_LOCK:
        if not _DRIVE_UPLOAD_QUEUE:
            return 0
        now = time.time()
        for idx, item in enumerate(_DRIVE_UPLOAD_QUEUE):
            item["next_try_at"] = now
            _DRIVE_UPLOAD_QUEUE[idx] = item
        _persist_drive_upload_queue_locked()
        return len(_DRIVE_UPLOAD_QUEUE)


def _dpapi_encrypt_text(plain_text):
    if os.name != "nt":
        return ""
    text = str(plain_text or "")
    if not text:
        return ""
    try:
        from ctypes import wintypes

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [
                ("cbData", wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_byte)),
            ]

        def _make_blob(data_bytes):
            if not data_bytes:
                return DATA_BLOB(0, None), None
            buf = (ctypes.c_byte * len(data_bytes)).from_buffer_copy(data_bytes)
            return DATA_BLOB(len(data_bytes), ctypes.cast(buf, ctypes.POINTER(ctypes.c_byte))), buf

        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        CRYPTPROTECT_UI_FORBIDDEN = 0x01

        in_blob, in_buf = _make_blob(text.encode("utf-8"))
        entropy_blob, entropy_buf = _make_blob(APP_NAME.encode("utf-8"))
        out_blob = DATA_BLOB()
        ok = crypt32.CryptProtectData(
            ctypes.byref(in_blob),
            None,
            ctypes.byref(entropy_blob),
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            ctypes.byref(out_blob),
        )
        if not ok:
            return ""
        try:
            encrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
            return base64.b64encode(encrypted).decode("ascii")
        finally:
            if out_blob.pbData:
                kernel32.LocalFree(out_blob.pbData)
            _ = in_buf, entropy_buf
    except Exception:
        return ""


def _dpapi_decrypt_text(cipher_b64):
    if os.name != "nt":
        return ""
    payload = str(cipher_b64 or "").strip()
    if not payload:
        return ""
    try:
        encrypted = base64.b64decode(payload.encode("ascii"), validate=False)
    except Exception:
        return ""
    if not encrypted:
        return ""
    try:
        from ctypes import wintypes

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [
                ("cbData", wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_byte)),
            ]

        def _make_blob(data_bytes):
            if not data_bytes:
                return DATA_BLOB(0, None), None
            buf = (ctypes.c_byte * len(data_bytes)).from_buffer_copy(data_bytes)
            return DATA_BLOB(len(data_bytes), ctypes.cast(buf, ctypes.POINTER(ctypes.c_byte))), buf

        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        CRYPTPROTECT_UI_FORBIDDEN = 0x01

        in_blob, in_buf = _make_blob(encrypted)
        entropy_blob, entropy_buf = _make_blob(APP_NAME.encode("utf-8"))
        out_blob = DATA_BLOB()
        ok = crypt32.CryptUnprotectData(
            ctypes.byref(in_blob),
            None,
            ctypes.byref(entropy_blob),
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            ctypes.byref(out_blob),
        )
        if not ok:
            return ""
        try:
            decrypted = ctypes.string_at(out_blob.pbData, out_blob.cbData)
            return decrypted.decode("utf-8", errors="replace")
        finally:
            if out_blob.pbData:
                kernel32.LocalFree(out_blob.pbData)
            _ = in_buf, entropy_buf
    except Exception:
        return ""


def _load_saved_login_credentials():
    payload = {"remember": True, "username": "", "password": ""}
    path = _get_login_credentials_path()
    if not os.path.exists(path):
        return payload
    try:
        with open(path, "r", encoding="utf-8") as handle:
            raw = json.load(handle) or {}
    except Exception:
        return payload
    if not isinstance(raw, dict):
        return payload
    remember = bool(raw.get("remember", True))
    username = str(raw.get("username") or "").strip()
    password = _dpapi_decrypt_text(raw.get("password_enc"))
    payload["remember"] = remember
    payload["username"] = username
    payload["password"] = password if remember else ""
    return payload


def _save_login_credentials(username, password):
    user = str(username or "").strip()
    pwd = str(password or "")
    cipher = _dpapi_encrypt_text(pwd)
    if not user or not cipher:
        return
    payload = {
        "version": 1,
        "remember": True,
        "username": user,
        "password_enc": cipher,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    path = _get_login_credentials_path()
    tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _clear_login_credentials():
    path = _get_login_credentials_path()
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def _load_offline_auth_store():
    path = _get_offline_auth_path()
    if not os.path.exists(path):
        return {"version": 1, "users": {}}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle) or {}
    except Exception:
        return {"version": 1, "users": {}}
    if not isinstance(data, dict):
        return {"version": 1, "users": {}}
    users = data.get("users")
    if not isinstance(users, dict):
        data["users"] = {}
    data.setdefault("version", 1)
    return data


def _save_offline_auth_store(data):
    path = _get_offline_auth_path()
    tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _load_drafts_store():
    path = _get_drafts_path()
    if not os.path.exists(path):
        return {"version": 1, "users": {}}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle) or {}
    except Exception:
        return {"version": 1, "users": {}}
    if not isinstance(data, dict):
        return {"version": 1, "users": {}}
    users = data.get("users")
    if not isinstance(users, dict):
        data["users"] = {}
    data.setdefault("version", 1)
    return data


def _save_drafts_store(data):
    path = _get_drafts_path()
    tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _extract_draft_company_name(cache_snapshot):
    if not isinstance(cache_snapshot, dict):
        return ""
    section_1 = cache_snapshot.get("section_1") or {}
    if isinstance(section_1, dict):
        company_name = (
            section_1.get("nombre_empresa")
            or section_1.get("empresa")
            or section_1.get("razon_social")
            or ""
        )
        return str(company_name).strip()
    return ""


def _extract_draft_company_key(cache_snapshot):
    if not isinstance(cache_snapshot, dict):
        return "sin_clave"
    section_1 = cache_snapshot.get("section_1") or {}
    if not isinstance(section_1, dict):
        return "sin_clave"
    nit = str(section_1.get("nit_empresa") or section_1.get("nit") or "").strip()
    if nit:
        return f"nit:{nit}"
    company_name = _extract_draft_company_name(cache_snapshot)
    if company_name:
        return f"empresa:{_normalize_ascii_text(company_name).lower()}"
    return "sin_clave"


def _resolve_form_meta(form_id):
    for item in get_forms():
        if str(item.get("id") or "") == str(form_id or ""):
            return item
    return {"id": str(form_id or ""), "name": str(form_id or "")}


def _iter_widget_paths(root):
    def _walk(node, prefix=""):
        children = list(node.winfo_children())
        for idx, child in enumerate(children):
            path = f"{prefix}.{idx}" if prefix else str(idx)
            yield path, child
            yield from _walk(child, path)

    yield from _walk(root)


def _widget_from_path(root, path):
    node = root
    if not path:
        return None
    try:
        for token in str(path).split("."):
            children = list(node.winfo_children())
            node = children[int(token)]
        return node
    except Exception:
        return None


def _is_descendant_of(widget, ancestor):
    if not widget or not ancestor:
        return False
    node = widget
    while node is not None:
        if node == ancestor:
            return True
        node = getattr(node, "master", None)
    return False


def _get_widget_value_for_snapshot(widget):
    try:
        if isinstance(widget, tk.Text):
            return widget.get("1.0", tk.END).rstrip("\n")
        if isinstance(widget, ttk.Combobox):
            return widget.get()
        if isinstance(widget, (tk.Entry, DateEntry)):
            state = str(widget.cget("state") or "")
            if state == "readonly":
                return None
            return widget.get()
    except Exception:
        return None
    return None


def _set_widget_value_from_snapshot(widget, value):
    try:
        if isinstance(widget, tk.Text):
            widget.delete("1.0", tk.END)
            widget.insert("1.0", str(value or ""))
            return True
        if isinstance(widget, ttk.Combobox):
            widget.set(str(value or ""))
            return True
        if isinstance(widget, (tk.Entry, DateEntry)):
            state = str(widget.cget("state") or "")
            if state == "readonly":
                return False
            widget.delete(0, tk.END)
            widget.insert(0, str(value or ""))
            return True
    except Exception:
        return False
    return False


def _collect_visible_input_snapshot(window):
    sticky_bar = getattr(window, "_sticky_actions_bar", None)
    rows = []
    for path, widget in _iter_widget_paths(window):
        if sticky_bar and _is_descendant_of(widget, sticky_bar):
            continue
        value = _get_widget_value_for_snapshot(widget)
        if value is None:
            continue
        rows.append(
            {
                "path": path,
                "class": widget.__class__.__name__,
                "value": value,
            }
        )
    return rows


def _apply_input_snapshot(window, snapshot_rows):
    if not isinstance(snapshot_rows, list):
        return 0
    applied = 0
    for row in snapshot_rows:
        if not isinstance(row, dict):
            continue
        widget = _widget_from_path(window, row.get("path"))
        if not widget:
            continue
        if _set_widget_value_from_snapshot(widget, row.get("value")):
            applied += 1
    return applied


def _snapshot_has_meaningful_values(snapshot_rows):
    if not isinstance(snapshot_rows, list):
        return False
    for row in snapshot_rows:
        if not isinstance(row, dict):
            continue
        if str(row.get("value") or "").strip():
            return True
    return False


def _cache_snapshot_has_meaningful_values(value):
    if isinstance(value, dict):
        for key, item in value.items():
            if str(key or "").startswith("_"):
                continue
            if _cache_snapshot_has_meaningful_values(item):
                return True
        return False
    if isinstance(value, list):
        return any(_cache_snapshot_has_meaningful_values(item) for item in value)
    return str(value or "").strip() != ""


def _get_draft_save_command(window):
    save_cmd = getattr(window, "_save_draft_command", None)
    if callable(save_cmd):
        return save_cmd
    form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
    if not form_id:
        return None
    module = FORM_MODULE_MAP.get(form_id)
    hub = getattr(window, "master", None)
    if (
        not hub
        or not hasattr(hub, "_save_current_form_draft")
        or not module
        or not hasattr(module, "get_form_cache")
        or not hasattr(module, "save_cache_to_file")
    ):
        return None
    form_meta = _resolve_form_meta(form_id)
    window._form_id = form_id
    window._form_name = str(form_meta.get("name") or form_id)
    window._save_draft_command = lambda w=window, h=hub: h._save_current_form_draft(w)
    return window._save_draft_command


def _consume_pending_draft_restore(window, form_id, module, section_routes, default_route):
    hub = getattr(window, "master", None)
    pending = getattr(hub, "_pending_draft_restore", None)
    if not isinstance(pending, dict):
        return False
    if str(pending.get("form_id") or "") != str(form_id or ""):
        return False

    hub._pending_draft_restore = None
    cache_snapshot = pending.get("cache")
    if not isinstance(cache_snapshot, dict):
        cache_snapshot = {}
    try:
        if hasattr(module, "clear_form_cache"):
            module.clear_form_cache()
        form_cache = getattr(module, "FORM_CACHE", None)
        if isinstance(form_cache, dict):
            form_cache.clear()
            form_cache.update(copy.deepcopy(cache_snapshot))
        if hasattr(module, "save_cache_to_file"):
            module.save_cache_to_file()
    except Exception:
        pass

    window._draft_restore_pending_ui_snapshot = pending.get("ui_snapshot")
    window._draft_restore_target_section = str(
        pending.get("ui_section") or cache_snapshot.get("_last_section") or ""
    ).strip()
    route = section_routes.get(window._draft_restore_target_section) or default_route
    if callable(route):
        try:
            window.after_idle(route)
        except Exception:
            route()
        return True
    return False


def _normalize_ascii_text(value):
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[\\/:*?\"<>|]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _sanitize_sheet_name(value, fallback="Hoja"):
    text = _normalize_ascii_text(value) or fallback
    text = text.replace("[", "").replace("]", "").replace(":", "")
    return text[:31] if len(text) > 31 else text


def _detect_mojibake_issues(project_root):
    issues = []
    include_roots = [
        os.path.join(project_root, "app.py"),
        os.path.join(project_root, "formularios"),
    ]
    for target in include_roots:
        if os.path.isfile(target):
            candidates = [target]
        elif os.path.isdir(target):
            candidates = []
            for root, _dirs, files in os.walk(target):
                for name in files:
                    if name.lower().endswith(".py"):
                        candidates.append(os.path.join(root, name))
        else:
            continue
        for path in candidates:
            try:
                with open(path, "r", encoding="utf-8", errors="replace") as handle:
                    for lineno, line in enumerate(handle, start=1):
                        # Evita falso positivo por la línea de definición de patrones.
                        if "_MOJIBAKE_PATTERNS" in line:
                            continue
                        if any(mark in line for mark in _MOJIBAKE_PATTERNS):
                            text = line.strip()
                            issues.append((path, lineno, text[:180]))
                            if len(issues) >= 200:
                                return issues
            except Exception:
                continue
    return issues


def _run_encoding_health_check():
    global _ENCODING_CHECK_DONE
    if _ENCODING_CHECK_DONE:
        return
    _ENCODING_CHECK_DONE = True
    root = os.path.dirname(os.path.abspath(__file__))
    issues = _detect_mojibake_issues(root)
    if not issues:
        return
    _log_capture("Posibles problemas de encoding/mojibake detectados:")
    for path, lineno, snippet in issues:
        rel = os.path.relpath(path, root)
        _log_capture(f"[ENCODING] {rel}:{lineno} -> {snippet}")


def _digits_only(value, max_len=None):
    cleaned = re.sub(r"\D+", "", str(value or ""))
    if max_len is not None:
        cleaned = cleaned[: int(max_len)]
    return cleaned


def _normalize_person_name(value):
    cleaned = "".join(ch for ch in str(value or "") if ch.isalpha() or ch.isspace())
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return " ".join(word.capitalize() for word in cleaned.split())


def _bind_numeric_entry(entry, max_len=None):
    def _on_key_release(_event=None):
        raw = _digits_only(entry.get(), max_len=max_len)
        if entry.get() == raw:
            return
        entry.delete(0, tk.END)
        entry.insert(0, raw)

    entry.bind("<KeyRelease>", _on_key_release)


def _bind_decimal_entry(entry):
    def _normalize_current(*, allow_trailing_separator):
        raw = _normalize_decimal_value(
            entry.get(),
            allow_trailing_separator=allow_trailing_separator,
        )
        if entry.get() == raw:
            return
        entry.delete(0, tk.END)
        entry.insert(0, raw)

    entry.bind(
        "<KeyRelease>",
        lambda _event=None: _normalize_current(allow_trailing_separator=True),
    )
    entry.bind(
        "<FocusOut>",
        lambda _event=None: _normalize_current(allow_trailing_separator=False),
        add="+",
    )


def _bind_name_entry(entry):
    def _on_key_release(_event=None):
        filtered = "".join(ch for ch in entry.get() if ch.isalpha() or ch.isspace())
        if filtered == entry.get():
            return
        entry.delete(0, tk.END)
        entry.insert(0, filtered)

    def _on_focus_out(_event=None):
        normalized = _normalize_person_name(entry.get())
        entry.delete(0, tk.END)
        entry.insert(0, normalized)

    entry.bind("<KeyRelease>", _on_key_release)
    entry.bind("<FocusOut>", _on_focus_out)


def _set_readonly_entry_value(entry, value):
    entry.configure(state="normal")
    entry.delete(0, tk.END)
    entry.insert(0, str(value or ""))
    entry.configure(state="readonly")


def _format_birthdate_text(value):
    digits = _digits_only(value, max_len=8)
    if len(digits) <= 2:
        formatted = digits
    elif len(digits) <= 4:
        formatted = f"{digits[:2]}/{digits[2:]}"
    else:
        formatted = f"{digits[:2]}/{digits[2:4]}/{digits[4:]}"
    return digits, formatted


def _calc_age_from_digits(digits, min_year=1900):
    if len(digits) != 8:
        return None
    try:
        day = int(digits[:2])
        month = int(digits[2:4])
        year = int(digits[4:])
        if year < int(min_year):
            return None
        birth_date = date(year, month, day)
        today = date.today()
        if birth_date > today:
            return None
    except Exception:
        return None
    age = today.year - birth_date.year
    if (today.month, today.day) < (birth_date.month, birth_date.day):
        age -= 1
    return age


def _refresh_age_from_date_entry(date_entry, age_entry, min_year=1900):
    digits, _ = _format_birthdate_text(date_entry.get())
    age = _calc_age_from_digits(digits, min_year=min_year)
    _set_readonly_entry_value(age_entry, "" if age is None else age)
    return age


def _bind_birthdate_entry(
    date_entry,
    age_entry,
    *,
    min_year=1900,
    mark_invalid=True,
    clear_invalid=False,
):
    state = {"updating": False}

    def _format_and_validate(_event=None):
        if state["updating"]:
            return
        state["updating"] = True
        digits, formatted = _format_birthdate_text(date_entry.get())
        date_entry.delete(0, tk.END)
        date_entry.insert(0, formatted)

        age = _calc_age_from_digits(digits, min_year=min_year)
        invalid_complete = len(digits) == 8 and age is None
        if invalid_complete and clear_invalid:
            date_entry.delete(0, tk.END)
            digits = ""
            if mark_invalid:
                date_entry.configure(bg="#FDE2E2")
            _set_readonly_entry_value(age_entry, "")
            state["updating"] = False
            return
        if mark_invalid:
            date_entry.configure(bg="#FDE2E2" if invalid_complete else "white")
        _set_readonly_entry_value(age_entry, "" if age is None else age)
        state["updating"] = False

    date_entry.bind("<KeyRelease>", _format_and_validate)
    date_entry.bind("<FocusOut>", _format_and_validate)


def _hash_password(password, iterations=PASSWORD_HASH_ITERATIONS):
    pwd = str(password or "")
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", pwd.encode("utf-8"), salt, iterations)
    salt_b64 = base64.urlsafe_b64encode(salt).decode("ascii").rstrip("=")
    digest_b64 = base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")
    return f"{PASSWORD_HASH_ALGO}${iterations}${salt_b64}${digest_b64}"


def _verify_password_hash_native(password, stored_hash):
    if not stored_hash or "$" not in str(stored_hash):
        return False
    try:
        algo, iter_s, salt_b64, digest_b64 = str(stored_hash).split("$", 3)
        if algo != PASSWORD_HASH_ALGO:
            return False
        iterations = int(iter_s)
        salt = base64.urlsafe_b64decode(salt_b64 + "=" * (-len(salt_b64) % 4))
        expected = base64.urlsafe_b64decode(digest_b64 + "=" * (-len(digest_b64) % 4))
        current = hashlib.pbkdf2_hmac(
            "sha256",
            str(password or "").encode("utf-8"),
            salt,
            iterations,
        )
        return hmac.compare_digest(current, expected)
    except Exception:
        return False


def _verify_password_hash_legacy(password, stored_hash):
    if not stored_hash or "$" not in str(stored_hash):
        return False
    try:
        algo, iter_s, salt_text, digest_b64 = str(stored_hash).split("$", 3)
        if algo != PASSWORD_HASH_ALGO:
            return False
        iterations = int(iter_s)
        expected = base64.b64decode(digest_b64 + "=" * (-len(digest_b64) % 4))
        current = hashlib.pbkdf2_hmac(
            "sha256",
            str(password or "").encode("utf-8"),
            str(salt_text or "").encode("utf-8"),
            iterations,
        )
        return hmac.compare_digest(current, expected)
    except Exception:
        return False


def _verify_password_hash(password, stored_hash):
    return _verify_password_hash_native(password, stored_hash) or _verify_password_hash_legacy(
        password,
        stored_hash,
    )


def _normalize_login_value(value):
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", "", text).lower().strip()
    return text


def _get_usage_exempt_logins():
    global _USAGE_EXEMPT_LOGINS_CACHE
    if _USAGE_EXEMPT_LOGINS_CACHE is not None:
        return _USAGE_EXEMPT_LOGINS_CACHE
    raw = str(os.getenv("USAGE_EXEMPT_LOGINS") or "").strip()
    if not raw:
        try:
            env = _load_env_file(".env") or {}
            raw = str(env.get("USAGE_EXEMPT_LOGINS") or "").strip()
        except Exception:
            raw = ""
    tokens = re.split(r"[,\n;]+", raw)
    _USAGE_EXEMPT_LOGINS_CACHE = {
        _normalize_login_value(token)
        for token in tokens
        if _normalize_login_value(token)
    }
    return _USAGE_EXEMPT_LOGINS_CACHE


def _password_candidates(password):
    raw = str(password or "")
    options = [raw]
    trimmed = raw.strip()
    if trimmed != raw:
        options.append(trimmed)
    return options


def _is_connectivity_exception(exc):
    if exc is None:
        return False
    root = exc
    if isinstance(root, RuntimeError) and getattr(root, "__cause__", None) is not None:
        root = root.__cause__
    if isinstance(root, urllib.error.HTTPError):
        code = int(getattr(root, "code", 0) or 0)
        return code >= 500 or code == 429
    if isinstance(root, urllib.error.URLError):
        return True
    if isinstance(root, TimeoutError):
        return True
    if isinstance(root, OSError):
        return True
    text = str(exc).lower()
    return "supabase no esta disponible" in text or "timed out" in text


def check_internet(timeout=3, log_enabled=False):
    started_at = time.perf_counter()

    def _result(ok, status_text, error_code="", detail=""):
        payload = {
            "ok": bool(ok),
            "status_text": str(status_text or "").strip(),
            "error_code": str(error_code or "").strip(),
            "detail": str(detail or "").strip(),
            "latency_ms": int((time.perf_counter() - started_at) * 1000),
        }
        if log_enabled:
            level = "INFO" if ok else "ERROR"
            log_app_event(
                f"[INTERNET_PROBE] ok={payload['ok']} status={payload['status_text']!r} "
                f"code={payload['error_code']!r} detail={payload['detail']!r} "
                f"latency_ms={payload['latency_ms']}",
                level=level,
            )
        return payload

    request = urllib.request.Request(
        "https://www.gstatic.com/generate_204",
        headers={"User-Agent": f"{APP_NAME}/1.0"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            status = int(getattr(response, "status", 204) or 204)
        return _result(True, "Conectado", "", f"http_status={status}")
    except urllib.error.HTTPError as exc:
        code = int(getattr(exc, "code", 0) or 0)
        if 200 <= code < 500:
            return _result(True, "Conectado", "", f"http_status={code}")
        return _result(False, f"HTTP {code}", "http_error", exc)
    except Exception as exc:
        return _result(False, "Sin internet", "connectivity", exc)


def _result_with_dependency_down(status_text):
    return {
        "ok": False,
        "status_text": status_text,
        "error_code": "dependency",
        "detail": "Sin internet",
        "latency_ms": 0,
    }


def probe_startup_services(log_enabled=False):
    internet = check_internet(log_enabled=log_enabled)
    if not internet.get("ok"):
        return {
            "internet": internet,
            "supabase": _result_with_dependency_down("Sin internet"),
            "drive": _result_with_dependency_down("Sin internet"),
        }
    return {
        "internet": internet,
        "supabase": probe_supabase_service(log_enabled=log_enabled),
        "drive": drive_upload.probe_drive_service(log_enabled=log_enabled),
    }


def _is_seguimiento_form(form_name):
    return "seguimiento" in str(form_name or "").strip().lower()


def _build_company_workbook_path(individual_excel_path, company_name):
    company_safe = _normalize_ascii_text(company_name) or "Empresa"
    base_dir = os.path.dirname(individual_excel_path)
    parent_name = _normalize_ascii_text(os.path.basename(base_dir))
    folder = base_dir if parent_name.lower() == company_safe.lower() else os.path.join(base_dir, company_safe)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, f"{company_safe}.xlsx")


def _append_sheet_to_company_workbook(individual_excel_path, company_name, form_name):
    if not individual_excel_path or not os.path.exists(individual_excel_path):
        raise RuntimeError("No existe el Excel individual para consolidar.")
    if _is_seguimiento_form(form_name):
        return individual_excel_path

    try:
        import win32com.client as win32  # pyright: ignore[reportMissingModuleSource]
    except ImportError as exc:
        raise RuntimeError("Falta pywin32 para consolidar hojas por empresa.") from exc

    company_workbook_path = _build_company_workbook_path(individual_excel_path, company_name)
    sheet_base = _sanitize_sheet_name(form_name, fallback="Formulario")

    if not os.path.exists(company_workbook_path):
        shutil.copy2(individual_excel_path, company_workbook_path)
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = None
        try:
            wb = excel.Workbooks.Open(os.path.abspath(company_workbook_path))
            existing_names = {str(wb.Worksheets(i).Name) for i in range(1, wb.Worksheets.Count + 1)}
            first_sheet = wb.Worksheets(1)
            next_name = sheet_base
            suffix = 2
            while next_name in existing_names and next_name != str(first_sheet.Name):
                next_name = _sanitize_sheet_name(f"{sheet_base} {suffix}", fallback=sheet_base)
                suffix += 1
            first_sheet.Name = next_name
            wb.Save()
            return company_workbook_path
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                excel.Quit()
            except Exception:
                pass

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    src_wb = None
    dst_wb = None
    try:
        src_wb = excel.Workbooks.Open(os.path.abspath(individual_excel_path))
        dst_wb = excel.Workbooks.Open(os.path.abspath(company_workbook_path))
        existing_names = {str(dst_wb.Worksheets(i).Name) for i in range(1, dst_wb.Worksheets.Count + 1)}

        # In COM automation, cross-workbook copy is reliable with "Before".
        src_wb.Worksheets(1).Copy(Before=dst_wb.Worksheets(1))
        new_sheet = dst_wb.Worksheets(1)

        next_name = sheet_base
        suffix = 2
        while next_name in existing_names:
            next_name = _sanitize_sheet_name(f"{sheet_base} {suffix}", fallback=sheet_base)
            suffix += 1
        new_sheet.Name = next_name
        dst_wb.Save()
        return company_workbook_path
    finally:
        try:
            if src_wb is not None:
                src_wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if dst_wb is not None:
                dst_wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass


def _maximize_window(window):
    screen_w = None
    screen_h = None
    try:
        screen_w = int(window.winfo_screenwidth() or 0)
        screen_h = int(window.winfo_screenheight() or 0)
    except Exception:
        screen_w = screen_h = 0
    try:
        window.state("zoomed")
    except tk.TclError:
        try:
            window.attributes("-zoomed", True)
        except tk.TclError:
            if screen_w and screen_h:
                window.geometry(f"{screen_w}x{screen_h}+0+0")
    try:
        window.update_idletasks()
        if screen_w and screen_h:
            cur_w = int(window.winfo_width() or 0)
            cur_h = int(window.winfo_height() or 0)
            if cur_w < int(screen_w * 0.9) or cur_h < int(screen_h * 0.9):
                window.geometry(f"{screen_w}x{screen_h}+0+0")
    except Exception:
        pass


def _finish_with_loading(loading, message, open_target=None, open_prompt=None):
    loading.set_status("Listo")
    loading.set_progress(100)
    try:
        loading.window.grab_release()
    except tk.TclError:
        pass
    if open_target:
        open_file = messagebox.askyesno(
            "Listo",
            f"{message}\n\n{open_prompt or '¿Quieres abrirlo?'}",
        )
        if open_file:
            try:
                if str(open_target).startswith("http"):
                    webbrowser.open(open_target)
                else:
                    os.startfile(open_target)
            except Exception as exc:
                messagebox.showerror(
                    "Error",
                    f"No se pudo abrir el destino.\n{exc}",
                )
    else:
        messagebox.showinfo("Listo", message)
    loading.close()


def _focus_window(window):
    try:
        window.lift()
        window.focus_force()
        window.attributes("-topmost", True)
        window.after(150, lambda: window.attributes("-topmost", False))
    except tk.TclError:
        return


def _return_to_hub(window):
    hub = window.master if isinstance(window.master, HubWindow) else None
    if not hub:
        return
    try:
        hub.deiconify()
        hub.lift()
        hub.focus_force()
    except tk.TclError:
        return


def _finalize_export_flow(window, loading, completion_result):
    result = dict(completion_result or {})
    status = str(result.get("status") or "").strip()
    output_path = str(result.get("output_path") or "").strip()
    remote_url = str(result.get("remote_url") or "").strip()
    error = str(result.get("error") or "").strip()
    hub = window.master if isinstance(window.master, HubWindow) else None

    if hub and status in {"synced", "pending", "failed", "local"}:
        try:
            hub._delete_window_draft(window)
        except Exception as exc:
            _log_capture(
                f"[DRAFT] delete_after_finalize_failed form={getattr(window, '_form_id', '')} status={status} err={exc}"
            )

    if status == "synced":
        open_target = remote_url or output_path
        message = "Formulario completado.\nEl archivo quedó publicado en Google Drive."
        if remote_url:
            message = f"{message}\nLink: {remote_url}"
        elif output_path:
            message = f"{message}\nCopia local: {output_path}"
        _finish_with_loading(
            loading,
            message,
            open_target=open_target,
            open_prompt=(
                "¿Quieres abrirlo en Google Drive?"
                if remote_url
                else "¿Quieres abrir la copia local?"
            ),
        )
        return

    if status == "pending":
        message = (
            "Formulario completado.\n"
            "La copia local quedó guardada y la subida a Google Drive quedó pendiente.\n"
            "La información del formulario se conserva para soporte o reintento manual.\n"
            f"Archivo local: {output_path}"
        )
        if error:
            message = f"{message}\nDetalle: {error}"
        _finish_with_loading(
            loading,
            message,
            open_target=output_path,
            open_prompt="¿Quieres abrir la copia local?",
        )
        return

    if status == "failed":
        message = (
            "Formulario completado.\n"
            "No se pudo publicar el archivo en Google Drive.\n"
            "La información del formulario se conserva para soporte o reintento manual.\n"
            f"Archivo local: {output_path}"
        )
        if error:
            message = f"{message}\nDetalle: {error}"
        _finish_with_loading(
            loading,
            message,
            open_target=output_path,
            open_prompt="¿Quieres abrir la copia local?",
        )
        return

    if status == "local":
        _finish_with_loading(
            loading,
            "Formato completado. ¿Quieres abrir la copia local?",
            open_target=output_path,
            open_prompt="¿Quieres abrir la copia local?",
        )
        return

    _finish_with_loading(
        loading,
        f"Formulario completado.\nArchivo: {output_path}",
        open_target=output_path,
        open_prompt="¿Quieres abrir el archivo?",
    )


def _clear_sticky_actions(window):
    after_id = getattr(window, "_sticky_actions_after_id", None)
    if after_id:
        try:
            window.after_cancel(after_id)
        except Exception:
            pass
        try:
            window._sticky_actions_after_id = None
        except Exception:
            pass
    bar = getattr(window, "_sticky_actions_bar", None)
    if bar and bar.winfo_exists():
        try:
            bar.destroy()
        except Exception:
            pass
    try:
        window._sticky_actions_bar = None
    except Exception:
        pass
    try:
        window._sticky_actions_source = None
    except Exception:
        pass
    try:
        window._sticky_actions_buttons = []
    except Exception:
        pass


def _install_sticky_actions(frame):
    try:
        window = frame.winfo_toplevel()
    except Exception:
        return
    if not isinstance(window, (tk.Tk, tk.Toplevel)):
        return

    source_buttons = [w for w in frame.winfo_children() if isinstance(w, ttk.Button)]
    if not source_buttons:
        _clear_sticky_actions(window)
        return

    _clear_sticky_actions(window)

    bar = tk.Frame(window, bg=COLOR_LIGHT_BG, bd=1, relief="solid")
    inner = tk.Frame(bar, bg=COLOR_LIGHT_BG)
    inner.pack(padx=10, pady=8)

    sticky_buttons = []
    for src in source_buttons:
        clone = ttk.Button(inner, text=src.cget("text"), command=src.invoke)
        side = "left"
        padx = (8, 0)
        try:
            info = src.pack_info()
            side = info.get("side", "left")
        except Exception:
            pass
        clone.pack(side=side, padx=padx)
        sticky_buttons.append((src, clone))

    bar.place(relx=0.5, rely=1.0, anchor="s", y=-10)

    window._sticky_actions_bar = bar
    window._sticky_actions_source = frame
    window._sticky_actions_buttons = sticky_buttons
    window._sticky_actions_after_id = None

    def _sync():
        source = getattr(window, "_sticky_actions_source", None)
        bar_widget = getattr(window, "_sticky_actions_bar", None)
        pairs = getattr(window, "_sticky_actions_buttons", [])
        if not source or not source.winfo_exists() or not bar_widget or not bar_widget.winfo_exists():
            _clear_sticky_actions(window)
            return
        for src_btn, clone_btn in pairs:
            if not src_btn.winfo_exists() or not clone_btn.winfo_exists():
                continue
            try:
                clone_btn.configure(text=src_btn.cget("text"))
                src_state = str(src_btn.cget("state") or "normal")
                clone_btn.state(["disabled"] if src_state == "disabled" else ["!disabled"])
            except Exception:
                continue
        try:
            window._sticky_actions_after_id = window.after(250, _sync)
        except Exception:
            _clear_sticky_actions(window)

    _sync()


def _pack_actions(frame, pad_y=(8, FORM_PADY), pad_x=True):
    try:
        window = frame.winfo_toplevel()
        save_cmd = _get_draft_save_command(window)
        if callable(save_cmd):
            has_save = False
            for child in frame.winfo_children():
                if isinstance(child, ttk.Button) and str(child.cget("text")).strip().lower() == "guardar borrador":
                    has_save = True
                    break
            if not has_save:
                ttk.Button(
                    frame,
                    text="Guardar borrador",
                    command=save_cmd,
                ).pack(side="left", padx=(8, 0))
    except Exception:
        pass

    padx = FORM_PADX if pad_x else 0
    # Keep action buttons grouped and centered to avoid corner placement on small screens.
    frame.pack(anchor="center", pady=pad_y, padx=padx)
    # Duplicate actions in a sticky dock so they remain visible even on long scroll sections.
    try:
        frame.after_idle(lambda: _install_sticky_actions(frame))
    except Exception:
        pass


def _build_wizard_actions(
    parent,
    *,
    back_command=None,
    primary_command=None,
    primary_text="Continuar",
    left_buttons=None,
    right_buttons=None,
    pad_y=(8, FORM_PADY),
    pad_x=True,
):
    actions = tk.Frame(parent, bg=COLOR_LIGHT_BG)
    _pack_actions(actions, pad_y=pad_y, pad_x=pad_x)

    left_specs = []
    if callable(back_command):
        left_specs.append(("Regresar", back_command))
    left_specs.extend(list(left_buttons or []))

    for idx, spec in enumerate(left_specs):
        if not spec:
            continue
        text, command = spec[0], spec[1]
        ttk.Button(actions, text=text, command=command).pack(
            side="left",
            padx=(8, 0) if idx else 0,
        )

    if callable(primary_command):
        ttk.Button(actions, text=primary_text, command=primary_command).pack(side="right")

    for spec in list(right_buttons or []):
        if not spec:
            continue
        text, command = spec[0], spec[1]
        ttk.Button(actions, text=text, command=command).pack(side="right", padx=(0, 8))

    return actions


def _section1_build_search(self, parent, include_tipo_visita=False):
    search_w = ENTRY_W_LONG
    try:
        sw = int(self.winfo_screenwidth() or 0)
        if sw and sw <= 1366:
            search_w = 22
        elif sw and sw <= 1600:
            search_w = 25
    except Exception:
        pass
    frame = tk.Frame(parent, bg=COLOR_LIGHT_BG)
    frame.pack(fill="x", padx=FORM_PADX, pady=(8, FORM_PADY))

    current_row = 0
    if include_tipo_visita:
        tk.Label(
            frame,
            text="Tipo de visita",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=current_row, column=0, sticky="w", padx=(0, 8))

        self.fields["tipo_visita"] = ttk.Combobox(
            frame,
            values=["Presentacion", "Reactivacion"],
            state="readonly",
            width=22,
        )
        self.fields["tipo_visita"].grid(row=current_row, column=1, sticky="w", padx=(0, 24))
        self.fields["tipo_visita"].set("Presentacion")
        current_row += 1

    tk.Label(
        frame,
        text="Número de NIT",
        font=FONT_LABEL,
        bg=COLOR_LIGHT_BG,
    ).grid(row=current_row, column=0, sticky="w", padx=(0, 8))

    self.fields["nit_empresa"] = tk.Entry(frame, width=search_w)
    self.fields["nit_empresa"].grid(row=current_row, column=1, sticky="w")

    search_nit_btn = ttk.Button(
        frame,
        text="Buscar por NIT",
        command=lambda: self._search_company("nit"),
    )
    search_nit_btn.grid(row=current_row, column=2, padx=12)
    current_row += 1

    tk.Label(
        frame,
        text="Nombre de la empresa",
        font=FONT_LABEL,
        bg=COLOR_LIGHT_BG,
    ).grid(row=current_row, column=0, sticky="w", padx=(0, 8))

    self.fields["nombre_busqueda"] = ttk.Combobox(frame, width=search_w)
    self.fields["nombre_busqueda"].grid(row=current_row, column=1, sticky="w")
    self.fields["nombre_busqueda"].bind(
        "<KeyRelease>",
        lambda _event: _section1_update_nombre_suggestions(self),
    )

    search_name_btn = ttk.Button(
        frame,
        text="Buscar por nombre",
        command=lambda: self._search_company("nombre"),
    )
    search_name_btn.grid(row=current_row, column=2, padx=12)

    self.status_label = tk.Label(
        frame,
        text="",
        font=FONT_SUBTITLE,
        fg=COLOR_TEAL,
        bg=COLOR_LIGHT_BG,
    )
    self.status_label.grid(row=current_row + 1, column=0, columnspan=3, sticky="w", pady=(6, 0))


def _section1_build_groups(self, parent, groups, labels, modalidad_options=None):
    readonly_w = ENTRY_W_XL
    try:
        sw = int(self.winfo_screenwidth() or 0)
        if sw and sw <= 1366:
            readonly_w = 42
        elif sw and sw <= 1600:
            readonly_w = 50
    except Exception:
        pass
    container = tk.Frame(parent, bg=COLOR_LIGHT_BG)
    container.pack(fill="both", expand=True)
    self._section1_labels = labels

    top_inputs = tk.Frame(container, bg=COLOR_LIGHT_BG)
    top_inputs.pack(fill="x", pady=(0, FORM_PADY))

    tk.Label(
        top_inputs,
        text="Fecha de la visita",
        font=FONT_LABEL,
        bg=COLOR_LIGHT_BG,
    ).grid(row=0, column=0, sticky="w", padx=(0, 8))
    self.fields["fecha_visita"] = DateEntry(
        top_inputs,
        width=ENTRY_W_MED,
        date_pattern="yyyy-mm-dd",
    )
    self.fields["fecha_visita"].delete(0, tk.END)
    self.fields["fecha_visita"].grid(row=0, column=1, sticky="w", padx=(0, 24))

    tk.Label(
        top_inputs,
        text="Modalidad",
        font=FONT_LABEL,
        bg=COLOR_LIGHT_BG,
    ).grid(row=0, column=2, sticky="w", padx=(0, 8))
    self.fields["modalidad"] = ttk.Combobox(
        top_inputs,
        values=modalidad_options or ["Virtual", "Presencial", "Mixto", "No aplica"],
        state="readonly",
        width=ENTRY_W_MED,
    )
    self.fields["modalidad"].grid(row=0, column=3, sticky="w")

    for title, color, field_ids in groups:
        group_label = tk.Label(
            container,
            text=title,
            bg=color,
            fg=COLOR_PURPLE,
            font=FONT_LABEL,
        )
        group_frame = tk.LabelFrame(
            container,
            labelwidget=group_label,
            bg=color,
            padx=12,
            pady=8,
            bd=1,
        )
        group_frame.pack(fill="x", pady=8)
        group_frame.grid_columnconfigure(1, weight=1)

        for row, field_id in enumerate(field_ids):
            label_text = self._label_for_field(field_id)
            tk.Label(
                group_frame,
                text=label_text,
                font=FONT_LABEL,
                bg=color,
            ).grid(row=row, column=0, sticky="w", padx=6, pady=ROW_PADY)

            entry = tk.Entry(group_frame, state="readonly", width=readonly_w)
            entry.grid(row=row, column=1, sticky="w", padx=6, pady=ROW_PADY)
            self.fields[field_id] = entry


def _section1_build_actions(self, parent):
    actions = tk.Frame(parent, bg=COLOR_LIGHT_BG)
    _pack_actions(actions)
    self.continue_btn = ttk.Button(
        actions,
        text="Continuar",
        command=self._confirm_and_continue,
        state="disabled",
    )
    self.continue_btn.pack(side="right")


def _create_vscroll(parent, command):
    return tk.Scrollbar(parent, orient="vertical", command=command, width=SCROLLBAR_WIDTH)


def _build_scrollable_section_shell(parent, owner=None):
    parent.grid_rowconfigure(0, weight=1)
    parent.grid_columnconfigure(0, weight=1)

    canvas = tk.Canvas(parent, bg=COLOR_LIGHT_BG, highlightthickness=0)
    scrollbar = _create_vscroll(parent, canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    content = tk.Frame(canvas, bg=COLOR_LIGHT_BG)
    content.bind(
        "<Configure>",
        lambda _event: canvas.configure(scrollregion=canvas.bbox("all")),
    )

    content_window = canvas.create_window((0, 0), window=content, anchor="nw")
    canvas.bind(
        "<Configure>",
        lambda event: canvas.itemconfigure(content_window, width=event.width),
    )

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")

    if owner and hasattr(owner, "_bind_mousewheel"):
        try:
            owner._bind_mousewheel(canvas, content)
        except Exception:
            pass

    return content


def _build_scrollable_content(parent, owner=None):
    return _build_scrollable_section_shell(parent, owner)


def _get_required_modalidad(window):
    fields = getattr(window, "fields", {}) or {}
    widget = fields.get("modalidad")
    modalidad = widget.get().strip() if widget else ""
    if modalidad:
        return modalidad
    messagebox.showerror("Campo obligatorio", "Debes seleccionar una modalidad para continuar.")
    try:
        if widget:
            widget.focus_set()
    except Exception:
        pass
    return None


def _get_required_fecha_visita(window):
    fields = getattr(window, "fields", {}) or {}
    widget = fields.get("fecha_visita")
    fecha_visita = widget.get().strip() if widget else ""
    if fecha_visita:
        return fecha_visita
    messagebox.showerror("Campo obligatorio", "Debes seleccionar una fecha de visita para continuar.")
    try:
        if widget:
            widget.focus_set()
    except Exception:
        pass
    return None


def _iter_widget_tree(root):
    stack = [root]
    while stack:
        node = stack.pop()
        yield node
        try:
            children = list(node.winfo_children())
        except Exception:
            children = []
        stack.extend(children)


def _attach_dictation_for_section(window, form_id, section_name):
    if not window or not form_id:
        return
    _clear_section_dictation_button(window)
    container = getattr(window, "section_container", None) or window
    text_widgets = []
    for widget in _iter_widget_tree(container):
        if not isinstance(widget, tk.Text):
            continue
        try:
            if str(widget.cget("state")) == "disabled":
                continue
            height = int(widget.cget("height") or 0)
            if height < 4:
                continue
        except Exception:
            continue
        text_widgets.append(widget)

    for idx, widget in enumerate(text_widgets, start=1):
        field_id = f"{section_name}:text_{idx}"
        try:
            attach_dictation(
                widget,
                form_id=form_id,
                field_id=field_id,
                session_provider=lambda: _supabase_get_access_token(".env"),
                log_fn=_log_capture,
                show_controls=False,
            )
        except Exception as exc:
            _log_capture(
                f"[DICTATION] attach_failed form={form_id} section={section_name} "
                f"field={field_id} err={exc}"
            )
    if text_widgets:
        _install_section_dictation_button(window, text_widgets)


def _clear_section_dictation_button(window):
    after_id = getattr(window, "_section_dictation_after_id", None)
    if after_id:
        try:
            window.after_cancel(after_id)
        except Exception:
            pass
        try:
            window._section_dictation_after_id = None
        except Exception:
            pass
    btn = getattr(window, "_section_dictation_button", None)
    if btn and btn.winfo_exists():
        try:
            btn.destroy()
        except Exception:
            pass
    try:
        window._section_dictation_button = None
    except Exception:
        pass
    try:
        window._section_dictation_widgets = []
    except Exception:
        pass


def _resolve_section_dictation_helper(window):
    widgets = []
    for widget in list(getattr(window, "_section_dictation_widgets", []) or []):
        try:
            if widget.winfo_exists():
                widgets.append(widget)
        except Exception:
            continue
    window._section_dictation_widgets = widgets
    if not widgets:
        return None

    active_helper = None
    processing_helper = None
    for widget in widgets:
        helper = getattr(widget, "_dictation_helper", None)
        if helper is None:
            continue
        if getattr(helper, "_is_processing", False):
            processing_helper = helper
            break
        if getattr(helper, "_is_recording", False):
            active_helper = helper

    if processing_helper is not None:
        return processing_helper
    if active_helper is not None:
        return active_helper

    focused = None
    try:
        focused = window.focus_get()
    except Exception:
        focused = None
    if isinstance(focused, tk.Text) and focused in widgets:
        return getattr(focused, "_dictation_helper", None)
    return getattr(widgets[0], "_dictation_helper", None)


def _refresh_section_dictation_button(window):
    button = getattr(window, "_section_dictation_button", None)
    title = getattr(window, "header_title", None)
    if not button or not title:
        return
    try:
        if not button.winfo_exists() or not title.winfo_exists():
            return
    except Exception:
        return

    helper = _resolve_section_dictation_helper(window)
    if helper is None:
        try:
            button.place_forget()
        except Exception:
            pass
        return

    try:
        title.update_idletasks()
        parent = title.master
        parent.update_idletasks()
        x = title.winfo_x() + title.winfo_width() + 12
        y = title.winfo_y() - 2
        button.place(x=x, y=max(0, y))
    except Exception:
        pass

    if getattr(helper, "_is_processing", False):
        button.configure(text="🎤 Procesando...", state="disabled")
    elif getattr(helper, "_is_recording", False):
        button.configure(text="🎤 Detener", state="normal")
    else:
        button.configure(text="🎤 Dictar", state="normal" if helper._can_dictate() else "disabled")

    try:
        window._section_dictation_after_id = window.after(
            250,
            lambda w=window: _refresh_section_dictation_button(w),
        )
    except Exception:
        window._section_dictation_after_id = None


def _install_section_dictation_button(window, text_widgets):
    title = getattr(window, "header_title", None)
    if not title or not title.winfo_exists():
        return
    parent = title.master
    button = tk.Button(
        parent,
        text="🎤 Dictar",
        font=("Segoe UI Emoji", 10, "bold"),
        cursor="hand2",
        padx=8,
        pady=2,
        command=lambda w=window: _on_section_dictation_click(w),
    )
    window._section_dictation_button = button
    window._section_dictation_widgets = list(text_widgets)
    if not getattr(window, "_section_dictation_bound", False):
        title.bind("<Configure>", lambda _e, w=window: _refresh_section_dictation_button(w), add="+")
        parent.bind("<Configure>", lambda _e, w=window: _refresh_section_dictation_button(w), add="+")
        window._section_dictation_bound = True
    _refresh_section_dictation_button(window)


def _on_section_dictation_click(window):
    helper = _resolve_section_dictation_helper(window)
    if helper is None:
        return
    try:
        helper.text.focus_set()
    except Exception:
        pass
    helper._on_toggle()
    _refresh_section_dictation_button(window)


_ASISTENTES_PROF_CACHE = {
    "loaded_at": 0.0,
    "nombres": [],
    "cargos": [],
    "name_to_cargo": {},
}
_ASESORES_AGENCIA_CACHE = {
    "loaded_at": 0.0,
    "nombres": [],
}
_ASISTENTES_PROF_CACHE_TTL = 300


def _asistentes_norm(value):
    text = _normalize_ascii_text(value)
    return re.sub(r"\s+", " ", text).strip().lower()


def _dedupe_keep_order(values):
    seen = set()
    result = []
    for raw in values or []:
        value = str(raw or "").strip()
        if not value:
            continue
        key = value.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


def _get_asistentes_profesionales_catalog(force=False):
    global _ASISTENTES_PROF_CACHE
    now = time.time()
    cached = _ASISTENTES_PROF_CACHE
    if (
        not force
        and cached.get("nombres")
        and (now - float(cached.get("loaded_at") or 0.0)) < _ASISTENTES_PROF_CACHE_TTL
    ):
        return cached

    try:
        rows = _supabase_get_paged(
            "profesionales",
            {
                "select": "nombre_profesional,cargo_profesional",
                "order": "nombre_profesional.asc",
            },
            env_path=".env",
            page_size=500,
            max_pages=20,
        )
    except Exception as exc:
        _log_capture(f"[ASISTENTES] no se pudo leer profesionales: {exc}")
        return cached

    nombres = []
    cargos = []
    name_to_cargo = {}
    for row in rows or []:
        nombre = str((row or {}).get("nombre_profesional") or "").strip()
        cargo = str((row or {}).get("cargo_profesional") or "").strip()
        if nombre:
            nombres.append(nombre)
            norm_name = _asistentes_norm(nombre)
            if norm_name and cargo and norm_name not in name_to_cargo:
                name_to_cargo[norm_name] = cargo
        if cargo:
            cargos.append(cargo)

    _ASISTENTES_PROF_CACHE = {
        "loaded_at": now,
        "nombres": _dedupe_keep_order(nombres),
        "cargos": _dedupe_keep_order(cargos),
        "name_to_cargo": name_to_cargo,
    }
    return _ASISTENTES_PROF_CACHE


def _get_asesores_agencia_catalog(force=False):
    global _ASESORES_AGENCIA_CACHE
    now = time.time()
    cached = _ASESORES_AGENCIA_CACHE
    if (
        not force
        and cached.get("nombres")
        and (now - float(cached.get("loaded_at") or 0.0)) < _ASISTENTES_PROF_CACHE_TTL
    ):
        return cached

    try:
        rows = _supabase_get_paged(
            "asesores",
            {
                "select": "nombre",
                "order": "nombre.asc",
            },
            env_path=".env",
            page_size=500,
            max_pages=20,
        )
    except Exception as exc:
        _log_capture(f"[ASESORES] no se pudo leer asesores: {exc}")
        return cached

    nombres = []
    for row in rows or []:
        nombre = str((row or {}).get("nombre") or "").strip()
        if nombre:
            nombres.append(nombre)

    _ASESORES_AGENCIA_CACHE = {
        "loaded_at": now,
        "nombres": _dedupe_keep_order(nombres),
    }
    return _ASESORES_AGENCIA_CACHE


def _create_asistente_inputs(parent, width, use_catalog=False, catalog=None):
    if use_catalog:
        nombre_widget = ttk.Combobox(parent, width=width, state="normal")
        cargo_widget = ttk.Combobox(parent, width=width, state="normal")
        _configure_asistente_widgets(nombre_widget, cargo_widget, catalog=catalog)
        return nombre_widget, cargo_widget
    return tk.Entry(parent, width=width), tk.Entry(parent, width=width)


def _create_asesor_agencia_inputs(parent, width, catalog=None, default_cargo="Asesor Agencia"):
    nombre_widget = ttk.Combobox(parent, width=width, state="normal")
    nombres = list((catalog or {}).get("nombres") or [])
    nombre_widget.configure(values=nombres)
    _bind_editable_combobox_filter(nombre_widget, nombres)

    cargo_widget = tk.Entry(parent, width=width)
    cargo_widget.insert(0, default_cargo)

    def _ensure_default_cargo(_event=None):
        if not cargo_widget.get().strip():
            cargo_widget.delete(0, tk.END)
            cargo_widget.insert(0, default_cargo)

    nombre_widget.bind("<<ComboboxSelected>>", _ensure_default_cargo, add="+")
    nombre_widget.bind("<FocusOut>", _ensure_default_cargo, add="+")
    return nombre_widget, cargo_widget


def _set_input_value(widget, value):
    text = str(value or "")
    if isinstance(widget, ttk.Combobox):
        widget.set(text)
        return
    widget.delete(0, tk.END)
    if text:
        widget.insert(0, text)


def _get_input_value(widget):
    try:
        return widget.get().strip()
    except Exception:
        return ""


def _bind_editable_combobox_filter(widget, values):
    options = list(values or [])
    if not options:
        return

    def _refresh(_event=None):
        typed = widget.get().strip()
        if not typed:
            widget.configure(values=options)
            return
        needle = _asistentes_norm(typed)
        filtered = [item for item in options if needle in _asistentes_norm(item)]
        widget.configure(values=filtered or options)

    widget.configure(values=options)
    widget.bind("<KeyRelease>", _refresh, add="+")
    widget.bind("<Button-1>", lambda _e: widget.configure(values=options), add="+")


def _configure_asistente_widgets(nombre_widget, cargo_widget, catalog=None):
    catalog = catalog or _get_asistentes_profesionales_catalog()
    nombres = list(catalog.get("nombres") or [])
    cargos = list(catalog.get("cargos") or [])
    name_to_cargo = dict(catalog.get("name_to_cargo") or {})

    nombre_widget.configure(values=nombres, state="normal")
    cargo_widget.configure(values=cargos, state="normal")
    _bind_editable_combobox_filter(nombre_widget, nombres)
    _bind_editable_combobox_filter(cargo_widget, cargos)

    def _sync_cargo(_event=None):
        selected_name = _asistentes_norm(nombre_widget.get())
        suggested = name_to_cargo.get(selected_name)
        if suggested:
            cargo_widget.set(suggested)

    nombre_widget.bind("<<ComboboxSelected>>", _sync_cargo, add="+")
    nombre_widget.bind("<FocusOut>", _sync_cargo, add="+")


class FormMousewheelMixin:
    def _bind_mousewheel(self, canvas, target):
        def _is_descendant(widget, ancestor):
            current = widget
            while current is not None:
                if current == ancestor:
                    return True
                try:
                    current = current.master
                except Exception:
                    return False
            return False

        def _is_wheel_blocked(widget):
            current = widget
            while current is not None:
                try:
                    if isinstance(current, ttk.Combobox):
                        return True
                    if isinstance(current, tk.Spinbox):
                        return True
                    if current.winfo_class() in {"TSpinbox", "Spinbox"}:
                        return True
                except Exception:
                    return False
                try:
                    current = current.master
                except Exception:
                    return False
            return False

        def _on_mousewheel(event):
            widget = getattr(event, "widget", None)
            if widget is None or not _is_descendant(widget, target):
                return None
            if _is_wheel_blocked(widget):
                return "break"

            if getattr(event, "delta", 0):
                delta = int(-1 * (event.delta / 120))
                if delta == 0:
                    delta = -1 if event.delta > 0 else 1
                canvas.yview_scroll(delta, "units")
            else:
                num = getattr(event, "num", None)
                if num == 4:
                    canvas.yview_scroll(-3, "units")
                elif num == 5:
                    canvas.yview_scroll(3, "units")
            return "break"

        try:
            target.bind_all("<MouseWheel>", _on_mousewheel, add="+")
            target.bind_all("<Button-4>", _on_mousewheel, add="+")
            target.bind_all("<Button-5>", _on_mousewheel, add="+")
        except Exception:
            pass


class LoadingDialog:
    def __init__(self, parent, title="Guardando"):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.configure(bg=COLOR_LIGHT_BG)
        self.window.geometry("420x160")
        self.window.transient(parent)
        self.window.grab_set()

        self.status_label = tk.Label(
            self.window,
            text="Iniciando...",
            bg=COLOR_LIGHT_BG,
            fg="#333333",
            font=FONT_LABEL,
        )
        self.status_label.pack(pady=(24, 8))

        self.progress = ttk.Progressbar(self.window, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=24, pady=(0, 16))
        self._center()
        self.window.update_idletasks()

    def _center(self):
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def exists(self):
        try:
            return bool(self.window and self.window.winfo_exists())
        except tk.TclError:
            return False

    def set_status(self, text):
        if not self.exists():
            return
        self.status_label.config(text=text)
        self.window.update_idletasks()

    def set_progress(self, value):
        if not self.exists():
            return
        self.progress["value"] = value
        self.window.update_idletasks()

    def close(self):
        if not self.exists():
            return
        try:
            self.window.destroy()
        except tk.TclError:
            return


class FinalizeProcessError(RuntimeError):
    def __init__(self, stage, cause):
        self.stage = str(stage or "").strip() or "finalizando el formulario"
        self.cause = cause
        message = str(cause).strip() if cause is not None else ""
        super().__init__(message or self.stage)


def _safe_widget_after(widget, callback):
    try:
        if widget and widget.winfo_exists():
            widget.after(0, callback)
    except tk.TclError:
        return


def _update_loading_async(loading, *, status=None, progress=None):
    if loading is None:
        return

    def _apply():
        if not loading.exists():
            return
        if status is not None:
            loading.set_status(status)
        if progress is not None:
            loading.set_progress(progress)

    _safe_widget_after(getattr(loading, "window", None), _apply)


def _close_loading_async(loading):
    if loading is None:
        return
    _safe_widget_after(getattr(loading, "window", None), loading.close)


def _build_finalize_error_message(form_name, stage, exc):
    label = str(form_name or "el formulario").strip()
    step = str(stage or "finalizando el formulario").strip()
    detail = str(exc).strip() or repr(exc)
    return f"No se pudo finalizar {label}.\n\nEtapa: {step}.\nDetalle: {detail}"


def _raise_finalize_stage(stage, func):
    try:
        return func()
    except FinalizeProcessError:
        raise
    except Exception as exc:
        raise FinalizeProcessError(stage, exc) from exc


def _clear_form_cache_safe(module):
    if hasattr(module, "clear_cache_file"):
        module.clear_cache_file()
    if hasattr(module, "clear_form_cache"):
        module.clear_form_cache()


def _should_clear_form_cache_after_delivery(completion_result):
    status = str((completion_result or {}).get("status") or "").strip().lower()
    return status in {"synced", "local"}


def _start_background_finalization(
    window,
    loading,
    *,
    form_name,
    company_name,
    form_id,
    worker_fn,
    post_delivery_fn=None,
):
    if getattr(window, "_finalize_in_progress", False):
        messagebox.showinfo(
            "Finalización",
            "Ya hay una finalización en curso para este formulario.",
            parent=window,
        )
        return

    window._finalize_in_progress = True

    def _finish_success(completion_result):
        try:
            _finalize_export_flow(
                window,
                loading,
                completion_result,
            )
            _return_to_hub(window)
            try:
                window.destroy()
            except tk.TclError:
                pass
        finally:
            try:
                window._finalize_in_progress = False
            except Exception:
                pass

    def _finish_error(exc):
        try:
            stage = exc.stage if isinstance(exc, FinalizeProcessError) else "finalizando el formulario"
            detail = exc.cause if isinstance(exc, FinalizeProcessError) else exc
            message = _build_finalize_error_message(form_name, stage, detail)
            loading.close()
            messagebox.showerror("Finalización", message, parent=window)
        finally:
            try:
                window._finalize_in_progress = False
            except Exception:
                pass

    def _worker():
        pythoncom = None
        com_initialized = False
        module = FORM_MODULE_MAP.get(form_id)
        original_cache_snapshot = {}
        review_result = None
        restore_original_cache = False
        completed_successfully = False
        try:
            try:
                import pythoncom as _pythoncom  # pyright: ignore[reportMissingImports]

                pythoncom = _pythoncom
                pythoncom.CoInitialize()
                com_initialized = True
            except ImportError:
                pythoncom = None

            if module and hasattr(module, "get_form_cache"):
                try:
                    original_cache_snapshot = copy.deepcopy(module.get_form_cache() or {})
                except Exception:
                    original_cache_snapshot = {}

            if module and original_cache_snapshot:
                _update_loading_async(
                    loading,
                    status="Revisando ortografía...",
                    progress=45,
                )
                review_result = text_review.review_export_cache(form_id, original_cache_snapshot)
                _log_capture(
                    "[OPENAI_REVIEW] "
                    f"form={form_id} status={review_result.status} "
                    f"reviewed_count={review_result.reviewed_count} "
                    f"elapsed_ms={review_result.elapsed_ms} "
                    f"reason={review_result.reason!r}"
                )
                if review_result.status == "reviewed":
                    _set_module_cache_snapshot(module, review_result.cache)
                    restore_original_cache = True
                    _update_loading_async(
                        loading,
                        status="Ortografía revisada. Generando Excel...",
                        progress=50,
                    )
                elif review_result.status in {"skipped", "failed"}:
                    _update_loading_async(
                        loading,
                        status="Revisión ortográfica omitida. Generando Excel...",
                        progress=50,
                    )

            export_result = worker_fn()
            drive_job = None
            if isinstance(export_result, dict):
                output_path = str(export_result.get("output_path") or "").strip()
                drive_job = export_result.get("drive_job")
            else:
                output_path = export_result
            if not output_path:
                raise FinalizeProcessError(
                    "generando el Excel",
                    RuntimeError("No se encontró el archivo de Excel generado."),
                )
            if not os.path.exists(output_path):
                raise FinalizeProcessError(
                    "verificando el archivo generado",
                    RuntimeError(f"No se encontró el archivo generado:\n{output_path}"),
                )
            hub = window.master if isinstance(window.master, HubWindow) else None
            if hub and form_id:
                hub.track_form_finished(form_id)
            if hub:
                _update_loading_async(
                    loading,
                    status="Publicando en Google Drive...",
                    progress=92,
                )
                completion_result = hub.finalize_form_delivery(
                    output_path,
                    form_name=form_name,
                    company_name=company_name,
                    drive_job=drive_job,
                )
            else:
                completion_result = {
                    "status": "local",
                    "output_path": output_path,
                    "remote_url": "",
                    "drive_file_id": "",
                    "error": "",
                }
            completed_successfully = True
            if post_delivery_fn is not None and _should_clear_form_cache_after_delivery(completion_result):
                try:
                    post_delivery_fn()
                except Exception:
                    pass
        except Exception as exc:
            _safe_widget_after(window, lambda exc=exc: _finish_error(exc))
        else:
            _safe_widget_after(window, lambda result=completion_result: _finish_success(result))
        finally:
            if restore_original_cache and module:
                try:
                    current_cache = module.get_form_cache() if hasattr(module, "get_form_cache") else {}
                except Exception:
                    current_cache = {}
                if not (completed_successfully and not current_cache):
                    try:
                        _set_module_cache_snapshot(module, original_cache_snapshot)
                    except Exception:
                        pass
            if com_initialized and pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    threading.Thread(target=_worker, daemon=True).start()


def get_forms():
    return [
        presentacion_programa.register_form(),
        evaluacion_accesibilidad.register_form(),
        condiciones_vacante.register_form(),
        seleccion_incluyente.register_form(),
        contratacion_incluyente.register_form(),
        induccion_organizacional.register_form(),
        induccion_operativa.register_form(),
        sensibilizacion.register_form(),
        seguimientos.register_form(),
    ]


class Section1Window(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Presentacion Programa - Seccion 1")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = presentacion_programa

        self.company_data = None
        self.fields = {}

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "presentacion_programa",
            presentacion_programa,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_3": self._show_section_3,
                "section_3_item_8": self._show_section_4,
                "section_4": self._show_section_4,
                "section_5": self._show_section_5,
            },
            self._show_section_1,
        ):
            return True
        if not presentacion_programa.cache_file_exists():
            return False
        resume = messagebox.askyesno(
            "Reanudar",
            "Se encontró un formulario en progreso. ¿Deseas continuar donde lo dejaste?",
        )
        if not resume:
            presentacion_programa.clear_cache_file()
            presentacion_programa.clear_form_cache()
            return False
        presentacion_programa.load_cache_from_file()
        last_section = presentacion_programa.get_form_cache().get("_last_section")
        if last_section == "section_1":
            self._show_section_2()
        elif last_section == "section_2":
            self._show_section_3()
        elif last_section in {"section_3", "section_3_item_8"}:
            self._show_section_4()
        elif last_section in {"section_4", "section_5"}:
            self._show_section_5()
        else:
            self._show_section_1()
        return True

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))

        self.header_title = tk.Label(
            header,
            text="1. DATOS GENERALES",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        self.header_title.pack(anchor="w")

        self.header_subtitle = tk.Label(
            header,
            text="Busca empresa por NIT y confirma datos.",
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        self.header_subtitle.pack(anchor="w", pady=(4, 0))

    def _build_section_container(self):
        self.section_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)

    def _clear_section_container(self):
        for child in self.section_container.winfo_children():
            child.destroy()

    def _show_section_1(self):
        self._clear_section_container()
        self.header_title.config(text="1. DATOS GENERALES")
        self.header_subtitle.config(text="Busca empresa por NIT y confirma datos.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._build_actions(content)
    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. TEMARIO")
        self.header_subtitle.config(
            text="Por favor, explique en su totalidad el temario a ser cubierto en la reuni\u00f3n."
        )
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        header = tk.Frame(content, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", pady=(8, 12))
        tk.Label(
            header,
            text="#",
            font=FONT_LABEL,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
            width=4,
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="Tema",
            font=FONT_LABEL,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=1, sticky="w")

        for idx, item in enumerate(presentacion_programa.SECTION_2["items"], start=1):
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=str(idx),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                width=4,
            ).grid(row=0, column=0, sticky="nw", padx=8, pady=8)

            tk.Label(
                row,
                text=item,
                font=("Arial", 10),
                bg="white",
                fg="#333333",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=1, sticky="w", padx=8, pady=8)

        _build_wizard_actions(
            content,
            back_command=self._show_section_1,
            primary_command=self._show_section_3,
        )
    def _show_section_3(self):
        self._clear_section_container()
        self.header_title.config(text="3. DESCRIPCI\u00d3N DE LOS TEMAS")
        self.header_subtitle.config(text="Describe los temas tratados.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        header = tk.Frame(content, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", pady=(8, 12))
        tk.Label(
            header,
            text="#",
            font=FONT_LABEL,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
            width=4,
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="Tema",
            font=FONT_LABEL,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=1, sticky="w")

        self.section3_check_vars = {}
        for item in presentacion_programa.SECTION_3["items"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=str(item["id"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                width=4,
            ).grid(row=0, column=0, sticky="nw", padx=8, pady=8)

            body = tk.Frame(row, bg="white")
            body.grid(row=0, column=1, sticky="ew", padx=8, pady=8)
            body.grid_columnconfigure(0, weight=1)

            tk.Label(
                body,
                text=item["title"],
                font=FONT_SECTION,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, sticky="w")

            if item.get("type") == "checkboxes":
                checks = tk.Frame(body, bg="white")
                checks.grid(row=1, column=0, sticky="w", pady=(6, 0))
                for label, default_value in item["content"].items():
                    var = tk.BooleanVar(value=default_value)
                    self.section3_check_vars[label] = var
                    tk.Checkbutton(
                        checks,
                        text=label,
                        variable=var,
                        bg="white",
                        anchor="w",
                        justify="left",
                        wraplength=720,
                    ).pack(anchor="w")
            else:
                tk.Label(
                    body,
                    text=item["content"],
                    font=("Arial", 10),
                    bg="white",
                    fg="#333333",
                    wraplength=760,
                    justify="left",
                ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        cached_checks = presentacion_programa.get_form_cache().get("section_3_item_8", {})
        for label, var in self.section3_check_vars.items():
            if label in cached_checks:
                var.set(bool(cached_checks.get(label)))

        _build_wizard_actions(
            content,
            back_command=self._show_section_2,
            primary_command=self._confirm_section_3,
        )
    def _confirm_section_3(self):
        if not self.section3_check_vars:
            self._show_section_4()
            return
        values = {key: var.get() for key, var in self.section3_check_vars.items()}
        try:
            presentacion_programa.confirm_section_3_item8(values)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_4()

    def _show_section_4(self):
        self._clear_section_container()
        self.header_title.config(text="4. ACUERDOS Y OBSERVACIONES DE LA REUNI\u00d3N")
        self.header_subtitle.config(text="Registra acuerdos y observaciones.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)
        form_container = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        form_container.pack(fill="both", expand=True, pady=(8, 12))

        tk.Label(
            form_container,
            text="Acuerdos y observaciones de la reuni\u00f3n",
            font=FONT_LABEL,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w")

        template_actions = tk.Frame(form_container, bg=COLOR_LIGHT_BG)
        template_actions.pack(fill="x", pady=(8, 4))
        for idx, (template_key, label) in enumerate(presentacion_programa.RUTA_INCLUSION_TEMPLATE_BUTTONS):
            btn = ttk.Button(
                template_actions,
                text=label,
                command=lambda key=template_key: self._insert_section4_template(key),
            )
            btn.grid(row=idx // 3, column=idx % 3, sticky="w", padx=(0, 8), pady=(0, 8))

        self.section4_text = tk.Text(
            form_container,
            height=10,
            wrap="word",
        )
        self.section4_text.pack(fill="x", pady=(6, 16))

        cached_notes = presentacion_programa.get_form_cache().get("section_4", {}).get(
            "acuerdos_observaciones"
        )
        if cached_notes:
            self.section4_text.delete("1.0", tk.END)
            self.section4_text.insert("1.0", cached_notes)

        _build_wizard_actions(
            section_frame,
            back_command=self._show_section_3,
            primary_command=self._confirm_section_4,
        )

    def _show_section_5(self):
        self._clear_section_container()
        self.header_title.config(text="5. ASISTENTES")
        self.header_subtitle.config(text="Registra asistentes y agrega filas si aplica.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        form_container = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        form_container.pack(fill="both", expand=True, pady=(8, 12))

        asistentes_frame = tk.Frame(form_container, bg=COLOR_LIGHT_BG)
        asistentes_frame.pack(fill="x")

        tk.Label(
            asistentes_frame,
            text="Asistentes",
            font=FONT_LABEL,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        tk.Label(
            asistentes_frame,
            text="Nombre completo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=1, column=0, sticky="w", padx=(0, 8))
        tk.Label(
            asistentes_frame,
            text="Cargo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=1, column=1, sticky="w", padx=(0, 8))

        self.add_asistente_btn = ttk.Button(
            asistentes_frame,
            text="Agregar asistente",
            command=self._add_asistente_row,
        )

        self._asesores_agencia_catalog = _get_asesores_agencia_catalog()
        self.section5_entries = []
        self.section5_frame = asistentes_frame
        cached_asistentes = presentacion_programa.get_form_cache().get("section_5", [])
        self._render_section5_asistentes(cached_asistentes)

        _build_wizard_actions(
            section_frame,
            back_command=self._show_section_4,
            primary_command=self._confirm_section_5,
            primary_text="Finalizar",
        )

    def _insert_section4_template(self, template_key):
        template_text = presentacion_programa.RUTA_INCLUSION_TEMPLATES.get(template_key, "").strip()
        if not template_text:
            return
        current_text = self.section4_text.get("1.0", tk.END).strip()
        if current_text:
            self.section4_text.insert(tk.END, "\n\n")
        self.section4_text.insert(tk.END, template_text)
        self.section4_text.focus_set()
        self.section4_text.see(tk.END)

    def _confirm_section_4(self):
        notes = self.section4_text.get("1.0", tk.END).strip()
        try:
            presentacion_programa.confirm_section_4(notes)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_5()

    def _confirm_section_5(self):
        asistentes = []
        for nombre_entry, cargo_entry in self.section5_entries:
            asistentes.append(
                {
                    "nombre": nombre_entry.get().strip(),
                    "cargo": cargo_entry.get().strip(),
                }
            )
        try:
            presentacion_programa.confirm_section_5(asistentes)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Guardando Excel...")
        loading.set_progress(30)
        cache_snapshot = presentacion_programa.get_form_cache()
        cache = cache_snapshot
        section_1 = cache.get("section_1", {})
        visit_type = (section_1.get("tipo_visita") or "Presentacion").strip()
        form_name = (
            "Reactivacion Programa" if visit_type.lower() == "reactivacion" else "Presentacion Programa"
        )
        company_name = section_1.get("nombre_empresa")
        _start_background_finalization(
            self,
            loading,
            form_name=form_name,
            company_name=company_name,
            form_id="presentacion_programa",
            worker_fn=lambda: _raise_finalize_stage(
                "guardando el Excel",
                presentacion_programa.export_to_excel,
            ),
        )

    def _add_asistente_row(self):
        max_items = presentacion_programa.SECTION_5.get("max_items", 10)
        if len(self.section5_entries) >= max_items:
            messagebox.showinfo("Asistentes", f"Máximo {max_items} asistentes.")
            return
        rows = self._get_section5_asistentes_values()
        if rows:
            # Insert new empty row before the last (asesor agencia) row,
            # preserving the asesor agencia's entered values.
            rows.insert(len(rows) - 1, {"nombre": "", "cargo": ""})
        else:
            rows = [{"nombre": "", "cargo": ""}]
        self._render_section5_asistentes(rows)

    def _get_section5_asistentes_values(self):
        values = []
        for nombre_entry, cargo_entry in self.section5_entries:
            values.append(
                {
                    "nombre": _get_input_value(nombre_entry),
                    "cargo": _get_input_value(cargo_entry),
                }
            )
        return values

    def _render_section5_asistentes(self, values=None):
        rows = list(values or [])
        while len(rows) < 3:
            rows.append({"nombre": "", "cargo": ""})

        for nombre_entry, cargo_entry in self.section5_entries:
            nombre_entry.destroy()
            cargo_entry.destroy()
        self.section5_entries = []

        for idx, entry in enumerate(rows):
            row = 2 + idx
            is_first = idx == 0
            is_last = idx == len(rows) - 1
            if is_first and not is_last:
                nombre_entry, cargo_entry = _create_asistente_inputs(
                    self.section5_frame,
                    40,
                    use_catalog=True,
                    catalog=_get_asistentes_profesionales_catalog(),
                )
            elif is_last:
                nombre_entry, cargo_entry = _create_asesor_agencia_inputs(
                    self.section5_frame,
                    40,
                    catalog=getattr(self, "_asesores_agencia_catalog", None),
                )
            else:
                nombre_entry, cargo_entry = _create_asistente_inputs(
                    self.section5_frame,
                    40,
                    use_catalog=False,
                )
            nombre_entry.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
            cargo_entry.grid(row=row, column=1, sticky="w", padx=(0, 8), pady=4)
            _set_input_value(nombre_entry, entry.get("nombre", ""))
            cargo_value = entry.get("cargo", "")
            if is_last and not cargo_value:
                cargo_value = "Asesor Agencia"
            _set_input_value(cargo_entry, cargo_value)
            self.section5_entries.append((nombre_entry, cargo_entry))

        self.add_asistente_btn.grid(
            row=2 + len(self.section5_entries) + 1,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(8, 0),
        )
    def _build_search(self, parent):
        _section1_build_search(self, parent, include_tipo_visita=True)

    def _build_groups(self, parent):
        container = tk.Frame(parent, bg=COLOR_LIGHT_BG)
        container.pack(fill="both", expand=True)
        readonly_w = ENTRY_W_XL
        try:
            sw = int(self.winfo_screenwidth() or 0)
            if sw and sw <= 1366:
                readonly_w = 42
            elif sw and sw <= 1600:
                readonly_w = 50
        except Exception:
            pass

        groups = [
            ("Información de Empresa", COLOR_GROUP_EMPRESA, [
                "nombre_empresa",
                "direccion_empresa",
                "correo_1",
                "contacto_empresa",
                "telefono_empresa",
                "cargo",
                "ciudad_empresa",
                "sede_empresa",
                "caja_compensacion",
            ]),
            ("Información de Compensar", COLOR_GROUP_COMPENSAR, [
                "asesor",
                "correo_asesor",
            ]),
            ("Información de RECA", COLOR_GROUP_RECA, [
                "profesional_asignado",
                "correo_profesional",
            ]),
        ]

        top_inputs = tk.Frame(container, bg=COLOR_LIGHT_BG)
        top_inputs.pack(fill="x", pady=(0, 12))

        tk.Label(
            top_inputs,
            text="Fecha de la visita",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.fields["fecha_visita"] = DateEntry(
            top_inputs,
            width=ENTRY_W_MED,
            date_pattern="yyyy-mm-dd",
        )
        self.fields["fecha_visita"].delete(0, tk.END)
        self.fields["fecha_visita"].grid(row=0, column=1, sticky="w", padx=(0, 24))

        tk.Label(
            top_inputs,
            text="Modalidad",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.fields["modalidad"] = ttk.Combobox(
            top_inputs,
            values=["Virtual", "Presencial", "Mixto", "No aplica"],
            state="readonly",
            width=ENTRY_W_MED,
        )
        self.fields["modalidad"].grid(row=0, column=3, sticky="w")

        for title, color, field_ids in groups:
            group_label = tk.Label(
                container,
                text=title,
                bg=color,
                fg=COLOR_PURPLE,
                font=FONT_LABEL,
            )
            group_frame = tk.LabelFrame(
                container,
                labelwidget=group_label,
                bg=color,
                padx=12,
                pady=8,
                bd=1,
            )
            group_frame.pack(fill="x", pady=8)
            group_frame.grid_columnconfigure(1, weight=1)

            for row, field_id in enumerate(field_ids):
                label_text = self._label_for_field(field_id)
                tk.Label(
                    group_frame,
                    text=label_text,
                    font=FONT_LABEL,
                    bg=color,
                ).grid(row=row, column=0, sticky="w", padx=6, pady=4)

                entry = tk.Entry(group_frame, state="readonly", width=readonly_w)
                entry.grid(row=row, column=1, sticky="w", padx=6, pady=4)
                self.fields[field_id] = entry

    def _build_actions(self, parent):
        _section1_build_actions(self, parent)

    def _label_for_field(self, field_id):
        labels = {
            "nombre_empresa": "Nombre de la empresa",
            "direccion_empresa": "Dirección de la empresa",
            "correo_1": "Correo electrónico",
            "contacto_empresa": "Contacto de la empresa",
            "telefono_empresa": "Teléfonos responsable empresa",
            "cargo": "Cargo responsable empresa",
            "ciudad_empresa": "Ciudad/Municipio",
            "sede_empresa": "Sede Compensar",
            "caja_compensacion": "Empresa afiliada a Caja de Compensación",
            "asesor": "Asesor fidelización",
            "correo_asesor": "Correo asesor",
            "profesional_asignado": "Profesional asignado RECA",
            "correo_profesional": "Correo profesional RECA",
        }
        return labels.get(field_id, field_id)

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _search_company(self, mode="nit"):
        nit = self.fields["nit_empresa"].get().strip()
        nombre = self.fields.get("nombre_busqueda").get().strip() if self.fields.get("nombre_busqueda") else ""
        if mode == "nit":
            if not nit:
                messagebox.showerror("Error", "Ingresa un NIT.")
                return
        elif mode == "nombre":
            if not nombre:
                messagebox.showerror("Error", "Ingresa el nombre de la empresa.")
                return
        else:
            messagebox.showerror("Error", "Tipo de búsqueda no válido.")
            return

        lookup = getattr(self, "_empresa_lookup", presentacion_programa)
        try:
            self.status_label.config(text="Buscando empresa...")
            self.update_idletasks()
            if mode == "nombre":
                company = lookup.get_empresa_by_nombre(nombre)
            else:
                company = lookup.get_empresa_by_nit(nit)
        except Exception as exc:
            self.status_label.config(text="")
            messagebox.showerror("Error", str(exc))
            return

        section_map = getattr(lookup, "SECTION_1_SUPABASE_MAP", presentacion_programa.SECTION_1_SUPABASE_MAP)
        if not company:
            self.company_data = None
            msg = "No se encontró empresa para ese nombre." if mode == "nombre" else "No se encontró empresa para ese NIT."
            self.status_label.config(text=msg)
            self.continue_btn.config(state="disabled")
            for key in section_map.keys():
                self._set_readonly_value(key, "")
            return

        if mode == "nombre":
            nit_value = company.get("nit_empresa")
            if nit_value:
                entry = self.fields.get("nit_empresa")
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, nit_value)

        self.company_data = company
        self.status_label.config(text="Empresa encontrada.")
        self.continue_btn.config(state="normal")
        for key in section_map.keys():
            self._set_readonly_value(key, company.get(key))

    def _confirm_and_continue(self):
        if not self.company_data:
            messagebox.showerror("Error", "Busca una empresa antes de confirmar.")
            return

        fecha_visita = _get_required_fecha_visita(self)
        if not fecha_visita:
            return
        modalidad = _get_required_modalidad(self)
        if not modalidad:
            return
        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad": modalidad,
            "nit_empresa": self.fields["nit_empresa"].get().strip(),
            "tipo_visita": self.fields["tipo_visita"].get().strip(),
        }
        try:
            presentacion_programa.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2()


def _section1_update_nombre_suggestions(self):
    entry = self.fields.get("nombre_busqueda")
    if not entry:
        return
    prefix = entry.get().strip()
    if len(prefix) < 2:
        entry["values"] = []
        return
    lookup = getattr(self, "_empresa_lookup", None)
    if not lookup or not hasattr(lookup, "get_empresas_by_nombre_prefix"):
        return
    try:
        suggestions = lookup.get_empresas_by_nombre_prefix(prefix)
    except Exception:
        suggestions = []
    entry["values"] = suggestions


class HubWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        _log_capture("==== Inicio de app ====")
        _log_capture(f"Python={sys.version.split()[0]} | platform={sys.platform} | cwd={os.getcwd()}")
        _log_capture(f"log_root={get_logs_root()}")
        _log_capture(f"log_path={_desktop_log_path()}")
        _run_encoding_health_check()
        try:
            removed = cleanup_stale_audio(ttl_hours=24)
            _log_capture(f"[DICTATION] cleanup_stale_audio removed={removed}")
        except Exception as exc:
            _log_capture(f"[DICTATION] cleanup_stale_audio_failed err={exc}")
        self.title(APP_NAME)
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("900x600")
        _maximize_window(self)

        self.current_user = None
        self.current_user_profile = {}
        self.login_frame = None
        self.header = None
        self.body = None
        self._toast_label = None
        self._toast_after_id = None
        self._session_info_label = None
        self._session_clock_after_id = None
        self.current_session_id = None
        self._form_event_ids = {}
        self._form_event_payloads = {}
        self._companies_all = []
        self._companies_by_id = {}
        self._companies_tree = None
        self._companies_search_var = None
        self._companies_sort_var = None
        self._version_var = tk.StringVar(value="Versión local: - | GitHub: -")
        self._version_check_thread = None
        self._drafts_btn = None
        self._refresh_db_btn = None
        self._sync_panel_btn = None
        self._net_status_label = None
        self._net_status_after_id = None
        self._is_online = False
        self._net_check_thread = None
        self._service_probe_cache = {
            "internet": {"ok": False, "status_text": "Sin verificar", "error_code": "", "detail": ""},
            "supabase": {"ok": False, "status_text": "Sin verificar", "error_code": "", "detail": ""},
            "drive": {"ok": False, "status_text": "Sin verificar", "error_code": "", "detail": ""},
        }
        self._startup_precheck_rows = {}
        self._startup_precheck_thread = None
        self._startup_precheck_watchdog_id = None
        self._startup_precheck_completed = False
        self._service_probe_cache_time = 0.0
        self._login_built = False
        self._login_btn = None

        _ensure_drive_upload_worker()
        self._configure_input_styles()
        self.protocol("WM_DELETE_WINDOW", self._on_app_close)
        self.after(0, self._build_login)

    def _configure_input_styles(self):
        self.option_add("*Entry.background", "white")
        self.option_add("*Entry.readonlyBackground", "#EDEDED")
        self.option_add("*Text.background", "white")
        for widget_class in ("TCombobox", "Spinbox", "TSpinbox"):
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                self.bind_class(widget_class, seq, lambda _e: "break")
        style = ttk.Style(self)
        style.configure("TEntry", fieldbackground="white")
        style.configure("TCombobox", fieldbackground="white", background="white")
        style.map("TCombobox", fieldbackground=[("readonly", "white")])

    def _build_startup_precheck_toast(self, parent):
        card = tk.Frame(parent, bg="white", bd=1, relief="solid", padx=8, pady=6)
        card.pack(anchor="w", fill="x", pady=(0, 12))

        header = tk.Frame(card, bg="white")
        header.pack(fill="x")
        tk.Label(
            header,
            text="Servicios",
            font=("Arial", 9, "bold"),
            fg=COLOR_PURPLE,
            bg="white",
        ).pack(side="left")
        ttk.Button(
            header,
            text="Revisar de nuevo",
            command=self._run_startup_precheck_async,
        ).pack(side="right")

        self._startup_precheck_status_var = tk.StringVar(value="Verificando servicios")
        self._startup_precheck_status_label = tk.Label(
            card,
            textvariable=self._startup_precheck_status_var,
            font=("Arial", 9),
            fg="#888888",
            bg="white",
            anchor="w",
        )
        self._startup_precheck_status_label.pack(anchor="w", pady=(4, 0))

    def _set_login_ready_state(self, ready, status_text=None):
        self._startup_precheck_completed = bool(ready)
        if self._login_btn is not None:
            try:
                self._login_btn.config(state=("normal" if ready else "disabled"))
            except tk.TclError:
                pass
        if self.login_status is not None:
            try:
                if status_text is not None:
                    self.login_status.config(text=status_text)
            except tk.TclError:
                pass

    def _apply_startup_precheck_results(self, result):
        self._service_probe_cache = dict(result or {})
        self._service_probe_cache_time = time.time()
        failed_service = ""
        for service_key in ("internet", "supabase", "drive"):
            payload = (result or {}).get(service_key) or {}
            if not bool(payload.get("ok")):
                failed_service = service_key.capitalize()
                break
        try:
            if failed_service:
                self._startup_precheck_status_var.set(failed_service)
                self._startup_precheck_status_label.config(fg="#B00020")
            else:
                self._startup_precheck_status_var.set("Todo correcto")
                self._startup_precheck_status_label.config(fg="#0A7D2E")
        except tk.TclError:
            pass
        self._set_login_ready_state(True, "")

    def _run_startup_precheck_async(self, log_enabled=False):
        if self._startup_precheck_thread and self._startup_precheck_thread.is_alive():
            return
        self._set_login_ready_state(False, "Verificando servicios...")
        try:
            self._startup_precheck_status_var.set("Verificando servicios")
            self._startup_precheck_status_label.config(fg="#888888")
        except tk.TclError:
            pass

        result = {}
        def _cancel_watchdog():
            if self._startup_precheck_watchdog_id is not None:
                try:
                    self.after_cancel(self._startup_precheck_watchdog_id)
                except tk.TclError:
                    pass
                self._startup_precheck_watchdog_id = None

        def _worker():
            result.update(probe_startup_services(log_enabled=log_enabled))

        self._startup_precheck_thread = threading.Thread(target=_worker, daemon=True)
        self._startup_precheck_thread.start()

        _cancel_watchdog()
        self._startup_precheck_watchdog_id = self.after(30000, self._startup_precheck_timeout)

        def _finish():
            if self._startup_precheck_thread and self._startup_precheck_thread.is_alive():
                self.after(120, _finish)
                return
            _cancel_watchdog()
            self._startup_precheck_thread = None
            self._apply_startup_precheck_results(result)

        self.after(120, _finish)

    def _startup_precheck_timeout(self):
        if self._startup_precheck_watchdog_id is not None:
            try:
                self.after_cancel(self._startup_precheck_watchdog_id)
            except tk.TclError:
                pass
            self._startup_precheck_watchdog_id = None
        try:
            self._startup_precheck_status_var.set("Timeout")
            self._startup_precheck_status_label.config(fg="#B00020")
        except tk.TclError:
            pass
        self._set_login_ready_state(True, "")

    def _build_login(self):
        if self._login_built:
            return
        self._login_built = True
        self.login_frame = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.login_frame.place(relx=0.5, rely=0.5, anchor="center")
        saved_login = _load_saved_login_credentials()

        title = tk.Label(
            self.login_frame,
            text="Iniciar sesión",
            font=("Arial", 20, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(0, 12))

        self._build_startup_precheck_toast(self.login_frame)

        form = tk.Frame(self.login_frame, bg=COLOR_LIGHT_BG)
        form.pack(anchor="w")

        tk.Label(
            form,
            text="Usuario",
            font=FONT_SECTION,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=(0, 8))
        self.login_user_entry = tk.Entry(form, width=30)
        self.login_user_entry.grid(row=0, column=1, sticky="w", pady=(0, 8))

        tk.Label(
            form,
            text="Contraseña",
            font=FONT_SECTION,
            bg=COLOR_LIGHT_BG,
        ).grid(row=1, column=0, sticky="w", padx=(0, 12), pady=(0, 8))
        self.login_pass_entry = tk.Entry(form, width=30, show="*")
        self.login_pass_entry.grid(row=1, column=1, sticky="w", pady=(0, 8))

        self.remember_login_var = tk.BooleanVar(value=bool(saved_login.get("remember", True)))
        remember_cb = tk.Checkbutton(
            form,
            text="Recordarme",
            variable=self.remember_login_var,
            bg=COLOR_LIGHT_BG,
            activebackground=COLOR_LIGHT_BG,
        )
        remember_cb.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 6))

        if saved_login.get("username"):
            self.login_user_entry.insert(0, str(saved_login.get("username")))
        if saved_login.get("password"):
            self.login_pass_entry.insert(0, str(saved_login.get("password")))

        self.login_status = tk.Label(
            self.login_frame,
            text="Verificando servicios...",
            font=("Arial", 10),
            fg="#555555",
            bg=COLOR_LIGHT_BG,
        )
        self.login_status.pack(anchor="w", pady=(2, 12))

        self._login_btn = ttk.Button(
            self.login_frame,
            text="Ingresar",
            command=self._handle_login,
            state="disabled",
        )
        self._login_btn.pack(anchor="w")

        forgot_btn = ttk.Button(
            self.login_frame,
            text="Olvide mi contraseña",
            command=self._show_forgot_password_info,
        )
        forgot_btn.pack(anchor="w", pady=(8, 0))
        self._run_startup_precheck_async(log_enabled=True)

    def _show_forgot_password_info(self):
        messagebox.showinfo(
            "Recuperacion de contraseña",
            "Para recuperar tu contraseña, comunicate con Aaron Uyaban\n"
            "Correo: admonusaid@recacolombia.org",
        )

    def _handle_login(self):
        username_input = self.login_user_entry.get().strip()
        username = _normalize_login_value(username_input)
        password = self.login_pass_entry.get()
        if not username or not password:
            messagebox.showerror("Error", "Ingresa usuario y contraseña.")
            return
        used_offline = False
        auth_exc = None
        try:
            self.login_status.config(text="Validando credenciales...")
            self.update_idletasks()
            user_row = self._authenticate_user(username, password)
        except Exception as exc:
            auth_exc = exc
            user_row = None
            if not _is_connectivity_exception(exc):
                self.login_status.config(text="")
                messagebox.showerror("Error", str(exc))
                return
        if not user_row:
            can_use_offline = bool(auth_exc)
            if not can_use_offline:
                try:
                    can_use_offline = not _supabase_ping(timeout=3)
                except Exception:
                    can_use_offline = True
            if can_use_offline:
                user_row = self._authenticate_user_offline(username, password)
                if user_row:
                    used_offline = True
                    self.login_status.config(text="Modo offline: sesión local")
        if not user_row:
            self.login_status.config(text="")
            if auth_exc:
                messagebox.showerror("Error", str(auth_exc))
                return
            messagebox.showerror("Error", "Usuario y contraseña incorrectos.")
            return
        if not used_offline and self._must_force_password_change(user_row, password):
            changed = self._prompt_force_password_change(user_row, password)
            if not changed:
                self.login_status.config(text="")
                messagebox.showwarning(
                    "Cambio requerido",
                    "Debes cambiar la contraseña para continuar.",
                )
                return
            # Reload profile to keep local state aligned.
            try:
                refreshed = _supabase_rpc("get_my_profesional_profile", {})
                if isinstance(refreshed, dict):
                    user_row = refreshed
            except Exception:
                pass
        self._cache_offline_user_auth(user_row, password)
        remember_enabled = True
        try:
            remember_enabled = bool(self.remember_login_var.get())
        except Exception:
            pass
        if remember_enabled:
            _save_login_credentials(username_input or username, password)
        else:
            _clear_login_credentials()
        self.current_user = (user_row.get("usuario_login") or username).strip()
        self.current_user_profile = user_row
        try:
            self._normalize_profesional_asignado()
        except Exception:
            pass
        self._start_usage_session()
        if self.login_frame:
            self.login_frame.destroy()
            self.login_frame = None
        self._build_header()
        self._build_body()

    def _authenticate_user(self, username, password):
        username_norm = _normalize_login_value(username)
        if not username_norm:
            return None
        _clear_supabase_session()
        try:
            resolved = _supabase_rpc(
                "resolve_login_email",
                {"p_login": username_norm},
                use_session=False,
            )
        except Exception as exc:
            raise RuntimeError(str(exc)) from exc

        if isinstance(resolved, str):
            email = resolved.strip()
        elif isinstance(resolved, dict):
            email = str(
                resolved.get("resolve_login_email")
                or resolved.get("email")
                or ""
            ).strip()
        else:
            email = ""
        if not email:
            return None
        _supabase_auth_password_login(email, password)
        profile = _supabase_rpc("get_my_profesional_profile", {})
        if isinstance(profile, dict) and profile.get("id"):
            profile["_auth_source"] = "jwt"
            return profile
        return None

    def _authenticate_user_offline(self, username, password):
        username_norm = _normalize_login_value(username)
        if not username_norm:
            return None
        store = _load_offline_auth_store()
        users = store.get("users", {})
        if not isinstance(users, dict):
            return None
        cached = users.get(username_norm)
        if not isinstance(cached, dict):
            return None
        pass_hash = (cached.get("usuario_pass_hash") or "").strip()
        if not pass_hash:
            return None
        for candidate in _password_candidates(password):
            if _verify_password_hash(candidate, pass_hash):
                return {
                    "id": cached.get("id"),
                    "usuario_login": cached.get("usuario_login") or username_norm,
                    "usuario_pass_hash": pass_hash,
                    "nombre_profesional": cached.get("nombre_profesional") or "",
                    "programa": cached.get("programa") or "",
                    "_auth_source": "offline",
                }
        return None

    def _cache_offline_user_auth(self, user_row, password):
        if not isinstance(user_row, dict):
            return
        login = _normalize_login_value(user_row.get("usuario_login") or "")
        if not login:
            return
        pass_hash = (user_row.get("usuario_pass_hash") or "").strip()
        if not pass_hash:
            # Compatibilidad con cuentas heredadas si el login fue exitoso.
            pass_hash = _hash_password(str(password or "").strip())
        if not pass_hash:
            return
        store = _load_offline_auth_store()
        users = store.get("users")
        if not isinstance(users, dict):
            users = {}
            store["users"] = users
        users[login] = {
            "id": user_row.get("id"),
            "usuario_login": user_row.get("usuario_login") or login,
            "usuario_pass_hash": pass_hash,
            "nombre_profesional": user_row.get("nombre_profesional") or "",
            "programa": user_row.get("programa") or "",
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        _save_offline_auth_store(store)

    def _must_force_password_change(self, user_row, current_password):
        _ = current_password
        return bool(user_row.get("auth_password_temp"))

    def _validate_new_password(self, new_password, current_password):
        pwd = str(new_password or "")
        if len(pwd) < 8:
            return False, "La nueva contraseña debe tener mínimo 8 caracteres."
        if pwd == str(current_password or ""):
            return False, "La nueva contraseña no puede ser igual a la actual."
        if pwd.isdigit():
            return False, "La contraseña no puede ser solo números."
        if not re.search(r"[A-Za-z]", pwd) or not re.search(r"\d", pwd):
            return False, "La contraseña debe incluir letras y números."
        return True, ""

    def _prompt_force_password_change(self, user_row, current_password):
        dialog = tk.Toplevel(self)
        dialog.title("Cambio obligatorio de contraseña")
        dialog.configure(bg=COLOR_LIGHT_BG)
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        dialog.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = tk.Frame(dialog, bg=COLOR_LIGHT_BG, padx=18, pady=14)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text="Debes cambiar tu contraseña para continuar.",
            font=("Arial", 11, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        tk.Label(frame, text="Nueva contraseña", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 6)
        )
        new_entry = tk.Entry(frame, width=30, show="*")
        new_entry.grid(row=1, column=1, sticky="w", pady=(0, 6))

        tk.Label(frame, text="Confirmar contraseña", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=(0, 6)
        )
        confirm_entry = tk.Entry(frame, width=30, show="*")
        confirm_entry.grid(row=2, column=1, sticky="w", pady=(0, 6))

        status = tk.Label(frame, text="", fg="#B00020", bg=COLOR_LIGHT_BG, font=("Arial", 9))
        status.grid(row=3, column=0, columnspan=2, sticky="w", pady=(2, 8))

        result = {"ok": False}

        def _save():
            new_pwd = new_entry.get()
            confirm_pwd = confirm_entry.get()
            is_valid, msg = self._validate_new_password(new_pwd, current_password)
            if not is_valid:
                status.config(text=msg)
                return
            if new_pwd != confirm_pwd:
                status.config(text="La confirmación no coincide.")
                return
            try:
                _supabase_auth_update_password(new_pwd)
                patch_result = _supabase_patch_with_queue(
                    "profesionales",
                    {"id": user_row.get("id")},
                    {
                        "usuario_pass_hash": _hash_password(new_pwd),
                        "usuario_pass": None,
                        "auth_password_temp": False,
                    },
                )
                if (patch_result or {}).get("status") not in {"synced", "queued"}:
                    raise RuntimeError("No se pudo guardar el estado de contraseña.")
            except Exception as exc:
                status.config(text=f"No se pudo actualizar contraseña: {exc}")
                return
            result["ok"] = True
            dialog.destroy()

        actions = tk.Frame(frame, bg=COLOR_LIGHT_BG)
        actions.grid(row=4, column=0, columnspan=2, sticky="e")
        ttk.Button(actions, text="Guardar", command=_save).pack(side="right")

        new_entry.focus_set()
        dialog.wait_window()
        return result["ok"]

    def _usage_upsert_async(self, table, row, on_conflict):
        try:
            _supabase_enqueue_upsert(table, [row], on_conflict=on_conflict)
        except Exception:
            return

    def _usage_upsert_sync(self, table, row, on_conflict):
        try:
            result = _supabase_upsert_with_queue(
                table,
                [row],
                on_conflict=on_conflict,
            )
            return (result or {}).get("status") in {"synced", "queued"}
        except Exception:
            return False

    def _should_track_usage(self):
        login = _normalize_login_value(
            self.current_user_profile.get("usuario_login") or self.current_user or ""
        )
        if not login:
            return True
        return login not in _get_usage_exempt_logins()

    def _start_usage_session(self):
        if not self._should_track_usage():
            return
        if self.current_session_id:
            return
        self.current_session_id = str(uuid.uuid4())
        now = self._get_colombia_now().isoformat()
        row = {
            "session_id": self.current_session_id,
            "usuario_login": (self.current_user_profile.get("usuario_login") or self.current_user or "").strip(),
            "nombre_profesional": (self.current_user_profile.get("nombre_profesional") or "").strip(),
            "programa": (self.current_user_profile.get("programa") or "").strip(),
            "login_at": now,
            "app_closed_at": None,
        }
        self._usage_upsert_async("utilizacion_il", row, on_conflict="session_id")

    def _mark_app_closed(self):
        if not self._should_track_usage():
            return
        if not self.current_session_id:
            return
        closed_at = self._get_colombia_now().isoformat()
        try:
            result = _supabase_patch_with_queue(
                "utilizacion_il",
                {"session_id": self.current_session_id},
                {"app_closed_at": closed_at},
            )
            if (result or {}).get("status") in {"synced", "queued"}:
                return
        except Exception:
            pass
        row = {"session_id": self.current_session_id, "app_closed_at": closed_at}
        self._usage_upsert_sync("utilizacion_il", row, on_conflict="session_id")

    def track_form_open(self, form_id, form_name):
        if not self._should_track_usage():
            return
        if not self.current_session_id:
            return
        event_id = str(uuid.uuid4())
        self._form_event_ids[form_id] = event_id
        row = {
            "event_id": event_id,
            "session_id": self.current_session_id,
            "usuario_login": (self.current_user_profile.get("usuario_login") or self.current_user or "").strip(),
            "form_id": form_id,
            "form_name": form_name,
            "opened_at": self._get_colombia_now().isoformat(),
            "finished_at": None,
        }
        self._form_event_payloads[form_id] = dict(row)
        self._usage_upsert_async("utilizacion_il_eventos", row, on_conflict="event_id")

    def track_form_finished(self, form_id):
        if not self._should_track_usage():
            return
        if not self.current_session_id:
            return
        event_id = self._form_event_ids.get(form_id)
        if not event_id:
            return
        finished_at = self._get_colombia_now().isoformat()
        try:
            result = _supabase_patch_with_queue(
                "utilizacion_il_eventos",
                {"event_id": event_id},
                {"finished_at": finished_at},
            )
            if (result or {}).get("status") in {"synced", "queued"}:
                self._form_event_ids.pop(form_id, None)
                self._form_event_payloads.pop(form_id, None)
                return
        except Exception:
            pass

        payload = dict(self._form_event_payloads.get(form_id) or {})
        payload.update(
            {
                "event_id": event_id,
                "session_id": self.current_session_id,
                "finished_at": finished_at,
            }
        )
        if not self._usage_upsert_sync("utilizacion_il_eventos", payload, on_conflict="event_id"):
            pass
        self._form_event_ids.pop(form_id, None)
        self._form_event_payloads.pop(form_id, None)

    def create_form_completion_record(self, form_name, company_name, output_path=None):
        if not self._should_track_usage():
            return ""
        usuario_login = (self.current_user_profile.get("usuario_login") or self.current_user or "").strip()
        nombre_usuario = (self.current_user_profile.get("nombre_profesional") or self.current_user or "").strip()
        now_col = self._get_colombia_now()
        registro_id = str(uuid.uuid4())
        row = {
            "registro_id": registro_id,
            "session_id": self.current_session_id,
            "usuario_login": usuario_login,
            "nombre_usuario": nombre_usuario,
            "nombre_formato": (form_name or "").strip(),
            "nombre_empresa": (company_name or "").strip(),
            "path_formato": "",
            "drive_file_id": "",
            "finalizado_at_colombia": now_col.strftime("%Y-%m-%d %H:%M:%S"),
            "finalizado_at_iso": now_col.isoformat(),
            "upload_status": "pending",
            "upload_error": "",
            "upload_attempted_at": None,
            "uploaded_at": None,
        }
        if output_path:
            _log_capture(
                f"create_form_completion_record form={form_name} company={company_name} output={output_path}"
            )
        try:
            self._usage_upsert_sync("formatos_finalizados_il", row, on_conflict="registro_id")
        except Exception as exc:
            _log_capture(f"create_form_completion_record failed registro_id={registro_id} err={exc}")
        return registro_id

    def finalize_form_delivery(self, output_path, *, form_name, company_name, drive_job=None):
        job = {
            "registro_id": self.create_form_completion_record(
                form_name,
                company_name,
                output_path=output_path,
            ),
            "form_name": form_name,
            "company_name": company_name,
            "local_excel_path": output_path,
        }
        if isinstance(drive_job, dict):
            job.update(drive_job)
        result = _perform_drive_upload_attempt(job)
        result["output_path"] = output_path
        result["registro_id"] = str(job.get("registro_id") or "").strip()
        if not result["registro_id"]:
            return result
        if result.get("status") == "pending":
            attempts = 1
            _enqueue_drive_upload_job(
                {
                    **job,
                    "attempts": attempts,
                    "last_error": result.get("error") or "",
                    "next_try_at": time.time() + _next_drive_retry_delay_seconds(attempts),
                }
            )
        elif result.get("status") == "synced":
            with _DRIVE_UPLOAD_LOCK:
                _remove_drive_job_locked(job)
                _persist_drive_upload_queue_locked()
                _persist_drive_failed_queue_locked()
        elif result.get("status") == "failed":
            with _DRIVE_UPLOAD_LOCK:
                _remove_drive_job_locked(job)
                failed_job = dict(job)
                failed_job["attempts"] = 1
                _store_failed_drive_job_locked(failed_job, result.get("error"))
                _persist_drive_upload_queue_locked()
                _persist_drive_failed_queue_locked()
        return result

    def _on_app_close(self):
        _log_capture("_on_app_close: cerrando app")
        if self._session_clock_after_id:
            try:
                self.after_cancel(self._session_clock_after_id)
            except tk.TclError:
                pass
            self._session_clock_after_id = None
        if self._net_status_after_id:
            try:
                self.after_cancel(self._net_status_after_id)
            except tk.TclError:
                pass
            self._net_status_after_id = None
        self._mark_app_closed()
        self.after(250, self.destroy)

    def _get_colombia_now(self):
        try:
            return datetime.now(ZoneInfo("America/Bogota"))
        except Exception:
            return datetime.now()

    def _update_session_clock(self):
        if not self._session_info_label:
            return
        now = self._get_colombia_now()
        nombre = (self.current_user_profile.get("nombre_profesional") or self.current_user or "-").strip()
        programa = (self.current_user_profile.get("programa") or "-").strip()
        usuario = (self.current_user_profile.get("usuario_login") or self.current_user or "-").strip()
        session_text = (
            f"Sesión activa\n"
            f"Nombre: {nombre}\n"
            f"Programa: {programa}\n"
            f"Usuario: {usuario}\n"
            f"COL: {now.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        self._session_info_label.config(text=session_text)
        self._session_clock_after_id = self.after(1000, self._update_session_clock)

    def set_version_info(self, local_version, remote_version):
        local = (local_version or "-").strip()
        remote = (remote_version or "-").strip()
        _log_capture(f"set_version_info local={local} remote={remote}")
        self._version_var.set(f"Versión local: {local} | GitHub: {remote}")

    def _refresh_version_info_async(self):
        if self._version_check_thread and self._version_check_thread.is_alive():
            _log_capture("_refresh_version_info_async omitido: hilo activo")
            return

        def _worker():
            local = get_version()
            remote = None
            _log_capture(f"_refresh_version_info_async worker start local={local}")
            try:
                remote, _assets = get_latest_release_assets()
                _log_capture(
                    f"_refresh_version_info_async worker success remote={remote} assets={list((_assets or {}).keys())}"
                )
            except Exception:
                remote = None
                _log_capture("_refresh_version_info_async worker error al obtener versión remota")
            self.after(0, lambda: self.set_version_info(local, remote))

        self._version_check_thread = threading.Thread(target=_worker, daemon=True)
        _log_capture("_refresh_version_info_async hilo iniciado")
        self._version_check_thread.start()

    def _open_update_page(self):
        _log_capture("_open_update_page: inicio de verificación manual")
        dialog = LoadingDialog(self, title="Verificando actualización")
        dialog.set_status("Consultando versión en GitHub...")
        dialog.set_progress(20)

        result = {"error": None, "local": None, "remote": None, "assets": None}

        def _worker():
            try:
                local = get_version()
                remote, assets = get_latest_release_assets()
                result["local"] = local
                result["remote"] = remote
                result["assets"] = assets
                _log_capture(
                    f"_open_update_page worker success local={local} remote={remote} assets={list((assets or {}).keys())}"
                )
            except Exception as exc:
                result["error"] = str(exc)
                _log_capture(f"_open_update_page worker error: {exc}")

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()

        def _check_done():
            if thread.is_alive():
                self.after(200, _check_done)
                return
            dialog.close()
            if result["error"]:
                _log_capture(f"_open_update_page resultado error: {result['error']}")
                messagebox.showerror("Actualización", f"No se pudo verificar: {result['error']}")
                return
            local = result["local"] or "0.0.0"
            remote = result["remote"]
            self.set_version_info(local, remote)
            if not remote:
                _log_capture("_open_update_page: remote vacío")
                messagebox.showerror("Actualización", "No se pudo obtener la versión remota.")
                return
            if not is_update_available(local, remote):
                _log_capture("_open_update_page: sin actualización disponible")
                messagebox.showinfo("Actualización", "Ya estás usando la última versión.")
                return
            _log_capture(f"_open_update_page: actualización disponible remote={remote} local={local}")
            confirm = messagebox.askyesno(
                "Actualización disponible",
                f"Hay una nueva versión disponible ({remote}).\n¿Deseas actualizar ahora?",
            )
            if not confirm:
                _log_capture("_open_update_page: usuario canceló actualización")
                return
            _log_capture("_open_update_page: usuario confirmó actualización")
            self._start_manual_update(result["assets"] or {})

        self.after(200, _check_done)

    def _start_manual_update(self, assets):
        _log_capture(f"_start_manual_update: inicio assets={list((assets or {}).keys())}")
        dialog = LoadingDialog(self, title="Descargando instalador")
        dialog.set_status("Preparando descarga...")
        dialog.set_progress(5)
        result = {"error": None, "path": None}
        progress_state = {"last": -1, "last_msg": ""}

        def _progress(message, value):
            try:
                value_int = int(value)
            except Exception:
                value_int = -1
            if (
                message != progress_state["last_msg"]
                or value_int >= progress_state["last"] + 10
                or value_int in {0, 100}
            ):
                _log_capture(f"_start_manual_update progress: {message} ({value_int})")
                progress_state["last"] = value_int
                progress_state["last_msg"] = message
            self.after(0, lambda: dialog.set_status(message))
            self.after(0, lambda: dialog.set_progress(value))

        def _worker():
            try:
                path = download_installer(assets, progress_callback=_progress)
                result["path"] = path
                _log_capture(f"_start_manual_update: instalador descargado en {path}")
                self.after(0, lambda: dialog.set_status("Instalando actualización..."))
                self.after(0, lambda: dialog.set_progress(100))
                run_installer(path, wait=True)
                _log_capture("_start_manual_update: run_installer finalizado")
            except Exception as exc:
                result["error"] = str(exc)
                _log_capture(f"_start_manual_update error: {exc}")

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()

        def _check_done():
            if thread.is_alive():
                self.after(300, _check_done)
                return
            dialog.close()
            if result["error"]:
                _log_capture(f"_start_manual_update resultado error: {result['error']}")
                messagebox.showerror("Actualización", f"No se pudo actualizar: {result['error']}")
                return
            _log_capture("_start_manual_update resultado OK: mostrando countdown")
            self._show_restart_countdown()

        self.after(300, _check_done)

    def _show_restart_countdown(self, seconds=5):
        modal = tk.Toplevel(self)
        modal.title("Actualización completada")
        modal.configure(bg=COLOR_LIGHT_BG)
        modal.transient(self)
        modal.grab_set()
        modal.resizable(False, False)

        body = tk.Frame(modal, bg=COLOR_LIGHT_BG, padx=18, pady=14)
        body.pack(fill="both", expand=True)
        tk.Label(
            body,
            text="Instalación terminada",
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 6))
        countdown = tk.Label(
            body,
            text="Reiniciando en 5 segundos...",
            font=("Arial", 10),
            fg=COLOR_TEAL,
            bg=COLOR_LIGHT_BG,
        )
        countdown.pack(anchor="w")

        modal.update_idletasks()
        w, h = 360, 140
        x = (modal.winfo_screenwidth() // 2) - (w // 2)
        y = (modal.winfo_screenheight() // 2) - (h // 2)
        modal.geometry(f"{w}x{h}+{x}+{y}")

        def _tick(remaining):
            if remaining <= 0:
                try:
                    modal.destroy()
                except Exception:
                    pass
                self._restart_app()
                return
            countdown.config(text=f"Reiniciando en {remaining} segundos...")
            modal.after(1000, lambda: _tick(remaining - 1))

        _tick(seconds)

    def _restart_app(self):
        try:
            if getattr(sys, "frozen", False):
                args = [sys.executable]
            else:
                args = [sys.executable, os.path.abspath(__file__)]
            _log_capture(f"_restart_app: relanzando con args={args}")
            subprocess.Popen(args, close_fds=True)
        except Exception:
            _log_capture("_restart_app: fallo al relanzar app")
            pass
        self.after(200, self.destroy)

    def _build_header(self):
        self.header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.header.pack(fill="x", padx=24, pady=(24, 8))
        self.header.grid_columnconfigure(0, weight=1)
        self.header.grid_columnconfigure(1, weight=0)

        left = tk.Frame(self.header, bg=COLOR_LIGHT_BG)
        left.grid(row=0, column=0, sticky="w")

        title = tk.Label(
            left,
            text="Hub de Formularios",
            font=("Arial", 20, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            left,
            text="Selecciona el formulario que necesitas diligenciar",
            font=("Arial", 11),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        subtitle.pack(anchor="w", pady=(4, 0))

        right = tk.Frame(
            self.header,
            bg="#EEF5FF",
            bd=1,
            relief="solid",
            padx=10,
            pady=8,
        )
        right.grid(row=0, column=1, sticky="ne", padx=(16, 0))
        self._session_info_label = tk.Label(
            right,
            text="Sesión activa",
            justify="left",
            anchor="w",
            font=("Arial", 10, "bold"),
            fg="#1F2A44",
            bg="#EEF5FF",
        )
        self._session_info_label.pack(anchor="w")
        tk.Label(
            right,
            textvariable=self._version_var,
            justify="left",
            anchor="w",
            font=("Arial", 9),
            fg="#333333",
            bg="#EEF5FF",
        ).pack(anchor="w", pady=(8, 2))
        ttk.Button(
            right,
            text="Actualizar aplicación",
            command=self._open_update_page,
        ).pack(anchor="e", pady=(2, 0))
        self._net_status_label = tk.Label(
            right,
            text="Estado: verificando...",
            justify="left",
            anchor="w",
            font=("Arial", 9, "bold"),
            fg="#1F2A44",
            bg="#EEF5FF",
        )
        self._net_status_label.pack(anchor="w", pady=(8, 0))
        self._sync_panel_btn = ttk.Button(
            right,
            text="Sincronización",
            command=self._open_sync_panel,
        )
        self._sync_panel_btn.pack(anchor="e", pady=(4, 0))
        if self._session_clock_after_id:
            try:
                self.after_cancel(self._session_clock_after_id)
            except tk.TclError:
                pass
            self._session_clock_after_id = None
        if self._net_status_after_id:
            try:
                self.after_cancel(self._net_status_after_id)
            except tk.TclError:
                pass
            self._net_status_after_id = None
        self._update_session_clock()
        self._start_network_status_monitor()
        self._refresh_version_info_async()

    def _start_network_status_monitor(self):
        if self._net_check_thread and self._net_check_thread.is_alive():
            self._net_status_after_id = self.after(1500, self._start_network_status_monitor)
            return

        result = {
            "services": {},
            "supabase_pending": 0,
            "supabase_failed": 0,
            "drive_pending": 0,
            "drive_failed": 0,
        }
        _cache_age = time.time() - (self._service_probe_cache_time or 0.0)
        _use_cached = _cache_age < 15 and bool(self._service_probe_cache.get("internet"))

        def _worker():
            if _use_cached:
                result["services"] = {}
            else:
                result["services"] = probe_startup_services(log_enabled=False)
            supabase_stats = _get_supabase_write_queue_stats() or {}
            drive_stats = _get_drive_upload_queue_stats() or {}
            result["supabase_pending"] = int(supabase_stats.get("pending") or 0)
            result["supabase_failed"] = int(supabase_stats.get("failed") or 0)
            result["drive_pending"] = int(drive_stats.get("pending") or 0)
            result["drive_failed"] = int(drive_stats.get("failed") or 0)

        self._net_check_thread = threading.Thread(target=_worker, daemon=True)
        self._net_check_thread.start()

        def _finish():
            if self._net_check_thread and self._net_check_thread.is_alive():
                self._net_status_after_id = self.after(200, _finish)
                return
            services = result.get("services") or {}
            if isinstance(services, dict) and services:
                self._service_probe_cache = dict(services)
                self._service_probe_cache_time = time.time()
            internet_state = self._service_probe_cache.get("internet") or {}
            supabase_state = self._service_probe_cache.get("supabase") or {}
            drive_state = self._service_probe_cache.get("drive") or {}
            self._is_online = bool(internet_state.get("ok"))
            supabase_pending = int(result.get("supabase_pending") or 0)
            supabase_failed = int(result.get("supabase_failed") or 0)
            drive_pending = int(result.get("drive_pending") or 0)
            drive_failed = int(result.get("drive_failed") or 0)
            total_pending = supabase_pending + drive_pending
            total_failed = supabase_failed + drive_failed
            if self._net_status_label:
                state_text = "Online" if self._is_online else "Offline"
                color = "#0A7D2E" if self._is_online else "#B00020"
                supabase_text = "OK" if supabase_state.get("ok") else "Falla"
                drive_text = "OK" if drive_state.get("ok") else "Falla"
                self._net_status_label.config(
                    text=(
                        f"Internet: {state_text} | "
                        f"Supabase: {supabase_text} ({supabase_pending}/{supabase_failed}) | "
                        f"Drive: {drive_text} ({drive_pending}/{drive_failed})"
                    ),
                    fg=color,
                )
            if self._sync_panel_btn:
                self._sync_panel_btn.config(text=f"Sincronización ({total_pending}/{total_failed})")
            self._net_status_after_id = self.after(9000, self._start_network_status_monitor)

        _finish()

    def _open_sync_panel(self):
        modal = tk.Toplevel(self)
        modal.title("Estado de sincronización")
        modal.configure(bg=COLOR_LIGHT_BG)
        modal.transient(self)
        modal.grab_set()
        modal.geometry("980x620")

        frame = tk.Frame(modal, bg=COLOR_LIGHT_BG, padx=12, pady=10)
        frame.pack(fill="both", expand=True)

        internet_state = self._service_probe_cache.get("internet") or {}
        supabase_state = self._service_probe_cache.get("supabase") or {}
        drive_state = self._service_probe_cache.get("drive") or {}
        status_txt = "Online" if self._is_online else "Offline"
        status_color = "#0A7D2E" if self._is_online else "#B00020"
        header = tk.Label(
            frame,
            text=(
                f"Internet: {status_txt} | "
                f"Supabase: {'OK' if supabase_state.get('ok') else 'Falla'} | "
                f"Drive: {'OK' if drive_state.get('ok') else 'Falla'}"
            ),
            font=("Arial", 11, "bold"),
            fg=status_color,
            bg=COLOR_LIGHT_BG,
        )
        header.pack(anchor="w", pady=(0, 8))

        summary_lbl = tk.Label(
            frame,
            text="",
            font=("Arial", 9),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        summary_lbl.pack(anchor="w", pady=(0, 8))

        tk.Label(
            frame,
            text="Pendientes de envío",
            font=("Arial", 10, "bold"),
            fg="#1F2A44",
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 4))

        columns = ("origen", "op", "tabla", "intentos", "proximo", "error")
        pending_box = tk.Frame(frame, bg="white", bd=1, relief="solid")
        pending_box.pack(fill="both", expand=True)
        pending_scrollbar = tk.Scrollbar(pending_box, orient="vertical", width=SCROLLBAR_WIDTH)
        pending_scrollbar.pack(side="right", fill="y")
        pending_tree = ttk.Treeview(
            pending_box,
            columns=columns,
            show="headings",
            yscrollcommand=pending_scrollbar.set,
        )
        pending_scrollbar.config(command=pending_tree.yview)
        pending_tree.heading("origen", text="Origen")
        pending_tree.heading("op", text="Operación")
        pending_tree.heading("tabla", text="Tabla")
        pending_tree.heading("intentos", text="Intentos")
        pending_tree.heading("proximo", text="Próximo intento")
        pending_tree.heading("error", text="Último error")
        pending_tree.column("origen", width=90, anchor="w")
        pending_tree.column("op", width=90, anchor="w")
        pending_tree.column("tabla", width=170, anchor="w")
        pending_tree.column("intentos", width=80, anchor="center")
        pending_tree.column("proximo", width=170, anchor="w")
        pending_tree.column("error", width=330, anchor="w")
        pending_tree.pack(side="left", fill="both", expand=True)

        tk.Label(
            frame,
            text="Fallidos no reintentables",
            font=("Arial", 10, "bold"),
            fg="#1F2A44",
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(10, 4))

        failed_box = tk.Frame(frame, bg="white", bd=1, relief="solid")
        failed_box.pack(fill="both", expand=True)
        failed_scrollbar = tk.Scrollbar(failed_box, orient="vertical", width=SCROLLBAR_WIDTH)
        failed_scrollbar.pack(side="right", fill="y")
        failed_tree = ttk.Treeview(
            failed_box,
            columns=("origen", "op", "tabla", "intentos", "failed_at", "error"),
            show="headings",
            yscrollcommand=failed_scrollbar.set,
        )
        failed_scrollbar.config(command=failed_tree.yview)
        failed_tree.heading("origen", text="Origen")
        failed_tree.heading("op", text="Operación")
        failed_tree.heading("tabla", text="Tabla")
        failed_tree.heading("intentos", text="Intentos")
        failed_tree.heading("failed_at", text="Falló en")
        failed_tree.heading("error", text="Error")
        failed_tree.column("origen", width=90, anchor="w")
        failed_tree.column("op", width=90, anchor="w")
        failed_tree.column("tabla", width=170, anchor="w")
        failed_tree.column("intentos", width=80, anchor="center")
        failed_tree.column("failed_at", width=170, anchor="w")
        failed_tree.column("error", width=330, anchor="w")
        failed_tree.pack(side="left", fill="both", expand=True)

        def _fmt_epoch(value):
            try:
                ts = float(value or 0)
                if ts <= 0:
                    return "-"
                return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                return "-"

        def _reload_rows():
            for item in pending_tree.get_children():
                pending_tree.delete(item)
            for item in failed_tree.get_children():
                failed_tree.delete(item)

            internet_state = self._service_probe_cache.get("internet") or {}
            supabase_state = self._service_probe_cache.get("supabase") or {}
            drive_state = self._service_probe_cache.get("drive") or {}
            header.config(
                text=(
                    f"Internet: {'Online' if self._is_online else 'Offline'} | "
                    f"Supabase: {'OK' if supabase_state.get('ok') else 'Falla'} | "
                    f"Drive: {'OK' if drive_state.get('ok') else 'Falla'}"
                ),
                fg="#0A7D2E" if self._is_online else "#B00020",
            )

            supabase_pending_rows = _get_supabase_write_queue_snapshot(limit=500)
            supabase_failed_rows = _get_supabase_failed_writes_snapshot(limit=500)
            drive_pending_rows = _get_drive_upload_queue_snapshot(limit=500)
            drive_failed_rows = _get_drive_failed_uploads_snapshot(limit=500)

            pending_rows = []
            for row in supabase_pending_rows:
                pending_rows.append(
                    (
                        "Supabase",
                        row.get("op") or "-",
                        row.get("table") or "-",
                        int(row.get("attempts") or 0),
                        _fmt_epoch(row.get("next_try_at")),
                        (row.get("last_error") or "")[:280],
                    )
                )
            for row in drive_pending_rows:
                pending_rows.append(
                    (
                        "Drive",
                        _drive_upload_operation_label(row),
                        f"{row.get('form_name') or '-'} | {row.get('company_name') or '-'}",
                        int(row.get("attempts") or 0),
                        _fmt_epoch(row.get("next_try_at")),
                        (row.get("last_error") or "")[:280],
                    )
                )

            failed_rows = []
            for row in supabase_failed_rows:
                failed_rows.append(
                    (
                        "Supabase",
                        row.get("op") or "-",
                        row.get("table") or "-",
                        int(row.get("attempts") or 0),
                        _fmt_epoch(row.get("failed_at")),
                        (row.get("error") or "")[:280],
                    )
                )
            for row in drive_failed_rows:
                failed_rows.append(
                    (
                        "Drive",
                        _drive_upload_operation_label(row),
                        f"{row.get('form_name') or '-'} | {row.get('company_name') or '-'}",
                        int(row.get("attempts") or 0),
                        _fmt_epoch(row.get("failed_at")),
                        (row.get("error") or "")[:280],
                    )
                )

            summary_lbl.config(
                text=(
                    f"Supabase pendientes/fallidos: {len(supabase_pending_rows)}/{len(supabase_failed_rows)} | "
                    f"Drive pendientes/fallidos: {len(drive_pending_rows)}/{len(drive_failed_rows)} | "
                    f"Último check: Internet {internet_state.get('status_text') or '-'} / "
                    f"Supabase {supabase_state.get('status_text') or '-'} / "
                    f"Drive {drive_state.get('status_text') or '-'}"
                )
            )

            if not pending_rows:
                pending_tree.insert("", "end", values=("-", "-", "-", "-", "-", "Sin pendientes"))
            else:
                for row in pending_rows:
                    pending_tree.insert("", "end", values=row)

            if not failed_rows:
                failed_tree.insert("", "end", values=("-", "-", "-", "-", "-", "Sin fallidos"))
            else:
                for row in failed_rows:
                    failed_tree.insert("", "end", values=row)

        def _retry_now():
            supabase_count = _supabase_retry_all_queued_writes()
            drive_count = _drive_retry_all_queued_uploads()
            self.show_toast(
                f"Reintento forzado | Supabase: {supabase_count} | Drive: {drive_count}"
            )
            _reload_rows()
            self._start_network_status_monitor()

        actions = tk.Frame(frame, bg=COLOR_LIGHT_BG)
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Reintentar pendientes", command=_retry_now).pack(side="left")
        ttk.Button(actions, text="Actualizar", command=_reload_rows).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Cerrar", command=modal.destroy).pack(side="right")

        _reload_rows()

    def _norm_match(self, value):
        return _normalize_ascii_text(value).lower()

    def _build_profesional_aliases(self, full_name):
        full = str(full_name or "").strip()
        if not full:
            return set()
        parts = [p for p in full.split() if p]
        aliases = {self._norm_match(full)}
        if len(parts) >= 2:
            aliases.add(self._norm_match(f"{parts[0]} {parts[-1]}"))
        if len(parts) >= 3:
            aliases.add(self._norm_match(f"{parts[0]} {parts[-2]}"))
        # Alias de pares contiguos: "alejandra perez", "laura alejandra", etc.
        for idx in range(len(parts) - 1):
            aliases.add(self._norm_match(f"{parts[idx]} {parts[idx + 1]}"))
        # Nombres compuestos + apellidos (patron comun en Colombia).
        if len(parts) >= 4:
            given_names = parts[:-2]
            surnames = parts[-2:]
            for given in given_names:
                for surname in surnames:
                    aliases.add(self._norm_match(f"{given} {surname}"))
            aliases.add(self._norm_match(" ".join(surnames)))
        return aliases

    def _is_profesional_match(self, asignado_text, aliases):
        asignado_norm = self._norm_match(asignado_text)
        if not asignado_norm:
            return False
        if asignado_norm in aliases:
            return True
        # Permite match parcial solo para alias suficientemente descriptivos.
        for alias in aliases:
            if len(alias) < 7:
                continue
            if alias in asignado_norm or asignado_norm in alias:
                return True
        return False

    def _normalize_profesional_asignado(self):
        profesionales = _supabase_get_paged(
            "profesionales",
            {"select": "nombre_profesional"},
            page_size=1000,
            max_pages=20,
        )
        alias_map = {}
        for row in profesionales:
            nombre = (row.get("nombre_profesional") or "").strip()
            if not nombre:
                continue
            for alias in self._build_profesional_aliases(nombre):
                alias_map.setdefault(alias, nombre)

        empresas = _supabase_get_paged(
            "empresas",
            {
                "select": "id,profesional_asignado",
                "profesional_asignado": "not.is.null",
            },
            page_size=1000,
            max_pages=50,
        )
        updates = []
        for row in empresas:
            current = (row.get("profesional_asignado") or "").strip()
            if not current:
                continue
            key = self._norm_match(current)
            target = alias_map.get(key)
            if target and target != current:
                updates.append({"id": row.get("id"), "profesional_asignado": target})
        if updates:
            _supabase_upsert_with_queue("empresas", updates, on_conflict="id")

    def _get_assigned_companies(self):
        user_login = self._norm_match(self.current_user_profile.get("usuario_login") or self.current_user)
        full_name = self._norm_match(self.current_user_profile.get("nombre_profesional"))
        is_admin = bool(self.current_user_profile.get("is_admin"))
        can_view_all = (
            is_admin
            or
            user_login in {"test", "sanpac", "sarzam", "sarzambrano"}
            or "sandra pachon" in full_name
            or "sara zambrano" in full_name
        )

        def _fetch_empresas(select_clause):
            return _supabase_get_paged(
                "empresas",
                {"select": select_clause},
                page_size=1000,
                max_pages=50,
            )

        try:
            empresas = _fetch_empresas(
                "id,nombre_empresa,nit_empresa,ciudad_empresa,profesional_asignado,estado,comentarios_empresas"
            )
        except Exception:
            try:
                empresas = _fetch_empresas(
                    "id,nombre_empresa,nit_empresa,ciudad_empresa,profesional_asignado,estado,comentarios_empresas,comentarios"
                )
            except Exception:
                empresas = _fetch_empresas(
                    "id,nombre_empresa,nit_empresa,ciudad_empresa,profesional_asignado"
                )
            for row in empresas:
                row.setdefault("estado", "")
                row.setdefault("comentarios_empresas", "")
                if not row.get("comentarios_empresas"):
                    row["comentarios_empresas"] = row.get("comentarios_empresa") or row.get("comentarios") or ""
        if can_view_all:
            assigned = [row for row in empresas if (row.get("nombre_empresa") or "").strip()]
            assigned.sort(key=lambda r: self._norm_match(r.get("nombre_empresa") or ""))
            return assigned

        full_name = (self.current_user_profile.get("nombre_profesional") or "").strip()
        aliases = self._build_profesional_aliases(full_name)
        if not aliases:
            return []
        assigned = []
        for row in empresas:
            asignado = (row.get("profesional_asignado") or "").strip()
            if self._is_profesional_match(asignado, aliases):
                assigned.append(row)
        assigned.sort(key=lambda r: self._norm_match(r.get("nombre_empresa") or ""))
        return assigned

    def _get_company_estado_options(self):
        return [opt for opt in DEFAULT_EMPRESA_ESTADOS if str(opt).strip()]

    def _filtered_sorted_companies(self):
        term = self._norm_match(self._companies_search_var.get() if self._companies_search_var else "")
        items = []
        for row in self._companies_all:
            nit = (row.get("nit_empresa") or "").strip()
            empresa = (row.get("nombre_empresa") or "").strip()
            profesional = (row.get("profesional_asignado") or "").strip()
            haystack = self._norm_match(f"{nit} {empresa} {profesional}")
            if term and term not in haystack:
                continue
            items.append(row)

        mode = self._companies_sort_var.get() if self._companies_sort_var else "Empresa A-Z"
        if mode == "Empresa Z-A":
            items.sort(key=lambda r: self._norm_match(r.get("nombre_empresa") or ""), reverse=True)
        elif mode == "NIT menor-mayor":
            items.sort(key=lambda r: self._norm_match(r.get("nit_empresa") or ""))
        elif mode == "NIT mayor-menor":
            items.sort(key=lambda r: self._norm_match(r.get("nit_empresa") or ""), reverse=True)
        else:
            items.sort(key=lambda r: self._norm_match(r.get("nombre_empresa") or ""))
        return items

    def _render_companies(self, *_args):
        if not self._companies_tree:
            return
        for item in self._companies_tree.get_children():
            self._companies_tree.delete(item)
        self._companies_by_id = {}

        items = self._filtered_sorted_companies()
        if not items:
            self._companies_tree.insert("", "end", iid="__empty__", values=("-", "No hay empresas para mostrar.", "-"))
            return

        for idx, row in enumerate(items, start=1):
            row_id = str(row.get("id") or "")
            if not row_id:
                row_id = f"row_{idx}"
            if row_id in self._companies_by_id:
                row_id = f"{row_id}_{idx}"
            self._companies_by_id[row_id] = row
            self._companies_tree.insert(
                "",
                "end",
                iid=row_id,
                values=(
                    (row.get("nit_empresa") or "").strip(),
                    (row.get("nombre_empresa") or "").strip(),
                    (row.get("profesional_asignado") or "").strip(),
                ),
            )

    def _open_company_editor(self, company_row):
        company_id = company_row.get("id")
        if not company_id:
            return

        estado_options = self._get_company_estado_options()
        current_estado = (company_row.get("estado") or "").strip()

        modal = tk.Toplevel(self)
        modal.title("Actualizar estado de empresa")
        modal.configure(bg=COLOR_LIGHT_BG)
        modal.transient(self)
        modal.grab_set()
        modal.resizable(False, False)

        frame = tk.Frame(modal, bg=COLOR_LIGHT_BG, padx=14, pady=12)
        frame.pack(fill="both", expand=True)

        empresa = (company_row.get("nombre_empresa") or "").strip()
        nit = (company_row.get("nit_empresa") or "").strip()
        tk.Label(
            frame,
            text=f"Empresa: {empresa}  |  NIT: {nit}",
            font=("Arial", 10, "bold"),
            bg=COLOR_LIGHT_BG,
            fg="#333333",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        tk.Label(frame, text="Estado", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=(0, 8)
        )
        estado_var = tk.StringVar(
            value=current_estado if current_estado in estado_options else estado_options[0]
        )
        estado_combo = ttk.Combobox(
            frame,
            textvariable=estado_var,
            state="readonly",
            width=40,
            values=estado_options,
        )
        estado_combo.grid(row=1, column=1, sticky="w", pady=(0, 8))

        tk.Label(frame, text="Comentarios", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=2, column=0, sticky="nw", padx=(0, 10), pady=(0, 6)
        )
        comentarios_txt = tk.Text(frame, width=52, height=6, wrap="word")
        comentarios_txt.grid(row=2, column=1, sticky="w", pady=(0, 6))
        comentarios_txt.insert(
            "1.0",
            company_row.get("comentarios_empresas")
            or company_row.get("comentarios_empresa")
            or company_row.get("comentarios")
            or "",
        )

        status_lbl = tk.Label(frame, text="", font=("Arial", 9), fg="#B00020", bg=COLOR_LIGHT_BG)
        status_lbl.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 8))

        actions = tk.Frame(frame, bg=COLOR_LIGHT_BG)
        actions.grid(row=4, column=0, columnspan=2, sticky="e")

        def _save():
            estado = estado_var.get().strip()
            comentarios = comentarios_txt.get("1.0", tk.END).strip()
            if not estado:
                status_lbl.config(text="Selecciona un estado válido.")
                return
            try:
                last_exc = None
                last_status = "synced"
                for comments_col in ("comentarios_empresas", "comentarios_empresa", "comentarios", "comentario_empresa"):
                    try:
                        result = _supabase_patch_with_queue(
                            "empresas",
                            {"id": company_id},
                            {
                                "estado": estado,
                                comments_col: comentarios,
                            },
                        )
                        last_status = (result or {}).get("status") or "synced"
                        company_row["estado"] = estado
                        company_row["comentarios_empresas"] = comentarios
                        company_row["comentarios_empresa"] = comentarios
                        company_row["comentarios"] = comentarios
                        self._render_companies()
                        modal.destroy()
                        if last_status == "queued":
                            self.show_toast("Sin internet: cambio de empresa en cola")
                        return
                    except Exception as exc:
                        last_exc = exc
                raise last_exc
            except Exception as exc:
                status_lbl.config(text=f"No se pudo guardar: {exc}")
                return

        ttk.Button(actions, text="Cancelar", command=modal.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(actions, text="Guardar", command=_save).pack(side="right")

        try:
            estado_combo.focus_set()
        except tk.TclError:
            pass

    def _on_company_double_click(self, _event=None):
        if not self._companies_tree:
            return
        item_id = self._companies_tree.focus()
        if not item_id or item_id == "__empty__":
            return
        row = self._companies_by_id.get(item_id)
        if not row:
            return
        self._open_company_editor(row)

    def _get_current_user_login(self):
        login = (self.current_user_profile.get("usuario_login") or self.current_user or "").strip()
        return login.lower()

    def _get_user_drafts(self):
        user_login = self._get_current_user_login()
        if not user_login:
            return []
        data = _load_drafts_store()
        users = data.get("users", {})
        drafts = users.get(user_login, [])
        if not isinstance(drafts, list):
            return []
        return [item for item in drafts if isinstance(item, dict)]

    def _refresh_drafts_badge(self):
        if not self._drafts_btn:
            return
        count = len(self._get_user_drafts())
        self._drafts_btn.config(text=f"Borradores ({count})")

    def _capture_window_draft_state(self, window, module):
        module.save_cache_to_file()
        cache_snapshot = copy.deepcopy(module.get_form_cache() or {})
        ui_section = str(
            getattr(window, "_current_section", "")
            or cache_snapshot.get("_last_section")
            or "section_1"
        ).strip()
        ui_snapshot = _collect_visible_input_snapshot(window)
        if ui_section:
            cache_snapshot["_last_section"] = ui_section
        return cache_snapshot, ui_snapshot, ui_section

    def _draft_state_has_content(self, cache_snapshot, ui_snapshot):
        return _cache_snapshot_has_meaningful_values(cache_snapshot) or _snapshot_has_meaningful_values(
            ui_snapshot
        )

    def _persist_form_draft(self, window, *, allow_empty=False, silent=False, toast_text=""):
        form_id = getattr(window, "_form_id", "") or ""
        form_name = getattr(window, "_form_name", "") or form_id
        module = FORM_MODULE_MAP.get(form_id)
        if not module:
            if not silent:
                messagebox.showinfo("Guardar", "Este formulario no tiene guardado disponible.")
            return False
        if not hasattr(module, "get_form_cache") or not hasattr(module, "save_cache_to_file"):
            if not silent:
                messagebox.showinfo("Guardar", "No se pudo guardar este formulario.")
            return False

        try:
            cache_snapshot, ui_snapshot, ui_section = self._capture_window_draft_state(window, module)
        except Exception as exc:
            if silent:
                _log_capture(f"[DRAFT] capture_failed form={form_id} err={exc}")
                return False
            messagebox.showerror("Guardar", f"No se pudo leer el formulario actual: {exc}")
            return False

        if not allow_empty and not self._draft_state_has_content(cache_snapshot, ui_snapshot):
            if not silent:
                messagebox.showinfo("Guardar", "Aún no hay datos para guardar.")
            return False

        try:
            fingerprint_payload = {
                "cache": cache_snapshot,
                "ui_section": ui_section,
                "ui_snapshot": ui_snapshot,
            }
            fingerprint = hashlib.sha1(
                json.dumps(
                    fingerprint_payload,
                    ensure_ascii=False,
                    sort_keys=True,
                    default=str,
                ).encode("utf-8")
            ).hexdigest()
        except Exception:
            fingerprint = ""
        if fingerprint and fingerprint == getattr(window, "_draft_last_fingerprint", ""):
            if not silent:
                self.show_toast("Borrador ya guardado")
            return False

        user_login = self._get_current_user_login()
        if not user_login:
            if not silent:
                messagebox.showerror("Guardar", "No hay una sesión activa.")
            return False

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        company_name = _extract_draft_company_name(cache_snapshot) or "Sin empresa"
        session_key = getattr(window, "_draft_session_key", "") or uuid.uuid4().hex
        window._draft_session_key = session_key
        company_key = _extract_draft_company_key(cache_snapshot)
        if company_key == "sin_clave":
            company_key = f"sesion:{form_id}:{session_key}"

        data = _load_drafts_store()
        users = data.setdefault("users", {})
        drafts = users.setdefault(user_login, [])
        if not isinstance(drafts, list):
            drafts = []
            users[user_login] = drafts

        draft_id = str(getattr(window, "_draft_id", "") or "").strip()
        existing = None
        if draft_id:
            for item in drafts:
                if str(item.get("draft_id") or "") == draft_id:
                    existing = item
                    break
        if existing is None:
            for item in drafts:
                if (
                    str(item.get("form_id") or "") == form_id
                    and str(item.get("company_key") or "") == company_key
                ):
                    existing = item
                    break

        if existing is None:
            draft_id = draft_id or str(uuid.uuid4())
            existing = {
                "draft_id": draft_id,
                "form_id": form_id,
                "form_name": form_name,
                "company_key": company_key,
                "company_name": company_name,
                "draft_session_key": session_key,
                "created_at": now,
            }
            drafts.append(existing)
        else:
            draft_id = str(existing.get("draft_id") or draft_id or uuid.uuid4())
            existing["draft_id"] = draft_id

        existing["updated_at"] = now
        existing["last_section"] = ui_section or cache_snapshot.get("_last_section", "")
        existing["cache"] = cache_snapshot
        existing["company_key"] = company_key
        existing["company_name"] = company_name
        existing["ui_section"] = ui_section
        existing["ui_snapshot"] = ui_snapshot
        existing["draft_session_key"] = session_key

        try:
            _save_drafts_store(data)
        except Exception as exc:
            if silent:
                _log_capture(f"[DRAFT] save_failed form={form_id} draft_id={draft_id} err={exc}")
                return False
            messagebox.showerror("Guardar", f"No se pudo guardar el borrador: {exc}")
            return False

        window._draft_id = draft_id
        window._draft_last_fingerprint = fingerprint
        self._refresh_drafts_badge()
        if toast_text:
            self.show_toast(toast_text)
        return True

    def _schedule_window_draft_autosave(self, window, delay_ms=250):
        if not window or not window.winfo_exists():
            return
        after_id = getattr(window, "_draft_autosave_after_id", None)
        if after_id:
            try:
                window.after_cancel(after_id)
            except tk.TclError:
                pass

        def _run():
            window._draft_autosave_after_id = None
            self._persist_form_draft(
                window,
                allow_empty=True,
                silent=True,
                toast_text="",
            )

        try:
            window._draft_autosave_after_id = window.after(delay_ms, _run)
        except tk.TclError:
            window._draft_autosave_after_id = None

    def _install_form_autosave_bindings(self, window):
        sticky_bar = getattr(window, "_sticky_actions_bar", None)
        for _path, widget in _iter_widget_paths(window):
            if sticky_bar and _is_descendant_of(widget, sticky_bar):
                continue
            if getattr(widget, "_draft_autosave_bound", False):
                continue
            if isinstance(widget, tk.Text):
                widget.bind(
                    "<FocusOut>",
                    lambda _event=None, w=window: self._schedule_window_draft_autosave(w),
                    add="+",
                )
            elif isinstance(widget, ttk.Combobox):
                widget.bind(
                    "<<ComboboxSelected>>",
                    lambda _event=None, w=window: self._schedule_window_draft_autosave(w),
                    add="+",
                )
                widget.bind(
                    "<FocusOut>",
                    lambda _event=None, w=window: self._schedule_window_draft_autosave(w),
                    add="+",
                )
                widget.bind(
                    "<Return>",
                    lambda _event=None, w=window: self._schedule_window_draft_autosave(w),
                    add="+",
                )
            elif isinstance(widget, (tk.Entry, DateEntry)):
                state = str(widget.cget("state") or "")
                if state != "readonly":
                    widget.bind(
                        "<FocusOut>",
                        lambda _event=None, w=window: self._schedule_window_draft_autosave(w),
                        add="+",
                    )
                    widget.bind(
                        "<Return>",
                        lambda _event=None, w=window: self._schedule_window_draft_autosave(w),
                        add="+",
                    )
            else:
                continue
            widget._draft_autosave_bound = True

    def _delete_window_draft(self, window):
        draft_id = str(getattr(window, "_draft_id", "") or "").strip()
        user_login = self._get_current_user_login()
        if not draft_id or not user_login:
            return False
        data = _load_drafts_store()
        users = data.get("users", {})
        current = users.get(user_login, [])
        if not isinstance(current, list):
            return False
        updated = [row for row in current if str(row.get("draft_id") or "") != draft_id]
        if len(updated) == len(current):
            return False
        users[user_login] = updated
        _save_drafts_store(data)
        window._draft_id = ""
        window._draft_last_fingerprint = ""
        self._refresh_drafts_badge()
        return True

    def _clear_form_memory_caches(self):
        for module in FORM_MODULE_MAP.values():
            try:
                section_cache = getattr(module, "SECTION_1_CACHE", None)
                if isinstance(section_cache, dict):
                    section_cache.clear()
            except Exception:
                pass

    def _refresh_database_cache(self):
        if self._refresh_db_btn:
            self._refresh_db_btn.config(state="disabled", text="Actualizando...")

        def _worker():
            err = None
            rows = []
            try:
                self._clear_form_memory_caches()
                rows = self._get_assigned_companies()
            except Exception as exc:
                err = exc

            def _done():
                if self._refresh_db_btn:
                    self._refresh_db_btn.config(state="normal", text="Actualizar Base de Datos")
                if err:
                    messagebox.showwarning("Base de Datos", f"No se pudo actualizar: {err}")
                    return
                self._companies_all = rows
                self._render_companies()
                self.show_toast("Base de datos actualizada")

            try:
                self.after(0, _done)
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _save_current_form_draft(self, window, *, silent=False, allow_empty=False, toast_text="Borrador guardado"):
        self._persist_form_draft(
            window,
            allow_empty=allow_empty,
            silent=silent,
            toast_text="" if silent else toast_text,
        )

    def _open_draft_entry(self, draft):
        form_id = str(draft.get("form_id") or "")
        module = FORM_MODULE_MAP.get(form_id)
        if not module:
            messagebox.showerror("Borradores", "El formulario de este borrador ya no está disponible.")
            return
        cache_snapshot = draft.get("cache")
        if not isinstance(cache_snapshot, dict) or not cache_snapshot:
            messagebox.showerror("Borradores", "El borrador no tiene datos válidos.")
            return

        form_meta = next((item for item in get_forms() if item.get("id") == form_id), None)
        if not form_meta:
            messagebox.showerror("Borradores", "No se encontró el formulario en el HUB.")
            return
        self._pending_draft_restore = {
            "form_id": form_id,
            "draft_id": str(draft.get("draft_id") or "").strip(),
            "draft_session_key": str(draft.get("draft_session_key") or "").strip(),
            "cache": copy.deepcopy(cache_snapshot),
            "ui_section": str(draft.get("ui_section") or "").strip(),
            "ui_snapshot": copy.deepcopy(draft.get("ui_snapshot") or []),
        }
        window = self._open_form(form_meta)
        if not window:
            self._pending_draft_restore = None
            return
        if window:
            window._draft_id = str(draft.get("draft_id") or "").strip()
            window._draft_session_key = str(draft.get("draft_session_key") or "").strip() or uuid.uuid4().hex
        ui_snapshot = getattr(window, "_draft_restore_pending_ui_snapshot", None)
        if not isinstance(ui_snapshot, list) or not ui_snapshot:
            return

        def _try_apply(attempt=0):
            if not window.winfo_exists():
                return
            applied = _apply_input_snapshot(window, ui_snapshot)
            if applied > 0 or attempt >= 12:
                try:
                    window._draft_restore_pending_ui_snapshot = None
                except Exception:
                    pass
                return
            window.after(150, lambda: _try_apply(attempt + 1))

        window.after(150, _try_apply)

    def _open_drafts_window(self):
        drafts = self._get_user_drafts()
        modal = tk.Toplevel(self)
        modal.title("Borradores guardados")
        modal.configure(bg=COLOR_LIGHT_BG)
        modal.geometry("860x420")
        modal.transient(self)
        modal.grab_set()

        frame = tk.Frame(modal, bg=COLOR_LIGHT_BG, padx=14, pady=12)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text="Formularios guardados",
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 8))

        box = tk.Frame(frame, bg="white", bd=1, relief="solid")
        box.pack(fill="both", expand=True)
        yscroll = tk.Scrollbar(box, orient="vertical", width=SCROLLBAR_WIDTH)
        yscroll.pack(side="right", fill="y")

        tree = ttk.Treeview(
            box,
            columns=("form", "empresa", "seccion", "actualizado"),
            show="headings",
            yscrollcommand=yscroll.set,
        )
        tree.heading("form", text="Formulario")
        tree.heading("empresa", text="Empresa")
        tree.heading("seccion", text="Última sección")
        tree.heading("actualizado", text="Actualizado")
        tree.column("form", width=220, anchor="w")
        tree.column("empresa", width=280, anchor="w")
        tree.column("seccion", width=140, anchor="w")
        tree.column("actualizado", width=170, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        yscroll.config(command=tree.yview)

        draft_by_iid = {}
        for idx, item in enumerate(
            sorted(drafts, key=lambda d: str(d.get("updated_at") or ""), reverse=True),
            start=1,
        ):
            iid = f"draft_{idx}"
            draft_by_iid[iid] = item
            tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    str(item.get("form_name") or item.get("form_id") or ""),
                    str(item.get("company_name") or "Sin empresa"),
                    str(item.get("last_section") or ""),
                    str(item.get("updated_at") or item.get("created_at") or ""),
                ),
            )
        if not draft_by_iid:
            tree.insert("", "end", iid="__empty__", values=("-", "No hay borradores guardados.", "-", "-"))

        actions = tk.Frame(frame, bg=COLOR_LIGHT_BG)
        actions.pack(fill="x", pady=(8, 0))

        def _open_selected():
            sel = tree.focus()
            if not sel or sel == "__empty__":
                return
            draft = draft_by_iid.get(sel)
            if not draft:
                return
            modal.destroy()
            self._open_draft_entry(draft)

        def _delete_selected():
            sel = tree.focus()
            if not sel or sel == "__empty__":
                return
            draft = draft_by_iid.get(sel)
            if not draft:
                return
            draft_id = str(draft.get("draft_id") or "")
            if not draft_id:
                return
            if not messagebox.askyesno("Borradores", "¿Eliminar este borrador?"):
                return
            user_login = self._get_current_user_login()
            data = _load_drafts_store()
            users = data.get("users", {})
            current = users.get(user_login, [])
            users[user_login] = [row for row in current if str(row.get("draft_id") or "") != draft_id]
            _save_drafts_store(data)
            tree.delete(sel)
            draft_by_iid.pop(sel, None)
            self._refresh_drafts_badge()

        ttk.Button(actions, text="Cerrar", command=modal.destroy).pack(side="right")
        ttk.Button(actions, text="Eliminar", command=_delete_selected).pack(side="right", padx=(0, 8))
        ttk.Button(actions, text="Abrir", command=_open_selected).pack(side="right", padx=(0, 8))
        tree.bind("<Double-1>", lambda _e: _open_selected())

    def _build_body(self):
        self.body = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.body.pack(fill="both", expand=True, padx=14, pady=10)
        # Layout responsive: formularios (izquierda) mas ancho que empresas (derecha).
        self.body.grid_columnconfigure(0, weight=3)
        self.body.grid_columnconfigure(1, weight=2)
        self.body.grid_rowconfigure(0, weight=1)

        left = tk.Frame(self.body, bg=COLOR_LIGHT_BG)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.grid_rowconfigure(1, weight=1)
        left.grid_columnconfigure(0, weight=1)
        right = tk.Frame(self.body, bg=COLOR_LIGHT_BG)
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        right.grid_rowconfigure(2, weight=1)
        right.grid_columnconfigure(0, weight=1)

        left_header = tk.Frame(left, bg=COLOR_LIGHT_BG)
        left_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tk.Label(
            left_header,
            text="Formularios",
            font=("Arial", 13, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(side="left", anchor="w")
        self._drafts_btn = ttk.Button(
            left_header,
            text="Borradores (0)",
            command=self._open_drafts_window,
        )
        self._drafts_btn.pack(side="right")
        self._refresh_db_btn = ttk.Button(
            left_header,
            text="Actualizar Base de Datos",
            command=self._refresh_database_cache,
        )
        self._refresh_db_btn.pack(side="right", padx=(0, 8))
        self._refresh_drafts_badge()

        # Lista de formularios con scroll para pantallas pequenas.
        forms_box = tk.Frame(left, bg="white", bd=1, relief="solid")
        forms_box.grid(row=1, column=0, sticky="nsew")
        forms_canvas = tk.Canvas(forms_box, bg=COLOR_LIGHT_BG, highlightthickness=0)
        forms_scroll = _create_vscroll(forms_box, forms_canvas.yview)
        forms_canvas.configure(yscrollcommand=forms_scroll.set)
        forms_scroll.pack(side="right", fill="y")
        forms_canvas.pack(side="left", fill="both", expand=True)

        forms_content = tk.Frame(forms_canvas, bg=COLOR_LIGHT_BG)
        forms_window_id = forms_canvas.create_window((0, 0), window=forms_content, anchor="nw")
        forms_content.grid_columnconfigure(0, weight=1)

        def _sync_forms_scrollregion(_event=None):
            forms_canvas.configure(scrollregion=forms_canvas.bbox("all"))

        def _sync_forms_width(_event=None):
            try:
                forms_canvas.itemconfigure(forms_window_id, width=forms_canvas.winfo_width())
                forms_canvas.xview_moveto(0.0)
            except Exception:
                pass

        forms_content.bind("<Configure>", _sync_forms_scrollregion)
        forms_canvas.bind("<Configure>", _sync_forms_width)
        forms_canvas.after_idle(_sync_forms_width)

        def _on_forms_wheel(event):
            try:
                forms_canvas.xview_moveto(0.0)
                if getattr(event, "delta", 0):
                    delta = int(-1 * (event.delta / 120))
                    if delta == 0:
                        delta = -1 if event.delta > 0 else 1
                    forms_canvas.yview_scroll(delta, "units")
                else:
                    num = getattr(event, "num", None)
                    if num == 4:
                        forms_canvas.yview_scroll(-3, "units")
                    elif num == 5:
                        forms_canvas.yview_scroll(3, "units")
            except Exception:
                pass
            return "break"

        for widget in (forms_canvas, forms_content):
            widget.bind("<MouseWheel>", _on_forms_wheel, add="+")
            widget.bind("<Button-4>", _on_forms_wheel, add="+")
            widget.bind("<Button-5>", _on_forms_wheel, add="+")
            widget.bind("<Shift-MouseWheel>", lambda _e: "break", add="+")
            widget.bind("<Enter>", lambda _e: forms_canvas.xview_moveto(0.0), add="+")

        forms = get_forms()
        if not forms:
            tk.Label(
                forms_content,
                text="No hay formularios disponibles.",
                font=("Arial", 12),
                bg=COLOR_LIGHT_BG,
                fg="#555555",
            ).pack(anchor="w", padx=8, pady=8)
        else:
            for form in forms:
                card = tk.Frame(forms_content, bg="white", bd=1, relief="solid")
                card.pack(fill="x", pady=6, padx=8)
                card.grid_columnconfigure(0, weight=1)
                title = tk.Label(
                    card,
                    text=form["name"],
                    font=("Arial", 12, "bold"),
                    bg="white",
                    fg="#222222",
                    anchor="w",
                    justify="left",
                    wraplength=460,
                )
                title.grid(row=0, column=0, sticky="ew", padx=(12, 8), pady=8)
                action = ttk.Button(
                    card,
                    text="Abrir",
                    command=lambda f=form: self._open_form(f),
                )
                action.grid(row=0, column=1, sticky="e", padx=12, pady=8)
                for _w in (card, title, action):
                    _w.bind("<MouseWheel>", _on_forms_wheel, add="+")
                    _w.bind("<Button-4>", _on_forms_wheel, add="+")
                    _w.bind("<Button-5>", _on_forms_wheel, add="+")

        tk.Label(
            right,
            text="Empresas Asignadas",
            font=("Arial", 13, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        controls = tk.Frame(right, bg=COLOR_LIGHT_BG)
        controls.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        controls.grid_columnconfigure(1, weight=1)

        tk.Label(
            controls,
            text="Buscar:",
            font=("Arial", 10, "bold"),
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._companies_search_var = tk.StringVar()
        search_entry = tk.Entry(controls, textvariable=self._companies_search_var, width=24)
        search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        tk.Label(
            controls,
            text="Ordenar:",
            font=("Arial", 10, "bold"),
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=2, sticky="w", padx=(0, 6))
        self._companies_sort_var = tk.StringVar(value="Empresa A-Z")
        sort_combo = ttk.Combobox(
            controls,
            textvariable=self._companies_sort_var,
            state="readonly",
            width=17,
            values=["Empresa A-Z", "Empresa Z-A", "NIT menor-mayor", "NIT mayor-menor"],
        )
        sort_combo.grid(row=0, column=3, sticky="e")

        box = tk.Frame(right, bg="white", bd=1, relief="solid")
        box.grid(row=2, column=0, sticky="nsew")

        scrollbar = tk.Scrollbar(box, orient="vertical", width=SCROLLBAR_WIDTH)
        scrollbar.pack(side="right", fill="y")

        tree = ttk.Treeview(
            box,
            columns=("nit", "empresa", "profesional"),
            show="headings",
            yscrollcommand=scrollbar.set,
        )
        self._companies_tree = tree
        tree.heading("nit", text="NIT")
        tree.heading("empresa", text="Nombre Empresa")
        tree.heading("profesional", text="Profesional Asignado")
        tree.column("nit", width=140, anchor="w")
        tree.column("empresa", width=360, anchor="w")
        tree.column("profesional", width=200, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=tree.yview)

        try:
            self._companies_all = self._get_assigned_companies()
        except Exception as exc:
            self._companies_all = []
            messagebox.showwarning("Empresas", f"Error cargando empresas: {exc}")
        self._companies_search_var.trace_add("write", self._render_companies)
        sort_combo.bind("<<ComboboxSelected>>", self._render_companies)
        tree.bind("<Double-1>", self._on_company_double_click)
        self._render_companies()

    def _bind_form_runtime(self, window, form_meta):
        if not window or not form_meta:
            return
        form_id = str(form_meta.get("id") or "")
        form_name = str(form_meta.get("name") or form_id)
        window._form_id = form_id
        window._form_name = form_name
        module = FORM_MODULE_MAP.get(form_id)
        if module and hasattr(module, "get_form_cache") and hasattr(module, "save_cache_to_file"):
            window._save_draft_command = lambda w=window: self._save_current_form_draft(w)
        else:
            window._save_draft_command = None
        window._current_section = "section_1"
        window._draft_session_key = getattr(window, "_draft_session_key", "") or uuid.uuid4().hex
        window._draft_autosave_after_id = None
        window._draft_last_fingerprint = getattr(window, "_draft_last_fingerprint", "")
        for name in [n for n in dir(window) if n.startswith("_show_section")]:
            original = getattr(window, name, None)
            if not callable(original):
                continue
            if getattr(original, "_section_wrapped", False):
                continue

            def _make_wrapper(fn, method_name):
                def _wrapped(*args, **kwargs):
                    section = method_name.replace("_show_", "")
                    window._current_section = section
                    result = fn(*args, **kwargs)
                    try:
                        _attach_dictation_for_section(window, form_id, section)
                    except Exception as exc:
                        _log_capture(
                            f"[DICTATION] attach_wrapper_failed form={form_id} section={section} err={exc}"
                        )
                    try:
                        self._install_form_autosave_bindings(window)
                        self._schedule_window_draft_autosave(window, delay_ms=250)
                    except Exception as exc:
                        _log_capture(
                            f"[DRAFT] autosave_wrapper_failed form={form_id} section={section} err={exc}"
                        )
                    return result

                _wrapped._section_wrapped = True
                return _wrapped

            setattr(window, name, _make_wrapper(original, name))
        try:
            cache = module.get_form_cache() if module and hasattr(module, "get_form_cache") else {}
            if isinstance(cache, dict) and cache.get("_last_section"):
                window._current_section = str(cache.get("_last_section"))
        except Exception:
            pass
        try:
            _attach_dictation_for_section(window, form_id, getattr(window, "_current_section", "section_1"))
        except Exception as exc:
            _log_capture(
                f"[DICTATION] attach_initial_failed form={form_id} section={getattr(window, '_current_section', 'section_1')} err={exc}"
            )
        try:
            self._install_form_autosave_bindings(window)
            self._schedule_window_draft_autosave(window, delay_ms=350)
        except Exception as exc:
            _log_capture(
                f"[DRAFT] autosave_initial_failed form={form_id} section={getattr(window, '_current_section', 'section_1')} err={exc}"
            )

    def _open_form(self, form_meta):
        if form_meta["id"] == "presentacion_programa":
            window = Section1Window(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        if form_meta["id"] == "evaluacion_accesibilidad":
            window = EvaluacionAccesibilidadWindow(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        if form_meta["id"] == "condiciones_vacante":
            window = CondicionesVacanteWindow(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        if form_meta["id"] == "seleccion_incluyente":
            window = SeleccionIncluyenteWindow(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        if form_meta["id"] == "contratacion_incluyente":
            window = ContratacionIncluyenteWindow(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        if form_meta["id"] == "induccion_organizacional":
            window = InduccionOrganizacionalWindow(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        if form_meta["id"] == "induccion_operativa":
            window = InduccionOperativaWindow(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        if form_meta["id"] == "sensibilizacion":
            window = SensibilizacionWindow(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        if form_meta["id"] == "seguimientos":
            window = SeguimientosWindow(self)
            self._bind_form_runtime(window, form_meta)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window
        messagebox.showinfo("Formulario", f"Abrir formulario: {form_meta['name']}")
        return None

    def _ensure_toast(self):
        if self._toast_label is not None:
            return
        self._toast_label = tk.Label(
            self,
            text="",
            bg="#333333",
            fg="white",
            font=("Arial", 9, "bold"),
            padx=12,
            pady=6,
        )
        self._toast_label.place_forget()

    def _hide_toast(self):
        if self._toast_label:
            self._toast_label.place_forget()
        self._toast_after_id = None

    def show_toast(self, text, duration_ms=5000):
        self._ensure_toast()
        if self._toast_after_id is not None:
            self.after_cancel(self._toast_after_id)
            self._toast_after_id = None
        self._toast_label.config(text=text)
        self._toast_label.lift()
        self._toast_label.place(relx=1.0, rely=1.0, x=-24, y=-24, anchor="se")
        if duration_ms is not None:
            self._toast_after_id = self.after(duration_ms, self._hide_toast)

    def _toast_async(self, text, duration_ms=5000):
        self.after(0, lambda: self.show_toast(text, duration_ms))

        threading.Thread(target=_run, daemon=True).start()

class EvaluacionAccesibilidadWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Evaluacion de Accesibilidad - Seccion 1")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = evaluacion_accesibilidad

        self.company_data = None
        self.fields = {}

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "evaluacion_accesibilidad",
            evaluacion_accesibilidad,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_2_1": self._show_section_2_1,
                "section_2_2": self._show_section_2_2,
                "section_2_3": self._show_section_2_3,
                "section_2_4": self._show_section_2_4,
                "section_2_5": self._show_section_2_5,
                "section_2_6": self._show_section_2_6,
                "section_3": self._show_section_3,
                "section_4": self._show_section_4,
                "section_5": self._show_section_5,
                "section_6": self._show_section_6,
                "section_7": self._show_section_7,
                "section_8": self._show_section_8,
            },
            self._show_section_1,
        ):
            return True
        if not evaluacion_accesibilidad.cache_file_exists():
            return False
        resume = messagebox.askyesno(
            "Reanudar",
            "Se encontró una evaluación en progreso. ¿Deseas continuar donde lo dejaste?",
        )
        if not resume:
            evaluacion_accesibilidad.clear_cache_file()
            evaluacion_accesibilidad.clear_form_cache()
            return False
        evaluacion_accesibilidad.load_cache_from_file()
        last_section = evaluacion_accesibilidad.get_form_cache().get("_last_section")
        if last_section == "section_1":
            self._show_section_2()
        elif last_section == "section_2_1":
            self._show_section_2_2()
        elif last_section == "section_2_2":
            self._show_section_2_3()
        elif last_section == "section_2_3":
            self._show_section_2_4()
        elif last_section == "section_2_4":
            self._show_section_2_5()
        elif last_section == "section_2_5":
            self._show_section_2_6()
        elif last_section == "section_2_6":
            self._show_section_3()
        elif last_section == "section_3":
            self._show_section_4()
        elif last_section == "section_4":
            self._show_section_5()
        elif last_section == "section_5":
            self._show_section_6()
        elif last_section == "section_6":
            self._show_section_7()
        else:
            self._show_section_8()
        return True


    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))

        self.header_title = tk.Label(
            header,
            text="1. DATOS DE LA EMPRESA",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        self.header_title.pack(anchor="w")

        self.header_subtitle = tk.Label(
            header,
            text="Busca empresa por NIT y confirma datos.",
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        self.header_subtitle.pack(anchor="w", pady=(4, 0))

    def _build_section_container(self):
        self.section_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)

    def _clear_section_container(self):
        for child in self.section_container.winfo_children():
            child.destroy()

    def _clean_text(self, text):
        if not text:
            return ""
        replacements = {
            "\u00b6\u00a8": "\u00bf",
            "\u00c7?": "\u00cd",
            "\u00c7\u00ad": "\u00e1",
            "\u00c7\u00b8": "\u00e9",
            "\u00c7\u00f0": "\u00ed",
            "\u00c7\u00a7": "\u00fa",
            "\u00c7\u00b1": "\u00f1",
            "\u00c7\u00fc": "\u00f3",
            "\u0418": "\u00f3",
            "\u30f5": "\u00f1",
            "\u9685": "\u00bf",
            "\u30f4": "\u00ed",
        }
        for bad, good in replacements.items():
            text = text.replace(bad, good)
        return text

    def _get_accessible_options(self):
        return ["Sí", "No", "Parcial"]

    def _show_section_1(self):
        self._clear_section_container()
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._build_actions(content)

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. ACCESIBILIDAD F\u00cdSICA")
        self.header_subtitle.config(text="Completa movilidad y entorno urbano.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        title = tk.Label(
            content,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_2_1["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(8, 12))

        self.section2_1_fields = {}
        accesible_options = self._get_accessible_options()
        for question in evaluacion_accesibilidad.SECTION_2_1["questions"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=self._clean_text(question["label"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))

            field_id = question["id"]
            self.section2_1_fields[field_id] = {}

            if question["type"] == "accesible_con_observaciones":
                tk.Label(
                    row,
                    text="\u00bfEs accesible?",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=1, column=0, sticky="w", padx=8, pady=4)
                accesible = ttk.Combobox(
                    row,
                    values=accesible_options,
                    state="readonly",
                    width=ENTRY_W_MED,
                )
                accesible.grid(row=1, column=1, sticky="w", padx=4, pady=4)
                self.section2_1_fields[field_id]["accesible"] = accesible

                tk.Label(
                    row,
                    text="Observaciones",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=2, column=0, sticky="w", padx=8, pady=4)
                obs = tk.Entry(row, width=80)
                obs.grid(row=2, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_1_fields[field_id]["observaciones"] = obs

            elif question["type"] == "texto":
                entry = tk.Entry(row, width=90)
                entry.grid(row=1, column=0, columnspan=4, sticky="w", padx=8, pady=6)
                self.section2_1_fields[field_id]["texto"] = entry
            elif question["type"] == "lista":
                if question.get("has_accesible"):
                    tk.Label(
                        row,
                        text="\u00bfEs accesible?",
                        font=("Arial", 9, "bold"),
                        bg="white",
                    ).grid(row=1, column=0, sticky="w", padx=8, pady=4)
                    accesible = ttk.Combobox(
                        row,
                        values=accesible_options,
                        state="readonly",
                        width=ENTRY_W_MED,
                    )
                    accesible.grid(row=1, column=1, sticky="w", padx=4, pady=4)
                    self.section2_1_fields[field_id]["accesible"] = accesible

                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=2, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=2, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_1_fields[field_id]["lista"] = combo

        self._prefill_section_fields("section_2_1", self.section2_1_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_1).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2_1).pack(side="right")
    def _show_section_2_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. ACCESIBILIDAD F\u00cdSICA")
        self.header_subtitle.config(text="Completa accesibilidad general.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        title = tk.Label(
            content,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_2_2["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(8, 12))

        self.section2_2_fields = {}
        accesible_options = self._get_accessible_options()

        for question in evaluacion_accesibilidad.SECTION_2_2["questions"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=self._clean_text(question["label"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))

            field_id = question["id"]
            self.section2_2_fields[field_id] = {}
            current_row = 1

            if question.get("has_accesible") or question["type"] == "accesible_con_observaciones":
                tk.Label(
                    row,
                    text="\u00bfEs accesible?",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                accesible = ttk.Combobox(
                    row,
                    values=accesible_options,
                    state="readonly",
                    width=ENTRY_W_MED,
                )
                accesible.grid(row=current_row, column=1, sticky="w", padx=4, pady=4)
                self.section2_2_fields[field_id]["accesible"] = accesible
                current_row += 1

            if question["type"] == "accesible_con_observaciones":
                tk.Label(
                    row,
                    text="Observaciones",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                obs = tk.Entry(row, width=80)
                obs.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_2_fields[field_id]["observaciones"] = obs

            elif question["type"] == "texto":
                tk.Label(
                    row,
                    text="Detalle",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                entry = tk.Entry(row, width=90)
                entry.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_2_fields[field_id]["texto"] = entry

            elif question["type"] == "lista":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_2_fields[field_id]["lista"] = combo
                current_row += 1

                if question.get("has_observaciones"):
                    tk.Label(
                        row,
                        text="Observaciones",
                        font=("Arial", 9, "bold"),
                        bg="white",
                    ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                    obs = tk.Entry(row, width=80)
                    obs.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                    self.section2_2_fields[field_id]["observaciones"] = obs

            elif question["type"] == "lista_doble":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_2_fields[field_id]["lista"] = combo
                current_row += 1

                tk.Label(
                    row,
                    text="Selecci\u00f3n 2",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo_secondary = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options_secondary"]],
                    state="readonly",
                    width=80,
                )
                combo_secondary.grid(
                    row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                )
                self.section2_2_fields[field_id]["lista_secundaria"] = combo_secondary

            elif question["type"] == "lista_triple":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_2_fields[field_id]["lista"] = combo
                current_row += 1

                tk.Label(
                    row,
                    text="Selecci\u00f3n 2",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo_secondary = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options_secondary"]],
                    state="readonly",
                    width=80,
                )
                combo_secondary.grid(
                    row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                )
                self.section2_2_fields[field_id]["lista_secundaria"] = combo_secondary
                current_row += 1

                tk.Label(
                    row,
                    text="Selecci\u00f3n 3",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo_tertiary = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options_tertiary"]],
                    state="readonly",
                    width=80,
                )
                combo_tertiary.grid(
                    row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                )
                self.section2_2_fields[field_id]["lista_terciaria"] = combo_tertiary

        self._prefill_section_fields("section_2_2", self.section2_2_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2_2).pack(side="right")
    def _show_section_2_3(self):
        self._clear_section_container()
        self.header_title.config(text="2. ACCESIBILIDAD F\u00cdSICA")
        self.header_subtitle.config(text="Completa accesibilidad fisica.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        title = tk.Label(
            content,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_2_3["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(8, 12))

        self.section2_3_fields = {}
        accesible_options = self._get_accessible_options()

        for question in evaluacion_accesibilidad.SECTION_2_3["questions"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=self._clean_text(question["label"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))

            field_id = question["id"]
            self.section2_3_fields[field_id] = {}
            current_row = 1

            if question.get("has_accesible") or question["type"] == "accesible_con_observaciones":
                tk.Label(
                    row,
                    text="\u00bfEs accesible?",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                accesible = ttk.Combobox(
                    row,
                    values=accesible_options,
                    state="readonly",
                    width=ENTRY_W_MED,
                )
                accesible.grid(row=current_row, column=1, sticky="w", padx=4, pady=4)
                self.section2_3_fields[field_id]["accesible"] = accesible
                current_row += 1

            if question["type"] == "accesible_con_observaciones":
                tk.Label(
                    row,
                    text="Observaciones",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                obs = tk.Entry(row, width=80)
                obs.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_3_fields[field_id]["observaciones"] = obs

            elif question["type"] == "texto":
                tk.Label(
                    row,
                    text="Detalle",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                entry = tk.Entry(row, width=90)
                entry.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_3_fields[field_id]["texto"] = entry

            elif question["type"] == "lista":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_3_fields[field_id]["lista"] = combo

            elif question["type"] == "lista_doble":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_3_fields[field_id]["lista"] = combo
                current_row += 1

                tk.Label(
                    row,
                    text="Selecci\u00f3n 2",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo_secondary = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options_secondary"]],
                    state="readonly",
                    width=80,
                )
                combo_secondary.grid(
                    row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                )
                self.section2_3_fields[field_id]["lista_secundaria"] = combo_secondary

            elif question["type"] == "lista_triple":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_3_fields[field_id]["lista"] = combo
                current_row += 1

                tk.Label(
                    row,
                    text="Selecci\u00f3n 2",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo_secondary = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options_secondary"]],
                    state="readonly",
                    width=80,
                )
                combo_secondary.grid(
                    row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                )
                self.section2_3_fields[field_id]["lista_secundaria"] = combo_secondary
                current_row += 1

                tk.Label(
                    row,
                    text="Selecci\u00f3n 3",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo_tertiary = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options_tertiary"]],
                    state="readonly",
                    width=80,
                )
                combo_tertiary.grid(
                    row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                )
                self.section2_3_fields[field_id]["lista_terciaria"] = combo_tertiary

            elif question["type"] == "lista_multiple":
                option_sets = [
                    ("lista", "Selecci\u00f3n", question.get("options")),
                    ("lista_secundaria", "Selecci\u00f3n 2", question.get("options_secondary")),
                    ("lista_terciaria", "Selecci\u00f3n 3", question.get("options_tertiary")),
                    ("lista_cuaternaria", "Selecci\u00f3n 4", question.get("options_quaternary")),
                    ("lista_quinta", "Selecci\u00f3n 5", question.get("options_quinary")),
                ]
                for key, label, options in option_sets:
                    if not options:
                        continue
                    tk.Label(
                        row,
                        text=label,
                        font=("Arial", 9, "bold"),
                        bg="white",
                    ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                    combo = ttk.Combobox(
                        row,
                        values=[self._clean_text(opt) for opt in options],
                        state="readonly",
                        width=80,
                    )
                    combo.grid(
                        row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                    )
                    self.section2_3_fields[field_id][key] = combo
                    current_row += 1

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2_2).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2_3).pack(side="right")

    def _show_section_2_4(self):
        self._clear_section_container()
        self.header_title.config(text="2. ACCESIBILIDAD F\u00cdSICA")
        self.header_subtitle.config(text="Completa accesibilidad sensorial.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        title = tk.Label(
            content,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_2_4["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(8, 12))

        self.section2_4_fields = {}
        accesible_options = self._get_accessible_options()

        for question in evaluacion_accesibilidad.SECTION_2_4["questions"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=self._clean_text(question["label"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))

            field_id = question["id"]
            self.section2_4_fields[field_id] = {}
            current_row = 1

            if question.get("has_accesible"):
                tk.Label(
                    row,
                    text="\u00bfEs accesible?",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                accesible = ttk.Combobox(
                    row,
                    values=accesible_options,
                    state="readonly",
                    width=ENTRY_W_MED,
                )
                accesible.grid(row=current_row, column=1, sticky="w", padx=4, pady=4)
                self.section2_4_fields[field_id]["accesible"] = accesible
                current_row += 1

            if question["type"] == "accesible_con_observaciones":
                tk.Label(
                    row,
                    text="Observaciones",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                obs = tk.Entry(row, width=80)
                obs.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_4_fields[field_id]["observaciones"] = obs

            elif question["type"] == "lista":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_4_fields[field_id]["lista"] = combo
                current_row += 1

            elif question["type"] == "lista_multiple":
                option_sets = [
                    ("lista", "Selecci\u00f3n", question.get("options")),
                    ("lista_secundaria", "Selecci\u00f3n 2", question.get("options_secondary")),
                    ("lista_terciaria", "Selecci\u00f3n 3", question.get("options_tertiary")),
                    ("lista_cuaternaria", "Selecci\u00f3n 4", question.get("options_quaternary")),
                    ("lista_quinta", "Selecci\u00f3n 5", question.get("options_quinary")),
                ]
                for key, label, options in option_sets:
                    if not options:
                        continue
                    tk.Label(
                        row,
                        text=label,
                        font=("Arial", 9, "bold"),
                        bg="white",
                    ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                    combo = ttk.Combobox(
                        row,
                        values=[self._clean_text(opt) for opt in options],
                        state="readonly",
                        width=80,
                    )
                    combo.grid(
                        row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                    )
                    self.section2_4_fields[field_id][key] = combo
                    current_row += 1

            if question.get("text_observaciones"):
                tk.Label(
                    row,
                    text="Detalle",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                detail = tk.Entry(row, width=80)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_4_fields[field_id]["detalle"] = detail

        self._prefill_section_fields("section_2_4", self.section2_4_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2_3).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2_4).pack(side="right")

    def _show_section_2_5(self):
        self._clear_section_container()
        self.header_title.config(text="2. ACCESIBILIDAD F\u00cdSICA")
        self.header_subtitle.config(text="Completa accesibilidad intelectual - TEA.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        title = tk.Label(
            content,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_2_5["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(8, 12))

        self.section2_5_fields = {}
        accesible_options = self._get_accessible_options()

        for question in evaluacion_accesibilidad.SECTION_2_5["questions"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=self._clean_text(question["label"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))

            field_id = question["id"]
            self.section2_5_fields[field_id] = {}
            current_row = 1

            if question.get("has_accesible"):
                tk.Label(
                    row,
                    text="\u00bfEs accesible?",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                accesible = ttk.Combobox(
                    row,
                    values=accesible_options,
                    state="readonly",
                    width=ENTRY_W_MED,
                )
                accesible.grid(row=current_row, column=1, sticky="w", padx=4, pady=4)
                self.section2_5_fields[field_id]["accesible"] = accesible
                current_row += 1

            if question["type"] == "accesible_con_observaciones":
                tk.Label(
                    row,
                    text="Observaciones",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                obs = tk.Entry(row, width=80)
                obs.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_5_fields[field_id]["observaciones"] = obs

            elif question["type"] == "lista":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_5_fields[field_id]["lista"] = combo
                current_row += 1

            elif question["type"] == "lista_multiple":
                option_sets = [
                    ("lista", "Selecci\u00f3n", question.get("options")),
                    ("lista_secundaria", "Selecci\u00f3n 2", question.get("options_secondary")),
                    ("lista_terciaria", "Selecci\u00f3n 3", question.get("options_tertiary")),
                    ("lista_cuaternaria", "Selecci\u00f3n 4", question.get("options_quaternary")),
                    ("lista_quinta", "Selecci\u00f3n 5", question.get("options_quinary")),
                ]
                for key, label, options in option_sets:
                    if not options:
                        continue
                    tk.Label(
                        row,
                        text=label,
                        font=("Arial", 9, "bold"),
                        bg="white",
                    ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                    combo = ttk.Combobox(
                        row,
                        values=[self._clean_text(opt) for opt in options],
                        state="readonly",
                        width=80,
                    )
                    combo.grid(
                        row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                    )
                    self.section2_5_fields[field_id][key] = combo
                    current_row += 1

            if question.get("text_observaciones"):
                tk.Label(
                    row,
                    text="Detalle",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                detail = tk.Entry(row, width=80)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_5_fields[field_id]["detalle"] = detail

        self._prefill_section_fields("section_2_5", self.section2_5_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2_4).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2_5).pack(side="right")

    def _show_section_2_6(self):
        self._clear_section_container()
        self.header_title.config(text="2. ACCESIBILIDAD F\u00cdSICA")
        self.header_subtitle.config(text="Completa accesibilidad psicosocial.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        title = tk.Label(
            content,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_2_6["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(8, 12))

        self.section2_6_fields = {}
        accesible_options = self._get_accessible_options()

        for question in evaluacion_accesibilidad.SECTION_2_6["questions"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=self._clean_text(question["label"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))

            field_id = question["id"]
            self.section2_6_fields[field_id] = {}
            current_row = 1

            tk.Label(
                row,
                text="\u00bfEs accesible?",
                font=("Arial", 9, "bold"),
                bg="white",
            ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
            accesible = ttk.Combobox(
                row,
                values=accesible_options,
                state="readonly",
                width=ENTRY_W_MED,
            )
            accesible.grid(row=current_row, column=1, sticky="w", padx=4, pady=4)
            self.section2_6_fields[field_id]["accesible"] = accesible
            current_row += 1

            tk.Label(
                row,
                text="Selecci\u00f3n",
                font=("Arial", 9, "bold"),
                bg="white",
            ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
            combo = ttk.Combobox(
                row,
                values=[self._clean_text(opt) for opt in question["options"]],
                state="readonly",
                width=80,
            )
            combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
            self.section2_6_fields[field_id]["lista"] = combo
            current_row += 1

            if question.get("text_observaciones"):
                tk.Label(
                    row,
                    text="Detalle",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                detail = tk.Entry(row, width=80)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section2_6_fields[field_id]["detalle"] = detail

        self._prefill_section_fields("section_2_6", self.section2_6_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2_5).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2_6).pack(side="right")


    def _show_section_3(self):
        self._clear_section_container()
        self.header_title.config(text="3. CONDICIONES ORGANIZACIONALES")
        self.header_subtitle.config(text="Completa condiciones organizacionales.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        title = tk.Label(
            content,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_3["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(8, 12))

        self.section3_fields = {}
        default_accesible_options = self._get_accessible_options()

        for question in evaluacion_accesibilidad.SECTION_3["questions"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=self._clean_text(question["label"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))

            field_id = question["id"]
            self.section3_fields[field_id] = {}
            current_row = 1

            if question.get("has_accesible"):
                accesible_values = default_accesible_options
                tk.Label(
                    row,
                    text="\u00bfEs accesible?",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                accesible = ttk.Combobox(
                    row,
                    values=accesible_values,
                    state="readonly",
                    width=ENTRY_W_MED,
                )
                accesible.grid(row=current_row, column=1, sticky="w", padx=4, pady=4)
                self.section3_fields[field_id]["accesible"] = accesible
                current_row += 1

            if question["type"] == "accesible_con_observaciones":
                tk.Label(
                    row,
                    text="Observaciones",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                obs = tk.Entry(row, width=80)
                obs.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section3_fields[field_id]["observaciones"] = obs

            elif question["type"] == "lista":
                tk.Label(
                    row,
                    text="Selecci\u00f3n",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                combo = ttk.Combobox(
                    row,
                    values=[self._clean_text(opt) for opt in question["options"]],
                    state="readonly",
                    width=80,
                )
                combo.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section3_fields[field_id]["lista"] = combo
                current_row += 1

            elif question["type"] == "lista_multiple":
                option_sets = [
                    ("lista", "Selecci\u00f3n", question.get("options")),
                    ("lista_secundaria", "Selecci\u00f3n 2", question.get("options_secondary")),
                    ("lista_terciaria", "Selecci\u00f3n 3", question.get("options_tertiary")),
                    ("lista_cuaternaria", "Selecci\u00f3n 4", question.get("options_quaternary")),
                    ("lista_quinta", "Selecci\u00f3n 5", question.get("options_quinary")),
                ]
                for key, label, options in option_sets:
                    if not options:
                        continue
                    tk.Label(
                        row,
                        text=label,
                        font=("Arial", 9, "bold"),
                        bg="white",
                    ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                    combo = ttk.Combobox(
                        row,
                        values=[self._clean_text(opt) for opt in options],
                        state="readonly",
                        width=80,
                    )
                    combo.grid(
                        row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4
                    )
                    self.section3_fields[field_id][key] = combo
                    current_row += 1

            if question.get("text_observaciones"):
                tk.Label(
                    row,
                    text="Detalle",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                detail = tk.Entry(row, width=80)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="w", padx=4, pady=4)
                self.section3_fields[field_id]["detalle"] = detail

        self._prefill_section_fields("section_3", self.section3_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2_6).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_3).pack(side="right")
    def _confirm_section_2_5(self):
        payload = self._collect_section_fields(self.section2_5_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2_6()

    def _confirm_section_2_6(self):
        payload = self._collect_section_fields(self.section2_6_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_3()


    def _show_section_4(self):
        self._clear_section_container()
        self.header_title.config(text="4. CONCEPTO DE LA EVALUACION")
        self.header_subtitle.config(text="Resume el nivel de accesibilidad.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        counts, percentages, suggestion = self._calculate_accessible_summary()

        summary_frame = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        summary_frame.pack(fill="x", pady=(8, 16))

        tk.Label(
            summary_frame,
            text="Resumen de respuestas",
            font=FONT_SECTION,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        tk.Label(summary_frame, text="Respuesta", font=("Arial", 9, "bold"), bg=COLOR_LIGHT_BG).grid(
            row=1, column=0, sticky="w", padx=(0, 12)
        )
        tk.Label(summary_frame, text="Cantidad", font=("Arial", 9, "bold"), bg=COLOR_LIGHT_BG).grid(
            row=1, column=1, sticky="w", padx=(0, 12)
        )
        tk.Label(summary_frame, text="Porcentaje", font=("Arial", 9, "bold"), bg=COLOR_LIGHT_BG).grid(
            row=1, column=2, sticky="w"
        )

        rows = [("Si", "si"), ("No", "no"), ("Parcial", "parcial")]
        for idx, (label, key) in enumerate(rows, start=2):
            tk.Label(summary_frame, text=label, bg=COLOR_LIGHT_BG).grid(row=idx, column=0, sticky="w")
            tk.Label(summary_frame, text=str(counts[key]), bg=COLOR_LIGHT_BG).grid(
                row=idx, column=1, sticky="w"
            )
            tk.Label(summary_frame, text=f"{percentages[key]:.1f}%", bg=COLOR_LIGHT_BG).grid(
                row=idx, column=2, sticky="w"
            )

        tk.Label(
            summary_frame,
            text=f"Sugerido: {suggestion or 'Sin datos'}",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 0))

        selector_frame = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        selector_frame.pack(fill="x", pady=(8, 12))

        tk.Label(
            selector_frame,
            text="Nivel de accesibilidad",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.section4_level_var = tk.StringVar()
        level_combo = ttk.Combobox(
            selector_frame,
            textvariable=self.section4_level_var,
            values=evaluacion_accesibilidad.SECTION_4["options"],
            state="readonly",
            width=ENTRY_W_MED,
        )
        level_combo.grid(row=0, column=1, sticky="w")

        cached_level = evaluacion_accesibilidad.get_form_cache().get("section_4", {}).get("nivel_accesibilidad")
        if cached_level:
            self.section4_level_var.set(cached_level)
        elif suggestion:
            self.section4_level_var.set(suggestion)
        else:
            self.section4_level_var.set("")

        level_combo.bind("<<ComboboxSelected>>", self._update_section4_description)

        tk.Label(
            selector_frame,
            text="Descripción",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=1, column=0, sticky="nw", pady=(12, 0))

        self.section4_desc = tk.Text(selector_frame, height=6, wrap="word")
        self.section4_desc.grid(row=1, column=1, sticky="w", pady=(12, 0))
        self.section4_desc.configure(state="disabled")
        self._update_section4_description()

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_3).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_4).pack(side="right")
    def _confirm_section_3(self):
        payload = self._collect_section_fields(self.section3_fields)
        try:
            evaluacion_accesibilidad.confirm_section_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_4()


    def _normalize_accessible_value(self, value):
        if value is None:
            return ""
        value = value.strip().lower()
        if not value:
            return ""
        value = value.replace("?", "a").replace("?", "e").replace("?", "i").replace("?", "o").replace("?", "u")
        value = value.replace("í", "i")
        return value

    def _calculate_accessible_summary(self):
        cache = evaluacion_accesibilidad.get_form_cache()
        counts = {"si": 0, "no": 0, "parcial": 0}
        sections = [
            "section_2_1",
            "section_2_2",
            "section_2_3",
            "section_2_4",
            "section_2_5",
            "section_2_6",
            "section_3",
        ]
        for section_id in sections:
            section = cache.get(section_id, {})
            for key, value in section.items():
                if not key.endswith("_accesible"):
                    continue
                normalized = self._normalize_accessible_value(str(value))
                if not normalized:
                    continue
                if normalized == "si":
                    counts["si"] += 1
                elif normalized == "no":
                    counts["no"] += 1
                elif normalized == "parcial":
                    counts["parcial"] += 1
        total = counts["si"] + counts["no"] + counts["parcial"]
        percentages = {
            key: (counts[key] / total * 100) if total else 0
            for key in counts
        }
        suggestion = ""
        if total:
            si_pct = percentages["si"]
            if si_pct >= 86:
                suggestion = "Alto"
            elif si_pct >= 51:
                suggestion = "Medio"
            elif si_pct >= 1:
                suggestion = "Bajo"
        return counts, percentages, suggestion

    def _update_section4_description(self, *_):
        nivel = self.section4_level_var.get()
        descripcion = evaluacion_accesibilidad.SECTION_4["descriptions"].get(nivel, "")
        self.section4_desc.configure(state="normal")
        self.section4_desc.delete("1.0", tk.END)
        self.section4_desc.insert("1.0", descripcion)
        self.section4_desc.configure(state="disabled")


    def _confirm_section_4(self):
        nivel = self.section4_level_var.get().strip()
        descripcion = evaluacion_accesibilidad.SECTION_4["descriptions"].get(nivel, "")
        payload = {
            "nivel_accesibilidad": nivel,
            "descripcion": descripcion,
        }
        try:
            evaluacion_accesibilidad.confirm_section_4(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_5()


    def _show_section_5(self):
        self._clear_section_container()
        self.header_title.config(text="5. AJUSTES RAZONABLES")
        self.header_subtitle.config(text="Marca aplicacion y registra notas.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        title = tk.Label(
            content,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_5["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        title.pack(anchor="w", pady=(8, 12))

        self.section5_fields = {}
        aplica_options = [
            self._clean_text(option)
            for option in evaluacion_accesibilidad.SECTION_5["aplica_options"]
        ]

        for item in evaluacion_accesibilidad.SECTION_5["items"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=self._clean_text(item["label"]),
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, columnspan=3, sticky="w", padx=8, pady=(8, 2))

            tk.Label(
                row,
                text=self._clean_text(item["codes"]),
                font=("Arial", 9),
                bg="white",
                fg="#555555",
                wraplength=760,
                justify="left",
            ).grid(row=1, column=0, columnspan=3, sticky="w", padx=8)

            tk.Label(
                row,
                text="Ajustes sugeridos:",
                font=("Arial", 9, "bold"),
                bg="white",
            ).grid(row=2, column=0, sticky="w", padx=8, pady=(6, 2))
            tk.Label(
                row,
                text=self._clean_text(item["ajustes"]),
                font=("Arial", 9),
                bg="white",
                fg="#333333",
                wraplength=760,
                justify="left",
            ).grid(row=3, column=0, columnspan=3, sticky="w", padx=8)

            tk.Label(
                row,
                text="Aplica",
                font=("Arial", 9, "bold"),
                bg="white",
            ).grid(row=4, column=0, sticky="w", padx=8, pady=(8, 4))
            aplica_combo = ttk.Combobox(
                row,
                values=aplica_options,
                state="readonly",
                width=ENTRY_W_MED,
            )
            aplica_combo.grid(row=4, column=1, sticky="w", padx=4, pady=(8, 4))
            aplica_combo.set("No aplica")

            tk.Label(
                row,
                text="Nota:",
                font=("Arial", 9, "bold"),
                bg="white",
            ).grid(row=5, column=0, sticky="nw", padx=8, pady=(4, 8))
            nota_entry = tk.Entry(row, width=90)
            nota_entry.grid(row=5, column=1, columnspan=2, sticky="w", padx=4, pady=(4, 8))

            self.section5_fields[item["id"]] = {
                "lista": aplica_combo,
                "nota": nota_entry,
            }

        self._prefill_section_fields("section_5", self.section5_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_4).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_5).pack(side="right")

    def _confirm_section_5(self):
        payload = self._collect_section_fields(self.section5_fields)
        try:
            evaluacion_accesibilidad.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_6()


    def _show_section_6(self):
        self._clear_section_container()
        self.header_title.config(text="6. OBSERVACIONES")
        self.header_subtitle.config(text="Registra observaciones generales.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        tk.Label(
            section_frame,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_6["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(8, 8))

        self.section6_fields = {}
        for field in evaluacion_accesibilidad.SECTION_6["fields"]:
            tk.Label(
                section_frame,
                text=field["label"],
                font=FONT_LABEL,
                bg=COLOR_LIGHT_BG,
            ).pack(anchor="w", pady=(6, 2))
            text_box = tk.Text(section_frame, height=4, wrap="word")
            text_box.pack(fill="x", pady=(0, 12))
            self.section6_fields[field["id"]] = {"texto": text_box}

        self._prefill_section_fields("section_6", self.section6_fields)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_5).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_6).pack(side="right")

    def _confirm_section_6(self):
        payload = self._collect_section_fields(self.section6_fields)
        try:
            evaluacion_accesibilidad.confirm_section_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_7()

    def _show_section_7(self):
        self._clear_section_container()
        self.header_title.config(text="7. CARGOS COMPATIBLES")
        self.header_subtitle.config(text="Registra cargos compatibles.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        tk.Label(
            section_frame,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_7["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(8, 8))

        for line in evaluacion_accesibilidad.SECTION_7.get("instructions", []):
            tk.Label(
                section_frame,
                text=self._clean_text(line),
                font=("Arial", 9),
                bg=COLOR_LIGHT_BG,
                fg="#333333",
                wraplength=760,
                justify="left",
            ).pack(anchor="w")

        self.section7_fields = {}
        for field in evaluacion_accesibilidad.SECTION_7["fields"]:
            tk.Label(
                section_frame,
                text=field["label"],
                font=FONT_LABEL,
                bg=COLOR_LIGHT_BG,
            ).pack(anchor="w", pady=(8, 2))
            text_box = tk.Text(section_frame, height=4, wrap="word")
            text_box.pack(fill="x", pady=(0, 12))
            self.section7_fields[field["id"]] = {"texto": text_box}

        self._prefill_section_fields("section_7", self.section7_fields)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_6).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_7).pack(side="right")

    def _confirm_section_7(self):
        payload = self._collect_section_fields(self.section7_fields)
        try:
            evaluacion_accesibilidad.confirm_section_7(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_8()

    def _show_section_8(self):
        self._clear_section_container()
        self.header_title.config(text="8. ASISTENTES")
        self.header_subtitle.config(text="Registra asistentes.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        tk.Label(
            section_frame,
            text=self._clean_text(evaluacion_accesibilidad.SECTION_8["title"]),
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(8, 8))

        table = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        table.pack(fill="x", pady=(4, 8))

        tk.Label(
            table,
            text="Nombre completo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        tk.Label(
            table,
            text="Cargo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=1, sticky="w", padx=(0, 12))

        self.section8_entries = []
        self.section8_table = table
        self.add_asistente_btn = ttk.Button(
            table,
            text="Agregar asistente",
            command=self._add_section8_asistente_row,
        )
        self._asesores_agencia_catalog = _get_asesores_agencia_catalog()

        cached = evaluacion_accesibilidad.get_form_cache().get("section_8", [])
        self._render_section8_asistentes(cached)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_7).pack(side="left")
        ttk.Button(actions, text="Finalizar", command=self._confirm_section_8).pack(side="right")

    def _get_section8_asistentes_values(self):
        values = []
        for name_widget, role_widget in self.section8_entries:
            values.append(
                {
                    "nombre": _get_input_value(name_widget),
                    "cargo": _get_input_value(role_widget),
                }
            )
        return values

    def _render_section8_asistentes(self, values=None):
        rows = list(values or [])
        min_rows = evaluacion_accesibilidad.EXCEL_MAPPING.get("section_8", {}).get("base_rows", 4)
        while len(rows) < min_rows:
            rows.append({"nombre": "", "cargo": ""})

        for name_widget, role_widget in self.section8_entries:
            name_widget.destroy()
            role_widget.destroy()
        self.section8_entries = []

        for idx, entry in enumerate(rows):
            row_idx = idx + 1
            is_first = idx == 0
            is_last = idx == len(rows) - 1
            if is_first and not is_last:
                name_widget, role_widget = _create_asistente_inputs(
                    self.section8_table,
                    ENTRY_W_WIDE,
                    use_catalog=True,
                    catalog=_get_asistentes_profesionales_catalog(),
                )
            elif is_last:
                name_widget, role_widget = _create_asesor_agencia_inputs(
                    self.section8_table,
                    ENTRY_W_WIDE,
                    catalog=getattr(self, "_asesores_agencia_catalog", None),
                )
            else:
                name_widget, role_widget = _create_asistente_inputs(
                    self.section8_table,
                    ENTRY_W_WIDE,
                    use_catalog=False,
                )
            name_widget.grid(row=row_idx, column=0, sticky="w", padx=(0, 12), pady=4)
            role_widget.grid(row=row_idx, column=1, sticky="w", padx=(0, 12), pady=4)
            _set_input_value(name_widget, entry.get("nombre", ""))
            cargo_value = entry.get("cargo", "")
            if is_last and not cargo_value:
                cargo_value = "Asesor Agencia"
            _set_input_value(role_widget, cargo_value)
            self.section8_entries.append((name_widget, role_widget))

        self.add_asistente_btn.grid(row=len(self.section8_entries) + 1, column=0, sticky="w", pady=(8, 0))

    def _add_section8_asistente_row(self):
        if len(self.section8_entries) >= evaluacion_accesibilidad.SECTION_8["max_items"]:
            messagebox.showinfo("Asistentes", "Máximo de asistentes alcanzado.")
            return
        rows = self._get_section8_asistentes_values()
        if rows:
            rows[-1] = {"nombre": "", "cargo": ""}
        rows.append({"nombre": "", "cargo": ""})
        self._render_section8_asistentes(rows)

    def _confirm_section_8(self):
        asistentes = []
        for name_widget, role_widget in self.section8_entries:
            nombre = name_widget.get().strip()
            cargo = role_widget.get().strip()
            if not nombre and not cargo:
                continue
            asistentes.append({"nombre": nombre, "cargo": cargo})
        try:
            evaluacion_accesibilidad.confirm_section_8(asistentes)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando exportación...")
        loading.set_progress(5)

        cache_snapshot = evaluacion_accesibilidad.get_form_cache()
        section_order = list(evaluacion_accesibilidad.EXCEL_MAPPING.keys())
        total_steps = len(section_order) or 1
        cache = cache_snapshot
        section_1 = cache.get("section_1", {})
        company_name = section_1.get("nombre_empresa")

        def _worker():
            def _on_progress(section_id):
                try:
                    idx = section_order.index(section_id) + 1
                except ValueError:
                    idx = 1
                _update_loading_async(
                    loading,
                    status=f"Guardando {section_id.replace('_', ' ')}...",
                    progress=5 + int((idx / total_steps) * 90),
                )

            reviewed_cache_snapshot = evaluacion_accesibilidad.get_form_cache()
            sheet_export = evaluacion_accesibilidad.build_google_sheet_export_payload(
                reviewed_cache_snapshot
            )
            output_path = _raise_finalize_stage(
                "guardando el Excel",
                lambda: evaluacion_accesibilidad.export_to_excel(progress_callback=_on_progress),
            )
            return {
                "output_path": output_path,
                "drive_job": {
                    "upload_kind": "evaluacion_sheet",
                    "remote_file_name": evaluacion_accesibilidad.build_output_base_name(cache_snapshot),
                    "sheet_export": sheet_export,
                },
            }

        _start_background_finalization(
            self,
            loading,
            form_name="Evaluacion Accesibilidad",
            company_name=company_name,
            form_id="evaluacion_accesibilidad",
            worker_fn=_worker,
        )


    def _set_widget_value(self, widget, value):
        if value is None:
            value = ""
        if isinstance(widget, ttk.Combobox):
            widget.set(value)
        elif isinstance(widget, tk.Text):
            widget.delete("1.0", tk.END)
            widget.insert("1.0", value)
        else:
            widget.delete(0, tk.END)
            widget.insert(0, value)


    def _prefill_section_fields(self, section_id, fields):
        cache = evaluacion_accesibilidad.get_form_cache().get(section_id, {})
        for field_id, widgets in fields.items():
            for key, widget in widgets.items():
                if key == "accesible":
                    cache_key = f"{field_id}_accesible"
                elif key == "observaciones":
                    cache_key = f"{field_id}_observaciones"
                elif key == "texto":
                    cache_key = field_id
                elif key == "lista":
                    cache_key = field_id
                elif key == "lista_secundaria":
                    cache_key = f"{field_id}_secundaria"
                elif key == "lista_terciaria":
                    cache_key = f"{field_id}_terciaria"
                elif key == "lista_cuaternaria":
                    cache_key = f"{field_id}_cuaternaria"
                elif key == "lista_quinta":
                    cache_key = f"{field_id}_quinary"
                elif key == "detalle":
                    cache_key = f"{field_id}_detalle"
                elif key == "nota":
                    cache_key = f"{field_id}_nota"
                else:
                    continue
                value = cache.get(cache_key, "")
                if key == "nota" and isinstance(value, str):
                    if value.lower().startswith("nota:"):
                        value = value[5:].lstrip()
                self._set_widget_value(widget, value)

    def _collect_section_fields(self, fields):
        payload = {}
        for field_id, widgets in fields.items():
            for key, widget in widgets.items():
                if isinstance(widget, ttk.Combobox):
                    value = widget.get().strip()
                elif isinstance(widget, tk.Text):
                    value = widget.get("1.0", tk.END).strip()
                else:
                    value = widget.get().strip()
                if key == "accesible":
                    payload[f"{field_id}_accesible"] = value
                elif key == "observaciones":
                    payload[f"{field_id}_observaciones"] = value
                elif key == "texto":
                    payload[field_id] = value
                elif key == "lista":
                    payload[field_id] = value
                elif key == "lista_secundaria":
                    payload[f"{field_id}_secundaria"] = value
                elif key == "lista_terciaria":
                    payload[f"{field_id}_terciaria"] = value
                elif key == "lista_cuaternaria":
                    payload[f"{field_id}_cuaternaria"] = value
                elif key == "lista_quinta":
                    payload[f"{field_id}_quinary"] = value
                elif key == "detalle":
                    payload[f"{field_id}_detalle"] = value
                elif key == "nota":
                    nota_value = value
                    if nota_value and not nota_value.startswith("Nota:"):
                        nota_value = f"Nota: {nota_value}"
                    elif nota_value == "":
                        nota_value = "Nota: "
                    payload[f"{field_id}_nota"] = nota_value
        return payload

    def _confirm_section_2_1(self):
        payload = self._collect_section_fields(self.section2_1_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_1(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2_2()

    def _confirm_section_2_2(self):
        payload = self._collect_section_fields(self.section2_2_fields)
        if "evaluacion_ergonomica_puestos" not in payload:
            widgets = self.section2_2_fields.get("evaluacion_ergonomica_puestos", {})
            combo = widgets.get("lista")
            if combo:
                payload["evaluacion_ergonomica_puestos"] = combo.get().strip()
            accesible_combo = widgets.get("accesible")
            if accesible_combo and "evaluacion_ergonomica_puestos_accesible" not in payload:
                payload["evaluacion_ergonomica_puestos_accesible"] = (
                    accesible_combo.get().strip()
                )
        try:
            evaluacion_accesibilidad.confirm_section_2_2(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2_3()

    def _confirm_section_2_3(self):
        payload = self._collect_section_fields(self.section2_3_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2_4()

    def _confirm_section_2_4(self):
        payload = self._collect_section_fields(self.section2_4_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_4(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2_5()

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        groups = [
            ('Información de Empresa', COLOR_GROUP_EMPRESA, ['nombre_empresa', 'direccion_empresa', 'correo_1', 'contacto_empresa', 'telefono_empresa', 'cargo', 'ciudad_empresa', 'sede_empresa', 'caja_compensacion']),
            ('Información de Compensar', COLOR_GROUP_COMPENSAR, ['asesor']),
            ('Información de RECA', COLOR_GROUP_RECA, ['profesional_asignado']),
        ]
        labels = {
            'nombre_empresa': 'Nombre de la empresa',
            'direccion_empresa': 'Dirección de la empresa',
            'correo_1': 'Correo electrónico',
            'contacto_empresa': 'Contacto de la empresa',
            'telefono_empresa': 'Teléfonos',
            'cargo': 'Cargo',
            'ciudad_empresa': 'Ciudad/Municipio',
            'sede_empresa': 'Sede Compensar',
            'caja_compensacion': 'Empresa afiliada a Caja de Compensación',
            'asesor': 'Asesor',
            'profesional_asignado': 'Profesional asignado RECA',
        }
        _section1_build_groups(
            self,
            parent,
            groups,
            labels,
            modalidad_options=["Presencial", "Virtual", "Mixta", "No aplica"],
        )

    def _build_actions(self, parent):
        _section1_build_actions(self, parent)

    def _label_for_field(self, field_id):
        return getattr(self, '_section1_labels', {}).get(field_id, field_id)

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _search_company(self, mode="nit"):
        nit = self.fields["nit_empresa"].get().strip()
        nombre = self.fields.get("nombre_busqueda").get().strip() if self.fields.get("nombre_busqueda") else ""
        if mode == "nit":
            if not nit:
                messagebox.showerror("Error", "Ingresa un NIT.")
                return
        elif mode == "nombre":
            if not nombre:
                messagebox.showerror("Error", "Ingresa el nombre de la empresa.")
                return
        else:
            messagebox.showerror("Error", "Tipo de búsqueda no válido.")
            return

        try:
            self.status_label.config(text="Buscando empresa...")
            self.update_idletasks()
            if mode == "nombre":
                company = evaluacion_accesibilidad.get_empresa_by_nombre(nombre)
            else:
                company = evaluacion_accesibilidad.get_empresa_by_nit(nit)
        except Exception as exc:
            self.status_label.config(text="")
            messagebox.showerror("Error", str(exc))
            return

        if not company:
            self.company_data = None
            msg = "No se encontró empresa para ese nombre." if mode == "nombre" else "No se encontró empresa para ese NIT."
            self.status_label.config(text=msg)
            self.continue_btn.config(state="disabled")
            for key in evaluacion_accesibilidad.SECTION_1_SUPABASE_MAP.keys():
                self._set_readonly_value(key, "")
            return

        if mode == "nombre":
            nit_value = company.get("nit_empresa")
            if nit_value:
                entry = self.fields.get("nit_empresa")
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, nit_value)

        self.company_data = company
        self.status_label.config(text="Empresa encontrada.")
        self.continue_btn.config(state="normal")
        for key in evaluacion_accesibilidad.SECTION_1_SUPABASE_MAP.keys():
            self._set_readonly_value(key, company.get(key))

    def _confirm_and_continue(self):
        if not self.company_data:
            messagebox.showerror("Error", "Busca una empresa antes de confirmar.")
            return

        fecha_visita = _get_required_fecha_visita(self)
        if not fecha_visita:
            return
        modalidad = _get_required_modalidad(self)
        if not modalidad:
            return
        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad": modalidad,
            "nit_empresa": self.fields["nit_empresa"].get().strip(),
        }
        try:
            evaluacion_accesibilidad.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2()

class CondicionesVacanteWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Condiciones de Vacante - Seccion 1")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = condiciones_vacante

        self.company_data = None
        self.fields = {}

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))

        self.header_title = tk.Label(
            header,
            text="1. DATOS GENERALES",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        self.header_title.pack(anchor="w")

        self.header_subtitle = tk.Label(
            header,
            text="Busca empresa por NIT y confirma datos.",
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        self.header_subtitle.pack(anchor="w", pady=(4, 0))

    def _build_section_container(self):
        self.section_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)

    def _clear_section_container(self):
        for child in self.section_container.winfo_children():
            child.destroy()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "condiciones_vacante",
            condiciones_vacante,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_2_1": self._show_section_2_1,
                "section_3": self._show_section_3,
                "section_4": self._show_section_4,
                "section_5": self._show_section_5,
                "section_6": self._show_section_6,
                "section_7": self._show_section_7,
                "section_8": self._show_section_8,
            },
            self._show_section_1,
        ):
            return True
        if not condiciones_vacante.cache_file_exists():
            return False
        resume = messagebox.askyesno(
            "Reanudar",
            "Se encontró un formulario en progreso. ¿Deseas continuar donde lo dejaste?",
        )
        if not resume:
            condiciones_vacante.clear_cache_file()
            condiciones_vacante.clear_form_cache()
            return False
        condiciones_vacante.load_cache_from_file()
        last_section = condiciones_vacante.get_form_cache().get("_last_section")
        if last_section == "section_1":
            self._show_section_2()
        elif last_section == "section_2":
            self._show_section_2_1()
        elif last_section == "section_2_1":
            self._show_section_3()
        elif last_section == "section_3":
            self._show_section_4()
        elif last_section == "section_4":
            self._show_section_5()
        elif last_section == "section_5":
            self._show_section_6()
        elif last_section == "section_6":
            self._show_section_7()
        elif last_section in {"section_7", "section_8"}:
            self._show_section_8()
        else:
            self._show_section_1()
        return True

    def _show_section_1(self):
        self._clear_section_container()
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._build_actions(content)

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        groups = [
            ('Información de Empresa', COLOR_GROUP_EMPRESA, ['nombre_empresa', 'direccion_empresa', 'correo_1', 'contacto_empresa', 'telefono_empresa', 'cargo', 'ciudad_empresa', 'sede_empresa', 'caja_compensacion']),
            ('Información de Compensar', COLOR_GROUP_COMPENSAR, ['asesor']),
            ('Información de RECA', COLOR_GROUP_RECA, ['profesional_asignado']),
        ]
        labels = {
            'nombre_empresa': 'Nombre de la empresa',
            'direccion_empresa': 'Dirección de la empresa',
            'correo_1': 'Correo electrónico',
            'contacto_empresa': 'Contacto de la empresa',
            'telefono_empresa': 'Teléfonos',
            'cargo': 'Cargo',
            'ciudad_empresa': 'Ciudad/Municipio',
            'sede_empresa': 'Sede Compensar',
            'caja_compensacion': 'Empresa afiliada a Caja de Compensación',
            'asesor': 'Asesor',
            'profesional_asignado': 'Profesional asignado RECA',
        }
        _section1_build_groups(self, parent, groups, labels)

    def _build_actions(self, parent):
        _section1_build_actions(self, parent)

    def _label_for_field(self, field_id):
        return getattr(self, '_section1_labels', {}).get(field_id, field_id)

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _search_company(self, mode="nit"):
        nit = self.fields["nit_empresa"].get().strip()
        nombre = self.fields.get("nombre_busqueda").get().strip() if self.fields.get("nombre_busqueda") else ""
        if mode == "nit":
            if not nit:
                messagebox.showerror("Error", "Ingresa un NIT.")
                return
        elif mode == "nombre":
            if not nombre:
                messagebox.showerror("Error", "Ingresa el nombre de la empresa.")
                return
        else:
            messagebox.showerror("Error", "Tipo de búsqueda no válido.")
            return

        try:
            self.status_label.config(text="Buscando empresa...")
            self.update_idletasks()
            if mode == "nombre":
                company = condiciones_vacante.get_empresa_by_nombre(nombre)
            else:
                company = condiciones_vacante.get_empresa_by_nit(nit)
        except Exception as exc:
            self.status_label.config(text="")
            messagebox.showerror("Error", str(exc))
            return

        if not company:
            self.company_data = None
            msg = "No se encontró empresa para ese nombre." if mode == "nombre" else "No se encontró empresa para ese NIT."
            self.status_label.config(text=msg)
            self.continue_btn.config(state="disabled")
            for key in condiciones_vacante.SECTION_1_SUPABASE_MAP.keys():
                self._set_readonly_value(key, "")
            return

        if mode == "nombre":
            nit_value = company.get("nit_empresa")
            if nit_value:
                entry = self.fields.get("nit_empresa")
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, nit_value)

        self.company_data = company
        self.status_label.config(text="Empresa encontrada.")
        self.continue_btn.config(state="normal")
        for key in condiciones_vacante.SECTION_1_SUPABASE_MAP.keys():
            self._set_readonly_value(key, company.get(key))

    def _confirm_and_continue(self):
        if not self.company_data:
            messagebox.showerror("Error", "Busca una empresa antes de confirmar.")
            return

        fecha_visita = _get_required_fecha_visita(self)
        if not fecha_visita:
            return
        modalidad = _get_required_modalidad(self)
        if not modalidad:
            return
        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad": modalidad,
            "nit_empresa": self.fields["nit_empresa"].get().strip(),
        }
        try:
            condiciones_vacante.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2()

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_2["title"])
        self.header_subtitle.config(text="Completa caracteristicas de la vacante.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        self.section2_fields = {}
        for field in condiciones_vacante.SECTION_2["fields"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=field["label"],
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)

            if field["type"] == "lista":
                widget = ttk.Combobox(
                    row,
                    values=field["options"],
                    state="readonly",
                    width=45,
                )
                widget.grid(row=0, column=1, sticky="w", padx=8, pady=8)
            else:
                widget = tk.Entry(row, width=48)
                widget.grid(row=0, column=1, sticky="w", padx=8, pady=8)

            self.section2_fields[field["id"]] = widget

            if field["id"] == "requiere_certificado":
                obs_row = tk.Frame(content, bg="white", bd=1, relief="solid")
                obs_row.pack(fill="x", pady=(0, 6))
                obs_row.grid_columnconfigure(1, weight=1)
                tk.Label(
                    obs_row,
                    text="Observaciones (Requiere certificado)",
                    font=("Arial", 9, "bold"),
                    bg="white",
                    fg="#222222",
                ).grid(row=0, column=0, sticky="w", padx=8, pady=6)
                obs_text = tk.Text(obs_row, height=3, wrap="word")
                obs_text.grid(row=0, column=1, sticky="we", padx=8, pady=6)
                self.section2_fields["requiere_certificado_observaciones"] = obs_text

        competencias_frame = tk.Frame(content, bg=COLOR_LIGHT_BG)
        competencias_frame.pack(fill="x", pady=(12, 16))
        tk.Label(
            competencias_frame,
            text="Competencias (auto-populadas)",
            font=FONT_LABEL,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 8))

        competencias_entries = tk.Frame(competencias_frame, bg=COLOR_LIGHT_BG)
        competencias_entries.pack(fill="x")

        self.competencia_entries = []
        for idx in range(8):
            entry = tk.Entry(competencias_entries, width=52, state="readonly")
            entry.grid(row=idx // 2, column=idx % 2, padx=8, pady=4, sticky="w")
            self.competencia_entries.append(entry)

        nivel_widget = self.section2_fields.get("nivel_cargo")
        if isinstance(nivel_widget, ttk.Combobox):
            nivel_widget.bind("<<ComboboxSelected>>", self._on_nivel_cargo_change)

        self._prefill_section2_fields()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_1).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2).pack(side="right")

    def _prefill_section2_fields(self):
        cache = condiciones_vacante.get_form_cache().get("section_2", {})
        for field_id, widget in self.section2_fields.items():
            value = cache.get(field_id, "")
            if isinstance(widget, ttk.Combobox):
                widget.set(value)
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", value)
            else:
                widget.delete(0, tk.END)
                widget.insert(0, value)
        nivel = cache.get("nivel_cargo")
        if nivel:
            self._populate_competencias(nivel)

    def _on_nivel_cargo_change(self, _event):
        nivel_widget = self.section2_fields.get("nivel_cargo")
        if not isinstance(nivel_widget, ttk.Combobox):
            return
        self._populate_competencias(nivel_widget.get())

    def _populate_competencias(self, nivel):
        values = condiciones_vacante.SECTION_2["competencias"].get(nivel, [])
        for idx, entry in enumerate(self.competencia_entries):
            entry.configure(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, values[idx] if idx < len(values) else "")
            entry.configure(state="readonly")

    def _confirm_section_2(self):
        payload = {}
        for field_id, widget in self.section2_fields.items():
            if isinstance(widget, ttk.Combobox):
                payload[field_id] = widget.get().strip()
            elif isinstance(widget, tk.Text):
                payload[field_id] = widget.get("1.0", tk.END).strip()
            else:
                payload[field_id] = widget.get().strip()
        competencias = [entry.get().strip() for entry in self.competencia_entries]
        for idx, value in enumerate(competencias, start=1):
            payload[f"competencia_{idx}"] = value
        try:
            condiciones_vacante.confirm_section_2(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2_1()

    def _show_section_2_1(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_2_1["title"])
        self.header_subtitle.config(text="Completa formacion academica.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        self.section2_1_fields = {}

        levels_frame = tk.Frame(content, bg=COLOR_LIGHT_BG)
        levels_frame.pack(fill="x", pady=(8, 12))
        tk.Label(
            levels_frame,
            text="Niveles educativos",
            font=FONT_LABEL,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 6))

        checkbox_container = tk.Frame(levels_frame, bg=COLOR_LIGHT_BG)
        checkbox_container.pack(anchor="w")
        for idx, (field_id, label, _cell) in enumerate(
            condiciones_vacante.SECTION_2_1["checkboxes"]
        ):
            var = tk.BooleanVar(value=False)
            cb = tk.Checkbutton(
                checkbox_container,
                text=label,
                variable=var,
                bg=COLOR_LIGHT_BG,
                anchor="w",
            )
            cb.grid(row=idx // 3, column=idx % 3, padx=8, pady=4, sticky="w")
            self.section2_1_fields[field_id] = var

        for field in condiciones_vacante.SECTION_2_1["fields"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(1, weight=1)

            tk.Label(
                row,
                text=field["label"],
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
                wraplength=760,
                justify="left",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)

            widget = None
            if field["type"] == "lista":
                widget = ttk.Combobox(
                    row,
                    values=field["options"],
                    state="readonly",
                    width=45,
                )
                widget.grid(row=0, column=1, sticky="w", padx=8, pady=8)
            elif field["type"] == "hora":
                widget = tk.Entry(row, width=22)
                widget.grid(row=0, column=1, sticky="w", padx=8, pady=8)
            elif field["type"] == "texto_largo":
                widget = tk.Text(row, height=3, wrap="word")
                widget.grid(row=0, column=1, sticky="we", padx=8, pady=8)
            else:
                widget = tk.Entry(row, width=48)
                widget.grid(row=0, column=1, sticky="w", padx=8, pady=8)

            self.section2_1_fields[field["id"]] = widget

        self._prefill_section2_1_fields()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2_1).pack(side="right")

    def _prefill_section2_1_fields(self):
        cache = condiciones_vacante.get_form_cache().get("section_2_1", {})
        for field_id, widget in self.section2_1_fields.items():
            value = cache.get(field_id, "")
            if isinstance(widget, tk.BooleanVar):
                widget.set(bool(value))
            elif isinstance(widget, ttk.Combobox):
                widget.set(value)
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", value)
            else:
                widget.delete(0, tk.END)
                widget.insert(0, value)

    def _confirm_section_2_1(self):
        payload = {}
        for field_id, widget in self.section2_1_fields.items():
            if isinstance(widget, tk.BooleanVar):
                payload[field_id] = bool(widget.get())
            elif isinstance(widget, ttk.Combobox):
                payload[field_id] = widget.get().strip()
            elif isinstance(widget, tk.Text):
                payload[field_id] = widget.get("1.0", tk.END).strip()
            else:
                payload[field_id] = widget.get().strip()
        try:
            condiciones_vacante.confirm_section_2_1(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_3()

    def _show_section_3(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_3["title"])
        self.header_subtitle.config(text="Selecciona nivel por habilidad.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        self.section3_fields = {}
        options = condiciones_vacante.SECTION_3["options"]

        bulk_actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        bulk_actions.pack(fill="x", pady=(0, 8))
        ttk.Button(
            bulk_actions,
            text='Marcar Todas "Alto"',
            command=lambda: self._set_section3_habilidades_nivel("Alto."),
        ).pack(side="right")
        ttk.Button(
            bulk_actions,
            text='Marcar Todas "Medio"',
            command=lambda: self._set_section3_habilidades_nivel("Medio."),
        ).pack(side="right", padx=(0, 8))
        ttk.Button(
            bulk_actions,
            text='Marcar Todas "Bajo"',
            command=lambda: self._set_section3_habilidades_nivel("Bajo."),
        ).pack(side="right", padx=(0, 8))

        for category in condiciones_vacante.SECTION_3["categories"]:
            tk.Label(
                content,
                text=category["title"],
                font=FONT_SECTION,
                fg=COLOR_PURPLE,
                bg=COLOR_LIGHT_BG,
            ).pack(anchor="w", pady=(8, 4))

            for field_id, label in category["items"]:
                row = tk.Frame(content, bg="white", bd=1, relief="solid")
                row.pack(fill="x", pady=4)
                row.grid_columnconfigure(1, weight=1)

                tk.Label(
                    row,
                    text=label,
                    font=FONT_LABEL,
                    bg="white",
                    fg="#222222",
                ).grid(row=0, column=0, sticky="w", padx=8, pady=6)

                combo = ttk.Combobox(
                    row,
                    values=options,
                    state="readonly",
                    width=ENTRY_W_MED,
                )
                combo.grid(row=0, column=1, sticky="w", padx=8, pady=6)
                self.section3_fields[field_id] = combo

            obs_row = tk.Frame(content, bg="white", bd=1, relief="solid")
            obs_row.pack(fill="x", pady=4)
            obs_row.grid_columnconfigure(1, weight=1)

            tk.Label(
                obs_row,
                text=category["observaciones_label"],
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=6)
            obs_text = tk.Text(obs_row, height=3, wrap="word")
            obs_text.grid(row=0, column=1, sticky="we", padx=8, pady=6)
            self.section3_fields[category["observaciones_id"]] = obs_text

        self._prefill_section3_fields()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2_1).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_3).pack(side="right")

    def _prefill_section3_fields(self):
        cache = condiciones_vacante.get_form_cache().get("section_3", {})
        for field_id, widget in self.section3_fields.items():
            value = cache.get(field_id, "")
            if isinstance(widget, ttk.Combobox):
                widget.set(value)
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", value)

    def _set_section3_habilidades_nivel(self, value):
        self._set_condiciones_habilidades_nivel(self.section3_fields, value)

    def _confirm_section_3(self):
        payload = {}
        for field_id, widget in self.section3_fields.items():
            if isinstance(widget, ttk.Combobox):
                payload[field_id] = widget.get().strip()
            elif isinstance(widget, tk.Text):
                payload[field_id] = widget.get("1.0", tk.END).strip()
        try:
            condiciones_vacante.confirm_section_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_4()

    def _show_section_4(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_4["title"])
        self.header_subtitle.config(text="Selecciona tiempo y frecuencia.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="both", expand=True)

        self.section4_fields = {}
        time_options = condiciones_vacante.SECTION_4["time_options"]
        frequency_options = condiciones_vacante.SECTION_4["frequency_options"]

        for field_id, label in condiciones_vacante.SECTION_4["fields"]:
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", pady=6)
            row.grid_columnconfigure(2, weight=1)

            tk.Label(
                row,
                text=label,
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=6)

            tk.Label(
                row,
                text="Tiempo de exposición",
                font=("Arial", 9),
                bg="white",
            ).grid(row=1, column=0, sticky="w", padx=8, pady=(0, 6))

            time_combo = ttk.Combobox(
                row,
                values=time_options,
                state="readonly",
                width=24,
            )
            time_combo.grid(row=1, column=1, sticky="w", padx=8, pady=(0, 6))
            self.section4_fields[f"{field_id}_tiempo"] = time_combo

            tk.Label(
                row,
                text="Frecuencia de exposición",
                font=("Arial", 9),
                bg="white",
            ).grid(row=2, column=0, sticky="w", padx=8, pady=(0, 6))

            frequency_combo = ttk.Combobox(
                row,
                values=frequency_options,
                state="readonly",
                width=ENTRY_W_MED,
            )
            frequency_combo.grid(row=2, column=1, sticky="w", padx=8, pady=(0, 6))
            self.section4_fields[f"{field_id}_frecuencia"] = frequency_combo

        self._prefill_section4_fields()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_3).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_4).pack(side="right")

    def _prefill_section4_fields(self):
        cache = condiciones_vacante.get_form_cache().get("section_4", {})
        for field_id, widget in self.section4_fields.items():
            widget.set(cache.get(field_id, ""))

    def _confirm_section_4(self):
        payload = {}
        for field_id, widget in self.section4_fields.items():
            payload[field_id] = widget.get().strip()
        try:
            condiciones_vacante.confirm_section_4(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_5()

    def _show_section_5(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_5["title"])
        self.header_subtitle.config(text="Selecciona nivel de riesgo y observaciones.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        self.section5_fields = {}
        options = condiciones_vacante.SECTION_5["options"]

        bulk_actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        bulk_actions.pack(fill="x", pady=(0, 8))
        ttk.Button(
            bulk_actions,
            text='Marcar Todas "Alto"',
            command=lambda: self._set_section5_habilidades_nivel("Alto."),
        ).pack(side="right")
        ttk.Button(
            bulk_actions,
            text='Marcar Todas "Medio"',
            command=lambda: self._set_section5_habilidades_nivel("Medio."),
        ).pack(side="right", padx=(0, 8))
        ttk.Button(
            bulk_actions,
            text='Marcar Todas "Bajo"',
            command=lambda: self._set_section5_habilidades_nivel("Bajo."),
        ).pack(side="right", padx=(0, 8))

        for category in condiciones_vacante.SECTION_5["categories"]:
            tk.Label(
                content,
                text=category["title"],
                font=FONT_SECTION,
                fg=COLOR_PURPLE,
                bg=COLOR_LIGHT_BG,
            ).pack(anchor="w", pady=(8, 4))

            for item in category["items"]:
                field_id = item[0]
                label = item[1]
                description = item[2] if len(item) > 2 else None

                row = tk.Frame(content, bg="white", bd=1, relief="solid")
                row.pack(fill="x", pady=4)
                row.grid_columnconfigure(1, weight=1)

                label_frame = tk.Frame(row, bg="white")
                label_frame.grid(row=0, column=0, sticky="w", padx=8, pady=6)

                tk.Label(
                    label_frame,
                    text=label,
                    font=FONT_LABEL,
                    bg="white",
                    fg="#222222",
                    wraplength=520,
                    justify="left",
                ).pack(anchor="w")

                if description:
                    tk.Label(
                        label_frame,
                        text=description,
                        font=("Arial", 9),
                        bg="white",
                        fg="#444444",
                        wraplength=520,
                        justify="left",
                    ).pack(anchor="w", pady=(4, 0))

                combo = ttk.Combobox(
                    row,
                    values=options,
                    state="readonly",
                    width=ENTRY_W_MED,
                )
                combo.grid(row=0, column=1, sticky="w", padx=8, pady=6)
                self.section5_fields[field_id] = combo

        obs_row = tk.Frame(content, bg="white", bd=1, relief="solid")
        obs_row.pack(fill="x", pady=8)
        obs_row.grid_columnconfigure(1, weight=1)

        tk.Label(
            obs_row,
            text=condiciones_vacante.SECTION_5["observaciones"]["label"],
            font=FONT_LABEL,
            bg="white",
            fg="#222222",
        ).grid(row=0, column=0, sticky="w", padx=8, pady=6)

        obs_text = tk.Text(obs_row, height=4, wrap="word")
        obs_text.grid(row=0, column=1, sticky="we", padx=8, pady=6)
        self.section5_fields[condiciones_vacante.SECTION_5["observaciones"]["id"]] = obs_text

        self._prefill_section5_fields()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_4).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_5).pack(side="right")

    def _prefill_section5_fields(self):
        cache = condiciones_vacante.get_form_cache().get("section_5", {})
        for field_id, widget in self.section5_fields.items():
            value = cache.get(field_id, "")
            if isinstance(widget, ttk.Combobox):
                widget.set(value)
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", value)

    def _set_section5_habilidades_nivel(self, value):
        self._set_condiciones_habilidades_nivel(self.section5_fields, value)

    def _set_condiciones_habilidades_nivel(self, fields, value):
        for widget in fields.values():
            if isinstance(widget, ttk.Combobox):
                widget.set(value)

    def _confirm_section_5(self):
        payload = {}
        for field_id, widget in self.section5_fields.items():
            if isinstance(widget, ttk.Combobox):
                payload[field_id] = widget.get().strip()
            elif isinstance(widget, tk.Text):
                payload[field_id] = widget.get("1.0", tk.END).strip()
        try:
            condiciones_vacante.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_6()

    def _show_section_6(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_6["title"])
        self.header_subtitle.config(text="Selecciona discapacidad y descripción.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        header = tk.Frame(content, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", pady=(8, 6))
        tk.Label(
            header,
            text="Discapacidad",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))
        tk.Label(
            header,
            text="Descripción sugerida",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=1, sticky="w")

        self.section6_rows = []
        self.section6_container = tk.Frame(content, bg=COLOR_LIGHT_BG)
        self.section6_container.pack(fill="x")
        self.disability_options = condiciones_vacante.SECTION_6["options"]
        self.disability_descriptions = condiciones_vacante.get_disability_descriptions()

        base_rows = condiciones_vacante.SECTION_6.get("base_rows", 4)
        for _ in range(base_rows):
            self._add_disability_row()

        add_row_btn = ttk.Button(
            content,
            text="Agregar discapacidad",
            command=self._add_disability_row,
        )
        add_row_btn.pack(anchor="w", pady=(6, 8))

        cached_rows = condiciones_vacante.get_form_cache().get("section_6", [])
        if cached_rows:
            for idx, entry in enumerate(cached_rows):
                if idx >= len(self.section6_rows):
                    self._add_disability_row()
                self._set_disability_row(self.section6_rows[idx], entry)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_5).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_6).pack(side="right")

    def _add_disability_row(self):
        row_index = len(self.section6_rows)
        row = tk.Frame(self.section6_container, bg="white", bd=1, relief="solid")
        row.pack(fill="x", pady=4)
        row.grid_columnconfigure(2, weight=1)

        combo = ttk.Combobox(
            row,
            values=self.disability_options,
            state="readonly",
            width=36,
        )
        combo.grid(row=0, column=0, sticky="w", padx=8, pady=6)

        descripcion = tk.Text(row, height=3, wrap="word", width=50, state="disabled")
        descripcion.grid(row=0, column=1, sticky="we", padx=8, pady=6)

        row_entry = {
            "combo": combo,
            "descripcion": descripcion,
        }
        self.section6_rows.append(row_entry)

        combo.bind(
            "<<ComboboxSelected>>",
            lambda _event, target=row_entry: self._update_disability_description(target),
        )

        return row_entry

    def _update_disability_description(self, row_entry):
        selection = row_entry["combo"].get().strip()
        key = condiciones_vacante.normalize_disability_key(selection)
        description = self.disability_descriptions.get(key, "")
        description_widget = row_entry["descripcion"]
        description_widget.configure(state="normal")
        description_widget.delete("1.0", tk.END)
        if description:
            description_widget.insert("1.0", description)
        description_widget.configure(state="disabled")

    def _set_disability_row(self, row_entry, values):
        discapacidad = (values or {}).get("discapacidad", "")
        descripcion = (values or {}).get("descripcion", "")
        if discapacidad:
            row_entry["combo"].set(discapacidad)
            self._update_disability_description(row_entry)
        if descripcion:
            row_entry["descripcion"].configure(state="normal")
            row_entry["descripcion"].delete("1.0", tk.END)
            row_entry["descripcion"].insert("1.0", descripcion)
            row_entry["descripcion"].configure(state="disabled")

    def _confirm_section_6(self):
        payload = []
        for row_entry in self.section6_rows:
            discapacidad = row_entry["combo"].get().strip()
            descripcion = row_entry["descripcion"].get("1.0", tk.END).strip()
            if discapacidad or descripcion:
                payload.append(
                    {
                        "discapacidad": discapacidad,
                        "descripcion": descripcion,
                    }
                )
        try:
            condiciones_vacante.confirm_section_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_7()

    def _show_section_7(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_7["title"])
        self.header_subtitle.config(text="Registra observaciones y recomendaciones.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        template_actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        template_actions.pack(fill="x", padx=24, pady=(12, 0))
        for idx, (template_key, label) in enumerate(condiciones_vacante.SECTION_7_TEMPLATE_BUTTONS):
            btn = ttk.Button(
                template_actions,
                text=label,
                command=lambda key=template_key: self._insert_condiciones_vacante_section7_template(key),
            )
            btn.grid(row=idx // 3, column=idx % 3, sticky="w", padx=(0, 8), pady=(0, 8))

        self.section7_text = tk.Text(section_frame, height=8, wrap="word")
        self.section7_text.pack(fill="x", padx=24, pady=(4, 12))

        cached = condiciones_vacante.get_form_cache().get("section_7", {})
        cached_text = cached.get(condiciones_vacante.SECTION_7["field_id"])
        if cached_text:
            self.section7_text.delete("1.0", tk.END)
            self.section7_text.insert("1.0", cached_text)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_6).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_7).pack(side="right")

    def _insert_condiciones_vacante_section7_template(self, template_key):
        template_text = condiciones_vacante.SECTION_7_TEMPLATES.get(template_key, "").strip()
        if not template_text:
            return
        current_text = self.section7_text.get("1.0", tk.END).strip()
        if current_text:
            self.section7_text.insert(tk.END, "\n\n")
        self.section7_text.insert(tk.END, template_text)
        self.section7_text.focus_set()
        self.section7_text.see(tk.END)

    def _confirm_section_7(self):
        payload = {
            condiciones_vacante.SECTION_7["field_id"]: self.section7_text.get("1.0", tk.END).strip()
        }
        try:
            condiciones_vacante.confirm_section_7(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_8()

    def _show_section_8(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_8["title"])
        self.header_subtitle.config(text="Registra asistentes.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        table = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        table.pack(fill="x", padx=24, pady=(12, 8))

        tk.Label(
            table,
            text="Nombre completo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        tk.Label(
            table,
            text="Cargo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=1, sticky="w")

        self.section8_rows = []
        self.section8_table = table
        self.section8_add_btn = ttk.Button(
            table,
            text="Agregar asistente",
            command=self._add_section8_asistente_row,
        )
        self._asesores_agencia_catalog = _get_asesores_agencia_catalog()

        cached_rows = condiciones_vacante.get_form_cache().get("section_8", [])
        self._render_section8_asistentes(cached_rows)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_7).pack(side="left")
        ttk.Button(actions, text="Finalizar", command=self._confirm_section_8).pack(side="right")

    def _get_section8_asistentes_values(self):
        values = []
        for nombre_entry, cargo_entry in self.section8_rows:
            values.append(
                {
                    "nombre": _get_input_value(nombre_entry),
                    "cargo": _get_input_value(cargo_entry),
                }
            )
        return values

    def _render_section8_asistentes(self, values=None):
        rows = list(values or [])
        min_rows = condiciones_vacante.SECTION_8.get("rows", 3)
        while len(rows) < min_rows:
            rows.append({"nombre": "", "cargo": ""})

        for nombre_entry, cargo_entry in self.section8_rows:
            nombre_entry.destroy()
            cargo_entry.destroy()
        self.section8_rows = []

        for idx, entry in enumerate(rows):
            row_idx = idx + 1
            is_first = idx == 0
            is_last = idx == len(rows) - 1
            if is_first and not is_last:
                nombre_entry, cargo_entry = _create_asistente_inputs(
                    self.section8_table,
                    ENTRY_W_WIDE,
                    use_catalog=True,
                    catalog=_get_asistentes_profesionales_catalog(),
                )
            elif is_last:
                nombre_entry, cargo_entry = _create_asesor_agencia_inputs(
                    self.section8_table,
                    ENTRY_W_WIDE,
                    catalog=getattr(self, "_asesores_agencia_catalog", None),
                )
            else:
                nombre_entry, cargo_entry = _create_asistente_inputs(
                    self.section8_table,
                    ENTRY_W_WIDE,
                    use_catalog=False,
                )
            nombre_entry.grid(row=row_idx, column=0, sticky="w", pady=4, padx=(0, 12))
            cargo_entry.grid(row=row_idx, column=1, sticky="w", pady=4)
            _set_input_value(nombre_entry, entry.get("nombre", ""))
            cargo_value = entry.get("cargo", "")
            if is_last and not cargo_value:
                cargo_value = "Asesor Agencia"
            _set_input_value(cargo_entry, cargo_value)
            self.section8_rows.append((nombre_entry, cargo_entry))

        self.section8_add_btn.grid(
            row=len(self.section8_rows) + 1,
            column=0,
            sticky="w",
            pady=(8, 0),
        )

    def _add_section8_asistente_row(self):
        rows = self._get_section8_asistentes_values()
        if rows:
            rows[-1] = {"nombre": "", "cargo": ""}
        rows.append({"nombre": "", "cargo": ""})
        self._render_section8_asistentes(rows)

    def _confirm_section_8(self):
        payload = []
        for nombre_entry, cargo_entry in self.section8_rows:
            payload.append(
                {
                    "nombre": nombre_entry.get().strip(),
                    "cargo": cargo_entry.get().strip(),
                }
            )
        try:
            condiciones_vacante.confirm_section_8(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Guardando Excel...")
        loading.set_progress(30)
        cache_snapshot = condiciones_vacante.get_form_cache()
        cache = cache_snapshot
        section_1 = cache.get("section_1", {})
        company_name = section_1.get("nombre_empresa")
        _start_background_finalization(
            self,
            loading,
            form_name="Revision Condicion",
            company_name=company_name,
            form_id="condiciones_vacante",
            worker_fn=lambda: _raise_finalize_stage(
                "guardando el Excel",
                condiciones_vacante.export_to_excel,
            ),
        )


class SeleccionIncluyenteWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Seleccion Incluyente - Seccion 1")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = seleccion_incluyente

        self.company_data = None
        self.fields = {}
        self.cedula_options = []

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "seleccion_incluyente",
            seleccion_incluyente,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_5": self._show_section_5,
                "section_6": self._show_section_6,
            },
            self._show_section_1,
        ):
            return True
        if not seleccion_incluyente.cache_file_exists():
            return False
        resume = messagebox.askyesno(
            "Reanudar",
            "Se encontró un formulario en progreso. ¿Deseas continuar donde lo dejaste?",
        )
        if not resume:
            seleccion_incluyente.clear_cache_file()
            seleccion_incluyente.clear_form_cache()
            return False
        seleccion_incluyente.load_cache_from_file()
        last_section = seleccion_incluyente.get_form_cache().get("_last_section")
        if last_section == "section_1":
            self._show_section_2()
        elif last_section == "section_2":
            self._show_section_5()
        elif last_section == "section_5":
            self._show_section_5()
        elif last_section == "section_6":
            self._show_section_6()
        else:
            self._show_section_1()
        return True

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))

        self.header_title = tk.Label(
            header,
            text="1. DATOS DE LA EMPRESA",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        self.header_title.pack(anchor="w")

        self.header_subtitle = tk.Label(
            header,
            text="Busca empresa por NIT y confirma datos.",
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        self.header_subtitle.pack(anchor="w", pady=(4, 0))

    def _build_section_container(self):
        self.section_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)

    def _clear_section_container(self):
        for child in self.section_container.winfo_children():
            child.destroy()

    def _clean_text(self, text):
        if not text:
            return ""
        replacements = {
            "\u00b6\u00a8": "\u00bf",
            "\u00c7?": "\u00cd",
            "\u00c7\u00ad": "\u00e1",
            "\u00c7\u00b8": "\u00e9",
            "\u00c7\u00f0": "\u00ed",
            "\u00c7\u00a7": "\u00fa",
            "\u00c7\u00b1": "\u00f1",
            "\u00c7\u00fc": "\u00f3",
            "\u0418": "\u00f3",
            "\u30f5": "\u00f1",
            "\u2cbe": "\u00f3",
            "\u00ef\u00bf\u00bd": "",
        }
        cleaned = str(text)
        for bad, good in replacements.items():
            cleaned = cleaned.replace(bad, good)
        return cleaned

    def _load_cedula_options(self):
        try:
            self.cedula_options = seleccion_incluyente.get_usuarios_reca_cedulas()
        except Exception:
            self.cedula_options = []

    def _filter_cedula_values(self, widget):
        raw = widget.get()
        normalized = re.sub(r"\D+", "", raw)
        options = self.cedula_options or []
        if normalized:
            filtered = [c for c in options if c and normalized in c]
        else:
            filtered = options
        widget.configure(values=filtered)

    def _format_date_for_ui(self, value):
        if not value:
            return ""
        raw = str(value).strip()
        if len(raw) >= 10 and "-" in raw:
            parts = raw[:10].split("-")
            if len(parts) == 3:
                return f"{parts[2]}/{parts[1]}/{parts[0]}"
        return raw

    def _apply_usuario_data(self, fields, data):
        mapping = {
            "nombre_usuario": "nombre_oferente",
            "certificado_porcentaje": "certificado_porcentaje",
            "discapacidad_detalle": "discapacidad",
            "telefono_oferente": "telefono_oferente",
            "fecha_nacimiento": "fecha_nacimiento",
            "cargo_oferente": "cargo_oferente",
            "contacto_emergencia": "nombre_contacto_emergencia",
            "parentesco": "parentesco",
            "telefono_emergencia": "telefono_emergencia",
            "resultado_certificado": "resultado_certificado",
            "pendiente_otros_oferentes": "pendiente_otros_oferentes",
            "cuenta_pension": "cuenta_pension",
            "tipo_pension": "tipo_pension",
        }
        for supa_key, field_id in mapping.items():
            value = data.get(supa_key)
            if value in (None, ""):
                continue
            widget = fields.get(field_id)
            if not widget:
                continue
            if supa_key == "fecha_nacimiento":
                value = self._format_date_for_ui(value)
            if supa_key == "discapacidad_detalle" and not value:
                continue
            if isinstance(widget, ttk.Combobox):
                widget.set(str(value))
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(value))
        fecha_widget = fields.get("fecha_nacimiento")
        edad_widget = fields.get("edad")
        if fecha_widget and edad_widget:
            self._format_birthdate(None, fecha_widget, edad_widget)

    def _on_cedula_selected(self, fields, widget):
        cedula = widget.get().strip()
        if not cedula:
            return
        normalized = re.sub(r"\D+", "", cedula)
        if normalized and normalized != cedula:
            widget.delete(0, tk.END)
            widget.insert(0, normalized)
        try:
            data = seleccion_incluyente.get_usuario_reca_by_cedula(normalized)
        except Exception:
            return
        if data:
            self._apply_usuario_data(fields, data)

    def _show_section_1(self):
        self._clear_section_container()
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._prefill_section_1()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self.continue_btn = ttk.Button(
            actions,
            text="Continuar",
            command=self._confirm_and_continue,
            state="disabled",
        )
        self.continue_btn.pack(side="right")

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        groups = [
            ('Información de Empresa', COLOR_GROUP_EMPRESA, ['nombre_empresa', 'direccion_empresa', 'correo_1', 'contacto_empresa', 'telefono_empresa', 'cargo', 'ciudad_empresa', 'sede_empresa', 'caja_compensacion']),
            ('Información de Compensar', COLOR_GROUP_COMPENSAR, ['asesor']),
            ('Información de RECA', COLOR_GROUP_RECA, ['profesional_asignado']),
        ]
        labels = {
            'nombre_empresa': 'Nombre de la empresa',
            'direccion_empresa': 'Dirección de la empresa',
            'correo_1': 'Correo electrónico',
            'contacto_empresa': 'Contacto de la empresa',
            'telefono_empresa': 'Teléfonos',
            'cargo': 'Cargo',
            'ciudad_empresa': 'Ciudad/Municipio',
            'sede_empresa': 'Sede Compensar',
            'caja_compensacion': 'Empresa afiliada a Caja de Compensación',
            'asesor': 'Asesor',
            'profesional_asignado': 'Profesional asignado RECA',
        }
        _section1_build_groups(self, parent, groups, labels)

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _search_company(self, mode="nit"):
        nit = self.fields["nit_empresa"].get().strip()
        nombre = self.fields.get("nombre_busqueda").get().strip() if self.fields.get("nombre_busqueda") else ""
        if mode == "nit":
            if not nit:
                messagebox.showerror("Error", "Ingresa un NIT.")
                return
        elif mode == "nombre":
            if not nombre:
                messagebox.showerror("Error", "Ingresa el nombre de la empresa.")
                return
        else:
            messagebox.showerror("Error", "Tipo de búsqueda no válido.")
            return

        try:
            self.status_label.config(text="Buscando empresa...")
            self.update_idletasks()
            if mode == "nombre":
                company = seleccion_incluyente.get_empresa_by_nombre(nombre)
            else:
                company = seleccion_incluyente.get_empresa_by_nit(nit)
        except Exception as exc:
            self.status_label.config(text="")
            messagebox.showerror("Error", str(exc))
            return

        if not company:
            self.company_data = None
            msg = "No se encontró empresa para ese nombre." if mode == "nombre" else "No se encontró empresa para ese NIT."
            self.status_label.config(text=msg)
            self.continue_btn.config(state="disabled")
            for key in seleccion_incluyente.SECTION_1_SUPABASE_MAP.keys():
                self._set_readonly_value(key, "")
            return

        if mode == "nombre":
            nit_value = company.get("nit_empresa")
            if nit_value:
                entry = self.fields.get("nit_empresa")
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, nit_value)

        self.company_data = company
        self.status_label.config(text="Empresa encontrada.")
        self.continue_btn.config(state="normal")
        for key in seleccion_incluyente.SECTION_1_SUPABASE_MAP.keys():
            self._set_readonly_value(key, company.get(key))

    def _confirm_and_continue(self):
        if not self.company_data:
            messagebox.showerror("Error", "Busca una empresa antes de confirmar.")
            return
        fecha_visita = _get_required_fecha_visita(self)
        if not fecha_visita:
            return
        modalidad = _get_required_modalidad(self)
        if not modalidad:
            return
        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad": modalidad,
            "nit_empresa": self.fields["nit_empresa"].get().strip(),
        }
        try:
            seleccion_incluyente.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2()

    def _label_for_field(self, field_id):
        return getattr(self, "_section1_labels", {}).get(field_id, field_id)

    def _prefill_section_1(self):
        cache = seleccion_incluyente.get_form_cache().get("section_1", {})
        if not cache:
            return
        self.company_data = cache
        self.fields["nit_empresa"].delete(0, tk.END)
        self.fields["nit_empresa"].insert(0, cache.get("nit_empresa", ""))
        self.fields["modalidad"].set(cache.get("modalidad", ""))
        fecha_value = cache.get("fecha_visita")
        if fecha_value:
            self.fields["fecha_visita"].set_date(fecha_value)
        for key in [
            "nombre_empresa",
            "direccion_empresa",
            "correo_1",
            "contacto_empresa",
            "telefono_empresa",
            "cargo",
            "ciudad_empresa",
            "sede_empresa",
            "caja_compensacion",
            "asesor",
            "profesional_asignado",
        ]:
            entry = self.fields.get(key)
            if not entry:
                continue
            entry.configure(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, cache.get(key, ""))
            entry.configure(state="readonly")

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. DATOS DEL OFERENTE")
        self.header_subtitle.config(text="Registra datos del oferente.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        self._load_cedula_options()
        content = _build_scrollable_content(section_frame, self)

        self.oferente_blocks = []
        self.oferente_frames = []
        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)

        field_meta = {field["id"]: field for field in seleccion_incluyente.SECTION_2["fields"]}

        def _create_widget(parent, field_id, width=30, text_height=4):
            meta = field_meta.get(field_id, {})
            if meta.get("type") == "lista":
                return ttk.Combobox(parent, values=meta.get("options", []), state="readonly", width=width)
            if meta.get("type") == "texto_largo":
                w = tk.Text(parent, width=width, height=text_height, wrap="word")
                return w
            if field_id == "cedula":
                widget = ttk.Combobox(parent, values=self.cedula_options, state="normal", width=width)
                self._apply_numeric_entry(widget)
                return widget
            return tk.Entry(parent, width=width)

        def _add_fields_grid(parent, field_ids, columns=2):
            fields = {}
            for idx, field_id in enumerate(field_ids):
                meta = field_meta.get(field_id, {})
                row = idx // columns
                col = (idx % columns) * 2
                label = tk.Label(
                    parent,
                    text=meta.get("label", field_id),
                    font=("Arial", 9, "bold"),
                    bg="white",
                    anchor="w",
                )
                label.grid(row=row, column=col, sticky="w", padx=6, pady=4)
                widget = _create_widget(parent, field_id, width=30)
                widget.grid(row=row, column=col + 1, sticky="w", padx=6, pady=4)
                if isinstance(widget, tk.Entry):
                    if field_id == "cedula":
                        self._apply_numeric_entry(widget)
                    if field_id == "certificado_porcentaje":
                        self._apply_decimal_entry(widget)
                    if field_id in {"telefono_oferente", "telefono_emergencia"}:
                        self._apply_numeric_entry(widget, max_len=10)
                    if field_id in {"nombre_oferente", "nombre_contacto_emergencia"}:
                        self._apply_name_entry(widget)
                fields[field_id] = widget
            return fields

        def _add_question_block(parent, title, field_ids, subitems=None):
            frame = tk.Frame(parent, bg="white")
            frame.pack(fill="x", pady=10)
            tk.Label(
                frame,
                text=title,
                font=("Arial", 9, "bold"),
                bg="white",
                anchor="w",
            ).grid(row=0, column=0, sticky="w", padx=6, pady=(2, 4))
            content_frame = tk.Frame(frame, bg="white")
            content_frame.grid(row=1, column=0, sticky="w", padx=18)
            fields = {}
            for idx, field_id in enumerate(field_ids):
                meta = field_meta.get(field_id, {})
                tk.Label(
                    content_frame,
                    text=meta.get("label", field_id),
                    font=("Arial", 8, "bold"),
                    bg="white",
                    anchor="w",
                ).grid(row=idx * 2, column=0, sticky="w")
                widget = _create_widget(content_frame, field_id, width=52, text_height=2)
                widget.grid(row=idx * 2 + 1, column=0, sticky="w", pady=(0, 2))
                fields[field_id] = widget
            if subitems:
                sub_frame = tk.Frame(content_frame, bg="white")
                sub_frame.grid(row=len(field_ids) * 2, column=0, sticky="w", pady=(6, 0))
                for row_idx, item in enumerate(subitems):
                    if len(item) == 4:
                        left_label, left_id, right_label, right_id = item
                        if left_label:
                            tk.Label(sub_frame, text=left_label, bg="white", anchor="w").grid(
                                row=row_idx, column=0, sticky="w", padx=(0, 8)
                            )
                        if left_id:
                            fields[left_id] = _create_widget(sub_frame, left_id, width=10)
                            fields[left_id].grid(row=row_idx, column=1, sticky="w", padx=4)
                        if right_label:
                            tk.Label(sub_frame, text=right_label, bg="white", anchor="w").grid(
                                row=row_idx, column=2, sticky="w", padx=(12, 8)
                            )
                        if right_id:
                            fields[right_id] = _create_widget(sub_frame, right_id, width=10)
                            fields[right_id].grid(row=row_idx, column=3, sticky="w", padx=4)
                    else:
                        label_text, req_id, cuenta_id = item
                        tk.Label(sub_frame, text=label_text, bg="white", anchor="w").grid(
                            row=row_idx, column=0, sticky="w", padx=(0, 8)
                        )
                        if req_id:
                            fields[req_id] = _create_widget(sub_frame, req_id, width=10)
                            fields[req_id].grid(row=row_idx, column=1, sticky="w", padx=4)
                        fields[cuenta_id] = _create_widget(sub_frame, cuenta_id, width=10)
                        fields[cuenta_id].grid(row=row_idx, column=2, sticky="w", padx=4)
            return fields

        def _add_activity_block(parent, title, nivel_id, observacion_id, nota_id, subitems):
            frame = tk.Frame(parent, bg="white")
            frame.pack(fill="x", pady=10)
            tk.Label(
                frame,
                text=title,
                font=("Arial", 9, "bold"),
                bg="white",
                anchor="w",
            ).grid(row=0, column=0, sticky="w", padx=6, pady=(2, 4))

            content_frame = tk.Frame(frame, bg="white")
            content_frame.grid(row=1, column=0, sticky="w", padx=18)
            fields = {}

            fields[nivel_id] = _create_widget(content_frame, nivel_id, width=ENTRY_W_LONG)
            fields[nivel_id].grid(row=0, column=0, sticky="w", pady=(0, 4))

            fields[observacion_id] = _create_widget(content_frame, observacion_id, width=52)
            fields[observacion_id].grid(row=1, column=0, sticky="w", pady=(0, 4))

            sub_frame = tk.Frame(content_frame, bg="white")
            sub_frame.grid(row=2, column=0, sticky="w", pady=(6, 0))
            for row_idx, (left_label, left_id, right_label, right_id) in enumerate(subitems):
                if left_label:
                    tk.Label(sub_frame, text=left_label, bg="white", anchor="w").grid(
                        row=row_idx, column=0, sticky="w", padx=(0, 4)
                    )
                if left_id:
                    fields[left_id] = _create_widget(sub_frame, left_id, width=10)
                    fields[left_id].grid(row=row_idx, column=1, sticky="w", padx=4)
                if right_label:
                    tk.Label(sub_frame, text=right_label, bg="white", anchor="w").grid(
                        row=row_idx, column=2, sticky="w", padx=(6, 4)
                    )
                if right_id:
                    fields[right_id] = _create_widget(sub_frame, right_id, width=10)
                    fields[right_id].grid(row=row_idx, column=3, sticky="w", padx=4)

            tk.Label(sub_frame, text="Nota:", font=("Arial", 8, "bold"), bg="white").grid(
                row=len(subitems), column=0, sticky="w", pady=(6, 0)
            )
            fields[nota_id] = tk.Entry(sub_frame, width=30)
            fields[nota_id].grid(row=len(subitems), column=1, columnspan=3, sticky="w", pady=(4, 0))
            return fields

        remove_btn = None

        def _refresh_oferente_numbers():
            for idx, fields in enumerate(self.oferente_blocks, start=1):
                numero_widget = fields.get("numero")
                if not numero_widget:
                    continue
                try:
                    numero_widget.configure(state="normal")
                    numero_widget.delete(0, tk.END)
                    numero_widget.insert(0, str(idx))
                finally:
                    numero_widget.configure(state="readonly")

        def _update_remove_button_state():
            if remove_btn is None:
                return
            state = "normal" if len(self.oferente_blocks) > 1 else "disabled"
            remove_btn.config(state=state)

        remove_btn = None

        def _get_shared_desarrollo_value():
            for entry_fields in self.oferente_blocks:
                widget = entry_fields.get("desarrollo_actividad")
                if isinstance(widget, tk.Text):
                    value = widget.get("1.0", tk.END).strip()
                    if value:
                        return value
            return ""

        def _set_text_widget_value(widget, value):
            if not isinstance(widget, tk.Text):
                return
            current = widget.get("1.0", tk.END).strip()
            if current == value:
                return
            widget.delete("1.0", tk.END)
            if value:
                widget.insert("1.0", value)

        def _sync_desarrollo_widgets(source_widget=None):
            if isinstance(source_widget, tk.Text):
                shared_value = source_widget.get("1.0", tk.END).strip()
            else:
                shared_value = _get_shared_desarrollo_value()
            for entry_fields in self.oferente_blocks:
                widget = entry_fields.get("desarrollo_actividad")
                if widget is source_widget:
                    continue
                _set_text_widget_value(widget, shared_value)

        def _bind_shared_desarrollo(widget):
            if not isinstance(widget, tk.Text):
                return
            widget.bind(
                "<FocusOut>",
                lambda _event=None, w=widget: _sync_desarrollo_widgets(w),
                add="+",
            )

        def _add_oferente_block():
            idx = len(self.oferente_blocks) + 1
            block = tk.LabelFrame(
                content,
                text=f"Oferente {idx}",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=12,
                pady=8,
            )
            block.pack(fill="x", pady=8, before=actions)

            fields = {}

            section2_frame = tk.LabelFrame(
                block,
                text="2. DATOS DEL OFERENTE",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section2_frame.pack(fill="x", pady=(0, 8))
            section2_frame.grid_columnconfigure(1, weight=1)

            section2_fields = [
                "numero",
                "nombre_oferente",
                "cedula",
                "certificado_porcentaje",
                "discapacidad",
                "telefono_oferente",
                "resultado_certificado",
                "cargo_oferente",
                "nombre_contacto_emergencia",
                "parentesco",
                "telefono_emergencia",
                "fecha_nacimiento",
                "edad",
                "pendiente_otros_oferentes",
                "lugar_firma_contrato",
                "fecha_firma_contrato",
                "cuenta_pension",
                "tipo_pension",
            ]
            fields.update(_add_fields_grid(section2_frame, section2_fields, columns=2))
            numero_widget = fields.get("numero")
            if numero_widget:
                numero_widget.delete(0, tk.END)
                numero_widget.insert(0, str(idx))
                numero_widget.configure(state="readonly")
            cedula_widget = fields.get("cedula")
            if isinstance(cedula_widget, ttk.Combobox):
                cedula_widget.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, f=fields, w=cedula_widget: self._on_cedula_selected(f, w),
                )
                cedula_widget.bind(
                    "<KeyRelease>",
                    lambda _e, w=cedula_widget: self._filter_cedula_values(w),
                )
                cedula_widget.bind(
                    "<FocusOut>",
                    lambda _e, f=fields, w=cedula_widget: self._on_cedula_selected(f, w),
                )
                cedula_widget.bind(
                    "<Return>",
                    lambda _e, f=fields, w=cedula_widget: self._on_cedula_selected(f, w),
                )
            fecha_widget = fields.get("fecha_nacimiento")
            edad_widget = fields.get("edad")
            if fecha_widget and edad_widget:
                edad_widget.configure(state="readonly")
                fecha_widget.bind(
                    "<KeyRelease>",
                    lambda event, fw=fecha_widget, ew=edad_widget: self._format_birthdate(event, fw, ew),
                )

            section3_frame = tk.LabelFrame(
                block,
                text="3. DESARROLLO DE LA ACTIVIDAD",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section3_frame.pack(fill="x", pady=(0, 8))
            fields["desarrollo_actividad"] = _create_widget(section3_frame, "desarrollo_actividad", width=80, text_height=6)
            fields["desarrollo_actividad"].pack(fill="x", padx=6, pady=6)
            _set_text_widget_value(fields["desarrollo_actividad"], _get_shared_desarrollo_value())
            _bind_shared_desarrollo(fields["desarrollo_actividad"])

            section41_frame = tk.LabelFrame(
                block,
                text="4.1 Condiciones medicas y de salud",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section41_frame.pack(fill="x", pady=(0, 8))

            fields.update(
                _add_question_block(
                    section41_frame,
                    "¿Toma medicamentos?",
                    ["medicamentos_nivel_apoyo", "medicamentos_conocimiento", "medicamentos_horarios", "medicamentos_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section41_frame,
                    "¿Presenta alguna alergia?",
                    ["alergias_nivel_apoyo", "alergias_tipo", "alergias_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section41_frame,
                    "¿Tiene algún tipo de restricción médica?",
                    ["restriccion_nivel_apoyo", "restriccion_conocimiento", "restriccion_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section41_frame,
                    "¿Asiste a controles médicos con especialista?",
                    ["controles_nivel_apoyo", "controles_asistencia", "controles_frecuencia", "controles_nota"],
                )
            )

            section42_frame = tk.LabelFrame(
                block,
                text="4.2 Habilidades basicas de la vida diaria",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section42_frame.pack(fill="x", pady=(0, 8))

            fields.update(
                _add_question_block(
                    section42_frame,
                    "¿Se desplaza por la ciudad de manera independiente?",
                    ["desplazamiento_nivel_apoyo", "desplazamiento_modo", "desplazamiento_transporte", "desplazamiento_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42_frame,
                    "¿Se le facilita ubicarse dentro de la ciudad?",
                    ["ubicacion_nivel_apoyo", "ubicacion_ciudad", "ubicacion_aplicaciones", "ubicacion_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42_frame,
                    "¿Reconoce y maneja el dinero?",
                    ["dinero_nivel_apoyo", "dinero_reconocimiento", "dinero_manejo", "dinero_medios", "dinero_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42_frame,
                    "Presentacion personal",
                    ["presentacion_nivel_apoyo", "presentacion_personal", "presentacion_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42_frame,
                    "¿Conoce y maneja algún apoyo de comunicación escrita?",
                    ["comunicacion_escrita_nivel_apoyo", "comunicacion_escrita_apoyo", "comunicacion_escrita_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42_frame,
                    "¿Conoce y maneja algún apoyo de comunicación verbal?",
                    ["comunicacion_verbal_nivel_apoyo", "comunicacion_verbal_apoyo", "comunicacion_verbal_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42_frame,
                    "¿A quién recurre al momento de tomar decisiones?",
                    ["decisiones_nivel_apoyo", "toma_decisiones", "toma_decisiones_nota"],
                )
            )
            fields.update(
                _add_activity_block(
                    section42_frame,
                    "¿Necesita apoyo en algunas de las siguientes actividades de la vida diaria?",
                    "aseo_nivel_apoyo",
                    "alimentacion",
                    "aseo_nota",
                    subitems=[
                        ("Criar y cuidado de ninos", "aseo_criar_apoyo", "Alimentacion", "aseo_alimentacion"),
                        ("Uso de los sistemas de comunicacion", "aseo_comunicacion_apoyo", "Movilidad funcional", "aseo_movilidad_funcional"),
                        ("Cuidado de las ayudas tecnicas personales", "aseo_ayudas_apoyo", "Higiene personal y aseo (Control de esfinter)", "aseo_higiene_aseo"),
                    ],
                )
            )
            fields.update(
                _add_activity_block(
                    section42_frame,
                    "¿Necesita apoyo en algunas de las siguientes actividades instrumentales de la vida diaria?",
                    "instrumentales_nivel_apoyo",
                    "instrumentales_actividades",
                    "instrumentales_nota",
                    subitems=[
                        ("Criar y cuidado de ninos", "instrumentales_criar_apoyo", "Manejo de tematicas financieras", "instrumentales_finanzas"),
                        ("Uso de los sistemas de comunicacion", "instrumentales_comunicacion_apoyo", "Cocina y limpieza", "instrumentales_cocina_limpieza"),
                        ("Movilidad en la comunidad", "instrumentales_movilidad_apoyo", "Crear y mantener un hogar", "instrumentales_crear_hogar"),
                        ("", None, "Cuidado de la salud y manutencion", "instrumentales_salud_cuenta_apoyo"),
                    ],
                )
            )
            fields.update(
                _add_question_block(
                    section42_frame,
                    "¿Necesita apoyo durante actividades laborales?",
                    ["actividades_nivel_apoyo", "actividades_apoyo", "actividades_nota"],
                    subitems=[
                        ("Actividades de esparcimiento con familia", "actividades_esparcimiento_apoyo", "Psicologico en salud", "actividades_esparcimiento_cuenta_apoyo"),
                        ("Complementarios médicos", "actividades_complementarios_apoyo", "Actividades académicas de hijos", "actividades_complementarios_cuenta_apoyo"),
                        ("Subsidios economicos para estudio de hijos", None, "", "actividades_subsidios_cuenta_apoyo"),
                    ],
                )
            )
            fields.update(
                _add_question_block(
                    section42_frame,
                    "¿Ha sufrido o vivido discriminación?",
                    ["discriminacion_nivel_apoyo", "discriminacion", "discriminacion_nota"],
                    subitems=[
                        ("Violencia fisica", "discriminacion_violencia_apoyo", "Acoso laboral", "discriminacion_violencia_cuenta_apoyo"),
                        ("Vulneracion de derechos", "discriminacion_vulneracion_apoyo", "Violencia psicosocial", "discriminacion_vulneracion_cuenta_apoyo"),
                    ],
                )
            )
            self.oferente_blocks.append(fields)
            _update_remove_button_state()
            self.oferente_frames.append(block)
            _refresh_oferente_numbers()
            _update_remove_button_state()

        def _remove_oferente_block():
            if len(self.oferente_blocks) <= 1:
                return
            self.oferente_blocks.pop()
            frame = self.oferente_frames.pop()
            frame.destroy()
            _refresh_oferente_numbers()
            _update_remove_button_state()

        def _prefill_section_2():
            cache = seleccion_incluyente.get_form_cache().get("section_2", [])
            if not cache:
                _add_oferente_block()
                return
            for _ in range(len(cache)):
                _add_oferente_block()
            for idx, entry in enumerate(cache):
                fields = self.oferente_blocks[idx]
                for key, widget in fields.items():
                    value = entry.get(key, "")
                    if isinstance(widget, ttk.Combobox):
                        widget.set(value)
                    elif isinstance(widget, tk.Text):
                        widget.delete("1.0", tk.END)
                        widget.insert("1.0", value)
                    else:
                        widget.delete(0, tk.END)
                        widget.insert(0, value)
            _sync_desarrollo_widgets()

        _prefill_section_2()

        ttk.Button(actions, text="Agregar oferente", command=_add_oferente_block).pack(
            side="left"
        )
        remove_btn = ttk.Button(
            actions, text="Eliminar ultimo oferente", command=_remove_oferente_block
        )
        remove_btn.pack(side="left", padx=8)
        _update_remove_button_state()
        ttk.Button(actions, text="Regresar", command=self._show_section_1).pack(
            side="left", padx=8
        )
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2).pack(
            side="right"
        )
    def _confirm_section_2(self):
        shared_desarrollo = ""
        for fields in self.oferente_blocks:
            widget = fields.get("desarrollo_actividad")
            if isinstance(widget, tk.Text):
                shared_desarrollo = widget.get("1.0", tk.END).strip()
                if shared_desarrollo:
                    break
        payload = []
        for fields in self.oferente_blocks:
            entry = {}
            for key, widget in fields.items():
                if isinstance(widget, ttk.Combobox):
                    entry[key] = widget.get().strip()
                elif isinstance(widget, tk.Text):
                    entry[key] = widget.get("1.0", tk.END).strip()
                else:
                    entry[key] = widget.get().strip()
            entry["desarrollo_actividad"] = shared_desarrollo
            payload.append(entry)
        try:
            seleccion_incluyente.confirm_section_2(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_5()

    def _show_section_5(self):
        self._clear_section_container()
        self.header_title.config(text="5. AJUSTES RAZONABLES / RECOMENDACIONES")
        self.header_subtitle.config(text="Completa ajustes y recomendaciones.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="both", expand=True)

        self.section5_fields = {}

        tk.Label(
            content,
            text="Ajustes razonables / recomendaciones",
            font=FONT_SECTION,
            bg=COLOR_LIGHT_BG,
            anchor="w",
        ).pack(anchor="w", pady=(8, 4))
        template_actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        template_actions.pack(fill="x", padx=4, pady=(0, 4))
        for idx, (template_key, label) in enumerate(seleccion_incluyente.AJUSTES_ENTREVISTA_TEMPLATE_BUTTONS):
            btn = ttk.Button(
                template_actions,
                text=label,
                command=lambda key=template_key: self._insert_seleccion_section5_template(key),
            )
            btn.grid(row=idx // 3, column=idx % 3, sticky="w", padx=(0, 8), pady=(0, 8))
        ajustes = tk.Text(content, height=8, width=TEXT_WIDE, wrap="word")
        ajustes.pack(fill="x", padx=4, pady=(0, 10))
        self.section5_fields["ajustes_recomendaciones"] = ajustes

        tk.Label(
            content,
            text="Nota",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
            anchor="w",
        ).pack(anchor="w", pady=(4, 2))
        nota = tk.Entry(content, width=80)
        nota.pack(anchor="w", padx=4, pady=(0, 16))
        self.section5_fields["nota"] = nota

        cache = seleccion_incluyente.get_form_cache().get("section_5", {})
        if cache:
            ajustes.delete("1.0", tk.END)
            ajustes.insert("1.0", cache.get("ajustes_recomendaciones", ""))
            nota.delete(0, tk.END)
            nota.insert(0, cache.get("nota", ""))

        _build_wizard_actions(
            content,
            back_command=self._show_section_2,
            primary_command=self._confirm_section_5,
        )

    def _insert_seleccion_section5_template(self, template_key):
        text_widget = self.section5_fields.get("ajustes_recomendaciones")
        if not text_widget:
            return
        template_text = seleccion_incluyente.AJUSTES_ENTREVISTA_TEMPLATES.get(template_key, "").strip()
        if not template_text:
            return
        current_text = text_widget.get("1.0", tk.END).strip()
        if current_text:
            text_widget.insert(tk.END, "\n\n")
        text_widget.insert(tk.END, template_text)
        text_widget.focus_set()
        text_widget.see(tk.END)

    def _confirm_section_5(self):
        payload = {
            "ajustes_recomendaciones": self.section5_fields["ajustes_recomendaciones"]
            .get("1.0", tk.END)
            .strip(),
            "nota": self.section5_fields["nota"].get().strip(),
        }
        try:
            seleccion_incluyente.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_6()

    def _show_section_6(self):
        self._clear_section_container()
        self.header_title.config(text=seleccion_incluyente.SECTION_6["title"])
        self.header_subtitle.config(text="Registra asistentes y agrega filas si aplica.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="both", expand=True)

        table = tk.Frame(content, bg=COLOR_LIGHT_BG)
        table.pack(fill="x", padx=4, pady=(0, 8))

        tk.Label(
            table,
            text="Nombre completo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        tk.Label(
            table,
            text="Cargo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=1, sticky="w")

        self.section6_rows = []
        asistentes_catalog = _get_asistentes_profesionales_catalog()
        for idx in range(seleccion_incluyente.SECTION_6["rows"]):
            nombre_entry, cargo_entry = _create_asistente_inputs(
                table,
                ENTRY_W_WIDE,
                use_catalog=(idx == 0),
                catalog=asistentes_catalog,
            )
            nombre_entry.grid(row=idx + 1, column=0, sticky="w", pady=4, padx=(0, 12))
            cargo_entry.grid(row=idx + 1, column=1, sticky="w", pady=4)
            self.section6_rows.append((nombre_entry, cargo_entry))

        def _add_asistente_row():
            row_idx = len(self.section6_rows) + 1
            nombre_entry, cargo_entry = _create_asistente_inputs(
                table,
                ENTRY_W_WIDE,
                use_catalog=False,
                catalog=asistentes_catalog,
            )
            nombre_entry.grid(row=row_idx, column=0, sticky="w", pady=4, padx=(0, 12))
            cargo_entry.grid(row=row_idx, column=1, sticky="w", pady=4)
            self.section6_rows.append((nombre_entry, cargo_entry))
            add_btn.grid(row=len(self.section6_rows) + 1, column=0, sticky="w", pady=(8, 0))

        add_btn = ttk.Button(
            table,
            text="Agregar asistente",
            command=_add_asistente_row,
        )
        add_btn.grid(row=len(self.section6_rows) + 1, column=0, sticky="w", pady=(8, 0))

        cached_rows = seleccion_incluyente.get_form_cache().get("section_6", [])
        while len(self.section6_rows) < len(cached_rows):
            _add_asistente_row()
        for idx, entry in enumerate(cached_rows):
            nombre_entry, cargo_entry = self.section6_rows[idx]
            nombre_entry.delete(0, tk.END)
            nombre_entry.insert(0, entry.get("nombre", ""))
            cargo_entry.delete(0, tk.END)
            cargo_entry.insert(0, entry.get("cargo", ""))

        _build_wizard_actions(
            content,
            back_command=self._show_section_5,
            primary_command=self._confirm_section_6,
            primary_text="Finalizar",
        )

    def _confirm_section_6(self):
        asistentes = []
        for nombre_entry, cargo_entry in self.section6_rows:
            asistentes.append(
                {
                    "nombre": nombre_entry.get().strip(),
                    "cargo": cargo_entry.get().strip(),
                }
            )
        try:
            seleccion_incluyente.confirm_section_6(asistentes)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Generando Excel...")
        loading.set_progress(40)
        cache_snapshot = seleccion_incluyente.get_form_cache()
        cache = cache_snapshot
        section_1 = cache.get("section_1", {})
        company_name = section_1.get("nombre_empresa")

        def _worker():
            output_path = _raise_finalize_stage(
                "generando el Excel",
                lambda: seleccion_incluyente.export_to_excel(clear_cache=False),
            )
            _update_loading_async(
                loading,
                status="Guardando en Supabase...",
                progress=70,
            )
            _raise_finalize_stage(
                "guardando en Supabase",
                seleccion_incluyente.sync_usuarios_reca,
            )
            return output_path

        _start_background_finalization(
            self,
            loading,
            form_name="Seleccion Incluyente",
            company_name=company_name,
            form_id="seleccion_incluyente",
            worker_fn=_worker,
            post_delivery_fn=lambda: _clear_form_cache_safe(seleccion_incluyente),
        )

    def _format_birthdate(self, _event, fecha_widget, edad_widget):
        digits, formatted = _format_birthdate_text(fecha_widget.get())
        fecha_widget.delete(0, tk.END)
        fecha_widget.insert(0, formatted)
        fecha_widget.icursor(tk.END)
        age = self._calculate_age(digits)
        _set_readonly_entry_value(edad_widget, "" if age is None else age)

    def _calculate_age(self, digits):
        return _calc_age_from_digits(digits, min_year=1900)

    def _apply_numeric_entry(self, entry, max_len=None):
        _bind_numeric_entry(entry, max_len=max_len)

    def _apply_decimal_entry(self, entry):
        _bind_decimal_entry(entry)

    def _apply_name_entry(self, entry):
        _bind_name_entry(entry)

class ContratacionIncluyenteWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Contratacion Incluyente - Seccion 1")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = contratacion_incluyente

        self.company_data = None
        self.fields = {}
        self.cedula_options = []

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))

        self.header_title = tk.Label(
            header,
            text="1. DATOS DE LA EMPRESA",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        self.header_title.pack(anchor="w")

        self.header_subtitle = tk.Label(
            header,
            text="Busca empresa por NIT y confirma datos.",
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        self.header_subtitle.pack(anchor="w", pady=(4, 0))

    def _build_section_container(self):
        self.section_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)

    def _clear_section_container(self):
        for child in self.section_container.winfo_children():
            child.destroy()

    def _normalize_name_value(self, value):
        return _normalize_person_name(value)

    def _apply_numeric_entry(self, entry, max_len=None):
        _bind_numeric_entry(entry, max_len=max_len)

    def _apply_decimal_entry(self, entry):
        _bind_decimal_entry(entry)

    def _apply_name_entry(self, entry):
        _bind_name_entry(entry)

    def _set_age_value(self, entry, value):
        _set_readonly_entry_value(entry, value)

    def _apply_date_entry(self, date_entry, age_entry):
        _bind_birthdate_entry(
            date_entry,
            age_entry,
            min_year=1900,
            mark_invalid=True,
            clear_invalid=True,
        )

    def _refresh_age_from_date(self, date_entry, age_entry):
        _refresh_age_from_date_entry(date_entry, age_entry, min_year=1900)

    def _load_cedula_options(self):
        try:
            self.cedula_options = contratacion_incluyente.get_usuarios_reca_cedulas()
        except Exception:
            self.cedula_options = []

    def _filter_cedula_values(self, widget):
        raw = widget.get()
        normalized = re.sub(r"\D+", "", raw)
        options = self.cedula_options or []
        if normalized:
            filtered = [c for c in options if c and normalized in c]
        else:
            filtered = options
        widget.configure(values=filtered)

    def _format_date_for_ui(self, value):
        if not value:
            return ""
        raw = str(value).strip()
        if len(raw) >= 10 and "-" in raw:
            parts = raw[:10].split("-")
            if len(parts) == 3:
                return f"{parts[2]}/{parts[1]}/{parts[0]}"
        return raw

    def _apply_usuario_data(self, fields, data):
        mapping = {
            "nombre_usuario": "nombre_oferente",
            "genero_usuario": "genero",
            "discapacidad_detalle": "discapacidad",
            "certificado_porcentaje": "certificado_porcentaje",
            "telefono_oferente": "telefono_oferente",
            "fecha_nacimiento": "fecha_nacimiento",
            "cargo_oferente": "cargo_oferente",
            "contacto_emergencia": "contacto_emergencia",
            "parentesco": "parentesco",
            "telefono_emergencia": "telefono_emergencia",
            "correo_oferente": "correo_oferente",
            "lgtbiq": "lgtbiq",
            "grupo_etnico": "grupo_etnico",
            "grupo_etnico_cual": "grupo_etnico_cual",
            "certificado_discapacidad": "certificado_discapacidad",
            "lugar_firma_contrato": "lugar_firma_contrato",
            "fecha_firma_contrato": "fecha_firma_contrato",
            "tipo_contrato": "tipo_contrato",
            "fecha_fin": "fecha_fin",
        }
        for supa_key, field_id in mapping.items():
            value = data.get(supa_key)
            if value in (None, ""):
                continue
            widget = fields.get(field_id)
            if not widget:
                continue
            if supa_key in {"fecha_nacimiento", "fecha_firma_contrato"}:
                value = self._format_date_for_ui(value)
            if supa_key == "discapacidad_detalle" and not value:
                continue
            if isinstance(widget, ttk.Combobox):
                widget.set(str(value))
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(value))
        fecha_widget = fields.get("fecha_nacimiento")
        edad_widget = fields.get("edad")
        if fecha_widget and edad_widget:
            self._refresh_age_from_date(fecha_widget, edad_widget)

    def _on_cedula_selected(self, fields, widget):
        cedula = widget.get().strip()
        if not cedula:
            return
        normalized = re.sub(r"\D+", "", cedula)
        if normalized and normalized != cedula:
            widget.delete(0, tk.END)
            widget.insert(0, normalized)
        try:
            data = contratacion_incluyente.get_usuario_reca_by_cedula(normalized)
        except Exception:
            return
        if data:
            self._apply_usuario_data(fields, data)

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "contratacion_incluyente",
            contratacion_incluyente,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_6": self._show_section_6,
                "section_7": self._show_section_7,
            },
            self._show_section_1,
        ):
            return True
        if not contratacion_incluyente.cache_file_exists():
            return False
        resume = messagebox.askyesno(
            "Reanudar",
            "Se encontró un formulario en progreso. ¿Deseas continuar donde lo dejaste?",
        )
        if not resume:
            contratacion_incluyente.clear_cache_file()
            contratacion_incluyente.clear_form_cache()
            return False
        contratacion_incluyente.load_cache_from_file()
        last_section = contratacion_incluyente.get_form_cache().get("_last_section")
        if last_section == "section_1":
            self._show_section_2()
        elif last_section == "section_2":
            self._show_section_6()
        elif last_section == "section_6":
            self._show_section_6()
        elif last_section == "section_7":
            self._show_section_7()
        else:
            self._show_section_1()
        return True

    def _show_section_1(self):
        self._clear_section_container()
        self.header_title.config(text="1. DATOS DE LA EMPRESA")
        self.header_subtitle.config(text="Busca empresa por NIT y confirma datos.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._build_actions(content)

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. DATOS DEL VINCULADO")
        self.header_subtitle.config(
            text="Completa la información del oferente. Puedes agregar más oferentes."
        )
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        self._load_cedula_options()
        content = _build_scrollable_content(section_frame, self)

        self.oferente_blocks = []
        self.oferente_frames = []
        self.section2_shared_desarrollo_widget = None

        def _add_fields_grid(parent, field_specs, columns=2):
            fields = {}
            for idx, spec in enumerate(field_specs):
                label = spec["label"]
                field_id = spec["id"]
                options = spec.get("options")
                width = spec.get("width", 28)
                row = idx // columns
                col = (idx % columns) * 2
                tk.Label(
                    parent,
                    text=label,
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=row, column=col, sticky="w", padx=(0, 6), pady=4)
                if field_id == "cedula":
                    widget = ttk.Combobox(
                        parent,
                        values=self.cedula_options,
                        state="normal",
                        width=width,
                    )
                elif options:
                    widget = ttk.Combobox(
                        parent,
                        values=options,
                        state="readonly",
                        width=width,
                    )
                else:
                    widget = tk.Entry(parent, width=width)
                widget.grid(row=row, column=col + 1, sticky="w", padx=(0, 12), pady=4)
                if field_id == "cedula":
                    self._apply_numeric_entry(widget)
                elif not options:
                    if field_id == "certificado_porcentaje":
                        self._apply_decimal_entry(widget)
                    elif field_id == "cedula":
                        self._apply_numeric_entry(widget)
                    if field_id in {"telefono_oferente", "telefono_emergencia"}:
                        self._apply_numeric_entry(widget, max_len=10)
                    if field_id in {"nombre_oferente", "contacto_emergencia"}:
                        self._apply_name_entry(widget)
                fields[field_id] = widget
            return fields

        def _add_question_block(parent, title, fields_def):
            frame = tk.Frame(parent, bg="white")
            frame.pack(fill="x", pady=(0, 8))
            tk.Label(
                frame,
                text=title,
                font=("Arial", 9, "bold"),
                bg="white",
            ).pack(anchor="w", padx=6, pady=(6, 4))
            inner = tk.Frame(frame, bg="white")
            inner.pack(fill="x", padx=6, pady=(0, 6))
            return _add_fields_grid(inner, fields_def, columns=2)

        def _add_oferente_block():
            idx = len(self.oferente_blocks) + 1
            block = tk.Frame(content, bg="white", bd=1, relief="solid")
            block.pack(fill="x", pady=8)
            self.oferente_frames.append(block)

            header = tk.Label(
                block,
                text=f"Oferente {idx}",
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
            )
            header.pack(anchor="w", padx=10, pady=(8, 4))

            fields = {}

            section2_frame = tk.LabelFrame(
                block,
                text="2. DATOS DEL VINCULADO",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section2_frame.pack(fill="x", padx=8, pady=(0, 8))

            row1 = tk.Frame(section2_frame, bg="white")
            row1.pack(fill="x", pady=(0, 6))
            fields.update(
                _add_fields_grid(
                    row1,
                    [
                        {"id": "numero", "label": "No", "width": 6},
                        {"id": "nombre_oferente", "label": "Nombre oferente", "width": 24},
                        {"id": "cedula", "label": "Cédula", "width": 14},
                        {"id": "certificado_porcentaje", "label": "Certificado %", "width": 10},
                        {
                            "id": "discapacidad",
                            "label": "Discapacidad",
                            "options": contratacion_incluyente.DISCAPACIDAD_OPTIONS,
                            "width": 26,
                        },
                        {"id": "telefono_oferente", "label": "Teléfono oferente", "width": 14},
                    ],
                    columns=3,
                )
            )

            row2 = tk.Frame(section2_frame, bg="white")
            row2.pack(fill="x", pady=(0, 6))
            fields.update(
                _add_fields_grid(
                    row2,
                    [
                        {
                            "id": "genero",
                            "label": "Género",
                            "options": contratacion_incluyente.GENERO_OPTIONS,
                            "width": 12,
                        },
                        {"id": "correo_oferente", "label": "Email", "width": 26},
                        {"id": "fecha_nacimiento", "label": "Fecha de nacimiento", "width": 12},
                        {"id": "edad", "label": "Edad", "width": 6},
                    ],
                    columns=2,
                )
            )
            if "edad" in fields:
                fields["edad"].configure(state="readonly")
            if "fecha_nacimiento" in fields and "edad" in fields:
                self._apply_date_entry(fields["fecha_nacimiento"], fields["edad"])
            cedula_widget = fields.get("cedula")
            if isinstance(cedula_widget, ttk.Combobox):
                cedula_widget.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, f=fields, w=cedula_widget: self._on_cedula_selected(f, w),
                )
                cedula_widget.bind(
                    "<KeyRelease>",
                    lambda _e, w=cedula_widget: self._filter_cedula_values(w),
                )
                cedula_widget.bind(
                    "<FocusOut>",
                    lambda _e, f=fields, w=cedula_widget: self._on_cedula_selected(f, w),
                )
                cedula_widget.bind(
                    "<Return>",
                    lambda _e, f=fields, w=cedula_widget: self._on_cedula_selected(f, w),
                )

            row3 = tk.Frame(section2_frame, bg="white")
            row3.pack(fill="x", pady=(0, 6))
            fields.update(
                _add_fields_grid(
                    row3,
                    [
                        {
                            "id": "lgtbiq",
                            "label": "LGTBIQ",
                            "options": contratacion_incluyente.LGTBIQ_OPTIONS,
                            "width": 16,
                        },
                        {
                            "id": "grupo_etnico",
                            "label": "Grupo étnico",
                            "options": contratacion_incluyente.GRUPO_ETNICO_OPTIONS,
                            "width": 16,
                        },
                        {
                            "id": "grupo_etnico_cual",
                            "label": "¿Cuál?",
                            "options": contratacion_incluyente.GRUPO_ETNICO_CUAL_OPTIONS,
                            "width": 20,
                        },
                    ],
                    columns=3,
                )
            )

            row4 = tk.Frame(section2_frame, bg="white")
            row4.pack(fill="x", pady=(0, 6))
            fields.update(
                _add_fields_grid(
                    row4,
                    [
                        {"id": "cargo_oferente", "label": "Cargo", "width": 18},
                        {"id": "contacto_emergencia", "label": "Contacto de emergencia", "width": 18},
                        {"id": "parentesco", "label": "Parentesco", "width": 12},
                        {"id": "telefono_emergencia", "label": "Teléfono", "width": 12},
                    ],
                    columns=2,
                )
            )

            row5 = tk.Frame(section2_frame, bg="white")
            row5.pack(fill="x")
            fields.update(
                _add_fields_grid(
                    row5,
                    [
                        {
                            "id": "certificado_discapacidad",
                            "label": "Certificado discapacidad",
                            "options": contratacion_incluyente.CERTIFICADO_DISCAPACIDAD_OPTIONS,
                            "width": 16,
                        },
                        {"id": "lugar_firma_contrato", "label": "Lugar de firma de contrato", "width": 18},
                        {"id": "fecha_firma_contrato", "label": "Fecha de firma de contrato", "width": 16},
                    ],
                    columns=3,
                )
            )

            section3_frame = tk.LabelFrame(
                block,
                text="3. DATOS ADICIONALES",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section3_frame.pack(fill="x", padx=8, pady=(0, 8))
            fields.update(
                _add_fields_grid(
                    section3_frame,
                    [
                        {
                            "id": "tipo_contrato",
                            "label": "Tipo de contrato",
                            "options": contratacion_incluyente.TIPO_CONTRATO_OPTIONS,
                            "width": 24,
                        },
                        {"id": "fecha_fin", "label": "Fecha de fin", "width": 14},
                    ],
                    columns=2,
                )
            )

            if idx == 1:
                section4_frame = tk.LabelFrame(
                    block,
                    text="4. DESARROLLO DE LA ACTIVIDAD",
                    bg="white",
                    fg="#222222",
                    font=FONT_LABEL,
                    padx=8,
                    pady=6,
                )
                section4_frame.pack(fill="x", padx=8, pady=(0, 8))
                fields["desarrollo_actividad"] = tk.Text(section4_frame, height=5, wrap="word")
                fields["desarrollo_actividad"].pack(fill="x", padx=6, pady=6)
                self.section2_shared_desarrollo_widget = fields["desarrollo_actividad"]

            section51_frame = tk.LabelFrame(
                block,
                text="5.1 CONDICIONES DE LA VACANTE",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section51_frame.pack(fill="x", padx=8, pady=(0, 8))

            fields.update(
                _add_question_block(
                    section51_frame,
                    "¿El vinculado lee el contrato de forma independiente?",
                    [
                        {
                            "id": "contrato_lee_nivel_apoyo",
                            "label": "Nivel de apoyo",
                            "options": contratacion_incluyente.NIVEL_APOYO_OPTIONS,
                            "width": 24,
                        },
                        {
                            "id": "contrato_lee_observacion",
                            "label": "Observación",
                            "options": contratacion_incluyente.OBS_LECTURA_CONTRATO_OPTIONS,
                            "width": 50,
                        },
                        {"id": "contrato_lee_nota", "label": "Nota", "width": 50},
                    ],
                )
            )
            fields.update(
                _add_question_block(
                    section51_frame,
                    "¿El contrato fue comprendido por el vinculado?",
                    [
                        {
                            "id": "contrato_comprendido_nivel_apoyo",
                            "label": "Nivel de apoyo",
                            "options": contratacion_incluyente.NIVEL_APOYO_OPTIONS,
                            "width": 24,
                        },
                        {
                            "id": "contrato_comprendido_observacion",
                            "label": "Observación",
                            "options": contratacion_incluyente.OBS_COMPRENDE_CONTRATO_OPTIONS,
                            "width": 50,
                        },
                        {"id": "contrato_comprendido_nota", "label": "Nota", "width": 50},
                    ],
                )
            )
            fields.update(
                _add_question_block(
                    section51_frame,
                    "¿Es claro para el vinculado el tipo de contrato a firmar?",
                    [
                        {
                            "id": "contrato_tipo_nivel_apoyo",
                            "label": "Nivel de apoyo",
                            "options": contratacion_incluyente.NIVEL_APOYO_OPTIONS,
                            "width": 24,
                        },
                        {
                            "id": "contrato_tipo_observacion",
                            "label": "Observación",
                            "options": contratacion_incluyente.OBS_TIPO_CONTRATO_OPTIONS,
                            "width": 50,
                        },
                        {
                            "id": "contrato_tipo_contrato",
                            "label": "Tipo de contrato",
                            "options": contratacion_incluyente.TIPO_CONTRATO_OPTIONS,
                            "width": 28,
                        },
                        {
                            "id": "contrato_jornada",
                            "label": "Jornada laboral",
                            "options": contratacion_incluyente.JORNADA_LABORAL_OPTIONS,
                            "width": 20,
                        },
                        {
                            "id": "contrato_clausulas",
                            "label": "Cláusulas",
                            "options": contratacion_incluyente.CLAUSULAS_CONTRATO_OPTIONS,
                            "width": 30,
                        },
                        {"id": "contrato_tipo_nota", "label": "Nota", "width": 50},
                    ],
                )
            )
            fields.update(
                _add_question_block(
                    section51_frame,
                    "Explicación de las condiciones salariales",
                    [
                        {
                            "id": "condiciones_salariales_nivel_apoyo",
                            "label": "Nivel de apoyo",
                            "options": contratacion_incluyente.NIVEL_APOYO_OPTIONS,
                            "width": 24,
                        },
                        {
                            "id": "condiciones_salariales_observacion",
                            "label": "Observación",
                            "options": contratacion_incluyente.OBS_CONDICIONES_SALARIALES_OPTIONS,
                            "width": 50,
                        },
                        {
                            "id": "condiciones_salariales_frecuencia_pago",
                            "label": "Frecuencia de pago",
                            "options": contratacion_incluyente.FRECUENCIA_PAGO_OPTIONS,
                            "width": 18,
                        },
                        {
                            "id": "condiciones_salariales_forma_pago",
                            "label": "Forma de pago",
                            "options": contratacion_incluyente.FORMA_PAGO_OPTIONS,
                            "width": 18,
                        },
                        {"id": "condiciones_salariales_nota", "label": "Nota", "width": 50},
                    ],
                )
            )

            section52_frame = tk.LabelFrame(
                block,
                text="5.2 PRESTACIONES DE LEY",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section52_frame.pack(fill="x", padx=8, pady=(0, 8))

            prestaciones = [
                ("Cesantías", "prestaciones_cesantias"),
                ("Auxilios de transporte", "prestaciones_auxilio_transporte"),
                ("Prima", "prestaciones_prima"),
                ("Seguridad Social (EPS, Pensión y ARL)", "prestaciones_seguridad_social"),
                ("Vacaciones", "prestaciones_vacaciones"),
                ("Auxilios y otros beneficios", "prestaciones_auxilios_beneficios"),
            ]
            for label, key_prefix in prestaciones:
                fields.update(
                    _add_question_block(
                        section52_frame,
                        label,
                        [
                            {
                                "id": f"{key_prefix}_nivel_apoyo",
                                "label": "Nivel de apoyo",
                                "options": contratacion_incluyente.NIVEL_APOYO_OPTIONS,
                                "width": 24,
                            },
                            {
                                "id": f"{key_prefix}_observacion",
                                "label": "Observación",
                                "options": contratacion_incluyente.OBS_PRESTACIONES_OPTIONS,
                                "width": 50,
                            },
                            {"id": f"{key_prefix}_nota", "label": "Nota", "width": 50},
                        ],
                    )
                )

            section53_frame = tk.LabelFrame(
                block,
                text="5.3 DEBERES Y DERECHOS DEL TRABAJADOR",
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section53_frame.pack(fill="x", padx=8, pady=(0, 8))

            fields.update(
                _add_question_block(
                    section53_frame,
                    "¿El vinculado tiene claro el conducto regular?",
                    [
                        {
                            "id": "conducto_regular_nivel_apoyo",
                            "label": "Nivel de apoyo",
                            "options": contratacion_incluyente.NIVEL_APOYO_OPTIONS,
                            "width": 24,
                        },
                        {
                            "id": "conducto_regular_observacion",
                            "label": "Conducto regular",
                            "options": contratacion_incluyente.OBS_CONDUCTO_REGULAR_OPTIONS,
                            "width": 50,
                        },
                        {
                            "id": "descargos_observacion",
                            "label": "Descargos",
                            "options": contratacion_incluyente.OBS_DESCARGOS_OPTIONS,
                            "width": 50,
                        },
                        {
                            "id": "tramites_observacion",
                            "label": "Trámites administrativos",
                            "options": contratacion_incluyente.OBS_TRAMITES_OPTIONS,
                            "width": 50,
                        },
                        {
                            "id": "permisos_observacion",
                            "label": "Permisos",
                            "options": contratacion_incluyente.OBS_PERMISOS_OPTIONS,
                            "width": 50,
                        },
                        {"id": "conducto_regular_nota", "label": "Nota", "width": 50},
                    ],
                )
            )
            fields.update(
                _add_question_block(
                    section53_frame,
                    "¿El vinculado tiene claras las causales de finalización de contrato?",
                    [
                        {
                            "id": "causales_fin_nivel_apoyo",
                            "label": "Nivel de apoyo",
                            "options": contratacion_incluyente.NIVEL_APOYO_OPTIONS,
                            "width": 24,
                        },
                        {
                            "id": "causales_fin_observacion",
                            "label": "Observación",
                            "options": contratacion_incluyente.OBS_CAUSALES_OPTIONS,
                            "width": 50,
                        },
                        {"id": "causales_fin_nota", "label": "Nota", "width": 50},
                    ],
                )
            )
            fields.update(
                _add_question_block(
                    section53_frame,
                    "¿El vinculado conoce las rutas de atención y/o denuncia?",
                    [
                        {
                            "id": "rutas_atencion_nivel_apoyo",
                            "label": "Nivel de apoyo",
                            "options": contratacion_incluyente.NIVEL_APOYO_OPTIONS,
                            "width": 24,
                        },
                        {
                            "id": "rutas_atencion_observacion",
                            "label": "Observación",
                            "options": contratacion_incluyente.OBS_RUTAS_OPTIONS,
                            "width": 50,
                        },
                        {"id": "rutas_atencion_nota", "label": "Nota", "width": 50},
                    ],
                )
            )

            numero_widget = fields.get("numero")
            if numero_widget:
                numero_widget.delete(0, tk.END)
                numero_widget.insert(0, str(idx))
                numero_widget.configure(state="readonly")

            self.oferente_blocks.append(fields)

        def _refresh_oferente_numbers():
            for idx, fields in enumerate(self.oferente_blocks, start=1):
                numero_widget = fields.get("numero")
                if numero_widget:
                    numero_widget.configure(state="normal")
                    numero_widget.delete(0, tk.END)
                    numero_widget.insert(0, str(idx))
                    numero_widget.configure(state="readonly")

        def _update_remove_button_state():
            if remove_btn is None:
                return
            if len(self.oferente_blocks) <= 1:
                remove_btn.configure(state="disabled")
            else:
                remove_btn.configure(state="normal")

        def _remove_oferente_block():
            if len(self.oferente_blocks) <= 1:
                return
            self.oferente_blocks.pop()
            frame = self.oferente_frames.pop()
            frame.destroy()
            _refresh_oferente_numbers()
            _update_remove_button_state()

        def _prefill_section_2():
            cache = contratacion_incluyente.get_form_cache().get("section_2", [])
            if not cache:
                _add_oferente_block()
                return
            for _ in range(len(cache)):
                _add_oferente_block()
            for idx, entry in enumerate(cache):
                fields = self.oferente_blocks[idx]
                for key, widget in fields.items():
                    value = entry.get(key, "")
                    if isinstance(widget, ttk.Combobox):
                        widget.set(value)
                    elif isinstance(widget, tk.Text):
                        widget.delete("1.0", tk.END)
                        widget.insert("1.0", value)
                    else:
                        widget.delete(0, tk.END)
                        widget.insert(0, value)
            shared_desarrollo = ""
            for entry in cache:
                shared_desarrollo = (entry.get("desarrollo_actividad") or "").strip()
                if shared_desarrollo:
                    break
            if self.section2_shared_desarrollo_widget is not None:
                self.section2_shared_desarrollo_widget.delete("1.0", tk.END)
                if shared_desarrollo:
                    self.section2_shared_desarrollo_widget.insert("1.0", shared_desarrollo)

        _prefill_section_2()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Agregar oferente", command=_add_oferente_block).pack(
            side="left"
        )
        remove_btn = ttk.Button(
            actions, text="Eliminar ultimo oferente", command=_remove_oferente_block
        )
        remove_btn.pack(side="left", padx=8)
        _update_remove_button_state()
        ttk.Button(actions, text="Regresar", command=self._show_section_1).pack(
            side="left", padx=8
        )
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2).pack(
            side="right"
        )

    def _show_section_6(self):
        self._clear_section_container()
        self.header_title.config(text="6. AJUSTES RAZONABLES")
        self.header_subtitle.config(text="Completa ajustes razonables.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="both", expand=True)

        self.section6_fields = {}

        tk.Label(
            content,
            text="Ajustes razonables / recomendaciones",
            font=FONT_SECTION,
            bg=COLOR_LIGHT_BG,
            anchor="w",
        ).pack(anchor="w", pady=(8, 4))
        ajustes = tk.Text(content, height=6, width=TEXT_WIDE, wrap="word")
        ajustes.pack(fill="x", padx=4, pady=(0, 16))
        self.section6_fields["ajustes_recomendaciones"] = ajustes

        cache = contratacion_incluyente.get_form_cache().get("section_6", {})
        if cache:
            ajustes.delete("1.0", tk.END)
            ajustes.insert("1.0", cache.get("ajustes_recomendaciones", ""))

        _build_wizard_actions(
            content,
            back_command=self._show_section_2,
            primary_command=self._confirm_section_6,
        )

    def _confirm_section_2(self):
        shared_desarrollo = ""
        if isinstance(getattr(self, "section2_shared_desarrollo_widget", None), tk.Text):
            shared_desarrollo = self.section2_shared_desarrollo_widget.get("1.0", tk.END).strip()
        payload = []
        for fields in self.oferente_blocks:
            entry = {}
            for key, widget in fields.items():
                if isinstance(widget, ttk.Combobox):
                    entry[key] = widget.get().strip()
                elif isinstance(widget, tk.Text):
                    entry[key] = widget.get("1.0", tk.END).strip()
                else:
                    entry[key] = widget.get().strip()
            entry["desarrollo_actividad"] = shared_desarrollo
            payload.append(entry)
        try:
            contratacion_incluyente.confirm_section_2(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_6()

    def _confirm_section_6(self):
        payload = {
            "ajustes_recomendaciones": self.section6_fields["ajustes_recomendaciones"]
            .get("1.0", tk.END)
            .strip(),
        }
        try:
            contratacion_incluyente.confirm_section_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_7()

    def _show_section_7(self):
        self._clear_section_container()
        self.header_title.config(text="7. ASISTENTES")
        self.header_subtitle.config(text="Registra asistentes y agrega filas si aplica.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="both", expand=True)

        table = tk.Frame(content, bg=COLOR_LIGHT_BG)
        table.pack(fill="x", padx=4, pady=(0, 8))

        tk.Label(
            table,
            text="Nombre completo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        tk.Label(
            table,
            text="Cargo",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=1, sticky="w")

        self.section7_rows = []
        asistentes_catalog = _get_asistentes_profesionales_catalog()

        def _add_asistente_row():
            row_idx = len(self.section7_rows) + 1
            nombre_entry, cargo_entry = _create_asistente_inputs(
                table,
                ENTRY_W_WIDE,
                use_catalog=(len(self.section7_rows) == 0),
                catalog=asistentes_catalog,
            )
            nombre_entry.grid(row=row_idx, column=0, sticky="w", pady=4, padx=(0, 12))
            cargo_entry.grid(row=row_idx, column=1, sticky="w", pady=4)
            self.section7_rows.append((nombre_entry, cargo_entry))

        _add_asistente_row()
        _add_asistente_row()
        _add_asistente_row()

        cached_rows = contratacion_incluyente.get_form_cache().get("section_7", [])
        for idx, entry in enumerate(cached_rows):
            if idx >= len(self.section7_rows):
                _add_asistente_row()
            nombre_entry, cargo_entry = self.section7_rows[idx]
            nombre_entry.delete(0, tk.END)
            nombre_entry.insert(0, entry.get("nombre", ""))
            cargo_entry.delete(0, tk.END)
            cargo_entry.insert(0, entry.get("cargo", ""))

        _build_wizard_actions(
            content,
            back_command=self._show_section_6,
            primary_command=self._confirm_section_7,
            primary_text="Finalizar",
            left_buttons=[("Agregar asistente", _add_asistente_row)],
        )

    def _confirm_section_7(self):
        asistentes = []
        for nombre_entry, cargo_entry in self.section7_rows:
            nombre = nombre_entry.get().strip()
            cargo = cargo_entry.get().strip()
            asistentes.append({"nombre": nombre, "cargo": cargo})
        try:
            contratacion_incluyente.confirm_section_7(asistentes)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Generando Excel...")
        loading.set_progress(40)
        cache = contratacion_incluyente.get_form_cache()
        section_1 = cache.get("section_1", {})
        company_name = section_1.get("nombre_empresa")

        def _worker():
            output_path = _raise_finalize_stage(
                "generando el Excel",
                lambda: contratacion_incluyente.export_to_excel(clear_cache=False),
            )
            _update_loading_async(
                loading,
                status="Guardando en Supabase...",
                progress=70,
            )
            _raise_finalize_stage(
                "guardando en Supabase",
                contratacion_incluyente.sync_usuarios_reca,
            )
            return output_path

        _start_background_finalization(
            self,
            loading,
            form_name="Contratacion Incluyente",
            company_name=company_name,
            form_id="contratacion_incluyente",
            worker_fn=_worker,
            post_delivery_fn=lambda: _clear_form_cache_safe(contratacion_incluyente),
        )

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        groups = [
            ('Información de Empresa', COLOR_GROUP_EMPRESA, ['nombre_empresa', 'direccion_empresa', 'correo_1', 'contacto_empresa', 'telefono_empresa', 'cargo', 'ciudad_empresa', 'sede_empresa', 'caja_compensacion']),
            ('Información de Compensar', COLOR_GROUP_COMPENSAR, ['asesor']),
            ('Información de RECA', COLOR_GROUP_RECA, ['profesional_asignado']),
        ]
        labels = {
            'nombre_empresa': 'Nombre de la empresa',
            'direccion_empresa': 'Dirección de la empresa',
            'correo_1': 'Correo electrónico',
            'contacto_empresa': 'Contacto de la empresa',
            'telefono_empresa': 'Teléfonos',
            'cargo': 'Cargo',
            'ciudad_empresa': 'Ciudad/Municipio',
            'sede_empresa': 'Sede Compensar',
            'caja_compensacion': 'Empresa afiliada a Caja de Compensación',
            'asesor': 'Asesor',
            'profesional_asignado': 'Profesional asignado RECA',
        }
        _section1_build_groups(self, parent, groups, labels)

    def _build_actions(self, parent):
        _section1_build_actions(self, parent)

    def _label_for_field(self, field_id):
        return getattr(self, '_section1_labels', {}).get(field_id, field_id)

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _search_company(self, mode="nit"):
        nit = self.fields["nit_empresa"].get().strip()
        nombre = self.fields.get("nombre_busqueda").get().strip() if self.fields.get("nombre_busqueda") else ""
        if mode == "nit":
            if not nit:
                messagebox.showerror("Error", "Ingresa un NIT.")
                return
        elif mode == "nombre":
            if not nombre:
                messagebox.showerror("Error", "Ingresa el nombre de la empresa.")
                return
        else:
            messagebox.showerror("Error", "Tipo de búsqueda no válido.")
            return

        try:
            self.status_label.config(text="Buscando empresa...")
            self.update_idletasks()
            if mode == "nombre":
                company = contratacion_incluyente.get_empresa_by_nombre(nombre)
            else:
                company = contratacion_incluyente.get_empresa_by_nit(nit)
        except Exception as exc:
            self.status_label.config(text="")
            messagebox.showerror("Error", str(exc))
            return

        if not company:
            self.company_data = None
            msg = "No se encontró empresa para ese nombre." if mode == "nombre" else "No se encontró empresa para ese NIT."
            self.status_label.config(text=msg)
            self.continue_btn.config(state="disabled")
            for key in contratacion_incluyente.SECTION_1_SUPABASE_MAP.keys():
                self._set_readonly_value(key, "")
            return

        if mode == "nombre":
            nit_value = company.get("nit_empresa")
            if nit_value:
                entry = self.fields.get("nit_empresa")
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, nit_value)

        self.company_data = company
        self.status_label.config(text="Empresa encontrada.")
        self.continue_btn.config(state="normal")
        for key in contratacion_incluyente.SECTION_1_SUPABASE_MAP.keys():
            self._set_readonly_value(key, company.get(key))

    def _confirm_and_continue(self):
        if not self.company_data:
            messagebox.showerror("Error", "Busca una empresa antes de confirmar.")
            return

        fecha_visita = _get_required_fecha_visita(self)
        if not fecha_visita:
            return
        modalidad = _get_required_modalidad(self)
        if not modalidad:
            return
        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad": modalidad,
            "nit_empresa": self.fields["nit_empresa"].get().strip(),
        }
        try:
            contratacion_incluyente.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2()


class InduccionOrganizacionalWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Induccion Organizacional - Seccion 1")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = induccion_organizacional

        self.company_data = None
        self.fields = {}
        self.cedula_options = []

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))

        self.header_title = tk.Label(
            header,
            text="1. DATOS GENERALES",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        self.header_title.pack(anchor="w")

        self.header_subtitle = tk.Label(
            header,
            text="Busca empresa por NIT y confirma datos.",
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        self.header_subtitle.pack(anchor="w", pady=(4, 0))

    def _build_section_container(self):
        self.section_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)

    def _clear_section_container(self):
        for child in self.section_container.winfo_children():
            child.destroy()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "induccion_organizacional",
            induccion_organizacional,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_3": self._show_section_3,
                "section_4": self._show_section_4,
                "section_5": self._show_section_5,
                "section_6": self._show_section_6,
            },
            self._show_section_1,
        ):
            return True
        if not induccion_organizacional.cache_file_exists():
            return False
        resume = messagebox.askyesno(
            "Reanudar",
            "Se encontró un formulario en progreso. ¿Deseas continuar donde lo dejaste?",
        )
        if not resume:
            induccion_organizacional.clear_cache_file()
            induccion_organizacional.clear_form_cache()
            return False
        induccion_organizacional.load_cache_from_file()
        last_section = induccion_organizacional.get_form_cache().get("_last_section")
        if last_section == "section_6":
            self._show_section_6()
        elif last_section == "section_5":
            self._show_section_5()
        elif last_section == "section_4":
            self._show_section_4()
        elif last_section == "section_3":
            self._show_section_3()
        elif last_section == "section_2":
            self._show_section_2()
        else:
            self._show_section_1()
        return True

    def _show_section_1(self):
        self._clear_section_container()
        self.header_title.config(text="1. DATOS GENERALES")
        self.header_subtitle.config(text="Busca empresa por NIT y confirma datos.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._prefill_section_1()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._close_to_hub).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_and_continue).pack(side="right")

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. DATOS DEL VINCULADO")
        self.header_subtitle.config(text="Registra uno o mas vinculados.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        self._load_cedula_options()
        content = _build_scrollable_content(section_frame, self)

        self.vinculado_blocks = []
        self.vinculado_frames = []

        def _create_widget(parent, field_id, width=30):
            if field_id == "cedula":
                return ttk.Combobox(
                    parent,
                    values=self.cedula_options,
                    state="normal",
                    width=width,
                )
            return tk.Entry(parent, width=width)

        def _add_fields_grid(parent, field_ids):
            fields = {}
            for idx, field_id in enumerate(field_ids):
                row = idx // 2
                col = (idx % 2) * 2
                meta = next(
                    (f for f in induccion_organizacional.SECTION_2["fields"] if f["id"] == field_id),
                    {"label": field_id},
                )
                tk.Label(
                    parent,
                    text=meta["label"],
                    font=FONT_LABEL,
                    bg=COLOR_LIGHT_BG,
                ).grid(row=row, column=col, sticky="w", padx=6, pady=(3, 2))
                widget = _create_widget(parent, field_id)
                widget.grid(row=row, column=col + 1, sticky="we", padx=6, pady=(3, 2))
                if field_id == "cedula":
                    widget.bind("<KeyRelease>", lambda _e, w=widget: self._filter_cedula_values(w))
                fields[field_id] = widget
            parent.grid_columnconfigure(1, weight=1)
            parent.grid_columnconfigure(3, weight=1)
            return fields

        def _apply_usuario_data(fields, data):
            mapping = {
                "nombre_usuario": "nombre_oferente",
                "cedula_usuario": "cedula",
                "telefono_oferente": "telefono_oferente",
                "cargo_oferente": "cargo_oferente",
            }
            for src, dest in mapping.items():
                value = data.get(src)
                if value in (None, ""):
                    continue
                widget = fields.get(dest)
                if not widget:
                    continue
                if isinstance(widget, ttk.Combobox):
                    widget.set(str(value))
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, str(value))

        def _on_cedula_selected(fields, widget):
            raw = widget.get().strip()
            if not raw:
                return
            normalized = re.sub(r"\D+", "", raw)
            if normalized and normalized != raw:
                widget.delete(0, tk.END)
                widget.insert(0, normalized)
            try:
                data = induccion_organizacional.get_usuario_reca_by_cedula(normalized)
            except Exception:
                return
            if data:
                _apply_usuario_data(fields, data)

        def _create_vinculado_block(index):
            card = tk.LabelFrame(
                content,
                text=f"Vinculado #{index + 1}",
                bg=COLOR_LIGHT_BG,
                padx=10,
                pady=8,
            )
            card.pack(fill="x", padx=FORM_PADX, pady=6)

            fields = _add_fields_grid(
                card,
                [
                    "nombre_oferente",
                    "cedula",
                    "telefono_oferente",
                    "cargo_oferente",
                ],
            )
            fields["numero"] = str(index + 1)

            cedula_widget = fields.get("cedula")
            if cedula_widget:
                cedula_widget.bind(
                    "<<ComboboxSelected>>",
                    lambda _e, f=fields, w=cedula_widget: _on_cedula_selected(f, w),
                )
                cedula_widget.bind(
                    "<FocusOut>",
                    lambda _e, f=fields, w=cedula_widget: _on_cedula_selected(f, w),
                )

            self.vinculado_blocks.append(fields)
            self.vinculado_frames.append(card)

        def _remove_last_vinculado():
            if len(self.vinculado_blocks) <= 1:
                return
            frame = self.vinculado_frames.pop()
            frame.destroy()
            self.vinculado_blocks.pop()

        def _add_vinculado():
            _create_vinculado_block(len(self.vinculado_blocks))

        _create_vinculado_block(0)
        cached_rows = induccion_organizacional.get_form_cache().get("section_2", [])
        for idx, row_data in enumerate(cached_rows):
            if idx >= len(self.vinculado_blocks):
                _add_vinculado()
            block = self.vinculado_blocks[idx]
            block["numero"] = str(idx + 1)
            for key in ["nombre_oferente", "cedula", "telefono_oferente", "cargo_oferente"]:
                widget = block.get(key)
                if not widget:
                    continue
                value = row_data.get(key, "")
                if isinstance(widget, ttk.Combobox):
                    widget.set(value)
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, value)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_1).pack(side="left")
        ttk.Button(actions, text="Agregar vinculado", command=_add_vinculado).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Eliminar ultimo", command=_remove_last_vinculado).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2).pack(
            side="right"
        )

    def _show_section_3(self):
        self._clear_section_container()
        self.header_title.config(text="3. DESARROLLO DEL PROCESO")
        self.header_subtitle.config(
            text="Completa visto, responsable, medio de socializacion y descripcion por cada tematica.",
        )

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        self.section3_fields = {}
        cached = induccion_organizacional.get_form_cache().get("section_3", {})

        for subsection in induccion_organizacional.SECTION_3["subsections"]:
            section_box = tk.LabelFrame(
                content,
                text=subsection["title"],
                bg=COLOR_LIGHT_BG,
                padx=10,
                pady=8,
            )
            section_box.pack(fill="x", padx=FORM_PADX, pady=8)
            section_box.grid_columnconfigure(0, weight=2)
            section_box.grid_columnconfigure(1, weight=1)
            section_box.grid_columnconfigure(2, weight=1)
            section_box.grid_columnconfigure(3, weight=1)
            section_box.grid_columnconfigure(4, weight=2)

            tk.Label(section_box, text="Tematica", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=0, column=0, sticky="w", padx=4, pady=(0, 6)
            )
            tk.Label(section_box, text="Visto", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=0, column=1, sticky="w", padx=4, pady=(0, 6)
            )
            tk.Label(section_box, text="Responsable", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=0, column=2, sticky="w", padx=4, pady=(0, 6)
            )
            tk.Label(
                section_box, text="Medio de socializacion", bg=COLOR_LIGHT_BG, font=FONT_LABEL
            ).grid(row=0, column=3, sticky="w", padx=4, pady=(0, 6))
            tk.Label(section_box, text="Descripción", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=0, column=4, sticky="w", padx=4, pady=(0, 6)
            )

            for idx, item in enumerate(subsection["items"], start=1):
                tk.Label(
                    section_box,
                    text=item["label"],
                    bg=COLOR_LIGHT_BG,
                    justify="left",
                    anchor="w",
                    wraplength=340,
                ).grid(row=idx, column=0, sticky="w", padx=4, pady=4)

                visto = ttk.Combobox(
                    section_box,
                    values=induccion_organizacional.VISTO_OPTIONS,
                    state="readonly",
                    width=14,
                )
                visto.grid(row=idx, column=1, sticky="we", padx=4, pady=4)

                responsable = tk.Entry(section_box, width=24)
                responsable.grid(row=idx, column=2, sticky="we", padx=4, pady=4)

                medio = ttk.Combobox(
                    section_box,
                    values=induccion_organizacional.MEDIO_SOCIALIZACION_OPTIONS,
                    state="readonly",
                    width=20,
                )
                medio.grid(row=idx, column=3, sticky="we", padx=4, pady=4)

                descripcion = tk.Entry(section_box, width=36)
                descripcion.grid(row=idx, column=4, sticky="we", padx=4, pady=4)

                item_cache = cached.get(item["id"], {}) if isinstance(cached, dict) else {}
                visto.set(item_cache.get("visto", ""))
                responsable.insert(0, item_cache.get("responsable", ""))
                medio.set(item_cache.get("medio_socializacion", ""))
                descripcion.insert(0, item_cache.get("descripcion", ""))

                self.section3_fields[item["id"]] = {
                    "visto": visto,
                    "responsable": responsable,
                    "medio_socializacion": medio,
                    "descripcion": descripcion,
                }

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_3).pack(side="right")

    def _show_section_4(self):
        self._clear_section_container()
        self.header_title.config(text="4. AJUSTES RAZONABLES AL PROCESO DE INDUCCION")
        self.header_subtitle.config(
            text="Selecciona el medio y se autocompleta la recomendacion.",
        )

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        self.section4_rows = []
        cached = induccion_organizacional.get_form_cache().get("section_4", [])
        row_labels = ["Ajuste 1", "Ajuste 2", "Ajuste 3"]

        def _on_medio_change(index):
            medio_widget, text_widget = self.section4_rows[index]
            medio = medio_widget.get().strip()
            recommendation = induccion_organizacional.SECTION_4_RECOMMENDATIONS.get(medio, "")
            text_widget.delete("1.0", tk.END)
            if recommendation:
                text_widget.insert("1.0", recommendation)

        for i in range(3):
            card = tk.LabelFrame(
                section_frame,
                text=row_labels[i],
                bg=COLOR_LIGHT_BG,
                padx=10,
                pady=8,
            )
            card.pack(fill="x", padx=FORM_PADX, pady=8)

            tk.Label(card, text="Medio", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=0, column=0, sticky="w", padx=4, pady=4
            )
            medio = ttk.Combobox(
                card,
                values=induccion_organizacional.SECTION_4_OPTIONS,
                state="readonly",
                width=65,
            )
            medio.grid(row=0, column=1, sticky="w", padx=4, pady=4)

            tk.Label(card, text="Recomendacion", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=1, column=0, sticky="nw", padx=4, pady=4
            )
            texto = tk.Text(card, width=95, height=8, wrap="word")
            texto.grid(row=1, column=1, sticky="we", padx=4, pady=4)

            medio.bind("<<ComboboxSelected>>", lambda _e, idx=i: _on_medio_change(idx))

            self.section4_rows.append((medio, texto))

            cached_entry = cached[i] if isinstance(cached, list) and i < len(cached) else {}
            medio_value = (cached_entry.get("medio") or "").strip()
            rec_value = (cached_entry.get("recomendacion") or "").strip()
            if medio_value:
                medio.set(medio_value)
            if rec_value:
                texto.insert("1.0", rec_value)
            elif medio_value:
                auto = induccion_organizacional.SECTION_4_RECOMMENDATIONS.get(medio_value, "")
                if auto:
                    texto.insert("1.0", auto)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_3).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_4).pack(side="right")

    def _show_section_5(self):
        self._clear_section_container()
        self.header_title.config(text="5. OBSERVACIONES")
        self.header_subtitle.config(text="Registra observaciones del proceso.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        tk.Label(
            section_frame,
            text="Observaciones",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", padx=FORM_PADX, pady=(8, 4))

        self.section5_text = tk.Text(section_frame, width=120, height=10, wrap="word")
        self.section5_text.pack(fill="x", padx=FORM_PADX, pady=(0, 8))

        cache = induccion_organizacional.get_form_cache().get("section_5", {})
        if cache.get("observaciones"):
            self.section5_text.insert("1.0", cache.get("observaciones", ""))

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_4).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_5).pack(side="right")

    def _show_section_6(self):
        self._clear_section_container()
        self.header_title.config(text="6. ASISTENTES")
        self.header_subtitle.config(text="Registra asistentes y agrega filas si aplica.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="x", padx=FORM_PADX, pady=(8, 8))

        self.section6_rows = []
        asistentes_catalog = _get_asistentes_profesionales_catalog()

        def _add_row(nombre="", cargo=""):
            row = tk.Frame(content, bg=COLOR_LIGHT_BG)
            row.pack(fill="x", pady=4)
            tk.Label(row, text="Nombre completo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 6)
            )
            nombre_entry, cargo_entry = _create_asistente_inputs(
                row,
                50,
                use_catalog=(len(self.section6_rows) == 0),
                catalog=asistentes_catalog,
            )
            nombre_entry.pack(side="left", padx=(0, 12))
            tk.Label(row, text="Cargo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 6)
            )
            cargo_entry.pack(side="left")
            if nombre:
                nombre_entry.insert(0, nombre)
            if cargo:
                cargo_entry.insert(0, cargo)
            self.section6_rows.append((row, nombre_entry, cargo_entry))

        def _remove_last():
            if len(self.section6_rows) <= 1:
                return
            row, _, _ = self.section6_rows.pop()
            row.destroy()

        cached_rows = induccion_organizacional.get_form_cache().get("section_6", [])
        if cached_rows:
            for item in cached_rows:
                _add_row(item.get("nombre", ""), item.get("cargo", ""))
        else:
            for _ in range(4):
                _add_row()

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_5).pack(side="left")
        ttk.Button(actions, text="Agregar asistente", command=_add_row).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Eliminar ultimo", command=_remove_last).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Finalizar", command=self._confirm_section_6).pack(side="right")

    def _load_cedula_options(self):
        try:
            self.cedula_options = induccion_organizacional.get_usuarios_reca_cedulas()
        except Exception:
            self.cedula_options = []

    def _filter_cedula_values(self, widget):
        raw = widget.get()
        normalized = re.sub(r"\D+", "", raw)
        options = self.cedula_options or []
        if normalized:
            filtered = [c for c in options if c and normalized in c]
        else:
            filtered = options
        widget.configure(values=filtered)

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        groups = [
            (
                "Información de Empresa",
                COLOR_GROUP_EMPRESA,
                [
                    "nombre_empresa",
                    "direccion_empresa",
                    "correo_1",
                    "contacto_empresa",
                    "telefono_empresa",
                    "cargo",
                    "ciudad_empresa",
                    "sede_empresa",
                    "caja_compensacion",
                ],
            ),
            ("Información de Compensar", COLOR_GROUP_COMPENSAR, ["asesor"]),
            ("Información de RECA", COLOR_GROUP_RECA, ["profesional_asignado"]),
        ]
        labels = {
            "nombre_empresa": "Nombre de la empresa",
            "direccion_empresa": "Dirección de la empresa",
            "correo_1": "Correo electrónico",
            "contacto_empresa": "Persona que atiende la visita",
            "telefono_empresa": "Teléfonos",
            "cargo": "Cargo",
            "ciudad_empresa": "Ciudad/Municipio",
            "sede_empresa": "Sede Compensar",
            "caja_compensacion": "Empresa afiliada a Caja de Compensación",
            "asesor": "Asesor",
            "profesional_asignado": "Profesional asignado RECA",
        }
        _section1_build_groups(self, parent, groups, labels)

    def _label_for_field(self, field_id):
        return getattr(self, "_section1_labels", {}).get(field_id, field_id)

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _search_company(self, mode="nit"):
        nit = self.fields["nit_empresa"].get().strip()
        nombre = (
            self.fields.get("nombre_busqueda").get().strip()
            if self.fields.get("nombre_busqueda")
            else ""
        )
        if mode == "nit":
            if not nit:
                messagebox.showerror("Error", "Ingresa un NIT.")
                return
        elif mode == "nombre":
            if not nombre:
                messagebox.showerror("Error", "Ingresa el nombre de la empresa.")
                return
        else:
            messagebox.showerror("Error", "Tipo de búsqueda no válido.")
            return

        try:
            self.status_label.config(text="Buscando empresa...")
            self.update_idletasks()
            if mode == "nombre":
                company = induccion_organizacional.get_empresa_by_nombre(nombre)
            else:
                company = induccion_organizacional.get_empresa_by_nit(nit)
        except Exception as exc:
            self.status_label.config(text="")
            messagebox.showerror("Error", str(exc))
            return

        if not company:
            self.company_data = None
            msg = (
                "No se encontró empresa para ese nombre."
                if mode == "nombre"
                else "No se encontró empresa para ese NIT."
            )
            self.status_label.config(text=msg)
            for key in induccion_organizacional.SECTION_1_SUPABASE_MAP.keys():
                self._set_readonly_value(key, "")
            return

        if mode == "nombre":
            nit_value = company.get("nit_empresa")
            if nit_value:
                entry = self.fields.get("nit_empresa")
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, nit_value)

        self.company_data = company
        self.status_label.config(text="Empresa encontrada.")
        for key in induccion_organizacional.SECTION_1_SUPABASE_MAP.keys():
            self._set_readonly_value(key, company.get(key))

    def _prefill_section_1(self):
        cache = induccion_organizacional.get_form_cache().get("section_1", {})
        if not cache:
            return
        self.company_data = cache
        self.fields["nit_empresa"].delete(0, tk.END)
        self.fields["nit_empresa"].insert(0, cache.get("nit_empresa", ""))
        self.fields["modalidad"].set(cache.get("modalidad", ""))
        fecha_value = cache.get("fecha_visita")
        if fecha_value:
            self.fields["fecha_visita"].set_date(fecha_value)
        for key in [
            "nombre_empresa",
            "direccion_empresa",
            "correo_1",
            "contacto_empresa",
            "telefono_empresa",
            "cargo",
            "ciudad_empresa",
            "sede_empresa",
            "caja_compensacion",
            "asesor",
            "profesional_asignado",
        ]:
            self._set_readonly_value(key, cache.get(key, ""))
        if hasattr(self, "continue_btn"):
            self.continue_btn.config(state="normal")

    def _confirm_and_continue(self):
        if not self.company_data:
            messagebox.showerror("Error", "Busca una empresa antes de confirmar.")
            return

        fecha_visita = _get_required_fecha_visita(self)
        if not fecha_visita:
            return
        modalidad = _get_required_modalidad(self)
        if not modalidad:
            return
        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad": modalidad,
            "nit_empresa": self.fields["nit_empresa"].get().strip(),
        }
        try:
            induccion_organizacional.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2()

    def _confirm_section_2(self):
        payload = []
        for idx, block in enumerate(self.vinculado_blocks):
            entry = {"numero": str(idx + 1)}
            for key in ["nombre_oferente", "cedula", "telefono_oferente", "cargo_oferente"]:
                widget = block.get(key)
                if not widget:
                    entry[key] = ""
                    continue
                if isinstance(widget, ttk.Combobox):
                    value = widget.get().strip()
                else:
                    value = widget.get().strip()
                if key == "cedula":
                    value = re.sub(r"\D+", "", value)
                entry[key] = value
            if any(entry.get(k) for k in ["nombre_oferente", "cedula", "telefono_oferente", "cargo_oferente"]):
                payload.append(entry)
        if not payload:
            messagebox.showerror("Error", "Registra al menos un vinculado.")
            return
        try:
            induccion_organizacional.confirm_section_2(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_3()

    def _confirm_section_3(self):
        payload = {}
        for item_id, widgets in self.section3_fields.items():
            payload[item_id] = {
                "visto": widgets["visto"].get().strip(),
                "responsable": widgets["responsable"].get().strip(),
                "medio_socializacion": widgets["medio_socializacion"].get().strip(),
                "descripcion": widgets["descripcion"].get().strip(),
            }
        try:
            induccion_organizacional.confirm_section_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_4()

    def _confirm_section_4(self):
        payload = []
        for medio_widget, text_widget in self.section4_rows:
            payload.append(
                {
                    "medio": medio_widget.get().strip(),
                    "recomendacion": text_widget.get("1.0", tk.END).strip(),
                }
            )
        try:
            induccion_organizacional.confirm_section_4(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_5()

    def _confirm_section_5(self):
        payload = {
            "observaciones": self.section5_text.get("1.0", tk.END).strip(),
        }
        try:
            induccion_organizacional.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_6()

    def _confirm_section_6(self):
        payload = []
        for _row, nombre_entry, cargo_entry in self.section6_rows:
            nombre = nombre_entry.get().strip()
            cargo = cargo_entry.get().strip()
            if nombre or cargo:
                payload.append({"nombre": nombre, "cargo": cargo})
        if not payload:
            messagebox.showerror("Error", "Registra al menos un asistente.")
            return
        try:
            induccion_organizacional.confirm_section_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._export_form()

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Exportando Excel...")
        loading.set_progress(35)

        cache_snapshot = induccion_organizacional.get_form_cache()
        section_1 = cache_snapshot.get("section_1", {})
        company_name = section_1.get("nombre_empresa")
        def _worker():
            output_path = _raise_finalize_stage(
                "exportando el Excel",
                lambda: induccion_organizacional.export_to_excel(clear_cache=False),
            )
            return output_path

        _start_background_finalization(
            self,
            loading,
            form_name="Induccion Organizacional",
            company_name=company_name,
            form_id="induccion_organizacional",
            worker_fn=_worker,
            post_delivery_fn=lambda: _clear_form_cache_safe(induccion_organizacional),
        )

    def _close_to_hub(self):
        _return_to_hub(self)
        self.destroy()


class InduccionOperativaWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Induccion Operativa - Seccion 1")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = induccion_operativa
        self.company_data = None
        self.fields = {}
        self.cedula_options = []

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))
        self.header_title = tk.Label(
            header,
            text="1. DATOS GENERALES",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        self.header_title.pack(anchor="w")
        self.header_subtitle = tk.Label(
            header,
            text="Busca empresa por NIT y confirma datos.",
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        self.header_subtitle.pack(anchor="w", pady=(4, 0))

    def _build_section_container(self):
        self.section_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)

    def _clear_section_container(self):
        for child in self.section_container.winfo_children():
            child.destroy()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "induccion_operativa",
            induccion_operativa,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_3": self._show_section_3,
                "section_4": self._show_section_4,
                "section_5": self._show_section_5,
                "section_6": self._show_section_6,
                "section_7": self._show_section_7,
                "section_8": self._show_section_8,
                "section_9": self._show_section_9,
            },
            self._show_section_1,
        ):
            return True
        if not induccion_operativa.cache_file_exists():
            return False
        resume = messagebox.askyesno(
            "Reanudar",
            "Se encontró un formulario en progreso. ¿Deseas continuar donde lo dejaste?",
        )
        if not resume:
            induccion_operativa.clear_cache_file()
            induccion_operativa.clear_form_cache()
            return False
        induccion_operativa.load_cache_from_file()
        last_section = induccion_operativa.get_form_cache().get("_last_section")
        if last_section == "section_9":
            self._show_section_9()
        elif last_section == "section_8":
            self._show_section_8()
        elif last_section == "section_7":
            self._show_section_7()
        elif last_section == "section_6":
            self._show_section_6()
        elif last_section == "section_5":
            self._show_section_5()
        elif last_section == "section_4":
            self._show_section_4()
        elif last_section == "section_3":
            self._show_section_3()
        elif last_section == "section_2":
            self._show_section_2()
        else:
            self._show_section_1()
        return True

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        groups = [
            (
                "Información de Empresa",
                COLOR_GROUP_EMPRESA,
                [
                    "nombre_empresa",
                    "direccion_empresa",
                    "correo_1",
                    "contacto_empresa",
                    "telefono_empresa",
                    "cargo",
                    "ciudad_empresa",
                    "sede_empresa",
                    "caja_compensacion",
                ],
            ),
            ("Información de Compensar", COLOR_GROUP_COMPENSAR, ["asesor"]),
            ("Información de RECA", COLOR_GROUP_RECA, ["profesional_asignado"]),
        ]
        labels = {
            "nombre_empresa": "Nombre de la empresa",
            "direccion_empresa": "Dirección de la empresa",
            "correo_1": "Correo electrónico",
            "contacto_empresa": "Persona que atiende la visita",
            "telefono_empresa": "Teléfonos",
            "cargo": "Cargo",
            "ciudad_empresa": "Ciudad/Municipio",
            "sede_empresa": "Sede Compensar",
            "caja_compensacion": "Empresa afiliada a Caja de Compensación",
            "asesor": "Asesor",
            "profesional_asignado": "Profesional asignado RECA",
        }
        _section1_build_groups(self, parent, groups, labels)

    def _label_for_field(self, field_id):
        return getattr(self, "_section1_labels", {}).get(field_id, field_id)

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _search_company(self, mode="nit"):
        nit = self.fields["nit_empresa"].get().strip()
        nombre = (
            self.fields.get("nombre_busqueda").get().strip()
            if self.fields.get("nombre_busqueda")
            else ""
        )
        if mode == "nit":
            if not nit:
                messagebox.showerror("Error", "Ingresa un NIT.")
                return
        elif mode == "nombre":
            if not nombre:
                messagebox.showerror("Error", "Ingresa el nombre de la empresa.")
                return
        else:
            messagebox.showerror("Error", "Tipo de búsqueda no válido.")
            return

        try:
            self.status_label.config(text="Buscando empresa...")
            self.update_idletasks()
            if mode == "nombre":
                company = induccion_operativa.get_empresa_by_nombre(nombre)
            else:
                company = induccion_operativa.get_empresa_by_nit(nit)
        except Exception as exc:
            self.status_label.config(text="")
            messagebox.showerror("Error", str(exc))
            return

        if not company:
            self.company_data = None
            msg = (
                "No se encontró empresa para ese nombre."
                if mode == "nombre"
                else "No se encontró empresa para ese NIT."
            )
            self.status_label.config(text=msg)
            for key in induccion_operativa.SECTION_1_SUPABASE_MAP.keys():
                self._set_readonly_value(key, "")
            return

        if mode == "nombre":
            nit_value = company.get("nit_empresa")
            if nit_value:
                entry = self.fields.get("nit_empresa")
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, nit_value)

        self.company_data = company
        self.status_label.config(text="Empresa encontrada.")
        for key in induccion_operativa.SECTION_1_SUPABASE_MAP.keys():
            self._set_readonly_value(key, company.get(key))

    def _prefill_section_1(self):
        cache = induccion_operativa.get_form_cache().get("section_1", {})
        if not cache:
            return
        self.company_data = cache
        self.fields["nit_empresa"].delete(0, tk.END)
        self.fields["nit_empresa"].insert(0, cache.get("nit_empresa", ""))
        self.fields["modalidad"].set(cache.get("modalidad", ""))
        fecha_value = cache.get("fecha_visita")
        if fecha_value:
            self.fields["fecha_visita"].set_date(fecha_value)
        for key in [
            "nombre_empresa",
            "direccion_empresa",
            "correo_1",
            "contacto_empresa",
            "telefono_empresa",
            "cargo",
            "ciudad_empresa",
            "sede_empresa",
            "caja_compensacion",
            "asesor",
            "profesional_asignado",
        ]:
            self._set_readonly_value(key, cache.get(key, ""))

    def _show_section_1(self):
        self._clear_section_container()
        self.header_title.config(text="1. DATOS GENERALES")
        self.header_subtitle.config(text="Busca empresa por NIT y confirma datos.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._prefill_section_1()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._close_to_hub).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_and_continue).pack(side="right")

    def _load_cedula_options(self):
        try:
            self.cedula_options = induccion_operativa.get_usuarios_reca_cedulas()
        except Exception:
            self.cedula_options = []

    def _filter_cedula_values(self, widget):
        raw = widget.get()
        normalized = re.sub(r"\D+", "", raw)
        options = self.cedula_options or []
        if normalized:
            filtered = [c for c in options if c and normalized in c]
        else:
            filtered = options
        widget.configure(values=filtered)

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. DATOS DEL VINCULADO")
        self.header_subtitle.config(text="Registra uno o mas vinculados.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        self._load_cedula_options()
        content = _build_scrollable_content(section_frame, self)

        self.vinculado_blocks = []
        self.vinculado_frames = []

        def _create_widget(parent, field_id, width=30):
            if field_id == "cedula":
                return ttk.Combobox(parent, values=self.cedula_options, state="normal", width=width)
            return tk.Entry(parent, width=width)

        def _apply_usuario_data(fields, data):
            mapping = {
                "nombre_usuario": "nombre_oferente",
                "cedula_usuario": "cedula",
                "telefono_oferente": "telefono_oferente",
                "cargo_oferente": "cargo_oferente",
            }
            for src, dest in mapping.items():
                value = data.get(src)
                if value in (None, ""):
                    continue
                widget = fields.get(dest)
                if not widget:
                    continue
                if isinstance(widget, ttk.Combobox):
                    widget.set(str(value))
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, str(value))

        def _on_cedula_selected(fields, widget):
            raw = widget.get().strip()
            if not raw:
                return
            normalized = re.sub(r"\D+", "", raw)
            if normalized and normalized != raw:
                widget.delete(0, tk.END)
                widget.insert(0, normalized)
            try:
                data = induccion_operativa.get_usuario_reca_by_cedula(normalized)
            except Exception:
                return
            if data:
                _apply_usuario_data(fields, data)

        def _create_vinculado_block(index):
            card = tk.LabelFrame(
                content,
                text=f"Vinculado #{index + 1}",
                bg=COLOR_LIGHT_BG,
                padx=10,
                pady=8,
            )
            card.pack(fill="x", padx=FORM_PADX, pady=6)
            card.grid_columnconfigure(1, weight=1)
            card.grid_columnconfigure(3, weight=1)

            fields = {}
            specs = [
                ("nombre_oferente", "Nombre completo"),
                ("cedula", "Cédula"),
                ("telefono_oferente", "Teléfono"),
                ("cargo_oferente", "Cargo"),
            ]
            for idx, (field_id, label) in enumerate(specs):
                row = idx // 2
                col = (idx % 2) * 2
                tk.Label(card, text=label, font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
                    row=row, column=col, sticky="w", padx=6, pady=(3, 2)
                )
                widget = _create_widget(card, field_id)
                widget.grid(row=row, column=col + 1, sticky="we", padx=6, pady=(3, 2))
                if field_id == "cedula":
                    widget.bind("<KeyRelease>", lambda _e, w=widget: self._filter_cedula_values(w))
                    widget.bind("<<ComboboxSelected>>", lambda _e, f=fields, w=widget: _on_cedula_selected(f, w))
                    widget.bind("<FocusOut>", lambda _e, f=fields, w=widget: _on_cedula_selected(f, w))
                fields[field_id] = widget
            fields["numero"] = str(index + 1)
            self.vinculado_blocks.append(fields)
            self.vinculado_frames.append(card)

        def _add_vinculado():
            _create_vinculado_block(len(self.vinculado_blocks))

        def _remove_last_vinculado():
            if len(self.vinculado_blocks) <= 1:
                return
            frame = self.vinculado_frames.pop()
            frame.destroy()
            self.vinculado_blocks.pop()

        _create_vinculado_block(0)

        cached_rows = induccion_operativa.get_form_cache().get("section_2", [])
        for idx, row_data in enumerate(cached_rows):
            if idx >= len(self.vinculado_blocks):
                _add_vinculado()
            block = self.vinculado_blocks[idx]
            for key in ["nombre_oferente", "cedula", "telefono_oferente", "cargo_oferente"]:
                widget = block.get(key)
                if not widget:
                    continue
                value = row_data.get(key, "")
                if isinstance(widget, ttk.Combobox):
                    widget.set(value)
                else:
                    widget.delete(0, tk.END)
                    widget.insert(0, value)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_1).pack(side="left")
        ttk.Button(actions, text="Agregar vinculado", command=_add_vinculado).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Eliminar ultimo", command=_remove_last_vinculado).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2).pack(side="right")

    def _show_section_3(self):
        self._clear_section_container()
        self.header_title.config(text="3. DESARROLLO DEL PROCESO DE INDUCCION OPERATIVA")
        self.header_subtitle.config(text="Registra ejecucion por actividad.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        self.section3_fields = {}
        cached = induccion_operativa.get_form_cache().get("section_3", {})

        table = tk.Frame(content, bg=COLOR_LIGHT_BG)
        table.pack(fill="x", padx=FORM_PADX, pady=(8, 4))
        table.grid_columnconfigure(0, weight=4, minsize=520)
        table.grid_columnconfigure(1, weight=1, minsize=150)
        table.grid_columnconfigure(2, weight=3, minsize=420)

        tk.Label(table, text="Actividad", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=0, sticky="w", padx=(4, 8), pady=(0, 6)
        )
        tk.Label(table, text="Ejecucion", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=1, sticky="w", padx=(4, 8), pady=(0, 6)
        )
        tk.Label(table, text="Observaciones", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=2, sticky="w", padx=(4, 8), pady=(0, 6)
        )

        for idx, item in enumerate(induccion_operativa.SECTION_3["items"], start=1):
            tk.Label(
                table,
                text=item["label"],
                bg=COLOR_LIGHT_BG,
                justify="left",
                anchor="w",
                wraplength=500,
            ).grid(row=idx, column=0, sticky="w", padx=(4, 8), pady=4)

            ejecucion = ttk.Combobox(
                table,
                values=induccion_operativa.SECTION_3_EJECUCION_OPTIONS,
                state="readonly",
                width=14,
            )
            ejecucion.grid(row=idx, column=1, sticky="we", padx=(4, 8), pady=4)

            observaciones = tk.Entry(table)
            observaciones.grid(row=idx, column=2, sticky="we", padx=(4, 8), pady=4)

            item_cache = cached.get(item["id"], {}) if isinstance(cached, dict) else {}
            ejecucion.set(item_cache.get("ejecucion", ""))
            observaciones.insert(0, item_cache.get("observaciones", ""))
            self.section3_fields[item["id"]] = {
                "ejecucion": ejecucion,
                "observaciones": observaciones,
            }

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_3).pack(side="right")

    def _show_section_4(self):
        self._clear_section_container()
        self.header_title.config(text="4. HABILIDADES SOCIOEMOCIONALES")
        self.header_subtitle.config(
            text="Registra nivel de apoyo, observaciones y nota por cada bloque.",
        )

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        cached = induccion_operativa.get_form_cache().get("section_4", {})
        cached_items = cached.get("items", {}) if isinstance(cached, dict) else {}
        cached_notes = cached.get("notes", {}) if isinstance(cached, dict) else {}

        self.section4_item_widgets = {}
        self.section4_note_widgets = {}

        for block in induccion_operativa.SECTION_4["blocks"]:
            card = tk.LabelFrame(
                content,
                text=block["title"],
                bg=COLOR_LIGHT_BG,
                padx=10,
                pady=8,
            )
            card.pack(fill="x", padx=FORM_PADX, pady=8)
            card.grid_columnconfigure(0, weight=3)
            card.grid_columnconfigure(1, weight=1)
            card.grid_columnconfigure(2, weight=3)

            tk.Label(card, text="Item", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
                row=0, column=0, sticky="w", padx=4, pady=(0, 6)
            )
            tk.Label(card, text="Nivel de apoyo", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
                row=0, column=1, sticky="w", padx=4, pady=(0, 6)
            )
            tk.Label(card, text="Observaciones", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
                row=0, column=2, sticky="w", padx=4, pady=(0, 6)
            )

            for idx, item in enumerate(block["items"], start=1):
                tk.Label(
                    card,
                    text=item["label"],
                    bg=COLOR_LIGHT_BG,
                    wraplength=420,
                    justify="left",
                    anchor="w",
                ).grid(row=idx, column=0, sticky="w", padx=4, pady=4)

                nivel = ttk.Combobox(
                    card,
                    values=induccion_operativa.SECTION_4_NIVEL_APOYO_OPTIONS,
                    state="readonly",
                    width=22,
                )
                nivel.grid(row=idx, column=1, sticky="we", padx=4, pady=4)

                obs_options = induccion_operativa.SECTION_4_OBSERVACIONES_OPTIONS.get(item["row"], [])
                if obs_options:
                    observaciones = ttk.Combobox(
                        card,
                        values=obs_options,
                        state="readonly",
                        width=60,
                    )
                else:
                    observaciones = tk.Entry(card, width=60)
                observaciones.grid(row=idx, column=2, sticky="we", padx=4, pady=4)

                item_cache = cached_items.get(item["id"], {})
                nivel.set(item_cache.get("nivel_apoyo", ""))
                if isinstance(observaciones, ttk.Combobox):
                    observaciones.set(item_cache.get("observaciones", ""))
                else:
                    observaciones.insert(0, item_cache.get("observaciones", ""))

                self.section4_item_widgets[item["id"]] = {
                    "nivel_apoyo": nivel,
                    "observaciones": observaciones,
                }

            note_row = len(block["items"]) + 1
            tk.Label(card, text="Nota", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
                row=note_row, column=0, sticky="w", padx=4, pady=(8, 4)
            )
            note_entry = tk.Entry(card, width=95)
            note_entry.grid(row=note_row, column=1, columnspan=2, sticky="we", padx=4, pady=(8, 4))
            note_entry.insert(0, cached_notes.get(block["id"], ""))
            self.section4_note_widgets[block["id"]] = note_entry

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_3).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_4).pack(side="right")

    def _show_section_5(self):
        self._clear_section_container()
        self.header_title.config(text="5. NIVEL DE APOYO REQUERIDO")
        self.header_subtitle.config(
            text="Completa nivel de apoyo requerido y observaciones para cada condicion.",
        )

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="x", padx=FORM_PADX, pady=(12, 8))
        content.grid_columnconfigure(0, weight=2)
        content.grid_columnconfigure(1, weight=2)
        content.grid_columnconfigure(2, weight=3)

        tk.Label(content, text="Condicion evaluada", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=0, sticky="w", padx=4, pady=(0, 6)
        )
        tk.Label(content, text="Nivel de apoyo requerido", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=1, sticky="w", padx=4, pady=(0, 6)
        )
        tk.Label(content, text="Observaciones", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=2, sticky="w", padx=4, pady=(0, 6)
        )

        cached = induccion_operativa.get_form_cache().get("section_5", {})
        self.section5_fields = {}

        for idx, row_cfg in enumerate(induccion_operativa.SECTION_5["rows"], start=1):
            tk.Label(
                content,
                text=row_cfg["label"],
                bg=COLOR_LIGHT_BG,
                justify="left",
                anchor="w",
                wraplength=360,
            ).grid(row=idx, column=0, sticky="w", padx=4, pady=4)

            nivel = ttk.Combobox(
                content,
                values=induccion_operativa.SECTION_5_NIVEL_OPTIONS,
                state="readonly",
                width=28,
            )
            nivel.grid(row=idx, column=1, sticky="we", padx=4, pady=4)

            observaciones = tk.Entry(content, width=60)
            observaciones.grid(row=idx, column=2, sticky="we", padx=4, pady=4)

            row_cache = cached.get(row_cfg["id"], {}) if isinstance(cached, dict) else {}
            nivel.set(row_cache.get("nivel_apoyo_requerido", ""))
            observaciones.insert(0, row_cache.get("observaciones", ""))

            self.section5_fields[row_cfg["id"]] = {
                "nivel_apoyo_requerido": nivel,
                "observaciones": observaciones,
            }

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_4).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_5).pack(side="right")

    def _show_section_6(self):
        self._clear_section_container()
        self.header_title.config(text="6. AJUSTES RAZONABLES REQUERIDOS")
        self.header_subtitle.config(text="Describe ajustes razonables.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        tk.Label(section_frame, text="Ajustes requeridos", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
            anchor="w", padx=FORM_PADX, pady=(8, 4)
        )
        self.section6_text = tk.Text(section_frame, width=120, height=8, wrap="word")
        self.section6_text.pack(fill="x", padx=FORM_PADX, pady=(0, 8))

        cached = induccion_operativa.get_form_cache().get("section_6", {})
        if cached.get("ajustes_requeridos"):
            self.section6_text.insert("1.0", cached.get("ajustes_requeridos", ""))

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_5).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_6).pack(side="right")

    def _show_section_7(self):
        self._clear_section_container()
        self.header_title.config(text="7. PRIMER SEGUIMIENTO ESTABLECIDO PARA EL VINCULADO")
        self.header_subtitle.config(text="Registra fecha del primer seguimiento.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        row = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        row.pack(fill="x", padx=FORM_PADX, pady=(12, 8))
        tk.Label(row, text="Fecha", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(side="left", padx=(0, 8))
        self.section7_date = DateEntry(row, width=ENTRY_W_MED, date_pattern="yyyy-mm-dd")
        self.section7_date.pack(side="left")

        cached = induccion_operativa.get_form_cache().get("section_7", {})
        fecha = cached.get("fecha_primer_seguimiento", "")
        if fecha:
            try:
                self.section7_date.set_date(fecha)
            except Exception:
                pass

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_6).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_7).pack(side="right")

    def _show_section_8(self):
        self._clear_section_container()
        self.header_title.config(text="8. OBSERVACIONES /RECOMENDACIONES")
        self.header_subtitle.config(text="Registra observaciones y recomendaciones.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        tk.Label(
            section_frame,
            text="Observaciones / Recomendaciones",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", padx=FORM_PADX, pady=(8, 4))
        self.section8_text = tk.Text(section_frame, width=120, height=8, wrap="word")
        self.section8_text.pack(fill="x", padx=FORM_PADX, pady=(0, 8))

        cached = induccion_operativa.get_form_cache().get("section_8", {})
        if cached.get("observaciones_recomendaciones"):
            self.section8_text.insert("1.0", cached.get("observaciones_recomendaciones", ""))

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_7).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_8).pack(side="right")

    def _show_section_9(self):
        self._clear_section_container()
        self.header_title.config(text="9. ASISTENTES")
        self.header_subtitle.config(text="Registra asistentes y agrega filas si aplica.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="x", padx=FORM_PADX, pady=(8, 8))
        self.section9_rows = []
        asistentes_catalog = _get_asistentes_profesionales_catalog()

        def _add_row(nombre="", cargo=""):
            row = tk.Frame(content, bg=COLOR_LIGHT_BG)
            row.pack(fill="x", pady=4)
            tk.Label(row, text="Nombre completo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 6)
            )
            nombre_entry, cargo_entry = _create_asistente_inputs(
                row,
                50,
                use_catalog=(len(self.section9_rows) == 0),
                catalog=asistentes_catalog,
            )
            nombre_entry.pack(side="left", padx=(0, 12))
            tk.Label(row, text="Cargo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 6)
            )
            cargo_entry.pack(side="left")
            if nombre:
                nombre_entry.insert(0, nombre)
            if cargo:
                cargo_entry.insert(0, cargo)
            self.section9_rows.append((row, nombre_entry, cargo_entry))

        def _remove_last():
            if len(self.section9_rows) <= 1:
                return
            row, _, _ = self.section9_rows.pop()
            row.destroy()

        cached_rows = induccion_operativa.get_form_cache().get("section_9", [])
        if cached_rows:
            for item in cached_rows:
                _add_row(item.get("nombre", ""), item.get("cargo", ""))
        else:
            for _ in range(4):
                _add_row()

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_8).pack(side="left")
        ttk.Button(actions, text="Agregar asistente", command=_add_row).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Eliminar ultimo", command=_remove_last).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Finalizar", command=self._confirm_section_9).pack(side="right")

    def _confirm_and_continue(self):
        if not self.company_data:
            messagebox.showerror("Error", "Busca una empresa antes de confirmar.")
            return
        fecha_visita = _get_required_fecha_visita(self)
        if not fecha_visita:
            return
        modalidad = _get_required_modalidad(self)
        if not modalidad:
            return
        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad": modalidad,
            "nit_empresa": self.fields["nit_empresa"].get().strip(),
        }
        try:
            induccion_operativa.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2()

    def _confirm_section_2(self):
        payload = []
        for idx, block in enumerate(self.vinculado_blocks):
            entry = {"numero": str(idx + 1)}
            for key in ["nombre_oferente", "cedula", "telefono_oferente", "cargo_oferente"]:
                widget = block.get(key)
                value = widget.get().strip() if widget else ""
                if key == "cedula":
                    value = re.sub(r"\D+", "", value)
                entry[key] = value
            if any(entry.get(k) for k in ["nombre_oferente", "cedula", "telefono_oferente", "cargo_oferente"]):
                payload.append(entry)
        if not payload:
            messagebox.showerror("Error", "Registra al menos un vinculado.")
            return
        try:
            induccion_operativa.confirm_section_2(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_3()

    def _confirm_section_3(self):
        payload = {}
        for item_id, widgets in self.section3_fields.items():
            payload[item_id] = {
                "ejecucion": widgets["ejecucion"].get().strip(),
                "observaciones": widgets["observaciones"].get().strip(),
            }
        try:
            induccion_operativa.confirm_section_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_4()

    def _confirm_section_4(self):
        payload = {"items": {}, "notes": {}}
        for item_id, widgets in self.section4_item_widgets.items():
            obs_widget = widgets["observaciones"]
            if isinstance(obs_widget, ttk.Combobox):
                obs_value = obs_widget.get().strip()
            else:
                obs_value = obs_widget.get().strip()
            payload["items"][item_id] = {
                "nivel_apoyo": widgets["nivel_apoyo"].get().strip(),
                "observaciones": obs_value,
            }
        for block_id, note_entry in self.section4_note_widgets.items():
            payload["notes"][block_id] = note_entry.get().strip()
        try:
            induccion_operativa.confirm_section_4(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_5()

    def _confirm_section_5(self):
        payload = {}
        for row_id, widgets in self.section5_fields.items():
            payload[row_id] = {
                "nivel_apoyo_requerido": widgets["nivel_apoyo_requerido"].get().strip(),
                "observaciones": widgets["observaciones"].get().strip(),
            }
        try:
            induccion_operativa.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_6()

    def _confirm_section_6(self):
        payload = {
            "ajustes_requeridos": self.section6_text.get("1.0", tk.END).strip(),
        }
        try:
            induccion_operativa.confirm_section_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_7()

    def _confirm_section_7(self):
        payload = {
            "fecha_primer_seguimiento": self.section7_date.get().strip(),
        }
        try:
            induccion_operativa.confirm_section_7(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_8()

    def _confirm_section_8(self):
        payload = {
            "observaciones_recomendaciones": self.section8_text.get("1.0", tk.END).strip(),
        }
        try:
            induccion_operativa.confirm_section_8(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_9()

    def _confirm_section_9(self):
        payload = []
        for _row, nombre_entry, cargo_entry in self.section9_rows:
            nombre = nombre_entry.get().strip()
            cargo = cargo_entry.get().strip()
            if nombre or cargo:
                payload.append({"nombre": nombre, "cargo": cargo})
        if not payload:
            messagebox.showerror("Error", "Registra al menos un asistente.")
            return
        try:
            induccion_operativa.confirm_section_9(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._export_form()

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Exportando Excel...")
        loading.set_progress(35)

        cache_snapshot = induccion_operativa.get_form_cache()
        section_1 = cache_snapshot.get("section_1", {})
        company_name = section_1.get("nombre_empresa")
        def _worker():
            output_path = _raise_finalize_stage(
                "exportando el Excel",
                lambda: induccion_operativa.export_to_excel(clear_cache=False),
            )
            return output_path

        _start_background_finalization(
            self,
            loading,
            form_name="Induccion Operativa",
            company_name=company_name,
            form_id="induccion_operativa",
            worker_fn=_worker,
            post_delivery_fn=lambda: _clear_form_cache_safe(induccion_operativa),
        )

    def _close_to_hub(self):
        _return_to_hub(self)
        self.destroy()


class SensibilizacionWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Sensibilizacion - Seccion 1")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = sensibilizacion
        self.company_data = None
        self.fields = {}

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))
        self.header_title = tk.Label(
            header,
            text="1. DATOS DE LA EMPRESA",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        )
        self.header_title.pack(anchor="w")
        self.header_subtitle = tk.Label(
            header,
            text="Busca empresa por NIT y confirma datos.",
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        )
        self.header_subtitle.pack(anchor="w", pady=(4, 0))

    def _build_section_container(self):
        self.section_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)

    def _clear_section_container(self):
        for child in self.section_container.winfo_children():
            child.destroy()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "sensibilizacion",
            sensibilizacion,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_3": self._show_section_3,
                "section_4": self._show_section_4,
                "section_5": self._show_section_5,
            },
            self._show_section_1,
        ):
            return True
        if not sensibilizacion.cache_file_exists():
            return False
        resume = messagebox.askyesno(
            "Reanudar",
            "Se encontró un formulario en progreso. ¿Deseas continuar donde lo dejaste?",
        )
        if not resume:
            sensibilizacion.clear_cache_file()
            sensibilizacion.clear_form_cache()
            return False
        sensibilizacion.load_cache_from_file()
        last_section = sensibilizacion.get_form_cache().get("_last_section")
        if last_section == "section_5":
            self._show_section_5()
        elif last_section == "section_4":
            self._show_section_4()
        elif last_section == "section_3":
            self._show_section_3()
        elif last_section == "section_2":
            self._show_section_2()
        else:
            self._show_section_1()
        return True

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        groups = [
            (
                "Información de Empresa",
                COLOR_GROUP_EMPRESA,
                [
                    "nombre_empresa",
                    "direccion_empresa",
                    "correo_1",
                    "contacto_empresa",
                    "telefono_empresa",
                    "cargo",
                    "ciudad_empresa",
                ],
            ),
            ("Información de Compensar", COLOR_GROUP_COMPENSAR, ["asesor", "sede_empresa"]),
        ]
        labels = {
            "nombre_empresa": "Nombre de la empresa",
            "direccion_empresa": "Dirección de la empresa",
            "correo_1": "Correo electrónico",
            "contacto_empresa": "Persona que atiende la visita",
            "telefono_empresa": "Teléfonos",
            "cargo": "Cargo",
            "ciudad_empresa": "Ciudad/Municipio",
            "asesor": "Asesor",
            "sede_empresa": "Sede Compensar",
        }
        _section1_build_groups(self, parent, groups, labels)

    def _label_for_field(self, field_id):
        return getattr(self, "_section1_labels", {}).get(field_id, field_id)

    def _set_readonly_value(self, field_id, value):
        entry = self.fields.get(field_id)
        if not entry:
            return
        entry.configure(state="normal")
        entry.delete(0, tk.END)
        entry.insert(0, value if value is not None else "")
        entry.configure(state="readonly")

    def _search_company(self, mode="nit"):
        nit = self.fields["nit_empresa"].get().strip()
        nombre = self.fields.get("nombre_busqueda").get().strip() if self.fields.get("nombre_busqueda") else ""
        if mode == "nit":
            if not nit:
                messagebox.showerror("Error", "Ingresa un NIT.")
                return
        elif mode == "nombre":
            if not nombre:
                messagebox.showerror("Error", "Ingresa el nombre de la empresa.")
                return
        else:
            messagebox.showerror("Error", "Tipo de búsqueda no válido.")
            return
        try:
            self.status_label.config(text="Buscando empresa...")
            self.update_idletasks()
            if mode == "nombre":
                company = sensibilizacion.get_empresa_by_nombre(nombre)
            else:
                company = sensibilizacion.get_empresa_by_nit(nit)
        except Exception as exc:
            self.status_label.config(text="")
            messagebox.showerror("Error", str(exc))
            return

        if not company:
            self.company_data = None
            msg = "No se encontró empresa para ese nombre." if mode == "nombre" else "No se encontró empresa para ese NIT."
            self.status_label.config(text=msg)
            for key in sensibilizacion.SECTION_1_SUPABASE_MAP.keys():
                self._set_readonly_value(key, "")
            return

        if mode == "nombre":
            nit_value = company.get("nit_empresa")
            if nit_value:
                entry = self.fields.get("nit_empresa")
                if entry:
                    entry.delete(0, tk.END)
                    entry.insert(0, nit_value)

        self.company_data = company
        self.status_label.config(text="Empresa encontrada.")
        for key in sensibilizacion.SECTION_1_SUPABASE_MAP.keys():
            self._set_readonly_value(key, company.get(key))

    def _prefill_section_1(self):
        cache = sensibilizacion.get_form_cache().get("section_1", {})
        if not cache:
            return
        self.company_data = cache
        self.fields["nit_empresa"].delete(0, tk.END)
        self.fields["nit_empresa"].insert(0, cache.get("nit_empresa", ""))
        self.fields["modalidad"].set(cache.get("modalidad", ""))
        fecha_value = cache.get("fecha_visita")
        if fecha_value:
            self.fields["fecha_visita"].set_date(fecha_value)
        for key in [
            "nombre_empresa",
            "direccion_empresa",
            "correo_1",
            "contacto_empresa",
            "telefono_empresa",
            "cargo",
            "ciudad_empresa",
            "asesor",
            "sede_empresa",
        ]:
            self._set_readonly_value(key, cache.get(key, ""))

    def _show_section_1(self):
        self._clear_section_container()
        self.header_title.config(text="1. DATOS DE LA EMPRESA")
        self.header_subtitle.config(text="Busca empresa por NIT y confirma datos.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._prefill_section_1()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._close_to_hub).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_1).pack(side="right")

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. PRESENTACION DE LOS TEMAS DE LA SENSIBILIZACION")
        self.header_subtitle.config(text="Describe los temas tratados.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        temas = [
            "Objetivo de la sensibilizacion y alcance general.",
            "Generalidades del concepto discapacidad.",
            "Tipos de discapacidad.",
            "Pautas de comunicacion e interaccion segun necesidad.",
            "Impacto en el clima laboral y recomendaciones de inclusion.",
        ]
        for idx, tema in enumerate(temas, start=1):
            row = tk.Frame(content, bg="white", bd=1, relief="solid")
            row.pack(fill="x", padx=FORM_PADX, pady=6)
            tk.Label(row, text=str(idx), bg="white", font=FONT_LABEL, width=3).pack(
                side="left", padx=8, pady=8
            )
            tk.Label(
                row,
                text=tema,
                bg="white",
                justify="left",
                anchor="w",
                wraplength=860,
            ).pack(side="left", fill="x", expand=True, padx=8, pady=8)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_1).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2).pack(side="right")

    def _show_section_3(self):
        self._clear_section_container()
        self.header_title.config(text="3. OBSERVACIONES")
        self.header_subtitle.config(text="Registra observaciones generales.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        tk.Label(section_frame, text="Observaciones", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
            anchor="w", padx=FORM_PADX, pady=(8, 4)
        )
        self.section3_text = tk.Text(section_frame, width=120, height=8, wrap="word")
        self.section3_text.pack(fill="x", padx=FORM_PADX, pady=(0, 8))
        cache = sensibilizacion.get_form_cache().get("section_3", {})
        if cache.get("observaciones"):
            self.section3_text.insert("1.0", cache.get("observaciones", ""))

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_3).pack(side="right")

    def _show_section_4(self):
        self._clear_section_container()
        self.header_title.config(text="4. REGISTRO FOTOGRAFICO")
        self.header_subtitle.config(text="Esta seccion se conserva para registro fotografico en el acta.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        tk.Label(
            section_frame,
            text="Esta seccion se conserva para registro fotografico en el acta.",
            bg=COLOR_LIGHT_BG,
            fg="#333333",
            font=FONT_SUBTITLE,
        ).pack(anchor="w", padx=FORM_PADX, pady=(12, 8))

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_3).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_4).pack(side="right")

    def _show_section_5(self):
        self._clear_section_container()
        self.header_title.config(text="5. ASISTENTES")
        self.header_subtitle.config(text="Registra asistentes y agrega filas si aplica.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        section_frame = _build_scrollable_content(section_frame, self)

        content = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        content.pack(fill="x", padx=FORM_PADX, pady=(8, 8))
        self.section5_rows = []
        asistentes_catalog = _get_asistentes_profesionales_catalog()

        def _add_row(nombre="", cargo=""):
            row = tk.Frame(content, bg=COLOR_LIGHT_BG)
            row.pack(fill="x", pady=4)
            tk.Label(row, text="Nombre completo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 6)
            )
            nombre_entry, cargo_entry = _create_asistente_inputs(
                row,
                50,
                use_catalog=(len(self.section5_rows) == 0),
                catalog=asistentes_catalog,
            )
            nombre_entry.pack(side="left", padx=(0, 12))
            tk.Label(row, text="Cargo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 6)
            )
            cargo_entry.pack(side="left")
            if nombre:
                nombre_entry.insert(0, nombre)
            if cargo:
                cargo_entry.insert(0, cargo)
            self.section5_rows.append((row, nombre_entry, cargo_entry))

        def _remove_last():
            if len(self.section5_rows) <= 1:
                return
            row, _, _ = self.section5_rows.pop()
            row.destroy()

        cached_rows = sensibilizacion.get_form_cache().get("section_5", [])
        if cached_rows:
            for item in cached_rows:
                _add_row(item.get("nombre", ""), item.get("cargo", ""))
        else:
            for _ in range(4):
                _add_row()

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_4).pack(side="left")
        ttk.Button(actions, text="Agregar asistente", command=_add_row).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Eliminar ultimo", command=_remove_last).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Finalizar", command=self._confirm_section_5).pack(side="right")

    def _confirm_section_1(self):
        if not self.company_data:
            messagebox.showerror("Error", "Busca una empresa antes de confirmar.")
            return
        fecha_visita = _get_required_fecha_visita(self)
        if not fecha_visita:
            return
        modalidad = _get_required_modalidad(self)
        if not modalidad:
            return
        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad": modalidad,
            "nit_empresa": self.fields["nit_empresa"].get().strip(),
        }
        try:
            sensibilizacion.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_2()

    def _confirm_section_2(self):
        try:
            sensibilizacion.confirm_section_2({})
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_3()

    def _confirm_section_3(self):
        payload = {"observaciones": self.section3_text.get("1.0", tk.END).strip()}
        try:
            sensibilizacion.confirm_section_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_4()

    def _confirm_section_4(self):
        try:
            sensibilizacion.confirm_section_4({})
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._show_section_5()

    def _confirm_section_5(self):
        payload = []
        for _row, nombre_entry, cargo_entry in self.section5_rows:
            nombre = nombre_entry.get().strip()
            cargo = cargo_entry.get().strip()
            if nombre or cargo:
                payload.append({"nombre": nombre, "cargo": cargo})
        if not payload:
            messagebox.showerror("Error", "Registra al menos un asistente.")
            return
        try:
            sensibilizacion.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self._export_form()

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Exportando Excel...")
        loading.set_progress(35)

        cache_snapshot = sensibilizacion.get_form_cache()
        section_1 = cache_snapshot.get("section_1", {})
        company_name = section_1.get("nombre_empresa")
        def _worker():
            output_path = _raise_finalize_stage(
                "exportando el Excel",
                lambda: sensibilizacion.export_to_excel(clear_cache=False),
            )
            return output_path

        _start_background_finalization(
            self,
            loading,
            form_name="Sensibilizacion",
            company_name=company_name,
            form_id="sensibilizacion",
            worker_fn=_worker,
            post_delivery_fn=lambda: _clear_form_cache_safe(sensibilizacion),
        )

    def _close_to_hub(self):
        _return_to_hub(self)
        self.destroy()


class SeguimientosWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Seguimientos - Inicio de Caso")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self.user_row = None
        self.case_path = None
        self.cedula_options = []
        self._filtered_cedulas = []

        self.status_var = tk.StringVar(value="Ingresa la cédula y busca el vinculado.")
        self.user_name_var = tk.StringVar(value="")
        self.user_phone_var = tk.StringVar(value="")
        self.user_role_var = tk.StringVar(value="")
        self.user_discapacidad_var = tk.StringVar(value="")
        self.path_var = tk.StringVar(value="")
        self.suggestion_var = tk.StringVar(value="")
        self.compensar_var = tk.StringVar(value="Si (Compensar)")

        self._build_header()
        self._build_body()
        self._load_cedulas()

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))

        tk.Label(
            header,
            text="SEGUIMIENTOS IL",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w")

        tk.Label(
            header,
            text=(
                "Flujo por persona: busca por cédula, detecta/crea archivo, "
                "y sugiere la pestaña siguiente."
            ),
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(2, 0))

    def _build_body(self):
        container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        container.pack(fill="both", expand=True, padx=FORM_PADX, pady=(8, FORM_PADY))

        search = tk.LabelFrame(
            container,
            text="Búsqueda de Vinculado",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        search.pack(fill="x")

        tk.Label(search, text="Cédula:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4
        )
        self.cedula_combo = ttk.Combobox(search, width=30, state="normal")
        self.cedula_combo.grid(row=0, column=1, sticky="w", pady=4)
        self.cedula_combo.bind("<KeyRelease>", self._filter_cedulas)

        ttk.Button(search, text="Buscar", command=self._buscar_vinculado).grid(
            row=0, column=2, sticky="w", padx=(12, 0), pady=4
        )

        tk.Label(
            search,
            text="¿Empresa afiliada a Compensar?",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self.compensar_combo = ttk.Combobox(
            search,
            textvariable=self.compensar_var,
            values=["Si (Compensar)", "No (Otro)"],
            state="readonly",
            width=30,
        )
        self.compensar_combo.grid(row=1, column=1, sticky="w", pady=4)

        info = tk.LabelFrame(
            container,
            text="Datos del Vinculado",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        info.pack(fill="x", pady=(12, 0))
        info.grid_columnconfigure(1, weight=1)

        self._add_info_row(info, 0, "Nombre:", self.user_name_var)
        self._add_info_row(info, 1, "Teléfono:", self.user_phone_var)
        self._add_info_row(info, 2, "Cargo:", self.user_role_var)
        self._add_info_row(info, 3, "Discapacidad:", self.user_discapacidad_var)

        status = tk.LabelFrame(
            container,
            text="Estado del Caso",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        status.pack(fill="x", pady=(12, 0))
        status.grid_columnconfigure(1, weight=1)

        self._add_info_row(status, 0, "Ruta archivo:", self.path_var)
        self._add_info_row(status, 1, "Sugerencia:", self.suggestion_var)
        tk.Label(
            status,
            textvariable=self.status_var,
            font=("Arial", 10),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        actions = tk.Frame(container, bg=COLOR_LIGHT_BG)
        _pack_actions(actions, pad_y=(14, FORM_PADY), pad_x=False)

        ttk.Button(actions, text="Regresar", command=self._close_to_hub).pack(side="left")
        self.create_btn = ttk.Button(
            actions,
            text="Crear / Actualizar Caso",
            command=self._crear_o_actualizar_caso,
            state="disabled",
        )
        self.create_btn.pack(side="right")
        self.edit_btn = ttk.Button(
            actions,
            text="Diligenciar",
            command=self._open_editor,
            state="disabled",
        )
        self.edit_btn.pack(side="right", padx=(0, 8))
        self.open_btn = ttk.Button(
            actions,
            text="Abrir archivo",
            command=self._abrir_archivo,
            state="disabled",
        )
        self.open_btn.pack(side="right", padx=(0, 8))

    def _add_info_row(self, parent, row, label, var):
        tk.Label(parent, text=label, font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=row,
            column=0,
            sticky="nw",
            padx=(0, 8),
            pady=2,
        )
        tk.Label(
            parent,
            textvariable=var,
            font=("Arial", 10),
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
            wraplength=730,
        ).grid(row=row, column=1, sticky="w", pady=2)

    def _load_cedulas(self):
        try:
            self.cedula_options = seguimientos.get_usuarios_reca_cedulas()
            self._filtered_cedulas = list(self.cedula_options)
            self.cedula_combo["values"] = self._filtered_cedulas
        except Exception as exc:
            self.status_var.set(f"No se pudieron cargar cédulas desde Supabase: {exc}")

    def _filter_cedulas(self, _event=None):
        raw = self.cedula_combo.get().strip()
        if not raw:
            self._filtered_cedulas = list(self.cedula_options)
        else:
            self._filtered_cedulas = [c for c in self.cedula_options if raw in str(c)]
        self.cedula_combo["values"] = self._filtered_cedulas[:50]

    def _buscar_vinculado(self):
        cedula = self.cedula_combo.get().strip()
        if not cedula:
            messagebox.showerror("Error", "Ingresa una cédula.")
            return
        try:
            row = seguimientos.get_usuario_reca_by_cedula(cedula)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        if not row:
            messagebox.showerror("Error", "No se encontró la cédula en usuarios_reca.")
            return

        self.user_row = row
        self.user_name_var.set(str(row.get("nombre_usuario") or ""))
        self.user_phone_var.set(str(row.get("telefono_oferente") or ""))
        self.user_role_var.set(str(row.get("cargo_oferente") or ""))
        discapacidad = row.get("discapacidad_detalle") or row.get("discapacidad_usuario") or ""
        self.user_discapacidad_var.set(str(discapacidad))

        normalized = re.sub(r"\D+", "", cedula)
        self.case_path = seguimientos.find_case_workbook(normalized, row.get("nombre_usuario"))
        self.path_var.set(self.case_path or "(No existe archivo aún)")
        self.create_btn.config(state="normal")
        self.open_btn.config(state="normal" if self.case_path else "disabled")
        self.edit_btn.config(state="normal" if self.case_path else "disabled")
        self._refresh_suggestion()

    def _refresh_suggestion(self):
        if not self.case_path:
            self.suggestion_var.set("Se sugiere crear caso y empezar por hoja base (9...)")
            self.status_var.set("No existe archivo para este vinculado.")
            return
        try:
            suggestion = seguimientos.suggest_next_step(self.case_path)
        except Exception as exc:
            self.suggestion_var.set("No fue posible leer sugerencia.")
            self.status_var.set(f"Archivo encontrado, pero falló lectura de estado: {exc}")
            return
        sheet = suggestion.get("sheet") or ""
        msg = suggestion.get("message") or ""
        self.suggestion_var.set(f"{sheet} - {msg}")
        self.status_var.set("Archivo encontrado y estado calculado correctamente.")

    def _crear_o_actualizar_caso(self):
        if not self.user_row:
            messagebox.showerror("Error", "Primero busca y selecciona una cédula válida.")
            return
        raw = self.cedula_combo.get().strip()
        cedula = re.sub(r"\D+", "", raw)
        if not cedula:
            messagebox.showerror("Error", "Cédula inválida.")
            return
        is_compensar = self.compensar_var.get().startswith("Si")
        try:
            result = seguimientos.ensure_case_workbook(cedula, self.user_row, is_compensar=is_compensar)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self.case_path = result.get("path")
        self.path_var.set(self.case_path or "")
        self.open_btn.config(state="normal" if self.case_path else "disabled")
        self.edit_btn.config(state="normal" if self.case_path else "disabled")
        created = bool(result.get("created"))
        max_seg = result.get("max_seguimientos")
        action = "creado" if created else "actualizado"
        self.status_var.set(
            f"Caso {action} correctamente. Límite de seguimientos: {max_seg}."
        )
        self._refresh_suggestion()

    def _abrir_archivo(self):
        if not self.case_path or not os.path.exists(self.case_path):
            messagebox.showerror("Error", "No hay archivo para abrir.")
            return
        try:
            os.startfile(self.case_path)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo abrir el archivo.\n{exc}")

    def _open_editor(self):
        if not self.case_path or not os.path.exists(self.case_path):
            messagebox.showerror("Error", "No hay archivo de seguimiento para diligenciar.")
            return
        editor = SeguimientoEditorWindow(self, self.case_path)
        _focus_window(editor)

    def _close_to_hub(self):
        _return_to_hub(self)
        self.destroy()


class SeguimientoEditorWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent, case_path):
        super().__init__(parent)
        self.owner = parent
        self.case_path = case_path
        self.title("Seguimientos - Diligenciamiento")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1200x820")
        _maximize_window(self)

        self.meta = seguimientos.get_case_meta(case_path)
        self.max_seg = int(self.meta.get("max_seguimientos") or 3)
        self.sheet_options = [seguimientos.SHEET_BASE] + [
            f"{seguimientos.SHEET_PREFIX}{i}" for i in range(1, self.max_seg + 1)
        ] + [seguimientos.SHEET_FINAL]
        suggestion = seguimientos.suggest_next_step(case_path)
        self.sheet_var = tk.StringVar(value=suggestion.get("sheet") or self.sheet_options[0])
        if self.sheet_var.get() not in self.sheet_options:
            self.sheet_var.set(self.sheet_options[0])
        self.status_var = tk.StringVar(value=suggestion.get("message") or "")

        self.base_vars = {}
        self.base_text = {}
        self.base_func_entries_1 = []
        self.base_func_entries_2 = []
        self.base_dates_1 = []
        self.base_dates_2 = []
        self.company_name_combo = None

        self.follow_vars = {}
        self.follow_text = {}
        self.follow_item_obs = []
        self.follow_item_auto = []
        self.follow_item_emp = []
        self.follow_emp_eval = []
        self.follow_emp_obs = []
        self.follow_asistentes = []

        self._build_header()
        self._build_controls()
        self._build_scroller()
        self._render_selected_sheet()

    def _build_header(self):
        header = tk.Frame(self, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=FORM_PADX, pady=(18, 8))
        tk.Label(
            header,
            text="DILIGENCIAMIENTO DE SEGUIMIENTO IL",
            font=FONT_TITLE,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w")
        tk.Label(
            header,
            text=f"Caso: {self.case_path}",
            font=("Arial", 10),
            bg=COLOR_LIGHT_BG,
            fg="#333333",
            wraplength=1120,
            justify="left",
        ).pack(anchor="w")

    def _build_controls(self):
        controls = tk.Frame(self, bg=COLOR_LIGHT_BG)
        controls.pack(fill="x", padx=FORM_PADX, pady=(0, 8))

        tk.Label(controls, text="Hoja:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
            side="left", padx=(0, 6)
        )
        self.sheet_combo = ttk.Combobox(
            controls,
            textvariable=self.sheet_var,
            values=self.sheet_options,
            state="readonly",
            width=45,
        )
        self.sheet_combo.pack(side="left")
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._render_selected_sheet())

        ttk.Button(controls, text="Guardar hoja", command=self._save_current_sheet).pack(
            side="right"
        )
        ttk.Button(controls, text="Cerrar", command=self.destroy).pack(side="right", padx=(0, 8))
        ttk.Button(controls, text="Abrir Excel", command=self._open_excel).pack(
            side="right", padx=(0, 8)
        )
        ttk.Button(controls, text="Refrescar", command=self._render_selected_sheet).pack(
            side="right", padx=(0, 8)
        )

        tk.Label(
            self,
            textvariable=self.status_var,
            font=("Arial", 10),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
            wraplength=1140,
        ).pack(fill="x", padx=FORM_PADX, pady=(0, 8))

    def _build_scroller(self):
        outer = tk.Frame(self, bg=COLOR_LIGHT_BG)
        outer.pack(fill="both", expand=True, padx=FORM_PADX, pady=(0, FORM_PADY))
        self.canvas = tk.Canvas(outer, bg=COLOR_LIGHT_BG, highlightthickness=0)
        self.v_scroll = _create_vscroll(outer, self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_scroll.pack(side="right", fill="y")
        self.content_frame = tk.Frame(self.canvas, bg=COLOR_LIGHT_BG)
        self.canvas_window_id = self.canvas.create_window(
            (0, 0), window=self.content_frame, anchor="nw"
        )
        self.content_frame.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(self.canvas_window_id, width=e.width),
        )
        self._bind_mousewheel(self.canvas, self.content_frame)

    def _clear_content(self):
        for child in self.content_frame.winfo_children():
            child.destroy()

    def _add_labeled_entry(self, parent, row, label, var, width=40):
        tk.Label(parent, text=label, font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=3
        )
        entry = tk.Entry(parent, textvariable=var, width=width)
        entry.grid(row=row, column=1, sticky="w", pady=3)
        return entry

    def _safe_set_date_widget(self, widget, value):
        text = str(value or "").strip()
        if not text:
            return
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
            try:
                widget.set_date(datetime.strptime(text, fmt).date())
                return
            except Exception:
                continue
        try:
            widget.set_date(text)
        except Exception:
            pass

    def _add_labeled_date(self, parent, row, label, var, width=20):
        tk.Label(parent, text=label, font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=3
        )
        date_entry = DateEntry(
            parent,
            textvariable=var,
            width=width,
            date_pattern="yyyy-mm-dd",
        )
        date_entry.grid(row=row, column=1, sticky="w", pady=3)
        self._safe_set_date_widget(date_entry, var.get())
        return date_entry

    def _update_empresa_nombre_suggestions(self, _event=None):
        combo = self.company_name_combo
        if not combo:
            return
        prefix = (
            self.base_vars.get("nombre_empresa").get().strip()
            if self.base_vars.get("nombre_empresa")
            else ""
        )
        if len(prefix) < 2:
            combo["values"] = ()
            return
        try:
            rows = seguimientos.get_empresas_by_nombre_prefix(prefix, limit=12)
        except Exception:
            rows = []
        values = []
        seen = set()
        for row in rows:
            name = str(row.get("nombre_empresa") or "").strip()
            if not name:
                continue
            key = name.casefold()
            if key in seen:
                continue
            seen.add(key)
            values.append(name)
        combo["values"] = values

    def _search_selected_or_typed_company_name(self, _event=None):
        combo = self.company_name_combo
        if not combo:
            return "break"
        typed = (
            self.base_vars.get("nombre_empresa").get().strip()
            if self.base_vars.get("nombre_empresa")
            else ""
        )
        if not typed:
            return "break"
        options = list(combo.cget("values") or [])
        if options:
            exact = next((v for v in options if str(v).strip().casefold() == typed.casefold()), None)
            chosen = exact or str(options[0]).strip()
            if chosen and chosen != typed:
                self.base_vars["nombre_empresa"].set(chosen)
        self._buscar_empresa_por_nombre()
        return "break"

    def _render_selected_sheet(self):
        self._clear_content()
        sheet = self.sheet_var.get()
        if sheet == seguimientos.SHEET_BASE:
            self._render_sheet_base()
        elif sheet.startswith(seguimientos.SHEET_PREFIX):
            match = re.search(r"(\d+)$", sheet)
            if not match:
                self.status_var.set("No se pudo identificar número de seguimiento.")
                return
            self._render_sheet_followup(int(match.group(1)))
        else:
            self._render_sheet_final()
        self.canvas.yview_moveto(0)

    def _render_sheet_base(self):
        payload = seguimientos.get_base_payload(self.case_path)
        self.base_vars = {
            k: tk.StringVar(value=str(payload.get(k, "")))
            for k in [
                "fecha_visita",
                "modalidad",
                "nombre_empresa",
                "ciudad_empresa",
                "direccion_empresa",
                "nit_empresa",
                "correo_1",
                "telefono_empresa",
                "contacto_empresa",
                "cargo",
                "asesor",
                "sede_empresa",
                "nombre_vinculado",
                "cedula",
                "telefono_vinculado",
                "correo_vinculado",
                "contacto_emergencia",
                "parentesco",
                "telefono_emergencia",
                "cargo_vinculado",
                "certificado_discapacidad",
                "certificado_porcentaje",
                "discapacidad",
                "tipo_contrato",
                "fecha_inicio_contrato",
                "fecha_fin_contrato",
            ]
        }

        top = tk.LabelFrame(
            self.content_frame,
            text="Hoja base - Datos generales",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        top.pack(fill="x", pady=(0, 10))
        top.grid_columnconfigure(1, weight=1)

        row = 0
        self._add_labeled_date(top, row, "Fecha visita:", self.base_vars["fecha_visita"], width=18)
        row += 1
        tk.Label(top, text="Modalidad:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=3
        )
        mod_combo = ttk.Combobox(
            top,
            textvariable=self.base_vars["modalidad"],
            values=seguimientos.MODALIDAD_OPTIONS,
            state="readonly",
            width=26,
        )
        mod_combo.grid(row=row, column=1, sticky="w", pady=3)
        row += 1
        tk.Label(top, text="Nombre empresa:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=3
        )
        self.company_name_combo = ttk.Combobox(
            top,
            textvariable=self.base_vars["nombre_empresa"],
            values=(),
            state="normal",
            width=58,
        )
        self.company_name_combo.grid(row=row, column=1, sticky="w", pady=3)
        self.company_name_combo.bind("<KeyRelease>", self._update_empresa_nombre_suggestions)
        self.company_name_combo.bind("<Button-1>", self._update_empresa_nombre_suggestions)
        self.company_name_combo.bind("<FocusIn>", self._update_empresa_nombre_suggestions)
        self.company_name_combo.bind("<<ComboboxSelected>>", self._search_selected_or_typed_company_name)
        self.company_name_combo.bind("<Return>", self._search_selected_or_typed_company_name)
        row += 1
        self._add_labeled_entry(top, row, "NIT:", self.base_vars["nit_empresa"], width=25)
        row += 1
        self._add_labeled_entry(top, row, "Ciudad/Municipio:", self.base_vars["ciudad_empresa"], width=40)
        row += 1
        self._add_labeled_entry(top, row, "Dirección:", self.base_vars["direccion_empresa"], width=70)
        row += 1
        self._add_labeled_entry(top, row, "Correo:", self.base_vars["correo_1"], width=60)
        row += 1
        self._add_labeled_entry(top, row, "Teléfonos:", self.base_vars["telefono_empresa"], width=40)
        row += 1
        self._add_labeled_entry(top, row, "Contacto empresa:", self.base_vars["contacto_empresa"], width=45)
        row += 1
        self._add_labeled_entry(top, row, "Cargo empresa:", self.base_vars["cargo"], width=45)
        row += 1
        self._add_labeled_entry(top, row, "Asesor:", self.base_vars["asesor"], width=40)
        row += 1
        self._add_labeled_entry(top, row, "Sede Compensar:", self.base_vars["sede_empresa"], width=30)
        row += 1

        search_actions = tk.Frame(top, bg=COLOR_LIGHT_BG)
        search_actions.grid(row=row, column=1, sticky="w", pady=(4, 0))
        ttk.Button(search_actions, text="Buscar empresa por NIT", command=self._buscar_empresa_por_nit).pack(
            side="left"
        )
        ttk.Button(
            search_actions,
            text="Buscar empresa por nombre",
            command=self._buscar_empresa_por_nombre,
        ).pack(side="left", padx=(8, 0))

        vinc = tk.LabelFrame(
            self.content_frame,
            text="Datos del vinculado",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        vinc.pack(fill="x", pady=(0, 10))
        r = 0
        for key, label in [
            ("nombre_vinculado", "Nombre"),
            ("cedula", "Cédula"),
            ("telefono_vinculado", "Teléfono"),
            ("correo_vinculado", "Correo"),
            ("cargo_vinculado", "Cargo"),
            ("contacto_emergencia", "Contacto emergencia"),
            ("parentesco", "Parentesco"),
            ("telefono_emergencia", "Teléfono emergencia"),
            ("certificado_discapacidad", "Certificado discapacidad (Si/No/No aplica)"),
            ("certificado_porcentaje", "Porcentaje certificado"),
            ("discapacidad", "Discapacidad"),
            ("tipo_contrato", "Tipo contrato"),
        ]:
            self._add_labeled_entry(vinc, r, f"{label}:", self.base_vars[key], width=50)
            r += 1
        self._add_labeled_date(
            vinc,
            r,
            "Fecha inicio contrato:",
            self.base_vars["fecha_inicio_contrato"],
            width=18,
        )
        r += 1
        self._add_labeled_date(
            vinc,
            r,
            "Fecha fin contrato:",
            self.base_vars["fecha_fin_contrato"],
            width=18,
        )
        r += 1

        txt = tk.LabelFrame(
            self.content_frame,
            text="Apoyos / Ajustes",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        txt.pack(fill="x", pady=(0, 10))
        self.base_text["apoyos_ajustes"] = tk.Text(txt, height=4, wrap="word")
        self.base_text["apoyos_ajustes"].pack(fill="x")
        attach_dictation(
            self.base_text["apoyos_ajustes"],
            form_id="seguimientos",
            field_id="base:apoyos_ajustes",
            session_provider=lambda: _supabase_get_access_token(".env"),
            log_fn=_log_capture,
        )
        self.base_text["apoyos_ajustes"].insert("1.0", str(payload.get("apoyos_ajustes", "")))

        funcs = tk.LabelFrame(
            self.content_frame,
            text="Funciones del vinculado",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        funcs.pack(fill="x", pady=(0, 10))

        tk.Label(funcs, text="Funciones 1 a 5", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=0, sticky="w"
        )
        tk.Label(funcs, text="Funciones 6 a 10", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=1, sticky="w", padx=(18, 0)
        )
        self.base_func_entries_1 = []
        self.base_func_entries_2 = []
        vals_1 = payload.get("funciones_1_5") or []
        vals_2 = payload.get("funciones_6_10") or []
        for i in range(5):
            e1 = tk.Entry(funcs, width=60)
            e1.grid(row=i + 1, column=0, sticky="w", pady=2)
            if i < len(vals_1):
                e1.insert(0, vals_1[i] or "")
            self.base_func_entries_1.append(e1)

            e2 = tk.Entry(funcs, width=60)
            e2.grid(row=i + 1, column=1, sticky="w", padx=(18, 0), pady=2)
            if i < len(vals_2):
                e2.insert(0, vals_2[i] or "")
            self.base_func_entries_2.append(e2)

        dates = tk.LabelFrame(
            self.content_frame,
            text="Fechas de seguimiento",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        dates.pack(fill="x", pady=(0, 10))

        self.base_dates_1 = []
        self.base_dates_2 = []
        d1 = payload.get("seguimiento_fechas_1_3") or []
        d2 = payload.get("seguimiento_fechas_4_6") or []
        for i in range(3):
            tk.Label(dates, text=f"Seguimiento {i+1}:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
                row=i, column=0, sticky="w", pady=2
            )
            e1 = DateEntry(dates, width=22, date_pattern="yyyy-mm-dd")
            e1.grid(row=i, column=1, sticky="w", padx=(0, 24), pady=2)
            if i < len(d1):
                self._safe_set_date_widget(e1, d1[i] or "")
            self.base_dates_1.append(e1)

            tk.Label(
                dates, text=f"Seguimiento {i+4}:", font=FONT_LABEL, bg=COLOR_LIGHT_BG
            ).grid(row=i, column=2, sticky="w", pady=2)
            e2 = DateEntry(dates, width=22, date_pattern="yyyy-mm-dd")
            e2.grid(row=i, column=3, sticky="w", pady=2)
            if i < len(d2):
                self._safe_set_date_widget(e2, d2[i] or "")
            if self.max_seg <= 3:
                e2.config(state="disabled")
            self.base_dates_2.append(e2)

        self.status_var.set("Editando hoja base.")

    def _render_sheet_followup(self, idx):
        payload = seguimientos.get_followup_payload(self.case_path, idx)
        self.follow_vars = {
            "modalidad": tk.StringVar(value=str(payload.get("modalidad") or "")),
            "seguimiento_numero": tk.StringVar(value=str(payload.get("seguimiento_numero") or idx)),
            "tipo_apoyo": tk.StringVar(value=str(payload.get("tipo_apoyo") or "")),
        }
        top = tk.LabelFrame(
            self.content_frame,
            text=f"SEGUIMIENTO PROCESO IL {idx}",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        top.pack(fill="x", pady=(0, 10))

        tk.Label(top, text="Modalidad:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Combobox(
            top,
            textvariable=self.follow_vars["modalidad"],
            values=seguimientos.MODALIDAD_OPTIONS,
            state="readonly",
            width=24,
        ).grid(row=0, column=1, sticky="w")
        tk.Label(top, text="Seguimiento #:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=2, sticky="w", padx=(24, 6)
        )
        ttk.Combobox(
            top,
            textvariable=self.follow_vars["seguimiento_numero"],
            values=[str(i) for i in range(1, self.max_seg + 1)],
            state="readonly",
            width=8,
        ).grid(row=0, column=3, sticky="w")

        items = tk.LabelFrame(
            self.content_frame,
            text="Evaluación desempeño (empleado / empresa)",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        items.pack(fill="x", pady=(0, 10))
        tk.Label(items, text="Ítem", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(row=0, column=0, sticky="w")
        tk.Label(items, text="Observación", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=1, sticky="w", padx=(8, 0)
        )
        tk.Label(items, text="Autoevaluación", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=2, sticky="w", padx=(8, 0)
        )
        tk.Label(items, text="Eval. empresa", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=3, sticky="w", padx=(8, 0)
        )

        self.follow_item_obs = []
        self.follow_item_auto = []
        self.follow_item_emp = []
        labels = payload.get("item_labels") or []
        obs_vals = payload.get("item_observaciones") or []
        auto_vals = payload.get("item_autoevaluacion") or []
        emp_vals = payload.get("item_eval_empresa") or []
        for i, label in enumerate(labels):
            tk.Label(items, text=label, bg=COLOR_LIGHT_BG, anchor="w", justify="left").grid(
                row=i + 1, column=0, sticky="w", pady=2
            )
            e_obs = tk.Entry(items, width=40)
            e_obs.grid(row=i + 1, column=1, sticky="w", padx=(8, 0), pady=2)
            if i < len(obs_vals):
                e_obs.insert(0, obs_vals[i] or "")
            self.follow_item_obs.append(e_obs)

            v_auto = tk.StringVar(value=auto_vals[i] if i < len(auto_vals) else "")
            c_auto = ttk.Combobox(
                items, textvariable=v_auto, values=seguimientos.EVAL_OPTIONS, state="readonly", width=20
            )
            c_auto.grid(row=i + 1, column=2, sticky="w", padx=(8, 0), pady=2)
            self.follow_item_auto.append(v_auto)

            v_emp = tk.StringVar(value=emp_vals[i] if i < len(emp_vals) else "")
            c_emp = ttk.Combobox(
                items, textvariable=v_emp, values=seguimientos.EVAL_OPTIONS, state="readonly", width=20
            )
            c_emp.grid(row=i + 1, column=3, sticky="w", padx=(8, 0), pady=2)
            self.follow_item_emp.append(v_emp)

        middle = tk.LabelFrame(
            self.content_frame,
            text="Tipo de apoyo y evaluación empresarial",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        middle.pack(fill="x", pady=(0, 10))
        tk.Label(middle, text="Tipo de apoyo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Combobox(
            middle,
            textvariable=self.follow_vars["tipo_apoyo"],
            values=seguimientos.TIPO_APOYO_OPTIONS,
            state="readonly",
            width=34,
        ).grid(row=0, column=1, sticky="w")

        tk.Label(middle, text="Ítem", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(row=1, column=0, sticky="w", pady=(8, 2))
        tk.Label(middle, text="Evaluación", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=1, sticky="w", padx=(8, 0), pady=(8, 2)
        )
        tk.Label(middle, text="Observación", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=2, sticky="w", padx=(8, 0), pady=(8, 2)
        )

        self.follow_emp_eval = []
        self.follow_emp_obs = []
        emp_labels = payload.get("empresa_item_labels") or []
        emp_eval_vals = payload.get("empresa_eval") or []
        emp_obs_vals = payload.get("empresa_observacion") or []
        for i, label in enumerate(emp_labels):
            tk.Label(middle, text=label, bg=COLOR_LIGHT_BG, anchor="w", justify="left").grid(
                row=i + 2, column=0, sticky="w", pady=2
            )
            v = tk.StringVar(value=emp_eval_vals[i] if i < len(emp_eval_vals) else "")
            ttk.Combobox(
                middle, textvariable=v, values=seguimientos.EVAL_OPTIONS, state="readonly", width=20
            ).grid(row=i + 2, column=1, sticky="w", padx=(8, 0), pady=2)
            self.follow_emp_eval.append(v)
            e = tk.Entry(middle, width=45)
            e.grid(row=i + 2, column=2, sticky="w", padx=(8, 0), pady=2)
            if i < len(emp_obs_vals):
                e.insert(0, emp_obs_vals[i] or "")
            self.follow_emp_obs.append(e)

        txt = tk.LabelFrame(
            self.content_frame,
            text="Situación encontrada y estrategias",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        txt.pack(fill="x", pady=(0, 10))
        template_actions = tk.Frame(txt, bg=COLOR_LIGHT_BG)
        template_actions.pack(fill="x", pady=(0, 8))
        ttk.Button(
            template_actions,
            text="Seguimiento y normativa",
            command=self._insert_followup_normativa_template,
        ).pack(side="left")
        self.follow_text["situacion_encontrada"] = tk.Text(txt, height=5, wrap="word")
        self.follow_text["situacion_encontrada"].pack(fill="x", pady=(0, 8))
        attach_dictation(
            self.follow_text["situacion_encontrada"],
            form_id="seguimientos",
            field_id=f"followup_{idx}:situacion_encontrada",
            session_provider=lambda: _supabase_get_access_token(".env"),
            log_fn=_log_capture,
        )
        self.follow_text["situacion_encontrada"].insert(
            "1.0", str(payload.get("situacion_encontrada") or "")
        )
        self.follow_text["estrategias_ajustes"] = tk.Text(txt, height=5, wrap="word")
        self.follow_text["estrategias_ajustes"].pack(fill="x")
        attach_dictation(
            self.follow_text["estrategias_ajustes"],
            form_id="seguimientos",
            field_id=f"followup_{idx}:estrategias_ajustes",
            session_provider=lambda: _supabase_get_access_token(".env"),
            log_fn=_log_capture,
        )
        self.follow_text["estrategias_ajustes"].insert(
            "1.0", str(payload.get("estrategias_ajustes") or "")
        )

        asist = tk.LabelFrame(
            self.content_frame,
            text="Asistentes",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        asist.pack(fill="x", pady=(0, 10))
        self.follow_asistentes = []
        asistentes_catalog = _get_asistentes_profesionales_catalog()
        asistentes = payload.get("asistentes") or []
        for i in range(4):
            row = tk.Frame(asist, bg=COLOR_LIGHT_BG)
            row.pack(fill="x", pady=3)
            tk.Label(row, text="Nombre:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(side="left")
            e_name, e_cargo = _create_asistente_inputs(
                row,
                45,
                use_catalog=(i == 0),
                catalog=asistentes_catalog,
            )
            e_name.pack(side="left", padx=(6, 12))
            tk.Label(row, text="Cargo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(side="left")
            e_cargo.pack(side="left", padx=(6, 0))
            if i < len(asistentes):
                e_name.insert(0, str(asistentes[i].get("nombre") or ""))
                e_cargo.insert(0, str(asistentes[i].get("cargo") or ""))
            self.follow_asistentes.append((e_name, e_cargo))

        self.status_var.set(f"Editando hoja de seguimiento {idx}.")

    def _render_sheet_final(self):
        card = tk.LabelFrame(
            self.content_frame,
            text="PONDERADO FINAL",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        card.pack(fill="x")
        tk.Label(
            card,
            text=(
                "Esta hoja es de cálculo consolidado. No se diligencia manualmente.\n"
                "Completa la hoja base y los seguimientos para actualizar sus resultados."
            ),
            bg=COLOR_LIGHT_BG,
            justify="left",
            anchor="w",
            font=("Arial", 10),
        ).pack(fill="x")
        self.status_var.set("Ponderado final es de solo revisión.")

    def _insert_followup_normativa_template(self):
        text_widget = self.follow_text.get("estrategias_ajustes")
        if not text_widget:
            return
        current_text = text_widget.get("1.0", tk.END).strip()
        if current_text:
            text_widget.insert(tk.END, "\n\n")
        text_widget.insert(tk.END, SEGUIMIENTO_NORMATIVA_TEMPLATE_TEXT)
        text_widget.focus_set()
        text_widget.see(tk.END)

    def _buscar_empresa_por_nit(self):
        nit = self.base_vars.get("nit_empresa").get().strip() if self.base_vars.get("nit_empresa") else ""
        if not nit:
            messagebox.showerror("Error", "Ingresa NIT para buscar empresa.")
            return
        try:
            company = seguimientos.get_empresa_by_nit(nit)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        if not company:
            messagebox.showinfo("Empresa", "No se encontró empresa para ese NIT.")
            return
        self._apply_company_data(company)

    def _buscar_empresa_por_nombre(self):
        name = (
            self.base_vars.get("nombre_empresa").get().strip()
            if self.base_vars.get("nombre_empresa")
            else ""
        )
        if not name:
            messagebox.showerror("Error", "Ingresa nombre de empresa para buscar.")
            return
        try:
            company = seguimientos.get_empresa_by_nombre(name)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        if not company:
            messagebox.showinfo("Empresa", "No se encontró empresa con ese nombre.")
            return
        self._apply_company_data(company)

    def _apply_company_data(self, company):
        mapping = {
            "nombre_empresa": "nombre_empresa",
            "ciudad_empresa": "ciudad_empresa",
            "direccion_empresa": "direccion_empresa",
            "nit_empresa": "nit_empresa",
            "correo_1": "correo_1",
            "telefono_empresa": "telefono_empresa",
            "contacto_empresa": "contacto_empresa",
            "cargo": "cargo",
            "asesor": "asesor",
            "sede_empresa": "sede_empresa",
        }
        for field, key in mapping.items():
            if field in self.base_vars:
                self.base_vars[field].set(str(company.get(key) or ""))
        self.status_var.set("Datos de empresa cargados desde Supabase.")

    def _collect_base_payload(self):
        payload = {k: v.get().strip() for k, v in self.base_vars.items()}
        payload["apoyos_ajustes"] = self.base_text["apoyos_ajustes"].get("1.0", tk.END).strip()
        payload["funciones_1_5"] = [e.get().strip() for e in self.base_func_entries_1]
        payload["funciones_6_10"] = [e.get().strip() for e in self.base_func_entries_2]
        payload["seguimiento_fechas_1_3"] = [e.get().strip() for e in self.base_dates_1]
        payload["seguimiento_fechas_4_6"] = [e.get().strip() for e in self.base_dates_2]
        return payload

    def _collect_followup_payload(self, idx):
        payload = {
            "modalidad": self.follow_vars["modalidad"].get().strip(),
            "seguimiento_numero": self.follow_vars["seguimiento_numero"].get().strip() or str(idx),
            "item_observaciones": [e.get().strip() for e in self.follow_item_obs],
            "item_autoevaluacion": [v.get().strip() for v in self.follow_item_auto],
            "item_eval_empresa": [v.get().strip() for v in self.follow_item_emp],
            "tipo_apoyo": self.follow_vars["tipo_apoyo"].get().strip(),
            "empresa_eval": [v.get().strip() for v in self.follow_emp_eval],
            "empresa_observacion": [e.get().strip() for e in self.follow_emp_obs],
            "situacion_encontrada": self.follow_text["situacion_encontrada"].get("1.0", tk.END).strip(),
            "estrategias_ajustes": self.follow_text["estrategias_ajustes"].get("1.0", tk.END).strip(),
            "asistentes": [
                {"nombre": n.get().strip(), "cargo": c.get().strip()} for n, c in self.follow_asistentes
            ],
        }
        return payload

    def _save_current_sheet(self):
        sheet = self.sheet_var.get()
        try:
            if sheet == seguimientos.SHEET_BASE:
                payload = self._collect_base_payload()
                seguimientos.save_base_payload(self.case_path, payload)
            elif sheet.startswith(seguimientos.SHEET_PREFIX):
                idx = int(re.search(r"(\d+)$", sheet).group(1))
                payload = self._collect_followup_payload(idx)
                seguimientos.save_followup_payload(self.case_path, idx, payload)
            else:
                messagebox.showinfo("Ponderado final", "Esta hoja no se diligencia manualmente.")
                return
        except PermissionError:
            messagebox.showerror(
                "Archivo en uso",
                "No se pudo guardar porque el Excel está abierto en otra aplicación.",
            )
            return
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        self.status_var.set(f"Guardado exitoso en hoja: {sheet}")
        if isinstance(self.owner, SeguimientosWindow):
            self.owner.case_path = self.case_path
            self.owner.path_var.set(self.case_path)
            self.owner._refresh_suggestion()

    def _open_excel(self):
        try:
            os.startfile(self.case_path)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo abrir el archivo.\n{exc}")


if __name__ == "__main__":
    if not _acquire_single_instance_mutex():
        _show_single_instance_warning()
        raise SystemExit(0)
    app = HubWindow()
    try:
        app.mainloop()
    except KeyboardInterrupt:
        try:
            app.destroy()
        except Exception:
            pass
    finally:
        _release_single_instance_mutex()










