"""
app.py — Punto de entrada y UI completa de la aplicación.

Este archivo contiene TODAS las ventanas Tkinter y la mayoría de la lógica de
navegación, autenticación y estado local. Es el archivo más grande (~20k líneas).

Estructura interna (buscar los marcadores # ── SECCIÓN):
  1. Imports y constantes globales           (líneas   1 – 225)
  2. Helpers: autosave y estado de formulario (línea  226)
  3. Helpers: instancia única y logs         (línea  523)
  4. Helpers: cola de subida a Drive         (línea  868)
  5. Helpers: credenciales DPAPI y auth local (línea 1450)
  6. Helpers: borradores y formularios       (línea 1643)
  7. Helpers: UI (widgets, feedback, wizard) (línea 1716)
  8. Helpers: test fill                      (línea 2225)
  9. Helpers: texto, encoding, entradas      (línea 2827)
 10. Helpers: autenticación y conectividad   (línea 3046)
 11. Helpers: ventanas y flujo de exportación (línea 3276)
 12. Helpers: labs / experimental            (línea 4058)
 13. Helpers: sección 1 (empresa)            (línea 4198)
 14. Helpers: scroll, dictado, asistentes    (línea 4471)
 15. Clases base: FormMousewheelMixin, LoadingDialog, Labs* (línea 4951)
 16. Finalization helpers                    (línea 5441)
 17. get_forms() — registro de formularios   (línea 5859)
 18. VENTANA: Section1Window                 (línea 5874)
 19. VENTANA: HubWindow                      (línea 6687)
 20. VENTANA: EvaluacionAccesibilidadWindow  (línea 10055)
 21. VENTANA: CondicionesVacanteWindow       (línea 12028)
 22. VENTANA: SeleccionIncluyenteWindow      (línea 13369)
 23. VENTANA: ContratacionIncluyenteWindow   (línea 14771)
 24. VENTANA: InduccionOrganizacionalWindow  (línea 15966)
 25. VENTANA: InduccionOperativaWindow       (línea 16794)
 26. VENTANA: SensibilizacionWindow          (línea 17724)
 27. VENTANA: SeguimientosWindow             (línea 18156)
 28. VENTANA: SeguimientoEditorWindow        (línea 19192)

Depende de: todos los demás módulos del proyecto.
"""
import threading
import errno
import re
import os
import time
import sys
import subprocess
import ctypes
import unicodedata
import shutil
import uuid
import tempfile
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
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry, Calendar

from formularios.presentacion_programa import presentacion_programa
from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.condiciones_vacante import condiciones_vacante
from formularios.condiciones_vacante.voice_section2 import (
    VOICE_FUNCTION_NAME as VACANCY_SECTION2_VOICE_FUNCTION,
    get_subsection_spec as get_vacancy_section2_spec,
    postprocess_extraction_payload as postprocess_vacancy_section2_extraction,
    summarize_candidate_updates as summarize_vacancy_section2_updates,
)
from formularios.seleccion_incluyente import seleccion_incluyente
from formularios.contratacion_incluyente import contratacion_incluyente
from formularios.induccion_organizacional import induccion_organizacional
from formularios.induccion_operativa import induccion_operativa
from formularios.sensibilizacion import sensibilizacion
from formularios.seguimientos import seguimientos
from formularios.interprete_lsc import interprete_lsc
import drive_upload
from dictation import (
    attach_dictation,
    cancel_recording,
    claim_recording,
    cleanup_stale_audio,
    start_recording,
    stop_and_submit_audio,
)
import text_review
import completion_payloads
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
    _ensure_roaming_service_account_file,
    _load_env_file,
    _extract_public_error_detail,
    _get_local_app_cache_dir,
    _normalize_decimal_value,
    _next_available_file_path,
    probe_supabase_service,
)
from formularios.finalize_validation import ValidationIssue, format_issues_for_message
from formularios import ui_feedback
from formularios.user_messages import map_exception_to_user_message
from version_info import get_version
from updater import (
    _repo_config as _updater_repo_config,
    download_installer,
    get_latest_release_assets,
    get_release_page_url,
    is_update_available,
)
from logging_utils import get_log_file_path, get_logs_root, log_app_event, log_labs_event

try:
    from formularios.seleccion_incluyente_labs import seleccion_incluyente as seleccion_incluyente_labs
    from formularios.seleccion_incluyente_labs.voice_section2 import (
        VOICE_FUNCTION_NAME as SELECTION_SECTION2_VOICE_FUNCTION,
        get_subsection_spec as get_selection_labs_subsection_spec,
        postprocess_extraction_payload as postprocess_selection_labs_extraction,
        summarize_candidate_updates as summarize_selection_labs_updates,
    )
except Exception:
    seleccion_incluyente_labs = None
    SELECTION_SECTION2_VOICE_FUNCTION = ""

    def get_selection_labs_subsection_spec(*_args, **_kwargs):
        raise RuntimeError("Labs deshabilitado.")

    def postprocess_selection_labs_extraction(*_args, **_kwargs):
        raise RuntimeError("Labs deshabilitado.")

    def summarize_selection_labs_updates(*_args, **_kwargs):
        raise RuntimeError("Labs deshabilitado.")


APP_NAME = "RECA Inclusion Laboral"
COLOR_PRIMARY = "#4B2E67"
COLOR_ACCENT = "#07B499"
COLOR_SUCCESS = "#0A7D2E"
COLOR_WARNING = "#B35300"
COLOR_DANGER = "#B00020"
COLOR_SURFACE = "#FFFFFF"
COLOR_FIELD_ERROR_BG = "#FDE2E2"
COLOR_FIELD_WARNING_BG = "#FFF4CC"
COLOR_TEXT_PRIMARY = "#23182F"
COLOR_TEXT_SECONDARY = "#5B5563"
COLOR_BORDER = "#D8D0E0"
COLOR_PURPLE = COLOR_PRIMARY
COLOR_TEAL = COLOR_ACCENT
COLOR_LIGHT_BG = COLOR_SURFACE
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
MAX_PASSWORD_LENGTH = 1024
OFFLINE_AUTH_STORE_VERSION = 2
OFFLINE_AUTH_TTL_DAYS = 30
DEFAULT_EMPRESA_ESTADOS = [
    "Activa",
    "Inactiva",
    "En pausa",
    "Cerrada",
    "No viable",
]
_MOJIBAKE_PATTERNS = ("Ãƒ", "Ã‚", "Ã¢â‚¬", "Ã¯Â¿Â½", "\ufffd", "Ã", "Ã‘", "ðŸ")
_ENCODING_CHECK_DONE = False
_TOAST_DURATIONS = {
    "success": 3500,
    "info": 4000,
    "warning": 6000,
    "error": 9000,
}
DRAFTS_FILE_NAME = "form_drafts_il.json"
COMPLETED_FORMS_FILE_NAME = "completed_forms_il.json"
OFFLINE_AUTH_FILE_NAME = "offline_auth_users.json"
LOGIN_CREDENTIALS_FILE_NAME = "login_credentials.json"
DRIVE_UPLOAD_QUEUE_FILE_NAME = "drive_upload_queue.json"
DRIVE_UPLOAD_FAILED_FILE_NAME = "drive_upload_failed.json"
FOLLOWUP_LOCAL_DRAFTS_FILE_NAME = "seguimientos_local_drafts.json"
COMPLETED_FORMS_RETENTION_DAYS = 30
TEST_FILL_DEFAULT_TEXT = "Pendiente"
COMPLETED_FORM_ID_ALIASES = {
    "condiciones_vacante_labs": "condiciones_vacante",
    "seleccion_incluyente_labs": "seleccion_incluyente",
}
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
"Los aprendices no integran la base de trabajadores de carácter permanente de la empresa, para efectos del cálculo de la cuota de empleo para personas en situación de discapacidad, prevista en el numeral 17 del artículo 57 del Código Sustantivo del Trabajo."

En consecuencia y a la luz de esta nueva norma, contratar aprendices con discapacidad no sirve para aumentar el número de personas con discapacidad computables dentro de la cuota de empleo exigida, por lo cual la cuota se calculará ahora sobre la base de trabajadores permanentes, y el decreto 0223 excluye a los aprendices de esa base.

Sin embargo, el Decreto genera un incentivo distinto en el numeral 2 del mismo artículo, donde establece que:
"La cuota de aprendices se reducirá en un 50% si las personas contratadas tienen una discapacidad comprobada no inferior al 25%", en cumplimiento del parágrafo del artículo 31 de la Ley 361 de 1997.

Es decir, que sí es posible contratar aprendices con discapacidad, pero el efecto jurídico directo es sobre la cuota de aprendices (Ley 789 de 2002), no sobre la cuota de empleo para personas con discapacidad (Ley 2466 de 2025 art. 57 num. 17 CST).
""".strip()
FORM_MODULE_MAP = {
    "presentacion_programa": presentacion_programa,
    "evaluacion_accesibilidad": evaluacion_accesibilidad,
    "condiciones_vacante": condiciones_vacante,
    "condiciones_vacante_labs": condiciones_vacante,
    "seleccion_incluyente": seleccion_incluyente,
    "seleccion_incluyente_labs": seleccion_incluyente,
    "contratacion_incluyente": contratacion_incluyente,
    "induccion_organizacional": induccion_organizacional,
    "induccion_operativa": induccion_operativa,
    "sensibilizacion": sensibilizacion,
}
WINDOW_CLASS_FORM_ID_MAP = {
    "Section1Window": "presentacion_programa",
    "EvaluacionAccesibilidadWindow": "evaluacion_accesibilidad",
    "CondicionesVacanteWindow": "condiciones_vacante",
    "CondicionesVacanteLabsWindow": "condiciones_vacante_labs",
    "SeleccionIncluyenteWindow": "seleccion_incluyente",
    "SeleccionIncluyenteLabsWindow": "seleccion_incluyente_labs",
    "ContratacionIncluyenteWindow": "contratacion_incluyente",
    "InduccionOrganizacionalWindow": "induccion_organizacional",
    "InduccionOperativaWindow": "induccion_operativa",
    "SensibilizacionWindow": "sensibilizacion",
    "SeguimientosWindow": "seguimientos",
    "LSCWindow": "interprete_lsc",
}


# ── HELPERS: Autosave, estado de formulario y borradores ────────────────────


def _autosave_section(module, section_key, collect_fn):
    """Guarda silenciosamente los datos de la seccion actual al cache local.
    Se llama al navegar hacia atras para no perder el trabajo."""
    try:
        payload = collect_fn()
        existing_payload = {}
        if hasattr(module, "get_form_cache"):
            try:
                existing_payload = (module.get_form_cache() or {}).get(section_key)
            except Exception:
                existing_payload = {}
        if (
            _cache_snapshot_has_meaningful_values(existing_payload)
            and not _cache_snapshot_has_meaningful_values(payload)
        ):
            module_id = (
                getattr(module, "FORM_ID", "")
                or getattr(module, "FORM_NAME", "")
                or getattr(module, "__name__", module.__class__.__name__)
            )
            _log_capture(
                f"[AUTOSAVE] preserve_existing_data form={module_id} section={section_key} "
                "reason=empty_payload_after_collect"
            )
            return
        try:
            module.set_section_cache(section_key, payload, source="autosave")
        except TypeError:
            module.set_section_cache(section_key, payload)
        module.save_cache_to_file()
    except Exception:
        pass


def _collect_flat_fields(fields):
    """Recolecta widgets o estructuras anidadas de widgets en un payload."""
    missing = object()

    def _collect_value(value):
        if isinstance(value, dict):
            nested_payload = {}
            for nested_key, nested_value in value.items():
                collected = _collect_value(nested_value)
                if collected is not missing:
                    nested_payload[nested_key] = collected
            return nested_payload
        if isinstance(value, list):
            nested_items = []
            for nested_value in value:
                collected = _collect_value(nested_value)
                if collected is not missing:
                    nested_items.append(collected)
            return nested_items
        if isinstance(value, tuple):
            nested_items = []
            for nested_value in value:
                collected = _collect_value(nested_value)
                if collected is not missing:
                    nested_items.append(collected)
            return tuple(nested_items)
        try:
            if isinstance(value, tk.Variable):
                return value.get()
            if isinstance(value, tk.Text):
                return value.get("1.0", tk.END).strip()
            if isinstance(value, (ttk.Combobox, tk.Entry)):
                return value.get().strip()
            if hasattr(value, "get"):
                raw = value.get()
                return raw.strip() if isinstance(raw, str) else raw
        except Exception:
            return missing
        return missing

    payload = {}
    for field_id, widget in fields.items():
        value = _collect_value(widget)
        if value is not missing:
            payload[field_id] = value
    return payload


def _attach_autoexpand(widget, min_h=3, max_h=20):
    """Hace que un tk.Text crezca automáticamente al escribir, hasta max_h líneas."""
    def _resize(_event=None):
        widget.update_idletasks()
        try:
            count = widget.count("1.0", "end-1c", "displaylines")
            if isinstance(count, tuple):
                count = count[0] if count else min_h
            count = int(count or min_h)
        except Exception:
            text = widget.get("1.0", "end-1c")
            count = max(1, text.count("\n") + 1)
        new_h = max(min_h, min(count, max_h))
        widget.config(height=new_h)
        try:
            widget.edit_modified(False)
        except Exception:
            pass
    widget.bind("<<Modified>>", lambda _event=None: widget.after_idle(_resize), add="+")
    widget.bind("<KeyRelease>", lambda _event=None: widget.after_idle(_resize), add="+")
    widget.bind("<<Paste>>", lambda _event=None: widget.after_idle(_resize), add="+")
    widget.bind("<<Cut>>", lambda _event=None: widget.after_idle(_resize), add="+")
    _attach_text_list_support(widget)
    widget.after_idle(_resize)


_TEXT_LIST_LINE_RE = re.compile(
    r"^(?P<indent>\s*)(?P<marker>[-*•]|\d+[.)])\s+(?P<body>.*\S.*)$"
)
_TEXT_LIST_DEFAULT_INDENT = "  "
_TEXT_LIST_TRAILING_PUNCTUATION = (".", "!", "?", ";", ":")


def _get_text_list_continuation(line_text):
    text = str(line_text or "").rstrip()
    if not text:
        return None
    match = _TEXT_LIST_LINE_RE.match(text)
    if not match:
        return None
    indent = match.group("indent") or ""
    marker = match.group("marker") or "-"
    if not indent:
        indent = _TEXT_LIST_DEFAULT_INDENT
    if marker in {"*", "•"}:
        marker = "-"
    if marker[0].isdigit():
        number_match = re.match(r"(\d+)([.)])", marker)
        if number_match:
            next_number = int(number_match.group(1)) + 1
            marker = f"{next_number}{number_match.group(2)}"
    return f"{indent}{marker} "


def _is_plain_text_list_candidate(line_text):
    stripped = str(line_text or "").strip()
    if not stripped:
        return False
    if _TEXT_LIST_LINE_RE.match(stripped):
        return False
    if len(stripped) > 90:
        return False
    if stripped.endswith(_TEXT_LIST_TRAILING_PUNCTUATION):
        return False
    words = stripped.split()
    if len(words) > 12:
        return False
    return any(any(ch.isalpha() for ch in word) for word in words)


def _normalize_text_list_blocks(text):
    lines = str(text or "").splitlines()
    if not lines:
        return str(text or "")
    normalized = list(lines)
    index = 0
    while index < len(normalized):
        line = normalized[index]
        match = _TEXT_LIST_LINE_RE.match(line)
        if match:
            indent = match.group("indent") or _TEXT_LIST_DEFAULT_INDENT
            marker = match.group("marker") or "-"
            body = str(match.group("body") or "").strip()
            if marker in {"*", "•"}:
                marker = "-"
            normalized[index] = f"{indent}{marker} {body}"
            index += 1
            continue

        if not _is_plain_text_list_candidate(line):
            index += 1
            continue

        prev_nonempty = ""
        prev_idx = index - 1
        while prev_idx >= 0:
            prev_line = str(normalized[prev_idx] or "").strip()
            if prev_line:
                prev_nonempty = prev_line
                break
            prev_idx -= 1
        if not prev_nonempty.endswith(":"):
            index += 1
            continue

        end_index = index
        while end_index < len(normalized) and _is_plain_text_list_candidate(normalized[end_index]):
            end_index += 1
        if end_index - index < 2:
            index += 1
            continue
        for line_idx in range(index, end_index):
            normalized[line_idx] = f"{_TEXT_LIST_DEFAULT_INDENT}- {str(normalized[line_idx]).strip()}"
        index = end_index
    return "\n".join(normalized)


def _attach_text_list_support(widget):
    if not isinstance(widget, tk.Text):
        return
    if getattr(widget, "_reca_text_list_support", False):
        return
    widget._reca_text_list_support = True

    def _handle_return(_event=None):
        try:
            line_start = widget.index("insert linestart")
            line_end = widget.index("insert lineend")
            line_text = widget.get(line_start, line_end)
        except Exception:
            return None
        continuation = _get_text_list_continuation(line_text)
        if continuation is None:
            return None
        try:
            widget.insert("insert", "\n" + continuation)
            widget.see("insert")
        except Exception:
            return None
        return "break"

    def _normalize_after_paste(_event=None):
        try:
            original_text = widget.get("1.0", "end-1c")
        except Exception:
            return
        normalized_text = _normalize_text_list_blocks(original_text)
        if normalized_text == original_text:
            return
        try:
            insert_index = widget.index("insert")
        except Exception:
            insert_index = None
        try:
            yview = widget.yview()
        except Exception:
            yview = None
        try:
            widget.delete("1.0", "end")
            widget.insert("1.0", normalized_text)
            if insert_index:
                widget.mark_set("insert", insert_index)
            if yview:
                widget.yview_moveto(yview[0])
            widget.see("insert")
        except Exception:
            return

    widget.bind("<Return>", _handle_return, add="+")
    widget.bind("<<Paste>>", lambda _event=None: widget.after_idle(_normalize_after_paste), add="+")


def _clear_local_resume_state(module):
    """Limpia el caché local temporal. Solo los borradores deben restaurar una sesión."""
    try:
        if hasattr(module, "clear_form_cache"):
            module.clear_form_cache()
    except Exception:
        pass
    try:
        if (
            hasattr(module, "cache_file_exists")
            and hasattr(module, "clear_cache_file")
            and module.cache_file_exists()
        ):
            module.clear_cache_file()
    except Exception:
        pass


def _collect_asistente_rows(rows):
    """Normaliza filas de asistentes y descarta las vacías."""
    values = []
    for row in rows:
        if not row:
            continue
        if len(row) >= 3:
            _, nombre_widget, cargo_widget = row[:3]
        elif len(row) >= 2:
            nombre_widget, cargo_widget = row[:2]
        else:
            continue
        try:
            nombre = _normalize_person_name(nombre_widget.get())
        except Exception:
            nombre = ""
        try:
            cargo = cargo_widget.get().strip()
        except Exception:
            cargo = ""
        if nombre or cargo:
            values.append({"nombre": nombre, "cargo": cargo})
    return values


# ── HELPERS: Instancia única (mutex), logs de app ───────────────────────────


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


def _log_labs(message, level="INFO"):
    try:
        log_labs_event(message, level=level)
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
    return _get_local_app_cache_dir()


def _get_drafts_path():
    return os.path.join(_get_local_cache_dir(), DRAFTS_FILE_NAME)


def _get_completed_forms_path():
    return os.path.join(_get_local_cache_dir(), COMPLETED_FORMS_FILE_NAME)


def _get_offline_auth_path():
    return os.path.join(_get_local_cache_dir(), OFFLINE_AUTH_FILE_NAME)


def _get_login_credentials_path():
    return os.path.join(_get_local_cache_dir(), LOGIN_CREDENTIALS_FILE_NAME)


def _get_drive_upload_queue_path():
    return os.path.join(_get_local_cache_dir(), DRIVE_UPLOAD_QUEUE_FILE_NAME)


def _get_drive_upload_failed_queue_path():
    return os.path.join(_get_local_cache_dir(), DRIVE_UPLOAD_FAILED_FILE_NAME)


def _get_followup_local_drafts_path():
    return os.path.join(_get_local_cache_dir(), FOLLOWUP_LOCAL_DRAFTS_FILE_NAME)


def _atomic_write_json_file(path, payload):
    tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _load_followup_local_drafts_store():
    path = _get_followup_local_drafts_path()
    if not os.path.exists(path):
        return {"cases": {}}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle) or {}
    except Exception:
        return {"cases": {}}
    if not isinstance(data, dict):
        return {"cases": {}}
    cases = data.get("cases")
    if not isinstance(cases, dict):
        cases = {}
    return {"cases": cases}


def _save_followup_local_drafts_store(data):
    path = _get_followup_local_drafts_path()
    folder = os.path.dirname(path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    payload = data if isinstance(data, dict) else {"cases": {}}
    cases = payload.get("cases")
    if not isinstance(cases, dict):
        payload["cases"] = {}
    _atomic_write_json_file(path, payload)


def _build_followup_local_case_key(case_target):
    if isinstance(case_target, dict):
        file_id = str(case_target.get("file_id") or "").strip()
        if file_id:
            return f"drive:{file_id}"
        try:
            encoded = json.dumps(case_target, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
        except Exception:
            encoded = repr(case_target).encode("utf-8", errors="ignore")
        return f"drive:{hashlib.sha256(encoded).hexdigest()}"
    path = os.path.abspath(str(case_target or "").strip())
    if not path:
        return ""
    if os.name == "nt":
        path = os.path.normcase(path)
    return f"path:{path}"


def _build_followup_local_draft_id(case_key, sheet_name):
    key = str(case_key or "").strip()
    sheet = str(sheet_name or "").strip()
    if not key or not sheet:
        return ""
    encoded = f"{key}|{sheet}".encode("utf-8", errors="ignore")
    return f"seguimientos:{hashlib.sha1(encoded).hexdigest()}"


def _get_followup_local_sheet_draft(case_target, sheet_name):
    case_key = _build_followup_local_case_key(case_target)
    sheet_key = str(sheet_name or "").strip()
    if not case_key or not sheet_key:
        return None
    store = _load_followup_local_drafts_store()
    case_entry = store.get("cases", {}).get(case_key)
    if not isinstance(case_entry, dict):
        return None
    sheets = case_entry.get("sheets")
    if not isinstance(sheets, dict):
        return None
    draft = sheets.get(sheet_key)
    return dict(draft or {}) if isinstance(draft, dict) else None


def _list_followup_local_sheet_drafts(case_target):
    case_key = _build_followup_local_case_key(case_target)
    if not case_key:
        return {}
    store = _load_followup_local_drafts_store()
    case_entry = store.get("cases", {}).get(case_key)
    if not isinstance(case_entry, dict):
        return {}
    sheets = case_entry.get("sheets")
    if not isinstance(sheets, dict):
        return {}
    items = {}
    for sheet_key, draft in sheets.items():
        if not isinstance(draft, dict):
            continue
        normalized_sheet = str(sheet_key or draft.get("sheet") or "").strip()
        if not normalized_sheet:
            continue
        items[normalized_sheet] = dict(draft)
    return items


def _save_followup_local_sheet_draft(case_target, request, *, metadata=None):
    case_key = _build_followup_local_case_key(case_target)
    request = dict(request or {})
    sheet_key = str(request.get("sheet") or "").strip()
    if not case_key or not sheet_key:
        return False
    payload = request.get("payload")
    if not isinstance(payload, dict):
        return False
    metadata = dict(metadata or {})
    store = _load_followup_local_drafts_store()
    cases = store.setdefault("cases", {})
    case_entry = cases.setdefault(case_key, {"sheets": {}})
    if not isinstance(case_entry, dict):
        case_entry = {"sheets": {}}
        cases[case_key] = case_entry
    sheets = case_entry.setdefault("sheets", {})
    if not isinstance(sheets, dict):
        sheets = {}
        case_entry["sheets"] = sheets
    now_iso = _get_colombia_now().isoformat()
    draft_id = _build_followup_local_draft_id(case_key, sheet_key)
    sheets[sheet_key] = {
        "draft_id": draft_id,
        "case_key": case_key,
        "sheet": sheet_key,
        "save_kind": str(request.get("save_kind") or "").strip(),
        "followup_index": request.get("followup_index"),
        "payload": copy.deepcopy(payload),
        "fingerprint": str(request.get("fingerprint") or "").strip(),
        "updated_at": now_iso,
        "metadata": copy.deepcopy(metadata),
    }
    case_entry["updated_at"] = now_iso
    case_entry["metadata"] = copy.deepcopy(metadata)
    _save_followup_local_drafts_store(store)
    return True


def _delete_followup_local_sheet_draft(case_target, sheet_name):
    case_key = _build_followup_local_case_key(case_target)
    sheet_key = str(sheet_name or "").strip()
    if not case_key or not sheet_key:
        return False
    store = _load_followup_local_drafts_store()
    cases = store.get("cases", {})
    if not isinstance(cases, dict):
        return False
    case_entry = cases.get(case_key)
    if not isinstance(case_entry, dict):
        return False
    sheets = case_entry.get("sheets")
    if not isinstance(sheets, dict) or sheet_key not in sheets:
        return False
    sheets.pop(sheet_key, None)
    if sheets:
        case_entry["updated_at"] = _get_colombia_now().isoformat()
    else:
        cases.pop(case_key, None)
    _save_followup_local_drafts_store(store)
    return True


def _delete_followup_local_sheet_draft_by_id(draft_id, *, user_login=""):
    draft_key = str(draft_id or "").strip()
    login = str(user_login or "").strip().lower()
    if not draft_key:
        return False
    store = _load_followup_local_drafts_store()
    cases = store.get("cases", {})
    if not isinstance(cases, dict):
        return False
    deleted = False
    for case_key, case_entry in list(cases.items()):
        if not isinstance(case_entry, dict):
            continue
        sheets = case_entry.get("sheets")
        if not isinstance(sheets, dict):
            continue
        for sheet_key, sheet_entry in list(sheets.items()):
            if not isinstance(sheet_entry, dict):
                continue
            if str(sheet_entry.get("draft_id") or "").strip() != draft_key:
                continue
            metadata = sheet_entry.get("metadata")
            entry_login = str((metadata or {}).get("user_login") or "").strip().lower()
            if login and entry_login and entry_login != login:
                continue
            sheets.pop(sheet_key, None)
            deleted = True
        if not sheets:
            cases.pop(case_key, None)
    if deleted:
        _save_followup_local_drafts_store(store)
    return deleted


def _list_followup_local_drafts_for_user(user_login):
    login = str(user_login or "").strip().lower()
    if not login:
        return []
    store = _load_followup_local_drafts_store()
    cases = store.get("cases", {})
    if not isinstance(cases, dict):
        return []
    items = []
    for case_entry in cases.values():
        if not isinstance(case_entry, dict):
            continue
        sheets = case_entry.get("sheets")
        if not isinstance(sheets, dict):
            continue
        for sheet_entry in sheets.values():
            if not isinstance(sheet_entry, dict):
                continue
            metadata = dict(sheet_entry.get("metadata") or {})
            entry_login = str(metadata.get("user_login") or "").strip().lower()
            if entry_login != login:
                continue
            sheet_name = str(sheet_entry.get("sheet") or "").strip()
            company_name = str(
                metadata.get("company_name")
                or metadata.get("case_label")
                or metadata.get("folder_name")
                or "Seguimientos"
            ).strip()
            items.append(
                {
                    "draft_id": str(sheet_entry.get("draft_id") or "").strip(),
                    "form_id": "seguimientos",
                    "form_name": "Seguimientos",
                    "company_name": company_name,
                    "last_section": sheet_name,
                    "updated_at": str(sheet_entry.get("updated_at") or ""),
                    "created_at": str(sheet_entry.get("updated_at") or ""),
                    "draft_type": "followup_local",
                    "sheet_name": sheet_name,
                    "save_kind": str(sheet_entry.get("save_kind") or "").strip(),
                    "case_record": copy.deepcopy(metadata.get("case_record") or {}),
                    "case_path": str(metadata.get("case_path") or ""),
                    "case_label": str(metadata.get("case_label") or company_name),
                }
            )
    items.sort(key=lambda item: str(item.get("updated_at") or ""), reverse=True)
    return items


def _get_colombia_now():
    try:
        return datetime.now(ZoneInfo("America/Bogota"))
    except Exception:
        return datetime.now()


def _parse_completed_datetime(value):
    text = str(value or "").strip()
    if not text:
        return None
    normalized = text[:-1] + "+00:00" if text.endswith("Z") else text
    try:
        return datetime.fromisoformat(normalized)
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _prune_completed_forms_entries(entries, *, now=None):
    if not isinstance(entries, list):
        return []
    current = now or _get_colombia_now()
    cutoff_ts = current.timestamp() - (COMPLETED_FORMS_RETENTION_DAYS * 86400)
    kept = []
    for item in entries:
        if not isinstance(item, dict):
            continue
        dt = _parse_completed_datetime(
            item.get("finalizado_at_iso")
            or item.get("payload_generated_at")
            or item.get("finalizado_at_colombia")
        )
        if dt is None or dt.timestamp() >= cutoff_ts:
            kept.append(item)
    kept.sort(
        key=lambda item: (
            _parse_completed_datetime(
                item.get("finalizado_at_iso")
                or item.get("payload_generated_at")
                or item.get("finalizado_at_colombia")
            ).timestamp()
            if _parse_completed_datetime(
                item.get("finalizado_at_iso")
                or item.get("payload_generated_at")
                or item.get("finalizado_at_colombia")
            )
            else 0.0
        ),
        reverse=True,
    )
    return kept


def _load_completed_forms_store():
    path = _get_completed_forms_path()
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
        users = data["users"]
    changed = False
    for login, entries in list(users.items()):
        pruned = _prune_completed_forms_entries(entries)
        if pruned != entries:
            users[login] = pruned
            changed = True
    data.setdefault("version", 1)
    if changed:
        try:
            _atomic_write_json_file(path, data)
        except Exception:
            pass
    return data


def _save_completed_forms_store(data):
    payload = dict(data or {})
    users = payload.get("users")
    if not isinstance(users, dict):
        users = {}
        payload["users"] = users
    for login, entries in list(users.items()):
        users[login] = _prune_completed_forms_entries(entries)
    payload.setdefault("version", 1)
    _atomic_write_json_file(_get_completed_forms_path(), payload)


def _normalize_completed_form_id(form_id):
    raw = str(form_id or "").strip()
    return COMPLETED_FORM_ID_ALIASES.get(raw, raw)


def _extract_completed_payload_raw(entry):
    payload_raw = entry.get("payload_raw")
    if isinstance(payload_raw, dict):
        return payload_raw
    if isinstance(payload_raw, str):
        try:
            parsed = json.loads(payload_raw)
        except Exception:
            return {}
        return parsed if isinstance(parsed, dict) else {}
    return {}


def _resolve_completed_restore_form_meta(entry):
    payload_raw = _extract_completed_payload_raw(entry)
    raw_form_id = payload_raw.get("form_id") or entry.get("form_id") or ""
    form_id = _normalize_completed_form_id(raw_form_id)
    if not form_id or form_id not in FORM_MODULE_MAP:
        return None
    form_meta = _resolve_form_meta(form_id)
    if not isinstance(form_meta, dict):
        return None
    if bool(form_meta.get("hidden")):
        return None
    return form_meta


def _store_completed_form_locally(row):
    if not isinstance(row, dict):
        return
    user_login = str(row.get("usuario_login") or "").strip().lower()
    if not user_login:
        return
    payload_raw = row.get("payload_raw")
    if not isinstance(payload_raw, dict):
        return
    cache_snapshot = payload_raw.get("cache_snapshot")
    if not isinstance(cache_snapshot, dict) or not cache_snapshot:
        return
    data = _load_completed_forms_store()
    users = data.setdefault("users", {})
    entries = users.setdefault(user_login, [])
    if not isinstance(entries, list):
        entries = []
        users[user_login] = entries
    entry = {
        "registro_id": str(row.get("registro_id") or "").strip(),
        "source_item_key": str(row.get("source_item_key") or "").strip(),
        "form_id": str(payload_raw.get("form_id") or "").strip(),
        "form_name": str(row.get("nombre_formato") or "").strip(),
        "company_name": str(row.get("nombre_empresa") or "").strip(),
        "output_path": str(row.get("path_formato") or "").strip(),
        "upload_status": str(row.get("upload_status") or "").strip(),
        "finalizado_at_iso": str(row.get("finalizado_at_iso") or "").strip(),
        "finalizado_at_colombia": str(row.get("finalizado_at_colombia") or "").strip(),
        "payload_generated_at": str(row.get("payload_generated_at") or "").strip(),
        "payload_raw": copy.deepcopy(payload_raw),
    }
    existing = None
    for item in entries:
        if not isinstance(item, dict):
            continue
        if entry["registro_id"] and str(item.get("registro_id") or "") == entry["registro_id"]:
            existing = item
            break
        if entry["source_item_key"] and str(item.get("source_item_key") or "") == entry["source_item_key"]:
            existing = item
            break
    if existing is None:
        entries.append(entry)
    else:
        existing.clear()
        existing.update(entry)
    _save_completed_forms_store(data)


def _sync_completed_forms_from_remote(user_login):
    login = str(user_login or "").strip().lower()
    if not login:
        return
    cutoff_iso = (_get_colombia_now() - timedelta(days=COMPLETED_FORMS_RETENTION_DAYS)).isoformat()
    try:
        rows = _supabase_get_paged(
            "formatos_finalizados_il",
            {
                "select": ",".join(
                    [
                        "registro_id",
                        "usuario_login",
                        "nombre_formato",
                        "nombre_empresa",
                        "path_formato",
                        "upload_status",
                        "finalizado_at_iso",
                        "finalizado_at_colombia",
                        "payload_generated_at",
                        "source_item_key",
                        "payload_raw",
                    ]
                ),
                "usuario_login": f"eq.{login}",
                "finalizado_at_iso": f"gte.{cutoff_iso}",
                "payload_raw": "not.is.null",
                "order": "finalizado_at_iso.desc",
            },
            page_size=100,
            max_pages=3,
        )
    except Exception as exc:
        _log_capture(f"sync_completed_forms_from_remote failed user={login} err={exc}")
        return
    for row in rows:
        try:
            _store_completed_form_locally(row)
        except Exception as exc:
            _log_capture(f"sync_completed_forms_from_remote store_failed user={login} err={exc}")


# ── HELPERS: Cola de subida a Google Drive (queue persistida en JSON) ────────


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


def _iter_exception_chain(exc):
    pending = [exc]
    visited = set()
    while pending:
        current = pending.pop(0)
        if current is None:
            continue
        marker = id(current)
        if marker in visited:
            continue
        visited.add(marker)
        yield current
        for attr in ("__cause__", "__context__"):
            value = getattr(current, attr, None)
            if isinstance(value, BaseException):
                pending.append(value)


def _is_transient_os_error(exc):
    if not isinstance(exc, OSError):
        return False
    if isinstance(exc, (ConnectionError, TimeoutError)):
        return True
    transient_errnos = {
        getattr(errno, "ECONNABORTED", None),
        getattr(errno, "ECONNREFUSED", None),
        getattr(errno, "ECONNRESET", None),
        getattr(errno, "EHOSTUNREACH", None),
        getattr(errno, "ENETDOWN", None),
        getattr(errno, "ENETRESET", None),
        getattr(errno, "ENETUNREACH", None),
        getattr(errno, "ETIMEDOUT", None),
    }
    transient_winerrors = {
        64,
        10051,
        10052,
        10053,
        10054,
        10060,
        10061,
    }
    err_no = getattr(exc, "errno", None)
    if isinstance(err_no, int) and err_no in transient_errnos:
        return True
    winerror = getattr(exc, "winerror", None)
    if isinstance(winerror, int) and winerror in transient_winerrors:
        return True
    return False


def _is_transient_drive_exception(exc):
    if exc is None:
        return False
    for root in _iter_exception_chain(exc):
        if isinstance(root, urllib.error.HTTPError):
            code = int(getattr(root, "code", 0) or 0)
            if code >= 500 or code == 429:
                return True
            if code in {400, 401, 403, 404}:
                return False
        if isinstance(root, urllib.error.URLError):
            return True
        if isinstance(root, TimeoutError):
            return True
        if _is_transient_os_error(root):
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
            "se ha anulado una conexion",
            "se ha anulado una conexión",
            "software en su equipo host",
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


def _build_drive_upload_result_from_exception(job, exc, *, attempted_at=None):
    attempted_at = str(attempted_at or _get_colombia_now().isoformat()).strip()
    error = str(exc).strip() or repr(exc)
    status = "pending" if _is_transient_drive_exception(exc) else "failed"
    _update_form_completion_upload_status(
        (job or {}).get("registro_id"),
        upload_status=status,
        upload_error=error,
        upload_attempted_at=attempted_at,
    )
    return {
        "status": status,
        "error": error,
        "attempted_at": attempted_at,
        "output_path": str((job or {}).get("local_excel_path") or "").strip(),
        "remote_url": "",
        "drive_file_id": "",
    }


def _drive_upload_operation_label(job):
    upload_kind = str((job or {}).get("upload_kind") or "excel_upload").strip().lower()
    if upload_kind == "evaluacion_sheet":
        return "sheet_copy"
    if upload_kind == "pdf_export":
        return "pdf_export"
    return "upload"


# Tipos de acta que tienen exportación a PDF habilitada
_PDF_EXPORT_ENABLED_TIPOS = {
    "presentacion_programa",
    "reactivacion_programa",
    "interprete_lsc",
    "condiciones_vacante",
    "seleccion_individual",
    "seleccion_grupal",
    "contratacion_individual",
    "contratacion_grupal",
    "induccion_organizacional",
    "induccion_operativa",
    "seguimiento",
}


def _enqueue_pdf_export_job(
    *,
    sheet_file_id,
    tipo_acta,
    fecha_servicio,
    acta_metadata,
    extra_name=None,
    pdf_folder_id,
    company_name="",
    registro_id="",
    selected_sheet_names=None,
    temp_parent_folder_id="",
):
    """Encola un job de exportación a PDF en la cola de uploads de Drive.

    A diferencia de ``_enqueue_drive_upload_job``, este job preserva todos los
    campos necesarios para la exportación a PDF y NO actualiza el estado en Supabase
    al completarse (el Google Sheet ya actualizó ese estado).
    """
    _ensure_drive_upload_worker()
    record = {
        "id": str(uuid.uuid4()),
        "upload_kind": "pdf_export",
        "sheet_file_id": str(sheet_file_id or "").strip(),
        "tipo_acta": str(tipo_acta or "").strip(),
        "fecha_servicio": str(fecha_servicio or "").strip(),
        "acta_metadata": acta_metadata if isinstance(acta_metadata, dict) else {},
        "extra_name": str(extra_name or "").strip() or None,
        "pdf_folder_id": str(pdf_folder_id or "").strip(),
        "company_name": str(company_name or "").strip(),
        "registro_id": str(registro_id or "").strip(),
        "selected_sheet_names": [
            str(name or "").strip()
            for name in list(selected_sheet_names or [])
            if str(name or "").strip()
        ],
        "temp_parent_folder_id": str(temp_parent_folder_id or "").strip(),
        "attempts": 0,
        "last_error": "",
        "next_try_at": float(time.time()),
    }
    with _DRIVE_UPLOAD_LOCK:
        _DRIVE_UPLOAD_QUEUE.append(record)
        _persist_drive_upload_queue_locked()
    return record["id"]


def _perform_pdf_export_attempt(job, attempted_at):
    """Ejecuta la exportación del Google Sheet a PDF con metadata RECA.

    No actualiza Supabase — el estado ya fue registrado por el job del Google Sheet.
    """
    from datetime import date as _date

    sheet_file_id = str((job or {}).get("sheet_file_id") or "").strip()
    tipo_acta = str((job or {}).get("tipo_acta") or "").strip()
    fecha_servicio_str = str((job or {}).get("fecha_servicio") or "").strip()
    acta_metadata = (job or {}).get("acta_metadata") or {}
    extra_name_raw = (job or {}).get("extra_name")
    extra_name = str(extra_name_raw or "").strip() or None
    pdf_folder_id = str((job or {}).get("pdf_folder_id") or "").strip()
    selected_sheet_names = [
        str(name or "").strip()
        for name in list((job or {}).get("selected_sheet_names") or [])
        if str(name or "").strip()
    ]
    temp_parent_folder_id = str((job or {}).get("temp_parent_folder_id") or "").strip()

    try:
        if not sheet_file_id:
            raise RuntimeError("Falta sheet_file_id en el job de PDF export.")
        if not tipo_acta:
            raise RuntimeError("Falta tipo_acta en el job de PDF export.")
        if not fecha_servicio_str:
            raise RuntimeError("Falta fecha_servicio en el job de PDF export.")

        try:
            fecha_servicio = _date.fromisoformat(fecha_servicio_str)
        except ValueError as exc:
            raise RuntimeError(f"Fecha de servicio inválida: {fecha_servicio_str!r}") from exc

        if not pdf_folder_id:
            pdf_folder_id = drive_upload._get_pdf_folder_id()

        company_name_for_pdf = str((job or {}).get("company_name") or "").strip()
        service = drive_upload._build_drive_service_for_pdf()
        upload_result = drive_upload.create_and_upload_acta_pdf(
            service=service,
            sheet_file_id=sheet_file_id,
            tipo_acta=tipo_acta,
            acta_metadata=acta_metadata,
            fecha_servicio=fecha_servicio,
            folder_id=pdf_folder_id,
            folder_name=company_name_for_pdf or None,
            extra=extra_name,
            selected_sheet_names=selected_sheet_names,
            temp_parent_folder_id=temp_parent_folder_id or pdf_folder_id,
        )
    except Exception as exc:
        return _build_drive_upload_result_from_exception(job, exc, attempted_at=attempted_at)

    uploaded_at = _get_colombia_now().isoformat()
    return {
        "status": "synced",
        "error": "",
        "attempted_at": attempted_at,
        "uploaded_at": uploaded_at,
        "output_path": "",
        "remote_url": str(upload_result.get("webViewLink") or ""),
        "drive_file_id": str(upload_result.get("file_id") or ""),
    }


def _perform_drive_upload_attempt(job):
    attempted_at = _get_colombia_now().isoformat()
    upload_kind = str((job or {}).get("upload_kind") or "excel_upload").strip().lower()

    # Delegación especial para PDF export (manejo propio, sin actualizar Supabase)
    if upload_kind == "pdf_export":
        return _perform_pdf_export_attempt(job, attempted_at)

    excel_path = str((job or {}).get("local_excel_path") or "").strip()
    company_name = str((job or {}).get("company_name") or "").strip()
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
                format_ranges=sheet_export.get("format_ranges") or [],
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
        return _build_drive_upload_result_from_exception(job, exc, attempted_at=attempted_at)

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


# ── HELPERS: Credenciales guardadas (Windows DPAPI) y auth local ─────────────


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
    payload = {
        "remember": True,
        "auto_login": False,
        "username": "",
        "password": "",
        "resolved_email": "",
    }
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
    auto_login = bool(raw.get("auto_login", False))
    username = str(raw.get("username") or "").strip()
    password = _dpapi_decrypt_text(raw.get("password_enc"))
    payload["remember"] = remember
    payload["auto_login"] = bool(remember and auto_login)
    payload["username"] = username
    payload["password"] = password if remember else ""
    payload["resolved_email"] = str(raw.get("resolved_email") or "").strip()
    return payload


def _save_login_credentials(username, password, resolved_email="", auto_login=False):
    user = str(username or "").strip()
    pwd = str(password or "")
    cipher = _dpapi_encrypt_text(pwd)
    if not user or not cipher:
        return
    payload = {
        "version": 1,
        "remember": True,
        "auto_login": bool(auto_login),
        "username": user,
        "password_enc": cipher,
        "resolved_email": str(resolved_email or "").strip(),
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
        return {"version": OFFLINE_AUTH_STORE_VERSION, "users": {}}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle) or {}
    except Exception:
        return {"version": OFFLINE_AUTH_STORE_VERSION, "users": {}}
    if not isinstance(data, dict):
        return {"version": OFFLINE_AUTH_STORE_VERSION, "users": {}}
    users = data.get("users")
    if not isinstance(users, dict):
        data["users"] = {}
    data["version"] = max(int(data.get("version") or 0), OFFLINE_AUTH_STORE_VERSION)
    return data


def _save_offline_auth_store(data):
    path = _get_offline_auth_path()
    tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


# ── HELPERS: Borradores de formulario (drafts.json) ─────────────────────────


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


def _form_supports_drafts(form_meta_or_id):
    if isinstance(form_meta_or_id, dict):
        value = form_meta_or_id.get("supports_drafts")
    else:
        value = _resolve_form_meta(form_meta_or_id).get("supports_drafts")
    if value is None:
        return True
    return bool(value)


# ── HELPERS: UI — widgets, feedback inline, wizard de progreso ───────────────


def _log_user_error(context, exc):
    _log_capture(f"[UI] context={context} err={_extract_public_error_detail(exc)}")
    return map_exception_to_user_message(context, exc)


def _button_style_for_kind(kind):
    key = str(kind or "").strip().lower()
    if key == "primary":
        return "Primary.TButton"
    if key == "danger":
        return "DangerOutline.TButton"
    return "Secondary.TButton"


def _safe_widget_state(widget):
    try:
        return str(widget.cget("state") or "")
    except Exception:
        return ""


def _safe_widget_text(widget):
    try:
        return str(widget.cget("text") or "")
    except Exception:
        return None


def _capture_widget_snapshots(widgets):
    seen = set()
    snapshots = []
    for widget in list(widgets or []):
        if widget is None:
            continue
        widget_id = id(widget)
        if widget_id in seen:
            continue
        seen.add(widget_id)
        snapshots.append(
            {
                "widget": widget,
                "state": _safe_widget_state(widget),
                "text": _safe_widget_text(widget),
            }
        )
    return snapshots


def _restore_widget_snapshots(snapshots):
    for snapshot in list(snapshots or []):
        widget = snapshot.get("widget")
        if widget is None:
            continue
        text = snapshot.get("text")
        state = snapshot.get("state")
        try:
            if text is not None:
                widget.configure(text=text)
        except Exception:
            pass
        try:
            if state:
                widget.configure(state=state)
        except Exception:
            pass


def _disable_widget(widget):
    if widget is None:
        return
    try:
        widget.configure(state="disabled")
    except Exception:
        return


def _set_window_busy_cursor(window, waiting):
    if window is None:
        return
    cursor = "watch" if waiting else ""
    targets = [window]
    try:
        master = window.master
        if master is not None:
            targets.append(master)
    except Exception:
        pass
    for target in targets:
        try:
            target.configure(cursor=cursor)
        except Exception:
            continue
    try:
        window.update_idletasks()
    except Exception:
        pass


def _run_async_ui_task(
    window,
    *,
    busy_attr,
    widgets=None,
    loading_button=None,
    loading_button_text=None,
    status_label=None,
    loading_text="",
    loading_state="loading",
    worker=None,
    on_success=None,
    on_error=None,
):
    if not callable(worker) or window is None:
        return False
    if bool(getattr(window, busy_attr, False)):
        return False
    setattr(window, busy_attr, True)
    tracked_widgets = list(widgets or [])
    if loading_button is not None:
        tracked_widgets.append(loading_button)
    snapshots = _capture_widget_snapshots(tracked_widgets)
    for widget in tracked_widgets:
        _disable_widget(widget)
    if loading_button is not None and loading_button_text:
        try:
            loading_button.configure(text=loading_button_text)
        except Exception:
            pass
    _set_window_busy_cursor(window, True)
    if status_label is not None and loading_text:
        ui_feedback.set_semantic_label(status_label, loading_text, state=loading_state)

    result = {"value": None, "error": None}

    def _worker():
        try:
            result["value"] = worker()
        except Exception as exc:
            result["error"] = exc

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()

    def _finish():
        if thread.is_alive():
            try:
                window.after(120, _finish)
            except Exception:
                pass
            return
        _restore_widget_snapshots(snapshots)
        _set_window_busy_cursor(window, False)
        setattr(window, busy_attr, False)
        error = result.get("error")
        if error is not None:
            if callable(on_error):
                on_error(error)
            return
        if callable(on_success):
            on_success(result.get("value"))

    try:
        window.after(120, _finish)
    except Exception:
        _restore_widget_snapshots(snapshots)
        _set_window_busy_cursor(window, False)
        setattr(window, busy_attr, False)
    return True


def _show_inline_feedback(window, text, *, state="error"):
    message = str(text or "").strip()
    if hasattr(window, "section_feedback_banner"):
        ui_feedback.set_banner(window, message, state=state)
        return True
    label = getattr(window, "status_label", None)
    if label is not None:
        ui_feedback.set_semantic_label(label, message, state=state)
        return True
    label = getattr(window, "status_label_widget", None)
    var = getattr(window, "status_var", None)
    if label is not None and var is not None:
        try:
            var.set(message)
        except Exception:
            pass
        try:
            label.config(fg=ui_feedback.state_color(state))
        except Exception:
            pass
        return True
    label = getattr(window, "login_status", None)
    if label is not None:
        ui_feedback.set_semantic_label(label, message, state=state)
        return True
    return False


def _clear_inline_feedback(window):
    if hasattr(window, "section_feedback_banner"):
        ui_feedback.clear_banner(window)
    label = getattr(window, "status_label", None)
    if label is not None:
        ui_feedback.set_semantic_label(label, "", state="info")


def _collect_field_widgets(value, mapping, current_field_id=None):
    if value is None:
        return
    if isinstance(value, tk.Widget):
        if current_field_id:
            mapping.setdefault(str(current_field_id), []).append(value)
        return
    if isinstance(value, dict):
        for key, item in value.items():
            next_field_id = current_field_id or (str(key or "").strip() if str(key or "").strip() else None)
            _collect_field_widgets(item, mapping, next_field_id)
        return
    if isinstance(value, (list, tuple, set)):
        for item in value:
            _collect_field_widgets(item, mapping, current_field_id)


def _section_widget_sources(window, section_id):
    sources = []
    section_key = str(section_id or "").strip()
    if section_key == "section_1":
        sources.append(getattr(window, "fields", None))
    attr_name = f"{section_key}_fields"
    sources.append(getattr(window, attr_name, None))
    if section_key == "section_2":
        sources.append(getattr(window, "oferente_blocks", None))
        sources.append(getattr(window, "vinculado_blocks", None))
    return [item for item in sources if item is not None]


def _register_section_feedback_fields(window, section_id):
    mapping = {}
    for source in _section_widget_sources(window, section_id):
        _collect_field_widgets(source, mapping)
    for field_id, widgets in mapping.items():
        for widget in list(widgets or []):
            ui_feedback.register_field(window, field_id, widget)
    return mapping


def _show_validation_issues_inline(window, issues, *, title="Revisa los campos marcados antes de continuar."):
    issue_list = [issue for issue in list(issues or []) if isinstance(issue, ValidationIssue)]
    if not issue_list:
        return False
    current_section = str(issue_list[0].section_id or "").strip()
    ui_feedback.clear_field_errors(window)
    section_widgets = _register_section_feedback_fields(window, current_section)
    section_issues = [issue for issue in issue_list if str(issue.section_id or "").strip() == current_section]
    for issue in section_issues:
        field_id = str(issue.field_id or "").strip()
        if not field_id:
            continue
        widgets = list(section_widgets.get(field_id) or [])
        for widget in widgets:
            ui_feedback.register_field(window, field_id, widget)
        ui_feedback.set_field_error(window, field_id, issue.message or "Campo obligatorio sin diligenciar.")
    summary = format_issues_for_message(section_issues or issue_list, title=title, limit=4)
    _show_inline_feedback(window, summary, state="error")
    if section_issues:
        preferred_fields = [str(issue.field_id or "").strip() for issue in section_issues if str(issue.field_id or "").strip()]
        ui_feedback.focus_first_invalid_field(window, preferred_fields)
    return True


def _update_wizard_progress(window, section_id=None):
    label = getattr(window, "header_progress_label", None)
    if label is None:
        return
    form_meta = getattr(window, "_form_meta", None) or _resolve_form_meta(getattr(window, "_form_id", ""))
    sections = list(form_meta.get("wizard_sections") or [])
    if not sections:
        sections = _discover_wizard_sections(window)
    total = int(form_meta.get("wizard_steps") or len(sections) or 0)
    current = str(section_id or getattr(window, "_current_section", "") or "section_1").strip()
    if not total:
        try:
            label.config(text="")
        except Exception:
            pass
        return
    if current in sections:
        index = sections.index(current) + 1
    else:
        index = min(max(1, len(sections) or 1), total)
    try:
        label.config(text=f"Sección {index} de {total}")
    except Exception:
        pass


def _natural_section_sort_key(section_id):
    tokens = re.split(r"(\d+)", str(section_id or ""))
    key = []
    for token in tokens:
        if not token:
            continue
        if token.isdigit():
            key.append((0, int(token)))
        else:
            key.append((1, token))
    return key


def _discover_wizard_sections(window):
    sections = []
    for name in dir(window):
        if not str(name).startswith("_show_section_"):
            continue
        candidate = str(name).replace("_show_", "", 1)
        if candidate not in sections:
            sections.append(candidate)
    return sorted(sections, key=_natural_section_sort_key)


def _init_wizard_header(window, *, title, subtitle):
    header = tk.Frame(window, bg=COLOR_LIGHT_BG)
    header.pack(fill="x", padx=FORM_PADX, pady=(24, 8))
    window.header_title = tk.Label(
        header,
        text=title,
        font=FONT_TITLE,
        fg=COLOR_PRIMARY,
        bg=COLOR_LIGHT_BG,
    )
    window.header_title.pack(anchor="w")
    window.header_subtitle = tk.Label(
        header,
        text=subtitle,
        font=FONT_SUBTITLE,
        fg="#333333",
        bg=COLOR_LIGHT_BG,
    )
    window.header_subtitle.pack(anchor="w", pady=(4, 0))
    window.header_progress_label = tk.Label(
        header,
        text="",
        font=("Arial", 9, "bold"),
        fg="#4f5b66",
        bg=COLOR_LIGHT_BG,
    )
    window.header_progress_label.pack(anchor="w", pady=(6, 0))
    _update_wizard_progress(window, getattr(window, "_current_section", "section_1"))


def _init_wizard_section_container(window):
    window.section_container = tk.Frame(window, bg=COLOR_LIGHT_BG)
    window.section_container.pack(fill="both", expand=True, padx=FORM_PADX, pady=8)
    banner = tk.Label(
        window.section_container,
        text="",
        font=("Arial", 9),
        fg=COLOR_DANGER,
        bg=COLOR_LIGHT_BG,
        justify="left",
        anchor="w",
        wraplength=920,
    )
    banner.pack(fill="x", pady=(0, 8))
    window.section_feedback_banner = banner
    ui_feedback.register_banner_label(window, banner)


def _ensure_wizard_runtime_widgets(window):
    header_title = getattr(window, "header_title", None)
    if header_title is not None and not getattr(window, "header_progress_label", None):
        header = getattr(header_title, "master", None)
        if header is not None:
            progress = tk.Label(
                header,
                text="",
                font=("Arial", 9, "bold"),
                fg="#4f5b66",
                bg=COLOR_LIGHT_BG,
            )
            progress.pack(anchor="w", pady=(6, 0))
            window.header_progress_label = progress
    container = getattr(window, "section_container", None)
    banner = getattr(window, "section_feedback_banner", None)
    banner_exists = False
    try:
        banner_exists = bool(banner is not None and banner.winfo_exists())
    except Exception:
        banner_exists = False
    if container is not None and not banner_exists:
        first_child = None
        try:
            children = list(container.winfo_children())
            first_child = children[0] if children else None
        except Exception:
            first_child = None
        banner = tk.Label(
            container,
            text="",
            font=("Arial", 9),
            fg=COLOR_DANGER,
            bg=COLOR_LIGHT_BG,
            justify="left",
            anchor="w",
            wraplength=920,
        )
        if first_child is not None:
            banner.pack(fill="x", pady=(0, 8), before=first_child)
        else:
            banner.pack(fill="x", pady=(0, 8))
        window.section_feedback_banner = banner
        ui_feedback.register_banner_label(window, banner)
    _update_wizard_progress(window, getattr(window, "_current_section", "section_1"))


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
            resolved = str(value or "")
            alias_map = getattr(widget, "_snapshot_value_aliases", None)
            if isinstance(alias_map, dict):
                direct = alias_map.get(resolved)
                if direct is None:
                    direct = next(
                        (
                            alias_value
                            for alias_key, alias_value in alias_map.items()
                            if str(alias_key).casefold() == resolved.casefold()
                        ),
                        None,
                    )
                if direct is not None:
                    resolved = str(direct or "")
            values = [str(item) for item in (widget.cget("values") or []) if str(item).strip()]
            if values:
                exact = next((item for item in values if item == resolved), None)
                if exact is None:
                    exact = next((item for item in values if item.casefold() == resolved.casefold()), None)
                if exact is not None:
                    resolved = exact
            widget.set(resolved)
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


# ── HELPERS: Test fill (relleno automático para QA) ─────────────────────────


def _login_allows_test_fill(login):
    enabled = str(os.getenv("RECA_ENABLE_TEST_FILL") or "").strip().lower()
    if enabled not in {"1", "true", "yes", "on"}:
        return False
    allowed_login = _normalize_login_value(os.getenv("RECA_TEST_FILL_LOGIN") or "")
    if not allowed_login:
        return False
    return _normalize_login_value(login) == allowed_login


def _pick_test_combobox_value(values):
    for value in list(values or []):
        text = str(value or "")
        if text.strip():
            return text
    return ""


def _get_test_fill_entry_value(kind="", max_len=None):
    normalized_kind = str(kind or "").strip().lower()
    if normalized_kind == "numeric":
        try:
            size = int(max_len)
        except Exception:
            size = 1
        size = max(1, size)
        return "1" * min(size, 4)
    if normalized_kind == "decimal":
        return "1"
    if normalized_kind == "birthdate":
        return "01/01/2000"
    return TEST_FILL_DEFAULT_TEXT


def _window_allows_test_fill(window):
    form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
    if not form_id:
        return False
    hub = getattr(window, "master", None)
    if hub is None:
        return False
    login = (
        getattr(hub, "current_user_profile", {}).get("usuario_login")
        or getattr(hub, "current_user", "")
    )
    return _login_allows_test_fill(login)


def _emit_test_fill_events(widget):
    try:
        if isinstance(widget, ttk.Combobox):
            widget.event_generate("<<ComboboxSelected>>")
            widget.event_generate("<FocusOut>")
            return
        if isinstance(widget, DateEntry):
            widget.event_generate("<<DateEntrySelected>>")
            widget.event_generate("<FocusOut>")
            return
        widget.event_generate("<FocusOut>")
    except Exception:
        return


def _fill_widget_for_test(widget):
    try:
        if not widget.winfo_exists() or not widget.winfo_ismapped():
            return False
    except Exception:
        return False

    try:
        if isinstance(widget, tk.Text):
            widget.delete("1.0", tk.END)
            widget.insert("1.0", TEST_FILL_DEFAULT_TEXT)
            _emit_test_fill_events(widget)
            return True
        if isinstance(widget, ttk.Combobox):
            state = str(widget.cget("state") or "")
            if state == "disabled":
                return False
            combo_value = _pick_test_combobox_value(widget.cget("values"))
            if combo_value:
                widget.set(combo_value)
            elif state != "readonly":
                widget.set(TEST_FILL_DEFAULT_TEXT)
            else:
                return False
            _emit_test_fill_events(widget)
            return True
        if isinstance(widget, DateEntry):
            state = str(widget.cget("state") or "")
            if state == "disabled":
                return False
            widget.set_date(_get_colombia_now().date())
            _emit_test_fill_events(widget)
            return True
        if isinstance(widget, (tk.Entry, ttk.Entry)):
            state = str(widget.cget("state") or "")
            if state in {"readonly", "disabled"}:
                return False
            value = _get_test_fill_entry_value(
                getattr(widget, "_test_fill_kind", ""),
                getattr(widget, "_test_fill_max_len", None),
            )
            widget.delete(0, tk.END)
            widget.insert(0, value)
            _emit_test_fill_events(widget)
            return True
    except Exception:
        return False
    return False


def _fill_current_section_with_test_data(window):
    if not _window_allows_test_fill(window):
        return False
    section_root = getattr(window, "section_container", None) or window
    changed_total = 0
    for _attempt in range(3):
        changed = 0
        for _path, widget in _iter_widget_paths(section_root):
            changed += 1 if _fill_widget_for_test(widget) else 0
        changed_total += changed
        try:
            window.update_idletasks()
        except Exception:
            pass
        if changed == 0:
            break
    autosave_fn = getattr(window, "_pending_autosave", None)
    if callable(autosave_fn):
        try:
            autosave_fn()
        except Exception as exc:
            _log_capture(
                f"[TEST_FILL] autosave_failed form={getattr(window, '_form_id', '')} "
                f"section={getattr(window, '_current_section', '')} err={exc}"
            )
    try:
        _refresh_form_save_status(window)
    except Exception:
        pass
    hub = getattr(window, "master", None)
    if changed_total and hub and hasattr(hub, "show_toast"):
        try:
            hub.show_toast("Seccion diligenciada en modo test")
        except Exception:
            pass
    return changed_total > 0


def _get_test_fill_command(window):
    if not _window_allows_test_fill(window):
        return None
    return lambda w=window: _fill_current_section_with_test_data(w)


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


_FORM_REQUIRED_SECTION_LABELS = {
    "induccion_organizacional": {"section_3": "Seccion 3"},
    "induccion_operativa": {"section_3": "Seccion 3"},
}


def _humanize_section_id(section_id):
    text = str(section_id or "").strip()
    if not text:
        return ""
    match = re.fullmatch(r"section_(\d+)", text)
    if match:
        return f"Seccion {match.group(1)}"
    return text.replace("_", " ").strip().title()


def _ensure_form_save_status_label(window):
    if not window or not hasattr(window, "winfo_exists") or not window.winfo_exists():
        return None
    status_var = getattr(window, "_save_status_var", None)
    if status_var is not None:
        return status_var
    status_var = tk.StringVar(window, value="Ultimo guardado: pendiente")
    label = tk.Label(
        window,
        textvariable=status_var,
        font=("Segoe UI", 9),
        fg="#4f5b66",
        bg=COLOR_LIGHT_BG,
        anchor="e",
        justify="right",
    )
    label.place(relx=1.0, x=-24, y=24, anchor="ne")
    window._save_status_var = status_var
    window._save_status_label = label
    window._save_status_timestamp = ""
    return status_var


def _render_save_status_text(saved_at, section_id="", source=""):
    ts = str(saved_at or "").strip()
    if not ts:
        return "Ultimo guardado: pendiente"
    detail = _humanize_section_id(section_id)
    if detail:
        return f"Ultimo guardado: {ts} | {detail}"
    return f"Ultimo guardado: {ts}"


def _update_window_save_status(window, saved_at, section_id="", source=""):
    status_var = _ensure_form_save_status_label(window)
    if status_var is None:
        return
    ts = str(saved_at or "").strip()
    current_ts = str(getattr(window, "_save_status_timestamp", "") or "").strip()
    if current_ts and ts and ts < current_ts:
        return
    if not ts and current_ts:
        return
    window._save_status_timestamp = ts
    status_var.set(_render_save_status_text(ts, section_id=section_id, source=source))


def _refresh_form_save_status(window):
    form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
    module = FORM_MODULE_MAP.get(form_id)
    if not module or not hasattr(module, "get_form_cache"):
        return
    try:
        cache_snapshot = module.get_form_cache() or {}
    except Exception:
        return
    if not isinstance(cache_snapshot, dict):
        return
    _update_window_save_status(
        window,
        cache_snapshot.get("_last_saved_at"),
        section_id=cache_snapshot.get("_last_saved_section"),
        source=cache_snapshot.get("_last_saved_source"),
    )


def _dropdown_prefix_key(value):
    normalized = _normalize_ascii_text(value).lower()
    if not normalized:
        return None
    if "no aplica" in normalized:
        return "no aplica"
    for prefix in ("0.", "1.", "2.", "3."):
        if normalized.startswith(prefix) or normalized == prefix[:1]:
            return prefix
    return None


def _resolve_prefixed_dropdown_value(source_value, target_values):
    prefix_key = _dropdown_prefix_key(source_value)
    if not prefix_key:
        return ""
    values = tuple(target_values or ())
    if prefix_key == "no aplica":
        for option in values:
            if "no aplica" in _normalize_ascii_text(option).lower():
                return option
        for option in values:
            if _normalize_ascii_text(option).lower().startswith("0."):
                return option
        return ""
    for option in values:
        if _normalize_ascii_text(option).lower().startswith(prefix_key):
            return option
    if prefix_key == "0.":
        for option in values:
            if "no aplica" in _normalize_ascii_text(option).lower():
                return option
    return ""


def _is_prefixed_dropdown_widget(widget):
    if widget is None:
        return False
    try:
        values = tuple(widget.cget("values"))
    except Exception:
        return False
    if not values:
        return False
    normalized_values = [_normalize_ascii_text(value).lower() for value in values]
    return any(
        value.startswith(("0.", "1.", "2.", "3.")) or "no aplica" in value
        for value in normalized_values
    )


def _bind_prefixed_dropdown_fields(fields_map, preferred_suffixes=("_nivel_apoyo", "_nivel_apoyo_requerido")):
    dropdown_entries = [
        (field_id, widget)
        for field_id, widget in (fields_map or {}).items()
        if _is_prefixed_dropdown_widget(widget)
    ]
    if len(dropdown_entries) < 2:
        return

    widgets = [widget for _, widget in dropdown_entries]
    preferred_widget = None
    for suffix in preferred_suffixes:
        preferred_widget = next(
            (widget for field_id, widget in dropdown_entries if field_id.endswith(suffix)),
            None,
        )
        if preferred_widget is not None:
            break
    if preferred_widget is None:
        preferred_widget = widgets[0]

    sync_state = {"active": False}

    def _sync_from(source_widget):
        if sync_state["active"]:
            return
        try:
            source_value = source_widget.get().strip()
        except Exception:
            return
        if not source_value:
            return
        sync_state["active"] = True
        try:
            for target_widget in widgets:
                if target_widget is source_widget:
                    continue
                resolved = _resolve_prefixed_dropdown_value(
                    source_value,
                    tuple(target_widget.cget("values")),
                )
                if resolved:
                    target_widget.set(resolved)
        finally:
            sync_state["active"] = False

    for widget in widgets:
        widget._nivel_apoyo_observacion_sync = lambda _event=None, w=preferred_widget: _sync_from(w)
        widget._prefixed_dropdown_sync = lambda _event=None, w=widget: _sync_from(w)
        widget.bind("<<ComboboxSelected>>", lambda _event, w=widget: _sync_from(w), add="+")
        widget.bind("<FocusOut>", lambda _event, w=widget: _sync_from(w), add="+")

    _sync_from(preferred_widget)


def _bind_prefixed_dropdown_subset(fields_map, field_ids):
    subset = {
        field_id: fields_map[field_id]
        for field_id in (field_ids or ())
        if field_id in (fields_map or {})
    }
    _bind_prefixed_dropdown_fields(subset)


def _find_dropdown_option(target_values, predicate):
    for option in tuple(target_values or ()):
        if predicate(_normalize_ascii_text(option).lower()):
            return option
    return ""


def _resolve_yes_no_dropdown_value(target_values, desired_state):
    normalized_desired = _normalize_ascii_text(desired_state).lower()
    return _find_dropdown_option(
        target_values,
        lambda option: option == normalized_desired,
    )


def _bind_selection_activity_dropdown_fields(
    fields_map,
    *,
    primary_field_id,
    secondary_field_id,
    dependent_field_ids=(),
):
    if not isinstance(fields_map, dict):
        return
    primary_widget = fields_map.get(primary_field_id)
    secondary_widget = fields_map.get(secondary_field_id)
    if not _is_prefixed_dropdown_widget(primary_widget) or not _is_prefixed_dropdown_widget(secondary_widget):
        return

    dependent_widgets = [
        fields_map[field_id]
        for field_id in (dependent_field_ids or ())
        if field_id in fields_map
    ]
    sync_state = {"active": False}

    def _set_widget_value(widget, value):
        try:
            widget.set(value)
        except Exception:
            return

    def _sync_from_primary(_event=None):
        if sync_state["active"]:
            return
        try:
            source_value = primary_widget.get().strip()
        except Exception:
            return
        if not source_value:
            return

        prefix_key = _dropdown_prefix_key(source_value)
        if not prefix_key:
            return

        sync_state["active"] = True
        try:
            resolved_secondary = _resolve_prefixed_dropdown_value(
                source_value,
                tuple(secondary_widget.cget("values")),
            )
            if resolved_secondary:
                _set_widget_value(secondary_widget, resolved_secondary)

            if prefix_key == "0.":
                dependent_value = "No"
            elif prefix_key == "no aplica":
                dependent_value = "No aplica"
            else:
                dependent_value = ""

            for widget in dependent_widgets:
                if not isinstance(widget, ttk.Combobox):
                    continue
                if dependent_value:
                    resolved = _resolve_yes_no_dropdown_value(
                        tuple(widget.cget("values")),
                        dependent_value,
                    )
                    _set_widget_value(widget, resolved or dependent_value)
                else:
                    _set_widget_value(widget, "")
        finally:
            sync_state["active"] = False

    primary_widget._selection_activity_sync = _sync_from_primary
    primary_widget._nivel_apoyo_observacion_sync = _sync_from_primary
    primary_widget.bind("<<ComboboxSelected>>", _sync_from_primary, add="+")
    primary_widget.bind("<FocusOut>", _sync_from_primary, add="+")


def _section_history_has_meaningful_payload(cache_snapshot, section_id):
    if not isinstance(cache_snapshot, dict):
        return False
    history_root = cache_snapshot.get("_section_history")
    if not isinstance(history_root, dict):
        return False
    entries = history_root.get(section_id)
    if not isinstance(entries, list):
        return False
    for entry in reversed(entries):
        if not isinstance(entry, dict):
            continue
        if _cache_snapshot_has_meaningful_values(entry.get("payload")):
            return True
    return False


def _find_guarded_missing_sections(form_id, cache_snapshot):
    labels = _FORM_REQUIRED_SECTION_LABELS.get(str(form_id or "").strip(), {})
    if not isinstance(cache_snapshot, dict) or not labels:
        return []
    missing = []
    for section_id, label in labels.items():
        if _cache_snapshot_has_meaningful_values(cache_snapshot.get(section_id)):
            continue
        if _section_history_has_meaningful_payload(cache_snapshot, section_id):
            missing.append((section_id, label))
    return missing


def _run_pending_section_autosave(window):
    autosave_fn = getattr(window, "_pending_autosave", None)
    if callable(autosave_fn):
        try:
            autosave_fn()
        except Exception as exc:
            _log_capture(
                f"[AUTOSAVE] flush_failed form={getattr(window, '_form_id', '')} "
                f"section={getattr(window, '_current_section', '')} err={exc}"
            )


def _guard_form_action(window, *, action_label):
    if str(action_label or "").strip().lower() == "cerrar":
        linked_state = getattr(window, "_linked_interpreter_state", None)
        if isinstance(linked_state, dict):
            status = str(linked_state.get("status") or "idle").strip().lower()
            show_section = linked_state.get("show_final_section")
            if status == "running":
                if callable(show_section):
                    try:
                        show_section()
                    except Exception:
                        pass
                messagebox.showerror(
                    "Acta de intérprete en proceso",
                    "No se puede cerrar esta acta mientras se está creando el acta de intérprete.\n\n"
                    "Espera a que termine o vuelve a la última sección para corregir antes de cerrar.",
                    parent=window,
                )
                return True
            if status == "failed":
                if callable(show_section):
                    try:
                        show_section()
                    except Exception:
                        pass
                messagebox.showerror(
                    "Acta de intérprete pendiente",
                    "No se puede cerrar esta acta porque falló la creación del acta de intérprete.\n\n"
                    "Corrige o vuelve a intentar el acta de intérprete antes de cerrar.",
                    parent=window,
                )
                return True
    form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
    module = FORM_MODULE_MAP.get(form_id)
    if not module or not hasattr(module, "get_form_cache"):
        return False
    _run_pending_section_autosave(window)
    if hasattr(module, "save_cache_to_file"):
        try:
            module.save_cache_to_file()
        except Exception:
            pass
    try:
        cache_snapshot = copy.deepcopy(module.get_form_cache() or {})
    except Exception:
        cache_snapshot = {}
    missing_sections = _find_guarded_missing_sections(form_id, cache_snapshot)
    if not missing_sections:
        return False
    labels = ", ".join(label for _, label in missing_sections)
    messagebox.showerror(
        "Datos incompletos",
        f"No se puede {action_label} este formulario porque {labels} quedo vacia despues "
        "de haber tenido informacion guardada.\n\n"
        "La ultima version buena sigue en el historial local de la seccion. "
        "Vuelve a esa seccion y revisa antes de continuar.",
        parent=window,
    )
    first_section = missing_sections[0][0]
    show_fn = getattr(window, f"_show_{first_section}", None)
    if callable(show_fn):
        try:
            show_fn()
        except Exception:
            pass
    return True


def _guard_form_finalization(window, *, loading=None):
    form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
    module = FORM_MODULE_MAP.get(form_id)
    if (
        not module
        or not hasattr(module, "get_form_cache")
        or not hasattr(module, "validate_before_finalize")
    ):
        return False
    _run_pending_section_autosave(window)
    if hasattr(module, "save_cache_to_file"):
        try:
            module.save_cache_to_file()
        except Exception:
            pass
    try:
        cache_snapshot = copy.deepcopy(module.get_form_cache() or {})
    except Exception:
        cache_snapshot = {}
    try:
        issues = list(module.validate_before_finalize(cache_snapshot) or [])
    except Exception as exc:
        _show_inline_feedback(
            window,
            _log_user_error("finalization", exc),
            state="error",
        )
        try:
            if loading is not None:
                loading.close()
        except Exception:
            pass
        return True
    if not issues:
        return False
    try:
        if loading is not None:
            loading.close()
    except Exception:
        pass
    first_issue = issues[0]
    show_fn = getattr(window, f"_show_{first_issue.section_id}", None)
    if callable(show_fn):
        try:
            show_fn()
        except Exception:
            pass
    _show_validation_issues_inline(window, issues)
    messagebox.showerror(
        "Finalización",
        format_issues_for_message(issues),
        parent=window,
    )
    return True


def _get_draft_save_command(window):
    save_cmd = getattr(window, "_save_draft_command", None)
    if callable(save_cmd):
        return save_cmd
    form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
    if not form_id:
        return None
    form_meta = _resolve_form_meta(form_id)
    if not _form_supports_drafts(form_meta):
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
    window._draft_restore_route_name = str(getattr(route, "__name__", "") or "").strip()
    if window._draft_restore_target_section:
        try:
            window._current_section = window._draft_restore_target_section
        except Exception:
            pass
    if not callable(route):
        return False
    if not getattr(window, "_runtime_sections_ready", False):
        return True
    try:
        route()
    except Exception:
        try:
            window.after_idle(route)
        except Exception:
            return False
    return True


# ── HELPERS: Normalización de texto, encoding, entradas numéricas ────────────


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
        os.path.join(project_root, "tests"),
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
    entry._test_fill_kind = "numeric"
    entry._test_fill_max_len = max_len

    def _on_key_release(_event=None):
        raw = _digits_only(entry.get(), max_len=max_len)
        if entry.get() == raw:
            return
        entry.delete(0, tk.END)
        entry.insert(0, raw)

    entry.bind("<KeyRelease>", _on_key_release)


def _bind_decimal_entry(entry):
    entry._test_fill_kind = "decimal"

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
    entry._test_fill_kind = "name"

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


def _bind_lsc_time_entry(entry, on_change=None):
    def _normalize_time(_event=None):
        normalized = interprete_lsc.normalize_time_value(entry.get())
        entry.delete(0, tk.END)
        if normalized:
            entry.insert(0, normalized)
        if callable(on_change):
            on_change()

    entry.bind("<FocusOut>", _normalize_time, add="+")


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
    date_entry._test_fill_kind = "birthdate"

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


# ── HELPERS: Autenticación — hash de contraseña, login offline, conectividad ─


def _hash_password(password, iterations=PASSWORD_HASH_ITERATIONS):
    pwd = str(password or "")
    if len(pwd) > MAX_PASSWORD_LENGTH:
        raise ValueError("La contraseña supera la longitud máxima permitida.")
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
    if len(raw) > MAX_PASSWORD_LENGTH:
        return []
    options = [raw]
    trimmed = raw.strip()
    if trimmed != raw:
        options.append(trimmed)
    return options


def _is_invalid_credentials_exception(exc):
    """Return True when Supabase rejects the email/password (401 / invalid_grant)."""
    if exc is None:
        return False
    root = exc.__cause__ if isinstance(exc, RuntimeError) and exc.__cause__ else exc
    if isinstance(root, urllib.error.HTTPError):
        if int(getattr(root, "code", 0) or 0) in (400, 401, 422):
            return True
    text = str(exc).lower()
    return "invalid login credentials" in text or "invalid_grant" in text or "email not confirmed" in text


def _is_profile_permission_exception(exc):
    if exc is None:
        return False
    root = exc.__cause__ if isinstance(exc, RuntimeError) and exc.__cause__ else exc
    if isinstance(root, urllib.error.HTTPError):
        if int(getattr(root, "code", 0) or 0) in (401, 403):
            return True
    text = str(exc).lower()
    return (
        "permission denied" in text
        or "42501" in text
        or "profesionales" in text
    ) and ("http 401" in text or "http 403" in text or "permission denied" in text)


def _offline_auth_entry_is_expired(entry):
    if not isinstance(entry, dict):
        return True
    try:
        cached_at = float(entry.get("cached_at") or 0)
    except Exception:
        cached_at = 0.0
    try:
        ttl_days = int(entry.get("ttl_days") or OFFLINE_AUTH_TTL_DAYS)
    except Exception:
        ttl_days = OFFLINE_AUTH_TTL_DAYS
    if cached_at <= 0:
        return True
    return (time.time() - cached_at) > max(1, ttl_days) * 86400


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


# ── HELPERS: Conectividad, probe de servicios, utilidades de ventana ─────────


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


def _get_services_badge_state(internet_state, supabase_state, drive_state, total_pending=0, total_failed=0):
    if int(total_failed or 0) > 0:
        return "● Error de sincronización", COLOR_DANGER
    if int(total_pending or 0) > 0:
        return f"● {int(total_pending or 0)} por sincronizar", COLOR_WARNING

    internet_ok = bool((internet_state or {}).get("ok"))
    if not internet_ok:
        return "● Sin conexión", COLOR_DANGER

    service_states = [supabase_state or {}, drive_state or {}]
    if all(bool(state.get("ok")) for state in service_states):
        return "● Conectado", COLOR_SUCCESS

    error_codes = {str((state or {}).get("error_code") or "").strip() for state in service_states if not state.get("ok")}
    if error_codes & {"credentials", "config", "folder_config", "missing_dependencies"}:
        return "● Configuración incompleta", COLOR_WARNING
    if error_codes & {"auth"}:
        return "● Credenciales inválidas", COLOR_WARNING
    return "● Servicios no disponibles", COLOR_WARNING


def probe_startup_services(log_enabled=False, require_drive_write=False):
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
        "drive": drive_upload.probe_drive_service(
            log_enabled=log_enabled,
            require_write=require_drive_write,
        ),
    }


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


def _find_chrome_executable():
    candidates = []
    local_app_data = str(os.getenv("LOCALAPPDATA") or "").strip()
    program_files = str(os.getenv("PROGRAMFILES") or "").strip()
    program_files_x86 = str(os.getenv("PROGRAMFILES(X86)") or "").strip()
    for base in (local_app_data, program_files, program_files_x86):
        if not base:
            continue
        candidates.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
    try:
        import winreg  # type: ignore

        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(
                    root,
                    r"Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                ) as key:
                    value, _ = winreg.QueryValueEx(key, "")
                    if value:
                        candidates.insert(0, str(value))
            except OSError:
                continue
    except Exception:
        pass
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return ""


def _open_url_prefer_chrome(url):
    target = str(url or "").strip()
    if not target:
        raise RuntimeError("No se indicó una URL para abrir.")
    chrome_path = _find_chrome_executable()
    if chrome_path:
        subprocess.Popen([chrome_path, target], close_fds=True)
        return
    webbrowser.open(target)


def _is_path_within_root(path, root):
    try:
        normalized_path = os.path.normcase(os.path.abspath(str(path or "").strip()))
        normalized_root = os.path.normcase(os.path.abspath(str(root or "").strip()))
        if not normalized_path or not normalized_root:
            return False
        return os.path.commonpath([normalized_path, normalized_root]) == normalized_root
    except Exception:
        return False


def _default_safe_open_roots():
    local_app_root = str(os.getenv("LOCALAPPDATA") or "").strip()
    roots = [
        _get_desktop_dir(),
        _get_local_cache_dir(),
        os.path.join(local_app_root, "RECA") if local_app_root else "",
        os.path.join(tempfile.gettempdir(), "reca_seguimientos_drive"),
        seguimientos._get_shared_root(),
        os.getcwd(),
    ]
    unique = []
    seen = set()
    for path in roots:
        candidate = str(path or "").strip()
        if not candidate:
            continue
        key = os.path.normcase(os.path.abspath(candidate))
        if key in seen:
            continue
        seen.add(key)
        unique.append(candidate)
    return unique


def _open_local_file_safely(path, allowed_roots=None):
    target = os.path.abspath(str(path or "").strip())
    if not target:
        raise RuntimeError("No se indicó un archivo para abrir.")
    if not os.path.exists(target):
        raise RuntimeError("No se encontró el archivo solicitado.")
    extension = os.path.splitext(target)[1].lower()
    if extension not in {".xlsx", ".xlsm", ".xls", ".pdf"}:
        raise RuntimeError("El tipo de archivo no está permitido para apertura directa.")
    roots = list(allowed_roots or ()) + _default_safe_open_roots()
    if not any(_is_path_within_root(target, root) for root in roots):
        raise RuntimeError("La ruta solicitada no está permitida para apertura directa.")
    os.startfile(target)


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
                    _open_url_prefer_chrome(open_target)
                else:
                    _open_local_file_safely(
                        open_target,
                        allowed_roots=[os.path.dirname(os.path.abspath(str(open_target)))],
                    )
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


def _resolve_hub_window(window):
    current = getattr(window, "master", None)
    while current is not None:
        if isinstance(current, HubWindow):
            return current
        current = getattr(current, "master", None)
    return None


def _return_to_hub(window):
    hub = _resolve_hub_window(window)
    if not hub:
        return
    try:
        hub.deiconify()
        hub.lift()
        hub.focus_force()
    except tk.TclError:
        return


def _show_acta_published_dialog(
    parent,
    *,
    sheet_url,
    company_name="",
    pdf_folder_url=None,
    dialog_title="Acta publicada",
    header_text="\u2705  \u00a1Acta publicada!",
    body_text=None,
    pdf_status_text=None,
    open_sheet_label="Abrir Google Sheet",
):
    """Dialog de éxito al publicar un acta.

    Muestra un mensaje amigable y permite abrir el Google Sheet y/o la
    carpeta de PDFs en Drive. Si ``pdf_folder_url`` se proporciona, ofrece
    tres acciones: abrir solo el Sheet, solo la carpeta de PDFs, o ambos.
    """
    dialog = tk.Toplevel(parent)
    dialog.title(str(dialog_title or "Acta publicada"))
    dialog.configure(bg=COLOR_SURFACE)
    dialog.resizable(False, False)
    dialog.grab_set()

    dialog_w = 440

    # ── Header ──────────────────────────────────────────────────────────
    header = tk.Frame(dialog, bg=COLOR_SUCCESS, height=58)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(
        header,
        text=str(header_text or "\u2705  \u00a1Acta publicada!"),
        font=("Arial", 14, "bold"),
        fg="#FFFFFF",
        bg=COLOR_SUCCESS,
    ).pack(expand=True)

    # ── Cuerpo ───────────────────────────────────────────────────────────
    # fill="x" (no expand) para que no consuma espacio que necesitan los botones
    body = tk.Frame(dialog, bg=COLOR_SURFACE, padx=24, pady=14)
    body.pack(fill="x")

    company_line = f" de {company_name}" if company_name else ""
    resolved_body_text = str(body_text or "").strip() or (
        f"El acta{company_line} quedó guardada en Google Sheets."
    )
    tk.Label(
        body,
        text=resolved_body_text,
        font=("Arial", 11),
        fg="#2D2D2D",
        bg=COLOR_SURFACE,
        wraplength=390,
        justify="left",
    ).pack(anchor="w")

    if pdf_folder_url:
        resolved_pdf_status = str(pdf_status_text or "").strip() or (
            "El PDF se est\u00e1 generando y estar\u00e1 disponible\nen la carpeta de Drive en unos segundos."
        )
        tk.Label(
            body,
            text=resolved_pdf_status,
            font=("Arial", 10),
            fg="#5B5563",
            bg=COLOR_SURFACE,
            justify="left",
        ).pack(anchor="w", pady=(8, 0))

    # ── Separador ────────────────────────────────────────────────────────
    tk.Frame(dialog, bg=COLOR_BORDER, height=1).pack(fill="x", padx=24)

    # ── Botones ──────────────────────────────────────────────────────────
    btn_area = tk.Frame(dialog, bg=COLOR_SURFACE, padx=24, pady=16)
    btn_area.pack(fill="x")

    def _open_sheet():
        if sheet_url:
            _open_url_prefer_chrome(sheet_url)
        dialog.destroy()

    def _open_pdf_folder():
        if pdf_folder_url:
            _open_url_prefer_chrome(pdf_folder_url)
        dialog.destroy()

    def _open_both():
        if sheet_url:
            _open_url_prefer_chrome(sheet_url)
        if pdf_folder_url:
            _open_url_prefer_chrome(pdf_folder_url)
        dialog.destroy()

    _BTN = dict(font=("Arial", 10, "bold"), relief="flat", padx=12, pady=8, cursor="hand2", bd=0)

    if pdf_folder_url:
        tk.Label(
            btn_area,
            text="\u00bfQu\u00e9 deseas abrir?",
            font=("Arial", 10, "bold"),
            fg="#2D2D2D",
            bg=COLOR_SURFACE,
        ).pack(anchor="w", pady=(0, 10))

        row = tk.Frame(btn_area, bg=COLOR_SURFACE)
        row.pack(anchor="w")

        tk.Button(
            row, text=str(open_sheet_label or "Abrir Google Sheet"),
            bg=COLOR_PRIMARY, fg="#FFFFFF",
            command=_open_sheet, **_BTN,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            row, text="Ver PDFs en Drive",
            bg=COLOR_ACCENT, fg="#FFFFFF",
            command=_open_pdf_folder, **_BTN,
        ).pack(side="left", padx=(0, 8))

        tk.Button(
            row, text="Abrir ambos",
            bg="#E5E7EB", fg="#374151",
            font=("Arial", 10), relief="flat", padx=10, pady=8, cursor="hand2", bd=0,
            command=_open_both,
        ).pack(side="left")
    else:
        row = tk.Frame(btn_area, bg=COLOR_SURFACE)
        row.pack(anchor="w")

        tk.Button(
            row, text=str(open_sheet_label or "Abrir Google Sheet"),
            bg=COLOR_PRIMARY, fg="#FFFFFF",
            command=_open_sheet, **_BTN,
        ).pack(side="left", padx=(0, 10))

        tk.Button(
            row, text="Cerrar",
            bg="#E5E7EB", fg="#374151",
            font=("Arial", 10), relief="flat", padx=10, pady=8, cursor="hand2", bd=0,
            command=dialog.destroy,
        ).pack(side="left")

    # Centrar una vez que todos los widgets están empaquetados
    dialog.update_idletasks()
    natural_h = dialog.winfo_reqheight()
    px = parent.winfo_rootx() + max(0, (parent.winfo_width() - dialog_w) // 2)
    py = parent.winfo_rooty() + max(0, (parent.winfo_height() - natural_h) // 2)
    dialog.geometry(f"{dialog_w}x{natural_h}+{px}+{py}")
    dialog.lift()
    dialog.focus_force()
    dialog.wait_window()


def _finalize_export_flow(window, loading, completion_result):
    result = dict(completion_result or {})
    status = str(result.get("status") or "").strip()
    output_path = str(result.get("output_path") or "").strip()
    remote_url = str(result.get("remote_url") or "").strip()
    error = str(result.get("error") or "").strip()
    hub = _resolve_hub_window(window)

    if hub and status in {"synced", "pending", "failed", "local"}:
        try:
            hub._delete_window_draft(window)
        except Exception as exc:
            _log_capture(
                f"[DRAFT] delete_after_finalize_failed form={getattr(window, '_form_id', '')} status={status} err={exc}"
            )

    if status == "synced":
        if remote_url:
            # Dialog amigable con acciones directas
            pdf_folder_id = str(result.get("pdf_folder_id") or "").strip()
            pdf_folder_url = (
                f"https://drive.google.com/drive/folders/{pdf_folder_id}"
                if pdf_folder_id
                else None
            )
            company_name = str(result.get("company_name") or "").strip()
            loading.set_status("Listo")
            loading.set_progress(100)
            try:
                loading.window.grab_release()
            except tk.TclError:
                pass
            loading.close()
            _show_acta_published_dialog(
                window,
                sheet_url=remote_url,
                company_name=company_name,
                pdf_folder_url=pdf_folder_url,
            )
        else:
            open_target = output_path
            message = "Formulario completado.\nEl archivo quedó publicado en Google Drive."
            if output_path:
                message = f"{message}\nArchivo local: {output_path}"
            _finish_with_loading(
                loading,
                message,
                open_target=open_target,
                open_prompt="¿Quieres abrir el archivo local?",
            )
        return

    if status == "pending":
        message = (
            "Formulario completado.\n"
            "Se conservó un archivo local y la subida a Google Drive quedó pendiente.\n"
            "La información del formulario se conserva para soporte o reintento manual.\n"
            f"Archivo local: {output_path}"
        )
        if error:
            message = f"{message}\nDetalle: {error}"
        _finish_with_loading(
            loading,
            message,
            open_target=output_path,
            open_prompt="¿Quieres abrir el archivo local?",
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
            open_prompt="¿Quieres abrir el archivo local?",
        )
        return

    if status == "local":
        _finish_with_loading(
            loading,
            "Formato completado. ¿Quieres abrir el archivo local?",
            open_target=output_path,
            open_prompt="¿Quieres abrir el archivo local?",
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
        current_buttons = [w for w in source.winfo_children() if isinstance(w, ttk.Button)]
        paired_sources = [src for src, _clone in pairs if src and src.winfo_exists()]
        if len(current_buttons) != len(paired_sources) or any(
            current is not paired for current, paired in zip(current_buttons, paired_sources)
        ):
            try:
                window.after_idle(lambda src=source: _install_sticky_actions(src))
            except Exception:
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
        test_cmd = _get_test_fill_command(window)
        if callable(test_cmd):
            has_test = False
            for child in frame.winfo_children():
                if isinstance(child, ttk.Button) and str(child.cget("text")).strip().lower() == "test":
                    has_test = True
                    break
            if not has_test:
                ttk.Button(
                    frame,
                    text="Test",
                    command=test_cmd,
                ).pack(side="left", padx=(8, 0))
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
        kind = spec[2] if len(spec) > 2 else "secondary"
        ttk.Button(actions, text=text, command=command, style=_button_style_for_kind(kind)).pack(
            side="left",
            padx=(8, 0) if idx else 0,
        )

    if callable(primary_command):
        ttk.Button(
            actions,
            text=primary_text,
            command=primary_command,
            style=_button_style_for_kind("primary"),
        ).pack(side="right")

    for spec in list(right_buttons or []):
        if not spec:
            continue
        text, command = spec[0], spec[1]
        kind = spec[2] if len(spec) > 2 else "secondary"
        ttk.Button(
            actions,
            text=text,
            command=command,
            style=_button_style_for_kind(kind),
        ).pack(side="right", padx=(0, 8))

    return actions


def _build_inline_error_label(parent, *, bg=COLOR_LIGHT_BG, wraplength=420):
    return tk.Label(
        parent,
        text="",
        font=("Arial", 9),
        fg=COLOR_DANGER,
        bg=bg,
        justify="left",
        anchor="w",
        wraplength=wraplength,
    )


def _section1_search_value(window, field_id):
    widget = (getattr(window, "fields", {}) or {}).get(field_id)
    return ui_feedback.get_widget_value(widget)


def _section1_validate_search(window, mode):
    ui_feedback.clear_field_error(window, "nit_empresa")
    ui_feedback.clear_field_error(window, "nombre_busqueda")
    _clear_inline_feedback(window)
    if mode == "nit":
        nit = _section1_search_value(window, "nit_empresa")
        if nit:
            return nit, ""
        ui_feedback.set_field_error(window, "nit_empresa", "Ingresa un NIT para buscar.")
        ui_feedback.focus_first_invalid_field(window, ["nit_empresa"])
        _show_inline_feedback(window, "Ingresa un NIT para buscar la empresa.", state="error")
        return "", ""
    if mode == "nombre":
        nombre = _section1_search_value(window, "nombre_busqueda")
        if nombre:
            return "", nombre
        ui_feedback.set_field_error(window, "nombre_busqueda", "Ingresa el nombre de la empresa.")
        ui_feedback.focus_first_invalid_field(window, ["nombre_busqueda"])
        _show_inline_feedback(window, "Ingresa el nombre de la empresa para buscar.", state="error")
        return "", ""
    _show_inline_feedback(window, "Tipo de búsqueda no válido.", state="error")
    return "", ""


def _apply_section1_company_result(window, *, lookup, company, mode):
    section_map = getattr(lookup, "SECTION_1_SUPABASE_MAP", presentacion_programa.SECTION_1_SUPABASE_MAP)
    if not company:
        window.company_data = None
        ui_feedback.set_semantic_label(
            window.status_label,
            "No se encontró empresa para ese nombre." if mode == "nombre" else "No se encontró empresa para ese NIT.",
            state="warning",
        )
        try:
            window.continue_btn.config(state="disabled")
        except Exception:
            pass
        for key in section_map.keys():
            window._set_readonly_value(key, "")
        return

    if mode == "nombre":
        nit_value = company.get("nit_empresa")
        entry = window.fields.get("nit_empresa")
        if nit_value and entry:
            entry.delete(0, tk.END)
            entry.insert(0, nit_value)
            entry._ui_placeholder_active = False

    window.company_data = company
    ui_feedback.set_semantic_label(window.status_label, "Empresa encontrada.", state="success")
    try:
        window.continue_btn.config(state="normal")
    except Exception:
        pass
    for key in section_map.keys():
        window._set_readonly_value(key, company.get(key))


def _run_section1_company_search(window, *, mode, lookup, button=None):
    nit, nombre = _section1_validate_search(window, mode)
    if mode == "nit" and not nit:
        return
    if mode == "nombre" and not nombre:
        return

    def _worker():
        if mode == "nombre":
            return lookup.get_empresa_by_nombre(nombre)
        return lookup.get_empresa_by_nit(nit)

    def _on_success(company):
        _apply_section1_company_result(window, lookup=lookup, company=company, mode=mode)

    def _on_error(exc):
        message = _log_user_error("company_search", exc)
        window.company_data = None
        try:
            window.continue_btn.config(state="disabled")
        except Exception:
            pass
        ui_feedback.set_semantic_label(window.status_label, message, state="error")
        _show_inline_feedback(window, message, state="error")

    _run_async_ui_task(
        window,
        busy_attr="_company_lookup_busy",
        widgets=[
            window.fields.get("nit_empresa"),
            window.fields.get("nombre_busqueda"),
            getattr(window, "search_nit_btn", None),
            getattr(window, "search_name_btn", None),
        ],
        loading_button=button,
        loading_button_text="Buscando...",
        status_label=getattr(window, "status_label", None),
        loading_text="Buscando empresa...",
        loading_state="loading",
        worker=_worker,
        on_success=_on_success,
        on_error=_on_error,
    )


def _confirm_section1_and_continue(window, *, confirm_fn, next_step, extra_inputs=None):
    ui_feedback.clear_field_errors(window)
    _clear_inline_feedback(window)
    if not getattr(window, "company_data", None):
        _show_inline_feedback(window, "Busca una empresa antes de continuar.", state="error")
        ui_feedback.focus_first_invalid_field(window, ["nit_empresa", "nombre_busqueda"])
        return

    fecha_visita = _get_required_fecha_visita(window)
    modalidad = _get_required_modalidad(window)
    if not fecha_visita or not modalidad:
        _show_inline_feedback(window, "Completa los campos obligatorios para continuar.", state="error")
        ui_feedback.focus_first_invalid_field(window, ["fecha_visita", "modalidad"])
        return

    user_inputs = {
        "fecha_visita": fecha_visita,
        "modalidad": modalidad,
        "nit_empresa": _section1_search_value(window, "nit_empresa"),
    }
    extra = extra_inputs() if callable(extra_inputs) else (extra_inputs or {})
    if isinstance(extra, dict):
        user_inputs.update(extra)
    try:
        confirm_fn(window.company_data, user_inputs)
    except Exception as exc:
        _show_inline_feedback(window, _log_user_error("section_confirm", exc), state="error")
        return
    _clear_inline_feedback(window)
    next_step()


def _build_lsc_context(window, *, module, source_form, oferentes=None):
    cache = {}
    if module is not None and hasattr(module, "get_form_cache"):
        try:
            cache = module.get_form_cache() or {}
        except Exception:
            cache = {}

    section_1 = cache.get("section_1", {}) if isinstance(cache, dict) else {}
    company_data = getattr(window, "company_data", None)
    empresa = (
        copy.deepcopy(section_1)
        if isinstance(section_1, dict) and section_1.get("nombre_empresa")
        else copy.deepcopy(company_data)
        if isinstance(company_data, dict)
        else {}
    )

    fields = getattr(window, "fields", {}) or {}
    fecha_visita = ""
    if isinstance(section_1, dict):
        fecha_visita = section_1.get("fecha_visita") or ""
    if not fecha_visita and fields.get("fecha_visita") is not None:
        fecha_visita = ui_feedback.get_widget_value(fields.get("fecha_visita"))

    context = {
        "empresa": empresa,
        "oferentes": list(oferentes or []),
        "source_form": source_form,
    }
    if fecha_visita not in (None, ""):
        context["fecha_visita"] = fecha_visita
    return context


def _get_linked_interpreter_state(window):
    state = getattr(window, "_linked_interpreter_state", None)
    if isinstance(state, dict):
        return state
    state = {
        "status": "idle",
        "result": None,
        "error_message": "",
        "show_final_section": None,
        "main_finish_action": None,
        "pending_main_finalize": False,
        "pending_export_action": None,
        "wait_loading": None,
    }
    window._linked_interpreter_state = state
    return state


def _close_linked_interpreter_wait_dialog(window):
    state = _get_linked_interpreter_state(window)
    loading = state.get("wait_loading")
    state["wait_loading"] = None
    if loading is not None:
        _close_loading_async(loading)


def _show_linked_interpreter_wait_dialog(window):
    state = _get_linked_interpreter_state(window)
    loading = state.get("wait_loading")
    if loading is None or not loading.exists():
        loading = LoadingDialog(window, title="Esperando acta de intérprete")
        state["wait_loading"] = loading
    loading.set_status("Se está terminando de crear el acta de intérprete...")
    loading.set_progress(55)
    return loading


def _restore_linked_parent_final_section(window):
    state = _get_linked_interpreter_state(window)
    show_section = state.get("show_final_section")
    if callable(show_section):
        try:
            show_section()
        except Exception:
            pass
    _focus_window(window)


def _ask_linked_interpreter_next_action(parent):
    dialog = tk.Toplevel(parent)
    dialog.title("Acta de intérprete")
    dialog.configure(bg=COLOR_LIGHT_BG)
    dialog.resizable(False, False)
    dialog.transient(parent)
    dialog.grab_set()

    result = {"value": "correct"}

    body = tk.Frame(dialog, bg=COLOR_LIGHT_BG, padx=24, pady=20)
    body.pack(fill="both", expand=True)

    tk.Label(
        body,
        text="Se está creando el acta de intérprete.",
        font=FONT_SECTION,
        fg=COLOR_PURPLE,
        bg=COLOR_LIGHT_BG,
        anchor="w",
    ).pack(fill="x")
    tk.Label(
        body,
        text="¿Deseas terminar esta acta o corregirla?",
        font=FONT_LABEL,
        bg=COLOR_LIGHT_BG,
        fg="#333333",
        justify="left",
        wraplength=380,
        anchor="w",
    ).pack(fill="x", pady=(8, 0))

    actions = tk.Frame(body, bg=COLOR_LIGHT_BG)
    actions.pack(fill="x", pady=(18, 0))

    def _choose(value):
        result["value"] = value
        dialog.destroy()

    ttk.Button(actions, text="Corregir", command=lambda: _choose("correct")).pack(side="right")
    ttk.Button(actions, text="Terminar acta", command=lambda: _choose("finish")).pack(
        side="right",
        padx=(0, 8),
    )

    dialog.protocol("WM_DELETE_WINDOW", lambda: _choose("correct"))
    dialog.update_idletasks()
    width = max(dialog.winfo_reqwidth(), 430)
    height = dialog.winfo_reqheight()
    x = parent.winfo_rootx() + max(0, (parent.winfo_width() - width) // 2)
    y = parent.winfo_rooty() + max(0, (parent.winfo_height() - height) // 2)
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    dialog.lift()
    dialog.focus_force()
    dialog.wait_window()
    return result["value"]


def _queue_or_run_main_form_export(window, export_action):
    state = _get_linked_interpreter_state(window)
    status = str(state.get("status") or "idle").strip().lower()
    if status == "failed":
        _restore_linked_parent_final_section(window)
        message = str(state.get("error_message") or "").strip() or (
            "No se pudo crear el acta de intérprete. Corrige o vuelve a intentar antes de finalizar esta acta."
        )
        _show_inline_feedback(window, message, state="error")
        messagebox.showerror("Acta de intérprete", message, parent=window)
        state["pending_main_finalize"] = False
        state["pending_export_action"] = None
        return False
    if status == "running":
        state["pending_main_finalize"] = True
        state["pending_export_action"] = export_action
        _show_linked_interpreter_wait_dialog(window)
        return False
    state["pending_main_finalize"] = False
    state["pending_export_action"] = None
    export_action()
    return True


def _handle_linked_interpreter_finished(window, *, status, result=None, error_message=""):
    state = _get_linked_interpreter_state(window)
    state["status"] = status
    state["result"] = result
    state["error_message"] = str(error_message or "").strip()

    pending_finalize = bool(state.get("pending_main_finalize"))
    export_action = state.get("pending_export_action")
    _close_linked_interpreter_wait_dialog(window)

    if status == "success":
        if pending_finalize and callable(export_action):
            state["pending_main_finalize"] = False
            state["pending_export_action"] = None
            _clear_inline_feedback(window)
            _safe_widget_after(window, export_action)
            return
        _show_inline_feedback(
            window,
            "El acta de intérprete quedó creada. Puedes terminar esta acta cuando quieras.",
            state="success",
        )
        return

    state["pending_main_finalize"] = False
    state["pending_export_action"] = None
    _restore_linked_parent_final_section(window)
    message = state["error_message"] or (
        "No se pudo crear el acta de intérprete. Corrige o vuelve a intentar antes de finalizar esta acta."
    )
    _show_inline_feedback(window, message, state="error")
    messagebox.showerror("Acta de intérprete", message, parent=window)


def _launch_linked_lsc_window(window, *, context, return_to_final_section, main_finish_action):
    state = _get_linked_interpreter_state(window)
    if str(state.get("status") or "").strip().lower() == "running":
        messagebox.showinfo(
            "Acta de intérprete",
            "Ya hay un acta de intérprete en proceso para esta acta.",
            parent=window,
        )
        return None

    _run_pending_section_autosave(window)
    form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
    module = FORM_MODULE_MAP.get(form_id)
    if module is not None and hasattr(module, "save_cache_to_file"):
        try:
            module.save_cache_to_file()
        except Exception:
            pass

    _close_linked_interpreter_wait_dialog(window)
    state["status"] = "idle"
    state["result"] = None
    state["error_message"] = ""
    state["show_final_section"] = return_to_final_section
    state["main_finish_action"] = main_finish_action
    state["pending_main_finalize"] = False
    state["pending_export_action"] = None

    def _on_started():
        state["status"] = "running"
        state["result"] = None
        state["error_message"] = ""
        _restore_linked_parent_final_section(window)
        _clear_inline_feedback(window)
        choice = _ask_linked_interpreter_next_action(window)
        if choice == "finish" and callable(main_finish_action):
            main_finish_action()
            return
        _show_inline_feedback(
            window,
            "La acta de intérprete sigue en creación. Puedes corregir esta acta y finalizar después.",
            state="info",
        )

    def _on_finished(*, status, result=None, error_message=""):
        _handle_linked_interpreter_finished(
            window,
            status=status,
            result=result,
            error_message=error_message,
        )

    return LSCWindow(
        window,
        context=context,
        linked_mode=True,
        parent_form=window,
        on_linked_export_started=_on_started,
        on_linked_export_finished=_on_finished,
    )


# ── HELPERS: Labs / funciones experimentales ────────────────────────────────


def _confirm_labs_experimental_warning(parent):
    accepted = {"value": False}
    modal = tk.Toplevel(parent)
    modal.title("Labs")
    modal.configure(bg=COLOR_LIGHT_BG)
    modal.geometry("720x360")
    modal.transient(parent)
    modal.grab_set()
    modal.resizable(False, False)

    shell = tk.Frame(modal, bg=COLOR_LIGHT_BG, padx=24, pady=20)
    shell.pack(fill="both", expand=True)

    tk.Label(
        shell,
        text="Función Experimental",
        font=("Arial", 20, "bold"),
        fg="#A40000",
        bg=COLOR_LIGHT_BG,
    ).pack(anchor="center", pady=(0, 12))

    warning_text = (
        "Este flujo de Labs está en desarrollo y es experimental.\n\n"
        "Puede contener errores, guardar información incompleta o producir resultados distintos al flujo oficial "
        "de Selección Incluyente.\n\n"
        "Úsalo solo para pruebas controladas y valida manualmente toda la información antes de continuar."
    )
    tk.Message(
        shell,
        text=warning_text,
        width=620,
        font=("Arial", 12),
        fg="#222222",
        bg=COLOR_LIGHT_BG,
        justify="center",
    ).pack(fill="x", pady=(0, 18))

    actions = tk.Frame(shell, bg=COLOR_LIGHT_BG)
    actions.pack(side="bottom")

    ttk.Button(actions, text="Cancelar", command=modal.destroy).pack(side="left", padx=(0, 10))

    def _accept():
        accepted["value"] = True
        modal.destroy()

    ttk.Button(actions, text="Acepto y continuar", command=_accept).pack(side="left")
    modal.protocol("WM_DELETE_WINDOW", modal.destroy)
    try:
        modal.wait_window()
    except Exception:
        pass
    _log_labs(
        f"experimental_warning accepted={bool(accepted['value'])}"
    )
    return bool(accepted["value"])


def _select_labs_flow(parent):
    selected = {"value": ""}
    modal = tk.Toplevel(parent)
    modal.title("Labs")
    modal.configure(bg=COLOR_LIGHT_BG)
    modal.geometry("760x360")
    modal.transient(parent)
    modal.grab_set()
    modal.resizable(False, False)

    shell = tk.Frame(modal, bg=COLOR_LIGHT_BG, padx=24, pady=20)
    shell.pack(fill="both", expand=True)

    tk.Label(
        shell,
        text="Selecciona el flujo experimental",
        font=("Arial", 18, "bold"),
        fg=COLOR_PURPLE,
        bg=COLOR_LIGHT_BG,
    ).pack(anchor="w", pady=(0, 16))

    options = tk.Frame(shell, bg=COLOR_LIGHT_BG)
    options.pack(fill="both", expand=True)

    def _build_option(title, description, form_id):
        card = tk.Frame(
            options,
            bg="white",
            bd=1,
            relief="solid",
            padx=14,
            pady=14,
        )
        card.pack(fill="x", pady=(0, 10))
        tk.Label(
            card,
            text=title,
            font=("Arial", 12, "bold"),
            bg="white",
            fg="#222222",
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            card,
            text=description,
            font=("Arial", 10),
            bg="white",
            fg="#444444",
            justify="left",
            wraplength=620,
            anchor="w",
        ).pack(fill="x", pady=(6, 10))

        def _choose():
            selected["value"] = form_id
            modal.destroy()

        ttk.Button(card, text="Abrir", command=_choose).pack(anchor="e")

    _build_option(
        "Seleccion Incluyente Labs",
        "Flujo experimental completo de Seleccion Incluyente con dictado por subsecciones.",
        "seleccion_incluyente_labs",
    )
    _build_option(
        "Condiciones de Vacante",
        "Abre la variante experimental de Condiciones de Vacante. El dictado actual esta en la seccion 2.",
        "condiciones_vacante_labs",
    )

    actions = tk.Frame(shell, bg=COLOR_LIGHT_BG)
    actions.pack(fill="x", pady=(8, 0))
    ttk.Button(actions, text="Cancelar", command=modal.destroy).pack(side="right")

    modal.protocol("WM_DELETE_WINDOW", modal.destroy)
    try:
        modal.wait_window()
    except Exception:
        pass
    return str(selected["value"] or "").strip()


# ── HELPERS: Sección 1 — búsqueda de empresa (compartida por todos los forms) ─


def _section1_build_search(self, parent, include_tipo_visita=False):
    search_w = 58
    try:
        sw = int(self.winfo_screenwidth() or 0)
        if sw and sw <= 1366:
            search_w = 48
        elif sw and sw <= 1600:
            search_w = 52
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
    ui_feedback.bind_placeholder(self.fields["nit_empresa"], "Ej: 900123456-7")

    self.search_nit_btn = ttk.Button(
        frame,
        text="Buscar por NIT",
        style="Secondary.TButton",
        command=lambda: self._search_company("nit"),
    )
    self.search_nit_btn.grid(row=current_row, column=2, padx=12)
    nit_error = _build_inline_error_label(frame)
    nit_error.grid(row=current_row + 1, column=1, columnspan=2, sticky="w", pady=(4, 0))
    ui_feedback.register_field(self, "nit_empresa", self.fields["nit_empresa"], error_label=nit_error)
    current_row += 2

    tk.Label(
        frame,
        text="Nombre de la empresa",
        font=FONT_LABEL,
        bg=COLOR_LIGHT_BG,
    ).grid(row=current_row, column=0, sticky="w", padx=(0, 8))

    self.fields["nombre_busqueda"] = ttk.Combobox(frame, width=search_w)
    self.fields["nombre_busqueda"].grid(row=current_row, column=1, sticky="w")
    ui_feedback.bind_placeholder(self.fields["nombre_busqueda"], "Escribe al menos 2 letras")
    self.fields["nombre_busqueda"].bind(
        "<KeyRelease>",
        lambda _event: _section1_update_nombre_suggestions(self, open_dropdown=True),
    )
    self.fields["nombre_busqueda"].bind(
        "<<ComboboxSelected>>",
        lambda _event: _section1_search_selected_company(self),
    )
    self.fields["nombre_busqueda"].bind(
        "<Return>",
        lambda _event: _section1_search_selected_company(self),
    )
    self.fields["nombre_busqueda"].bind(
        "<ButtonRelease-1>",
        lambda _event, widget=self.fields["nombre_busqueda"]: _restore_combobox_text_focus(widget),
        add="+",
    )
    self.fields["nombre_busqueda"].bind(
        "<Escape>",
        lambda _event: _hide_empresa_autocomplete_popup(self),
    )
    self.fields["nombre_busqueda"].bind(
        "<FocusOut>",
        lambda _event: self.after(150, lambda: _hide_empresa_autocomplete_popup(self)),
    )

    self.search_name_btn = ttk.Button(
        frame,
        text="Buscar por nombre",
        style="Secondary.TButton",
        command=lambda: self._search_company("nombre"),
    )
    self.search_name_btn.grid(row=current_row, column=2, padx=12)
    name_error = _build_inline_error_label(frame)
    name_error.grid(row=current_row + 1, column=1, columnspan=2, sticky="w", pady=(4, 0))
    ui_feedback.register_field(self, "nombre_busqueda", self.fields["nombre_busqueda"], error_label=name_error)

    self.status_label = tk.Label(
        frame,
        text="",
        font=FONT_SUBTITLE,
        fg=COLOR_ACCENT,
        bg=COLOR_LIGHT_BG,
        anchor="w",
        justify="left",
        wraplength=920,
    )
    self.status_label.grid(row=current_row + 2, column=0, columnspan=3, sticky="w", pady=(8, 0))


def _section1_build_groups(self, parent, groups, labels, modalidad_options=None, modalidad_aliases=None):
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
    fecha_error = _build_inline_error_label(top_inputs, wraplength=220)
    fecha_error.grid(row=1, column=1, sticky="w", pady=(4, 0), padx=(0, 24))
    ui_feedback.register_field(self, "fecha_visita", self.fields["fecha_visita"], error_label=fecha_error)

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
    if isinstance(modalidad_aliases, dict) and modalidad_aliases:
        self.fields["modalidad"]._snapshot_value_aliases = dict(modalidad_aliases)
    modalidad_error = _build_inline_error_label(top_inputs, wraplength=220)
    modalidad_error.grid(row=1, column=3, sticky="w", pady=(4, 0))
    ui_feedback.register_field(self, "modalidad", self.fields["modalidad"], error_label=modalidad_error)

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
            ui_feedback.register_field(self, field_id, entry)


def _section1_build_actions(self, parent):
    actions = tk.Frame(parent, bg=COLOR_LIGHT_BG)
    _pack_actions(actions)
    self.continue_btn = ttk.Button(
        actions,
        text="Continuar",
        style="Primary.TButton",
        command=self._confirm_and_continue,
        state="disabled",
    )
    self.continue_btn.pack(side="right")
    _refresh_section1_continue_button(self)


def _refresh_section1_continue_button(window):
    button = getattr(window, "continue_btn", None)
    if button is None:
        return
    try:
        if not button.winfo_exists():
            return
    except Exception:
        return
    try:
        button.config(state="normal" if getattr(window, "company_data", None) else "disabled")
    except Exception:
        pass


def _restore_section1_cached_state(window, module, *, include_company_name=True):
    if module is None or not hasattr(module, "get_form_cache"):
        return False
    try:
        cache = module.get_form_cache().get("section_1", {})
    except Exception:
        cache = {}
    if not isinstance(cache, dict) or not cache:
        return False

    fields = getattr(window, "fields", {}) or {}
    window.company_data = copy.deepcopy(cache)

    def _set_widget_value(widget, value):
        if widget is None:
            return
        if isinstance(widget, DateEntry):
            text = str(value or "").strip()
            try:
                widget.delete(0, tk.END)
            except Exception:
                pass
            if not text:
                return
            try:
                widget.set_date(text)
            except Exception:
                try:
                    widget.insert(0, text)
                except Exception:
                    pass
            return
        try:
            state = str(widget.cget("state") or "")
        except Exception:
            state = ""
        if state == "readonly" and isinstance(widget, tk.Entry):
            _set_readonly_entry_value(widget, value)
            return
        _set_input_value(widget, value)

    for field_id, value in cache.items():
        widget = fields.get(field_id)
        if widget is None:
            continue
        _set_widget_value(widget, value)

    if include_company_name:
        search_widget = fields.get("nombre_busqueda")
        if search_widget is not None:
            _set_widget_value(search_widget, cache.get("nombre_empresa", ""))

    status_label = getattr(window, "status_label", None)
    if status_label is not None:
        try:
            ui_feedback.set_semantic_label(status_label, "Empresa cargada desde el borrador.", state="success")
        except Exception:
            pass

    _refresh_section1_continue_button(window)
    return True


def _bind_tooltip(widget, text_provider, *, delay_ms=500):
    """Muestra un tooltip al hacer hover sobre widget.
    text_provider puede ser str o callable() -> str (se evalúa en el momento del show).
    """
    _state = {"after_id": None, "window": None}

    def _show():
        _state["after_id"] = None
        try:
            text = text_provider() if callable(text_provider) else str(text_provider)
            if not text:
                return
            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.wm_attributes("-topmost", True)
            x = widget.winfo_rootx() + widget.winfo_width() // 2
            y = widget.winfo_rooty() + widget.winfo_height() + 4
            win.wm_geometry(f"+{x}+{y}")
            tk.Label(
                win,
                text=text,
                bg="#2D2D2D",
                fg="white",
                font=("Arial", 8),
                padx=8,
                pady=4,
                wraplength=260,
            ).pack()
            _state["window"] = win
        except Exception:
            pass

    def _schedule(_event=None):
        _cancel()
        try:
            _state["after_id"] = widget.after(delay_ms, _show)
        except Exception:
            pass

    def _cancel(_event=None):
        if _state["after_id"] is not None:
            try:
                widget.after_cancel(_state["after_id"])
            except Exception:
                pass
            _state["after_id"] = None
        if _state["window"] is not None:
            try:
                _state["window"].destroy()
            except Exception:
                pass
            _state["window"] = None

    widget.bind("<Enter>", _schedule, add="+")
    widget.bind("<Leave>", _cancel, add="+")
    widget.bind("<Button>", _cancel, add="+")


# ── HELPERS: Scroll, dictado por voz, catálogos de asistentes ───────────────


def _create_vscroll(parent, command):
    return ttk.Scrollbar(parent, orient="vertical", command=command, style="Vertical.TScrollbar")


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
    modalidad = ui_feedback.get_widget_value(widget) if widget else ""
    if modalidad:
        ui_feedback.clear_field_error(window, "modalidad")
        return modalidad
    ui_feedback.set_field_error(window, "modalidad", "Selecciona la modalidad de la visita.")
    return None


def _get_required_fecha_visita(window):
    fields = getattr(window, "fields", {}) or {}
    widget = fields.get("fecha_visita")
    fecha_visita = widget.get().strip() if widget else ""
    if fecha_visita:
        ui_feedback.clear_field_error(window, "fecha_visita")
        return fecha_visita
    ui_feedback.set_field_error(window, "fecha_visita", "Selecciona la fecha de la visita.")
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
    disabled_sections = set(getattr(window, "_disable_section_dictation_sections", set()) or set())
    if str(section_name or "").strip() in disabled_sections:
        _clear_section_dictation_button(window)
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
    except tk.TclError:
        try:
            if getattr(window, "_section_dictation_button", None) is button:
                window._section_dictation_button = None
        except Exception:
            pass
        return
    except Exception:
        pass

    try:
        if getattr(helper, "_is_processing", False):
            button.configure(text="🎤 Procesando...", state="disabled")
        elif getattr(helper, "_is_recording", False):
            button.configure(text="🎤 Detener", state="normal")
        else:
            button.configure(text="🎤 Dictar", state="normal" if helper._can_dictate() else "disabled")
    except tk.TclError:
        try:
            if getattr(window, "_section_dictation_button", None) is button:
                window._section_dictation_button = None
        except Exception:
            pass
        return

    try:
        window._section_dictation_after_id = window.after(
            250,
            lambda w=window: _refresh_section_dictation_button(w),
        )
    except Exception:
        window._section_dictation_after_id = None


def _dictation_button_tooltip(window):
    helper = _resolve_section_dictation_helper(window)
    if helper is None:
        return ""
    if getattr(helper, "_is_processing", False):
        return "Procesando audio..."
    if getattr(helper, "_is_recording", False):
        return "Haz click para detener la grabación"
    if helper._can_dictate():
        return "Haz click para dictar en este campo de texto"
    return "Dictar no disponible: audio inactivo o campo bloqueado"


def _install_section_dictation_button(window, text_widgets):
    title = getattr(window, "header_title", None)
    if not title or not title.winfo_exists():
        return
    parent = title.master
    button = ttk.Button(
        parent,
        text="Dictar",
        style="Secondary.TButton",
        command=lambda w=window: _on_section_dictation_click(w),
    )
    _bind_tooltip(button, lambda w=window: _dictation_button_tooltip(w))
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


# ── HELPERS: Catálogos de asistentes y asesores ──────────────────────────────


_ASISTENTES_PROF_CACHE = {
    "loaded_at": 0.0,
    "nombres": [],
    "cargos": [],
    "name_to_cargo": {},
}
_INTERPRETES_CACHE = {
    "loaded_at": 0.0,
    "nombres": [],
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


def _get_interpretes_catalog(force=False):
    global _INTERPRETES_CACHE
    now = time.time()
    cached = _INTERPRETES_CACHE
    if (
        not force
        and cached.get("nombres")
        and (now - float(cached.get("loaded_at") or 0.0)) < _ASISTENTES_PROF_CACHE_TTL
    ):
        return cached

    rows = None
    for select_clause in ("nombre", "nombre_interprete"):
        try:
            rows = _supabase_get_paged(
                "interpretes",
                {
                    "select": select_clause,
                    "order": f"{select_clause}.asc",
                },
                env_path=".env",
                page_size=500,
                max_pages=20,
            )
            if rows is not None:
                break
        except Exception:
            rows = None
            continue

    if rows is None:
        _log_capture("[INTERPRETES] no se pudo leer catalogo de interpretes.")
        return cached

    nombres = []
    for row in rows or []:
        nombre = str((row or {}).get("nombre") or (row or {}).get("nombre_interprete") or "").strip()
        if nombre:
            nombres.append(_normalize_person_name(nombre))

    _INTERPRETES_CACHE = {
        "loaded_at": now,
        "nombres": _dedupe_keep_order(nombres),
    }
    return _INTERPRETES_CACHE


def _normalize_person_widget(widget):
    if widget is None:
        return ""
    normalized = _normalize_person_name(_get_input_value(widget))
    if normalized != _get_input_value(widget):
        _set_input_value(widget, normalized)
    return normalized


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
        _normalize_person_widget(nombre_widget)
        selected_name = _asistentes_norm(nombre_widget.get())
        suggested = name_to_cargo.get(selected_name)
        if suggested:
            cargo_widget.set(suggested)

    nombre_widget.bind("<<ComboboxSelected>>", _sync_cargo, add="+")
    nombre_widget.bind("<FocusOut>", _sync_cargo, add="+")


def _create_interprete_name_input(parent, width, catalog=None):
    nombre_widget = ttk.Combobox(parent, width=width, state="normal")
    nombres = list((catalog or _get_interpretes_catalog()).get("nombres") or [])
    nombre_widget.configure(values=nombres)
    _bind_editable_combobox_filter(nombre_widget, nombres)
    nombre_widget.bind("<<ComboboxSelected>>", lambda _e=None: _normalize_person_widget(nombre_widget), add="+")
    nombre_widget.bind("<FocusOut>", lambda _e=None: _normalize_person_widget(nombre_widget), add="+")
    return nombre_widget


# ── CLASES BASE: FormMousewheelMixin, LoadingDialog, Labs dialogs ─────────────


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


class LabsSection2VoiceDialog:
    def __init__(
        self,
        parent,
        *,
        subsection_key,
        section_label,
        candidate_index,
        form_id,
        session_provider,
        spec=None,
        function_name=None,
        section_id="section_2",
        record_label="Oferente",
        ui_namespace="Labs",
    ):
        self.parent = parent
        self.subsection_key = str(subsection_key or "").strip()
        self.section_label = str(section_label or self.subsection_key).strip() or self.subsection_key
        self.candidate_index = int(candidate_index)
        self.form_id = str(form_id or "").strip()
        self.session_provider = session_provider
        self.function_name = str(function_name or SELECTION_SECTION2_VOICE_FUNCTION).strip() or SELECTION_SECTION2_VOICE_FUNCTION
        self.section_id = str(section_id or "section_2").strip() or "section_2"
        self.record_label = str(record_label or "Oferente").strip() or "Oferente"
        self.ui_namespace = str(ui_namespace or "Labs").strip() or "Labs"
        self.result = None
        self._handle = None
        self._is_recording = False
        self._is_processing = False
        self._closed = False
        self._worker = None
        self.spec = dict(spec or get_selection_labs_subsection_spec(self.subsection_key))
        _log_labs(
            f"voice_dialog_open subsection={self.subsection_key} candidate_index={self.candidate_index}"
        )

        self.window = tk.Toplevel(parent)
        self.window.title(f"{self.ui_namespace} - {self.section_label}")
        self.window.configure(bg=COLOR_LIGHT_BG)
        self.window.geometry("760x560")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.protocol("WM_DELETE_WINDOW", self._on_cancel)

        container = tk.Frame(self.window, bg=COLOR_LIGHT_BG)
        container.pack(fill="both", expand=True, padx=18, pady=18)

        tk.Label(
            container,
            text=f"{self.section_label} - {self.record_label} {self.candidate_index}",
            font=FONT_SECTION,
            bg=COLOR_LIGHT_BG,
            fg=COLOR_PURPLE,
            anchor="w",
        ).pack(fill="x", pady=(0, 8))

        warning_box = tk.Frame(container, bg="#FFF4E5", bd=1, relief="solid")
        warning_box.pack(fill="x", pady=(0, 12))
        tk.Label(
            warning_box,
            text=(
                "Funcion experimental. Usa un solo audio por oferente y por subseccion. "
                "Si falta un dato, es mejor no decirlo que inventarlo."
            ),
            bg="#FFF4E5",
            fg="#7A4100",
            justify="left",
            wraplength=700,
            padx=12,
            pady=10,
        ).pack(fill="x")

        tk.Label(
            container,
            text="Instruccion general",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
            anchor="w",
        ).pack(fill="x")

        tk.Label(
            container,
            text=str(self.spec.get("script") or "").strip(),
            bg=COLOR_LIGHT_BG,
            fg="#333333",
            justify="left",
            wraplength=700,
            anchor="w",
        ).pack(fill="x", pady=(4, 10))

        tk.Label(
            container,
            text="Debes responder lo siguiente",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
            anchor="w",
        ).pack(fill="x")

        questions_text = tk.Text(container, height=8, wrap="word", bg="white")
        questions_text.pack(fill="x", pady=(4, 12))
        questions = self.spec.get("questions") or []
        questions_text.insert(
            "1.0",
            "\n".join(f"- {str(question).strip()}" for question in questions if str(question).strip()),
        )
        questions_text.configure(state="disabled")

        tk.Label(
            container,
            text="Ejemplo breve",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
            anchor="w",
        ).pack(fill="x")

        examples_text = tk.Text(container, height=5, wrap="word", bg="white")
        examples_text.pack(fill="both", expand=True, pady=(4, 12))
        examples = self.spec.get("examples") or []
        first_example = ""
        for example in examples:
            text = str(example).strip()
            if text:
                first_example = f"- {text}"
                break
        examples_text.insert("1.0", first_example)
        examples_text.configure(state="disabled")

        actions = tk.Frame(container, bg=COLOR_LIGHT_BG)
        actions.pack(fill="x")

        self.status_label = tk.Label(
            actions,
            text="Listo para grabar.",
            bg=COLOR_LIGHT_BG,
            fg="#444444",
            anchor="w",
        )
        self.status_label.pack(side="left")

        self.record_button = ttk.Button(actions, text="Iniciar grabacion", command=self._on_toggle)
        self.record_button.pack(side="right")
        self.cancel_button = ttk.Button(actions, text="Cancelar", command=self._on_cancel)
        self.cancel_button.pack(side="right", padx=(0, 8))

        self._center()

    def _center(self):
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def _set_status(self, text, color="#444444"):
        try:
            self.status_label.config(text=text, fg=color)
            self.window.update_idletasks()
        except tk.TclError:
            return

    def _on_toggle(self):
        if self._is_processing:
            return
        if self._is_recording:
            self._stop_and_submit_async()
        else:
            self._start_recording()

    def _start_recording(self):
        try:
            handle = start_recording(
                field_id=f"{self.subsection_key}:candidate_{self.candidate_index}",
                form_id=self.form_id,
            )
        except Exception as exc:
            self._set_status(f"No se pudo iniciar grabacion: {exc}", "#A40000")
            return
        if not claim_recording(handle):
            cancel_recording(handle)
            self._set_status("Ya hay otro dictado activo.", "#A40000")
            _log_labs(
                f"voice_record_start_blocked subsection={self.subsection_key} candidate_index={self.candidate_index}",
                level="WARN",
            )
            return
        self._handle = handle
        self._is_recording = True
        self.record_button.config(text="Detener y procesar")
        self._set_status("Grabando...", "#B35300")
        _log_labs(
            f"voice_record_started subsection={self.subsection_key} candidate_index={self.candidate_index} file={handle.file_path}"
        )

    def _stop_and_submit_async(self):
        if not self._handle:
            return
        self._is_recording = False
        self._is_processing = True
        self.record_button.config(state="disabled")
        self.cancel_button.config(state="disabled")
        self._set_status("Procesando audio...", "#1F4E79")
        _log_labs(
            f"voice_submit_start subsection={self.subsection_key} candidate_index={self.candidate_index}"
        )

        def _worker():
            try:
                jwt = str(self.session_provider() or "").strip()
            except Exception:
                jwt = ""
            result = stop_and_submit_audio(
                self._handle,
                jwt,
                function_name=self.function_name,
                language="es",
                extra_fields={
                    "form_id": self.form_id,
                    "section_id": self.section_id,
                    "subsection_key": self.subsection_key,
                    "candidate_index": str(self.candidate_index),
                },
            )
            try:
                self.window.after(0, lambda: self._finish_request(result))
            except tk.TclError:
                return

        self._worker = threading.Thread(target=_worker, daemon=True)
        self._worker.start()

    def _finish_request(self, result):
        self._is_processing = False
        self.record_button.config(state="normal", text="Iniciar grabacion")
        self.cancel_button.config(state="normal")
        self._handle = None
        if self._closed:
            return
        if not result.ok:
            self._set_status(result.error_message or "No fue posible procesar el audio.", "#A40000")
            _log_labs(
                f"voice_submit_failed subsection={self.subsection_key} candidate_index={self.candidate_index} "
                f"elapsed_ms={result.elapsed_ms} code={result.error_code} message={result.error_message}",
                level="ERROR",
            )
            messagebox.showerror(
                self.ui_namespace,
                result.error_message or "No fue posible procesar el audio.",
                parent=self.window,
            )
            return
        payload = result.payload or {}
        usage = payload.get("usage") or {}
        warnings = payload.get("warnings") or []
        _log_labs(
            f"voice_submit_ok subsection={self.subsection_key} candidate_index={self.candidate_index} "
            f"elapsed_ms={result.elapsed_ms} "
            f"transcription_chars={len(str(payload.get('transcription') or ''))} "
            f"warnings={len(warnings) if isinstance(warnings, list) else 0} "
            f"transcribe_model={usage.get('transcribe_model')} extract_model={usage.get('extract_model')}"
        )
        self.result = result.payload or {}
        self._closed = True
        self.window.destroy()

    def _on_cancel(self):
        self._closed = True
        if self._is_recording and self._handle is not None:
            try:
                cancel_recording(self._handle)
            except Exception:
                pass
            _log_labs(
                f"voice_dialog_cancel_recording subsection={self.subsection_key} candidate_index={self.candidate_index}",
                level="WARN",
            )
        self._handle = None
        try:
            self.window.destroy()
        except tk.TclError:
            return

    def show(self):
        self.window.wait_window()
        return self.result


class LabsSection2PreviewDialog:
    def __init__(
        self,
        parent,
        *,
        section_label,
        candidate_index,
        transcription,
        preview_lines,
        warnings,
        record_label="Oferente",
        ui_namespace="Labs",
    ):
        self.parent = parent
        self.result = False
        self.window = tk.Toplevel(parent)
        self.record_label = str(record_label or "Oferente").strip() or "Oferente"
        self.ui_namespace = str(ui_namespace or "Labs").strip() or "Labs"
        self.window.title(f"Preview {self.ui_namespace} - {section_label}")
        self.window.configure(bg=COLOR_LIGHT_BG)
        self.window.geometry("760x640")
        self.window.transient(parent)
        self.window.grab_set()
        self.window.protocol("WM_DELETE_WINDOW", self._on_cancel)

        container = tk.Frame(self.window, bg=COLOR_LIGHT_BG)
        container.pack(fill="both", expand=True, padx=18, pady=18)

        tk.Label(
            container,
            text=f"Preview de autollenado - {self.record_label} {candidate_index}",
            font=FONT_SECTION,
            bg=COLOR_LIGHT_BG,
            fg=COLOR_PURPLE,
            anchor="w",
        ).pack(fill="x", pady=(0, 10))

        tk.Label(container, text="Transcripcion", font=FONT_LABEL, bg=COLOR_LIGHT_BG, anchor="w").pack(fill="x")
        transcription_box = tk.Text(container, height=8, wrap="word", bg="white")
        transcription_box.pack(fill="x", pady=(4, 12))
        transcription_box.insert("1.0", transcription or "")
        transcription_box.configure(state="disabled")

        tk.Label(container, text="Campos detectados", font=FONT_LABEL, bg=COLOR_LIGHT_BG, anchor="w").pack(fill="x")
        extracted_box = tk.Text(container, height=12, wrap="word", bg="white")
        extracted_box.pack(fill="both", expand=True, pady=(4, 12))
        extracted_box.insert("1.0", "\n".join(preview_lines) if preview_lines else "No se detectaron campos aplicables.")
        extracted_box.configure(state="disabled")

        tk.Label(container, text="Advertencias", font=FONT_LABEL, bg=COLOR_LIGHT_BG, anchor="w").pack(fill="x")
        warnings_box = tk.Text(container, height=5, wrap="word", bg="white")
        warnings_box.pack(fill="x", pady=(4, 12))
        warnings_box.insert("1.0", "\n".join(f"- {item}" for item in warnings) if warnings else "Sin advertencias.")
        warnings_box.configure(state="disabled")

        actions = tk.Frame(container, bg=COLOR_LIGHT_BG)
        actions.pack(fill="x")
        ttk.Button(actions, text="Cancelar", command=self._on_cancel).pack(side="right")
        ttk.Button(actions, text="Aplicar", command=self._on_apply).pack(side="right", padx=(0, 8))
        self._center()

    def _center(self):
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def _on_apply(self):
        self.result = True
        self.window.destroy()

    def _on_cancel(self):
        self.result = False
        try:
            self.window.destroy()
        except tk.TclError:
            return

    def show(self):
        self.window.wait_window()
        return self.result


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


# ── HELPERS: Finalización de formularios (flujo Sheets + Drive + PDF) ────────


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


def _read_followup_case_state(case_target):
    try:
        suggestion = seguimientos.suggest_next_step(case_target)
        summary = seguimientos.describe_case(case_target)
        return {"suggestion": suggestion, "summary": summary, "error": None}
    except Exception as exc:
        _log_capture(f"followup_case_state_read_failed target={case_target!r} err={exc}")
        return {"suggestion": None, "summary": None, "error": exc}


def _format_followup_case_state_error(exc):
    if exc is None:
        return ""
    detail = str(exc).strip() or repr(exc)
    detail_lower = detail.lower()
    if _is_transient_drive_exception(exc) or "supabase no esta disponible" in detail_lower:
        return (
            "Archivo encontrado, pero no se pudo leer el estado actual del caso "
            "por una falla temporal de conexión."
        )
    return f"Archivo encontrado, pero falló lectura de estado: {detail}"


def _build_followup_suggestion_from_workflow(workflow):
    workflow = workflow or {}
    try:
        max_seguimientos = int(workflow.get("max_seguimientos") or 3)
    except Exception:
        max_seguimientos = 3
    return {
        "sheet": workflow.get("suggested_sheet") or workflow.get("base_sheet_name") or seguimientos.SHEET_BASE,
        "message": str(workflow.get("message") or "").strip(),
        "max_seguimientos": max_seguimientos,
    }


FOLLOWUP_STAGE_STATUS_LABELS = {
    "pending": "Pendiente",
    "not_started": "Pendiente",
    "in_progress": "En curso",
    "completed": "Completa",
    "review_only": "Solo lectura",
}


def _friendly_followup_sheet_title(sheet_name, workflow=None, *, base_sheet_name=None):
    current = str(sheet_name or "").strip()
    if not current:
        return ""
    workflow = workflow or {}
    stage_entries = list(workflow.get("stage_model") or workflow.get("sheet_progress") or [])
    for entry in stage_entries:
        if str((entry or {}).get("sheet_name") or "").strip() == current:
            return str((entry or {}).get("title") or (entry or {}).get("label") or current).strip()
    base_name = str(base_sheet_name or workflow.get("base_sheet_name") or seguimientos.SHEET_BASE).strip()
    if current == base_name:
        return "Ficha inicial del proceso"
    if current == seguimientos.SHEET_FINAL:
        return "Resultado final"
    match = re.search(r"(\d+)$", current)
    if match:
        return f"Seguimiento {int(match.group(1))}"
    return current


def _format_followup_stage_status(status, *, coverage_percent=None, is_suggested=False):
    normalized = str(status or "").strip() or "pending"
    label = FOLLOWUP_STAGE_STATUS_LABELS.get(normalized, normalized.replace("_", " ").title())
    if normalized in {"review_only", "pending", "not_started"}:
        text = label
    else:
        try:
            text = f"{label} · {int(coverage_percent or 0)}%"
        except Exception:
            text = label
    if is_suggested and normalized != "review_only":
        return f"{text} · Etapa sugerida"
    return text


def _get_followup_stage_palette(status, *, is_suggested=False):
    normalized = str(status or "").strip() or "pending"
    if normalized == "completed":
        return {"accent": COLOR_SUCCESS, "bg": "#EAF7EE", "muted": "#1E5E2E"}
    if normalized == "in_progress":
        return {"accent": "#2E6FD8", "bg": "#EAF1FD", "muted": "#234A87"}
    if normalized == "review_only":
        return {"accent": "#5F6B7A", "bg": "#F1F3F5", "muted": "#4A5563"}
    if is_suggested:
        return {"accent": COLOR_PURPLE, "bg": "#F4ECFF", "muted": "#5C2E8A"}
    return {"accent": "#8A8F98", "bg": "#F7F8FA", "muted": "#58616D"}


def _build_followup_window_flow_model(
    *,
    user_row=None,
    linked_company=None,
    compensar_choice="",
    case_record=None,
    workflow=None,
    summary=None,
    suggestion=None,
):
    workflow = dict(workflow or {})
    summary = dict(summary or {})
    suggestion = dict(suggestion or {})
    linked_company = dict(linked_company or {})
    user_row = dict(user_row or {})
    case_exists = bool(case_record)
    company_known = bool(
        str(linked_company.get("nombre_empresa") or "").strip()
        or str(linked_company.get("nit_empresa") or "").strip()
    )
    company_confirmed = company_known and (
        str(compensar_choice or "").strip().startswith("Si")
        or str(compensar_choice or "").strip().startswith("No")
        or bool(str(linked_company.get("caja_compensacion") or "").strip())
    )
    suggested_sheet = str(
        suggestion.get("sheet")
        or workflow.get("suggested_sheet")
        or workflow.get("base_sheet_name")
        or seguimientos.SHEET_BASE
    ).strip()
    suggested_title = _friendly_followup_sheet_title(suggested_sheet, workflow)
    total_followups = int(
        suggestion.get("max_seguimientos")
        or workflow.get("max_seguimientos")
        or 3
    )
    stage_entries = {
        str((entry or {}).get("sheet_name") or "").strip(): dict(entry or {})
        for entry in list(workflow.get("stage_model") or workflow.get("sheet_progress") or [])
    }
    base_sheet_name = str(workflow.get("base_sheet_name") or seguimientos.SHEET_BASE).strip()
    base_stage = dict(
        stage_entries.get(
            base_sheet_name,
            {
                "sheet_name": base_sheet_name,
                "title": "Ficha inicial del proceso",
                "status": "pending" if not case_exists else "not_started",
                "coverage_percent": 0,
                "is_suggested": suggested_sheet == base_sheet_name,
                "is_editable": bool(case_exists),
                "helper_text": "Se crea al confirmar la empresa del caso.",
            },
        )
    )
    if not case_exists:
        base_stage["status"] = "pending"
        base_stage["helper_text"] = "Se habilita cuando el caso tenga Google Sheet."
        base_stage["is_editable"] = False

    current_followup_title = ""
    current_followup_status = "pending"
    current_followup_helper = "Se habilita después de completar la ficha inicial del proceso."
    current_followup_sheet = ""
    current_followup_coverage = 0
    current_followup_editable = False
    current_followup_suggested = False
    if suggested_sheet and suggested_sheet not in {base_sheet_name, seguimientos.SHEET_FINAL}:
        suggested_entry = dict(stage_entries.get(suggested_sheet) or {})
        current_followup_title = str(
            suggested_entry.get("title")
            or _friendly_followup_sheet_title(suggested_sheet, workflow)
            or "Seguimiento actual"
        ).strip()
        current_followup_status = str(suggested_entry.get("status") or "not_started").strip()
        current_followup_helper = str(
            suggested_entry.get("helper_text")
            or suggestion.get("message")
            or ""
        ).strip()
        current_followup_sheet = suggested_sheet
        current_followup_coverage = int(suggested_entry.get("coverage_percent") or 0)
        current_followup_editable = bool(suggested_entry.get("is_editable"))
        current_followup_suggested = True
    completed_titles = list(summary.get("completed_sheets") or workflow.get("completed_sheets") or [])
    history_helper = ", ".join(completed_titles) if completed_titles else "Aún no hay seguimientos completos."
    final_entry = dict(
        stage_entries.get(
            seguimientos.SHEET_FINAL,
            {
                "sheet_name": seguimientos.SHEET_FINAL,
                "title": "Resultado final",
                "status": "pending" if not case_exists else "review_only",
                "coverage_percent": 0,
                "is_suggested": suggested_sheet == seguimientos.SHEET_FINAL,
                "is_editable": False,
                "helper_text": "Consolidado automático del caso.",
            },
        )
    )
    if not case_exists:
        final_entry["status"] = "pending"
        final_entry["helper_text"] = "Se actualiza automáticamente al completar la ficha y los seguimientos."

    return [
        {
            "stage_id": "identify_user",
            "title": "Identificar vinculado",
            "sheet_name": "",
            "status": "completed" if user_row else "pending",
            "coverage_percent": 100 if user_row else 0,
            "is_editable": True,
            "is_suggested": not bool(user_row),
            "helper_text": (
                str(user_row.get("nombre_usuario") or "").strip()
                if user_row
                else "Busca la cédula para cargar el caso."
            ),
        },
        {
            "stage_id": "confirm_company",
            "title": "Confirmar empresa",
            "sheet_name": "",
            "status": (
                "completed"
                if company_confirmed
                else ("in_progress" if user_row else "pending")
            ),
            "coverage_percent": 100 if company_confirmed else (40 if user_row and company_known else 0),
            "is_editable": True,
            "is_suggested": bool(user_row) and not company_confirmed,
            "helper_text": (
                "Empresa lista para crear o continuar el caso."
                if company_confirmed
                else (
                    "Confirma la empresa y si es Compensar para preparar el caso."
                    if user_row
                    else "Primero identifica el vinculado."
                )
            ),
        },
        {
            "stage_id": "base_process",
            "title": str(base_stage.get("title") or "Ficha inicial del proceso"),
            "sheet_name": str(base_stage.get("sheet_name") or ""),
            "status": str(base_stage.get("status") or "pending"),
            "coverage_percent": int(base_stage.get("coverage_percent") or 0),
            "is_editable": bool(base_stage.get("is_editable")),
            "is_suggested": bool(base_stage.get("is_suggested")),
            "helper_text": str(base_stage.get("helper_text") or ""),
        },
        {
            "stage_id": "current_followup",
            "title": "Seguimiento actual",
            "sheet_name": current_followup_sheet,
            "status": current_followup_status,
            "coverage_percent": current_followup_coverage,
            "is_editable": current_followup_editable,
            "is_suggested": current_followup_suggested,
            "helper_text": current_followup_title
            + (": " if current_followup_title and current_followup_helper else "")
            + current_followup_helper,
        },
        {
            "stage_id": "followup_history",
            "title": "Historial de seguimientos",
            "sheet_name": "",
            "status": "completed" if completed_titles else ("in_progress" if case_exists else "pending"),
            "coverage_percent": 100 if completed_titles else 0,
            "is_editable": False,
            "is_suggested": False,
            "helper_text": history_helper,
        },
        {
            "stage_id": "final_result",
            "title": str(final_entry.get("title") or "Resultado final"),
            "sheet_name": str(final_entry.get("sheet_name") or ""),
            "status": str(final_entry.get("status") or "pending"),
            "coverage_percent": int(final_entry.get("coverage_percent") or 0),
            "is_editable": bool(final_entry.get("is_editable")),
            "is_suggested": bool(final_entry.get("is_suggested")),
            "helper_text": str(final_entry.get("helper_text") or ""),
        },
    ]


def _format_followup_editor_open_error(exc):
    if exc is None:
        return ""
    detail = str(exc).strip() or repr(exc)
    detail_lower = detail.lower()
    if _is_transient_drive_exception(exc) or "supabase no esta disponible" in detail_lower:
        return "No se pudo abrir el editor del caso por una falla temporal de conexión."
    return f"No se pudo abrir el editor del caso: {detail}"


def _load_followup_editor_bootstrap(case_target):
    try:
        meta = seguimientos.get_case_meta(case_target)
        workflow = seguimientos.get_workflow_state(case_target)
    except Exception as exc:
        _log_capture(f"followup_editor_bootstrap_failed target={case_target!r} err={exc}")
        raise RuntimeError(_format_followup_editor_open_error(exc)) from exc
    return {
        "meta": dict(meta or {}),
        "workflow": dict(workflow or {}),
        "suggestion": _build_followup_suggestion_from_workflow(workflow),
    }


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
    on_success=None,
    on_error=None,
    close_window_on_success=True,
    return_to_hub_on_success=True,
    show_completion_ui=True,
    show_error_dialog=True,
):
    if getattr(window, "_finalize_in_progress", False):
        messagebox.showinfo(
            "Finalización",
            "Ya hay una finalización en curso para este formulario.",
            parent=window,
        )
        return False

    window._finalize_in_progress = True

    def _finish_success(completion_result):
        try:
            if show_completion_ui:
                _finalize_export_flow(
                    window,
                    loading,
                    completion_result,
                )
            else:
                _close_loading_async(loading)
            if callable(on_success):
                try:
                    on_success(completion_result)
                except Exception as callback_exc:
                    _log_capture(f"finalization_on_success_callback_failed form={form_id} err={callback_exc}")
            if return_to_hub_on_success:
                _return_to_hub(window)
            if close_window_on_success:
                try:
                    window._skip_close_guard = True
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
            _close_loading_async(loading)
            if show_error_dialog:
                messagebox.showerror("Finalización", message, parent=window)
            if callable(on_error):
                try:
                    on_error(exc, message)
                except Exception as callback_exc:
                    _log_capture(f"finalization_on_error_callback_failed form={form_id} err={callback_exc}")
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
        export_cache_snapshot = {}
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
                        status="Ortografía revisada. Preparando acta...",
                        progress=50,
                    )
                elif review_result.status in {"skipped", "failed"}:
                    _update_loading_async(
                        loading,
                        status="Revisión ortográfica omitida. Preparando acta...",
                        progress=50,
                    )

            export_result = worker_fn()
            drive_job = None
            already_in_drive = False
            if isinstance(export_result, dict):
                output_path = str(export_result.get("output_path") or "").strip()
                drive_job = export_result.get("drive_job")
                already_in_drive = bool(export_result.get("already_in_drive"))
            else:
                output_path = export_result
            if not output_path and not already_in_drive:
                raise FinalizeProcessError(
                    "generando el acta",
                    RuntimeError("No se generó el acta."),
                )
            if not already_in_drive and not os.path.exists(output_path):
                raise FinalizeProcessError(
                    "verificando el acta generada",
                    RuntimeError(f"No se encontró el acta generada:\n{output_path}"),
                )
            if module and hasattr(module, "get_form_cache"):
                try:
                    export_cache_snapshot = copy.deepcopy(module.get_form_cache() or {})
                except Exception:
                    export_cache_snapshot = {}
            if not export_cache_snapshot:
                if review_result is not None and getattr(review_result, "status", "") == "reviewed":
                    export_cache_snapshot = copy.deepcopy(getattr(review_result, "cache", {}) or {})
                else:
                    export_cache_snapshot = copy.deepcopy(original_cache_snapshot or {})
            hub = _resolve_hub_window(window)
            if hub and form_id:
                hub.track_form_finished(form_id)
            if hub:
                completion_payload = None
                if form_id and export_cache_snapshot:
                    try:
                        completion_payload = completion_payloads.build_completion_payload(
                            form_id,
                            form_name,
                            export_cache_snapshot,
                            output_path=output_path,
                            session_id=hub.current_session_id,
                            app_version=get_version(),
                            extra_context={"payload_source": "form_cache"},
                        )
                    except Exception as exc:
                        _log_capture(
                            f"build_completion_payload failed form={form_id} output={output_path} err={exc}"
                        )
                if already_in_drive:
                    _update_loading_async(
                        loading,
                        status="Acta publicada en Google Sheets.",
                        progress=95,
                    )
                    drive_file_id = ""
                    if isinstance(export_result, dict):
                        drive_file_id = str(
                            export_result.get("drive_file_id")
                            or export_result.get("file_id")
                            or ""
                        ).strip()
                    completion_result = hub.finalize_form_delivery(
                        output_path,
                        form_name=form_name,
                        company_name=company_name,
                        drive_job={
                            "status": "synced",
                            "drive_file_id": drive_file_id,
                            "remote_url": output_path,
                        },
                        completion_payload=completion_payload,
                    )
                    # Encolar exportación a PDF si el formulario lo soporta
                    if (
                        drive_file_id
                        and isinstance(export_result, dict)
                        and str(export_result.get("tipo_acta") or "") in _PDF_EXPORT_ENABLED_TIPOS
                    ):
                        try:
                            _pdf_folder_id = drive_upload._get_pdf_folder_id()
                            _pdf_registro_id = str(
                                (completion_result or {}).get("registro_id") or ""
                            ).strip()
                            _enqueue_pdf_export_job(
                                sheet_file_id=drive_file_id,
                                tipo_acta=str(export_result.get("tipo_acta") or ""),
                                fecha_servicio=str(export_result.get("fecha_servicio") or ""),
                                acta_metadata=export_result.get("acta_metadata") or {},
                                extra_name=export_result.get("extra_name"),
                                pdf_folder_id=_pdf_folder_id,
                                company_name=company_name,
                                registro_id=_pdf_registro_id,
                            )
                            # Pasar info al dialog de éxito para que ofrezca abrir carpeta de PDFs
                            if isinstance(completion_result, dict):
                                completion_result["pdf_folder_id"] = _pdf_folder_id
                                completion_result["company_name"] = company_name
                            _log_capture(
                                f"[PDF_EXPORT] job enqueued tipo={export_result.get('tipo_acta')!r} "
                                f"sheet_id={drive_file_id}"
                            )
                        except Exception as _pdf_err:
                            _log_capture(f"[PDF_EXPORT] failed to enqueue job: {_pdf_err}")
                else:
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
                        completion_payload=completion_payload,
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
    return True


_original_start_background_finalization = _start_background_finalization


def _start_background_finalization(
    window,
    loading,
    *,
    form_name,
    company_name,
    form_id,
    worker_fn,
    post_delivery_fn=None,
    on_success=None,
    on_error=None,
    close_window_on_success=True,
    return_to_hub_on_success=True,
    show_completion_ui=True,
    show_error_dialog=True,
):
    if _guard_form_action(window, action_label="finalizar"):
        try:
            if loading is not None:
                loading.close()
        except Exception:
            pass
        return False
    if _guard_form_finalization(window, loading=loading):
        return False
    return _original_start_background_finalization(
        window,
        loading,
        form_name=form_name,
        company_name=company_name,
        form_id=form_id,
        worker_fn=worker_fn,
        post_delivery_fn=post_delivery_fn,
        on_success=on_success,
        on_error=on_error,
        close_window_on_success=close_window_on_success,
        return_to_hub_on_success=return_to_hub_on_success,
        show_completion_ui=show_completion_ui,
        show_error_dialog=show_error_dialog,
    )


# ── REGISTRO DE FORMULARIOS ──────────────────────────────────────────────────


def get_forms():
    return [
        presentacion_programa.register_form(),
        evaluacion_accesibilidad.register_form(),
        condiciones_vacante.register_form(),
        {"id": "condiciones_vacante_labs", "name": "Condiciones de Vacante Labs", "hidden": True},
        seleccion_incluyente.register_form(),
        contratacion_incluyente.register_form(),
        induccion_organizacional.register_form(),
        induccion_operativa.register_form(),
        sensibilizacion.register_form(),
        seguimientos.register_form(),
        interprete_lsc.register_form(),
    ]


# ── VENTANA: Section1Window — Sección 1 compartida (búsqueda de empresa) ─────


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
        if presentacion_programa.cache_file_exists():
            _clear_local_resume_state(presentacion_programa)
        return False

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
        _fn = getattr(self, '_pending_autosave', None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
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
        _restore_section1_cached_state(self, presentacion_programa)
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        _attach_autoexpand(self.section4_text, 10, 30)

        cached_notes = presentacion_programa.get_form_cache().get("section_4", {}).get(
            "acuerdos_observaciones"
        )
        if cached_notes:
            self.section4_text.delete("1.0", tk.END)
            self.section4_text.insert("1.0", cached_notes)

        self._pending_autosave = lambda: _autosave_section(presentacion_programa, "section_4", lambda: {"acuerdos_observaciones": self.section4_text.get("1.0", tk.END).strip()})
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
        self._pending_autosave = lambda: _autosave_section(
            presentacion_programa,
            "section_5",
            lambda: self._get_section5_asistentes_values(),
        )

        _build_wizard_actions(
            section_frame,
            back_command=self._show_section_4,
            primary_command=self._confirm_section_5,
            primary_text="Finalizar",
            left_buttons=[("📞 Solicitar Intérprete LSC", self._open_lsc_window)],
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        _queue_or_run_main_form_export(self, self._export_form)

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando acta...")
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
                "preparando el acta",
                presentacion_programa.export_to_excel,
            ),
        )

    def _open_lsc_window(self):
        ctx = _build_lsc_context(
            self,
            module=presentacion_programa,
            source_form="presentacion_programa",
        )
        _launch_linked_lsc_window(
            self,
            context=ctx,
            return_to_final_section=self._show_section_5,
            main_finish_action=self._confirm_section_5,
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
        fecha_error = _build_inline_error_label(top_inputs, wraplength=220)
        fecha_error.grid(row=1, column=1, sticky="w", pady=(4, 0), padx=(0, 24))
        ui_feedback.register_field(self, "fecha_visita", self.fields["fecha_visita"], error_label=fecha_error)

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
        modalidad_error = _build_inline_error_label(top_inputs, wraplength=220)
        modalidad_error.grid(row=1, column=3, sticky="w", pady=(4, 0))
        ui_feedback.register_field(self, "modalidad", self.fields["modalidad"], error_label=modalidad_error)

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
                ui_feedback.register_field(self, field_id, entry)

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
        lookup = getattr(self, "_empresa_lookup", presentacion_programa)
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(self, mode=mode, lookup=lookup, button=target_button)

    def _confirm_and_continue(self):
        _confirm_section1_and_continue(
            self,
            confirm_fn=presentacion_programa.confirm_section_1,
            next_step=self._show_section_2,
            extra_inputs=lambda: {"tipo_visita": ui_feedback.get_widget_value(self.fields.get("tipo_visita"))},
        )


def _normalize_company_search_text(value):
    text = unicodedata.normalize("NFD", str(value or "").strip())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text.casefold()


def _is_combobox_posted(widget):
    """Return True if the ttk.Combobox dropdown is currently visible."""
    try:
        popdown = widget.tk.call("ttk::combobox::PopdownPath", widget)
        return bool(widget.tk.call("winfo", "ismapped", popdown))
    except Exception:
        return False


def _set_combobox_dropdown_open(widget, open_dropdown):
    if not isinstance(widget, ttk.Combobox):
        return
    try:
        if open_dropdown:
            # Only call Post if the dropdown is not already visible.
            # Calling Post when it is already open steals focus on every keystroke.
            if not _is_combobox_posted(widget):
                widget.tk.call("ttk::combobox::Post", widget)
                _restore_combobox_text_focus(widget)
        else:
            widget.tk.call("ttk::combobox::Unpost", widget)
            _restore_combobox_text_focus(widget)
    except Exception:
        pass


def _restore_combobox_text_focus(widget, move_cursor_to_end=False):
    if not isinstance(widget, ttk.Combobox):
        return

    def _apply_focus():
        try:
            widget.focus_set()
            if move_cursor_to_end:
                widget.icursor(tk.END)
        except Exception:
            pass

    try:
        widget.after(10, _apply_focus)
    except Exception:
        _apply_focus()


def _filter_company_name_suggestions(options, prefix):
    query = _normalize_company_search_text(prefix)
    if not query:
        return []
    starts = []
    contains = []
    for option in options or []:
        text = str(option or "").strip()
        if not text:
            continue
        normalized = _normalize_company_search_text(text)
        if normalized.startswith(query):
            starts.append(text)
        elif query in normalized:
            contains.append(text)
    return (starts + contains)[:50]


def _show_empresa_autocomplete_popup(self, entry_widget, suggestions):
    """Show (or update) a floating suggestion list below the entry widget."""
    popup = getattr(self, "_empresa_ac_popup", None)

    if popup is None or not popup.winfo_exists():
        popup = tk.Toplevel(self)
        popup.wm_overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.withdraw()

        outer = tk.Frame(popup, bd=1, relief="solid", bg="#aaaaaa")
        outer.pack(fill="both", expand=True)

        listbox = tk.Listbox(
            outer,
            height=8,
            takefocus=False,
            selectmode="browse",
            font=FONT_LABEL,
            activestyle="dotbox",
            borderwidth=0,
            highlightthickness=0,
        )
        listbox.pack(fill="both", expand=True)

        def _on_lb_select(event=None):
            try:
                idx = listbox.curselection()
                if not idx:
                    return
                value = listbox.get(idx[0])
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, value)
                _hide_empresa_autocomplete_popup(self)
                entry_widget.focus_set()
                _section1_search_selected_company(self)
            except Exception:
                pass

        listbox.bind("<ButtonRelease-1>", _on_lb_select)
        listbox.bind("<Return>", _on_lb_select)

        self._empresa_ac_popup = popup
        self._empresa_ac_listbox = listbox
    else:
        listbox = self._empresa_ac_listbox

    listbox.delete(0, tk.END)
    for s in suggestions:
        listbox.insert(tk.END, s)

    entry_widget.update_idletasks()
    x = entry_widget.winfo_rootx()
    y = entry_widget.winfo_rooty() + entry_widget.winfo_height()
    w = max(entry_widget.winfo_width(), 200)
    row_h = 20
    h = min(len(suggestions), 8) * row_h + 4
    popup.geometry(f"{w}x{h}+{x}+{y}")
    popup.deiconify()
    popup.lift()


def _hide_empresa_autocomplete_popup(self):
    popup = getattr(self, "_empresa_ac_popup", None)
    if popup and popup.winfo_exists():
        popup.withdraw()


def _section1_search_selected_company(self):
    entry = self.fields.get("nombre_busqueda")
    if not entry:
        return
    value = str(entry.get() or "").strip()
    if not value:
        return
    _hide_empresa_autocomplete_popup(self)
    self._search_company("nombre")
    _restore_combobox_text_focus(entry, move_cursor_to_end=True)


def _section1_update_nombre_suggestions(self, open_dropdown=False):
    entry = self.fields.get("nombre_busqueda")
    if not entry:
        return
    prefix = entry.get().strip()
    normalized_prefix = _normalize_company_search_text(prefix)
    if len(normalized_prefix) < 2:
        entry["values"] = []
        _hide_empresa_autocomplete_popup(self)
        return
    suggestions = []
    lookup = getattr(self, "_empresa_lookup", None)
    if lookup and hasattr(lookup, "get_empresas_by_nombre_prefix"):
        try:
            suggestions = lookup.get_empresas_by_nombre_prefix(prefix, limit=0)
        except TypeError:
            try:
                suggestions = lookup.get_empresas_by_nombre_prefix(prefix)
            except Exception:
                suggestions = []
        except Exception:
            suggestions = []
    if not suggestions:
        cache = getattr(self, "_empresa_names_cache", None)
        if cache:
            suggestions = _filter_company_name_suggestions(cache, prefix)
    entry["values"] = suggestions
    if open_dropdown and suggestions:
        _show_empresa_autocomplete_popup(self, entry, suggestions)
    else:
        _hide_empresa_autocomplete_popup(self)


# ── VENTANA: HubWindow — ventana principal / menú de formularios ──────────────


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
        self._empresa_names_cache = []
        self._companies_by_id = {}
        self._companies_tree = None
        self._companies_search_var = None
        self._companies_sort_var = None
        self._version_var = tk.StringVar(value="Versión local: - | GitHub: -")
        self._version_check_thread = None
        self._latest_update_snapshot = None
        self._startup_update_prompt_shown = False
        self._drafts_btn = None
        self._completed_btn = None
        self._refresh_db_btn = None
        self._refresh_db_status_label = None
        self._labs_btn = None
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
        self.auto_login_var = None
        self._auto_login_cb = None
        self._auto_login_attempted = False
        self._auto_login_in_progress = False
        self._auto_login_thread = None
        self._open_form_windows = {}
        self._form_action_buttons = {}

        _ensure_drive_upload_worker()
        self._configure_input_styles()
        self.protocol("WM_DELETE_WINDOW", self._on_app_close)
        self.after(0, self._build_login)

    def _configure_input_styles(self):
        try:
            ttk.Style(self).theme_use("clam")
        except Exception:
            pass
        self.option_add("*Label.foreground", COLOR_TEXT_PRIMARY)
        self.option_add("*Entry.background", "white")
        self.option_add("*Entry.readonlyBackground", "#EDEDED")
        self.option_add("*Text.background", "white")
        for widget_class in ("TCombobox", "Spinbox", "TSpinbox"):
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                self.bind_class(widget_class, seq, lambda _e: "break")
        style = ttk.Style(self)
        style.configure("TLabel", foreground=COLOR_TEXT_PRIMARY, background=COLOR_LIGHT_BG)
        style.configure(
            "TButton",
            padding=(12, 6),
            font=("Arial", 10, "bold"),
            foreground=COLOR_TEXT_PRIMARY,
            background="white",
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "TButton",
            background=[("active", "#F5F1F8"), ("disabled", "#EFEAF4")],
            foreground=[("disabled", "#8A8394")],
        )
        style.configure("TEntry", fieldbackground="white", foreground=COLOR_TEXT_PRIMARY)
        style.configure("TCombobox", fieldbackground="white", background="white", foreground=COLOR_TEXT_PRIMARY)
        style.map("TCombobox", fieldbackground=[("readonly", "white")])
        style.configure(
            "Primary.TButton",
            foreground="white",
            background=COLOR_PRIMARY,
            borderwidth=1,
            relief="flat",
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#3C2452"), ("disabled", "#D9D2E3")],
            foreground=[("disabled", "#F8F6FB")],
        )
        style.configure(
            "Secondary.TButton",
            foreground=COLOR_TEXT_PRIMARY,
            background="white",
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#F5F1F8"), ("disabled", "#EDEDED")],
            foreground=[("disabled", "#888888")],
        )
        style.configure(
            "DangerOutline.TButton",
            foreground=COLOR_DANGER,
            background="white",
            borderwidth=1,
            relief="solid",
        )
        style.map(
            "DangerOutline.TButton",
            background=[("active", "#FFF4E5"), ("disabled", "#EDEDED")],
            foreground=[("disabled", "#888888")],
        )
        style.configure("Error.TLabel", foreground=COLOR_DANGER, background=COLOR_LIGHT_BG)
        style.configure("Success.TLabel", foreground=COLOR_SUCCESS, background=COLOR_LIGHT_BG)
        style.configure("Hint.TLabel", foreground=COLOR_TEXT_SECONDARY, background=COLOR_LIGHT_BG)
        style.configure("Invalid.TEntry", fieldbackground=COLOR_FIELD_ERROR_BG)
        style.configure(
            "Invalid.TCombobox",
            fieldbackground=COLOR_FIELD_ERROR_BG,
            background=COLOR_FIELD_ERROR_BG,
        )
        style.map("Invalid.TCombobox", fieldbackground=[("readonly", COLOR_FIELD_ERROR_BG)])
        style.configure("Changed.TEntry", fieldbackground=COLOR_FIELD_WARNING_BG)
        style.configure(
            "Changed.TCombobox",
            fieldbackground=COLOR_FIELD_WARNING_BG,
            background=COLOR_FIELD_WARNING_BG,
        )
        style.map("Changed.TCombobox", fieldbackground=[("readonly", COLOR_FIELD_WARNING_BG)])
        style.configure(
            "Vertical.TScrollbar",
            troughcolor="#F3E5F5",
            background=COLOR_PRIMARY,
            borderwidth=0,
            arrowcolor="white",
            arrowsize=12,
        )
        style.map(
            "Vertical.TScrollbar",
            background=[("active", "#3C2452"), ("pressed", "#3C2452")],
        )

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
            fg="#6B6B6B",
            bg="white",
            anchor="w",
        )
        self._startup_precheck_status_label.pack(anchor="w", pady=(4, 0))

    def _set_login_ready_state(self, ready, status_text=None):
        self._startup_precheck_completed = bool(ready)
        if self._login_btn is not None:
            try:
                self._login_btn.config(state="normal")
            except tk.TclError:
                pass
        if self.login_status is not None:
            try:
                if status_text is not None:
                    self.login_status.config(text=status_text)
            except tk.TclError:
                pass
        if ready:
            self._maybe_attempt_auto_login()

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
            self._startup_precheck_status_label.config(fg="#6B6B6B")
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
            command=self._update_auto_login_toggle_state,
            bg=COLOR_LIGHT_BG,
            activebackground=COLOR_LIGHT_BG,
        )
        remember_cb.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 6))

        self.auto_login_var = tk.BooleanVar(
            value=bool(saved_login.get("remember", True) and saved_login.get("auto_login", False))
        )
        self._auto_login_cb = tk.Checkbutton(
            form,
            text="Ingresar automáticamente en este equipo",
            variable=self.auto_login_var,
            bg=COLOR_LIGHT_BG,
            activebackground=COLOR_LIGHT_BG,
        )
        self._auto_login_cb.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self._update_auto_login_toggle_state()

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
            style="Primary.TButton",
            command=self._handle_login,
        )
        self._login_btn.pack(anchor="w")
        self.login_user_entry.bind("<Return>", lambda _event: self._handle_login())
        self.login_pass_entry.bind("<Return>", lambda _event: self._handle_login())

        forgot_btn = ttk.Button(
            self.login_frame,
            text="Olvide mi contraseña",
            command=self._show_forgot_password_info,
        )
        forgot_btn.pack(anchor="w", pady=(8, 0))
        self._maybe_attempt_auto_login(ignore_precheck=True)
        self._run_startup_precheck_async(log_enabled=True)

    def _update_auto_login_toggle_state(self):
        remember_enabled = True
        try:
            remember_enabled = bool(self.remember_login_var.get())
        except Exception:
            pass
        if not remember_enabled and self.auto_login_var is not None:
            try:
                self.auto_login_var.set(False)
            except Exception:
                pass
        if self._auto_login_cb is not None:
            try:
                self._auto_login_cb.config(state=("normal" if remember_enabled else "disabled"))
            except tk.TclError:
                pass

    def _maybe_attempt_auto_login(self, ignore_precheck=False):
        if self._auto_login_attempted or self._auto_login_in_progress:
            return
        if (not ignore_precheck and not self._startup_precheck_completed) or self.login_frame is None:
            return
        remember_enabled = False
        auto_login_enabled = False
        try:
            remember_enabled = bool(self.remember_login_var.get())
        except Exception:
            pass
        try:
            auto_login_enabled = bool(self.auto_login_var.get()) if self.auto_login_var is not None else False
        except Exception:
            auto_login_enabled = False
        if not remember_enabled or not auto_login_enabled:
            return
        username = self.login_user_entry.get().strip() if self.login_user_entry is not None else ""
        password = self.login_pass_entry.get() if self.login_pass_entry is not None else ""
        if not username or not password:
            return
        self._auto_login_attempted = True
        self._auto_login_in_progress = True
        if self.login_status is not None:
            try:
                self.login_status.config(text="Intentando ingresar automáticamente...")
            except tk.TclError:
                pass
        self._start_auto_login_async(username_input=username, password=password)

    def _start_auto_login_async(self, *, username_input, password):
        username_input = str(username_input or "").strip()
        username = _normalize_login_value(username_input)
        saved_creds = _load_saved_login_credentials()
        cached_email = (
            saved_creds.get("resolved_email", "")
            if _normalize_login_value(saved_creds.get("username", "")) == username
            else ""
        )

        def _worker():
            try:
                auth_result = self._resolve_login_attempt(
                    username,
                    password,
                    cached_email=cached_email,
                )
                callback = lambda: self._complete_login_with_auth_result(
                    username_input=username_input,
                    username=username,
                    password=password,
                    auth_result=auth_result,
                    silent=True,
                    auto_login=True,
                )
            except Exception as exc:
                callback = lambda err=exc: self._handle_failed_login_attempt(
                    message=(
                        "No fue posible ingresar automáticamente. Verifica tu conexión o inicia sesión manualmente."
                        if _is_connectivity_exception(err)
                        else _log_user_error("login", err)
                    ),
                    silent=True,
                    clear_saved=bool(_is_invalid_credentials_exception(err)),
                )
            try:
                self.after(0, callback)
            except Exception:
                self._auto_login_in_progress = False

        self._auto_login_thread = threading.Thread(target=_worker, daemon=True)
        self._auto_login_thread.start()

    def _show_forgot_password_info(self):
        messagebox.showinfo(
            "Recuperacion de contraseña",
            "Para recuperar tu contraseña, comunicate con Aaron Uyaban\n"
            "Correo: admonusaid@recacolombia.org",
        )

    def _finalize_login_success(self, *, user_row, username_input, username, password):
        self._cache_offline_user_auth(user_row, password)
        remember_enabled = True
        auto_login_enabled = False
        try:
            remember_enabled = bool(self.remember_login_var.get())
        except Exception:
            pass
        try:
            auto_login_enabled = bool(self.auto_login_var.get()) if self.auto_login_var is not None else False
        except Exception:
            auto_login_enabled = False
        if remember_enabled:
            _save_login_credentials(
                username_input or username,
                password,
                resolved_email=str((user_row or {}).get("_resolved_email") or ""),
                auto_login=bool(remember_enabled and auto_login_enabled),
            )
        else:
            _clear_login_credentials()
        self.current_user = (user_row.get("usuario_login") or username).strip()
        self.current_user_profile = user_row
        if not bool((user_row or {}).get("_profile_fallback")):
            self._schedule_profesional_asignado_normalization()
        self._start_usage_session()
        if self.login_frame:
            self.login_frame.destroy()
            self.login_frame = None
        self._auto_login_in_progress = False
        self._build_header()
        self._build_body()

    def _schedule_profesional_asignado_normalization(self):
        if bool(getattr(self, "_profesional_normalization_in_progress", False)):
            return False
        self._profesional_normalization_in_progress = True

        def _worker():
            try:
                result = self._normalize_profesional_asignado()
                status = str((result or {}).get("status") or "").strip().lower()
                updated = int((result or {}).get("updated_rows") or 0)
                if status not in {"ok", "skipped"}:
                    _log_capture(
                        "[LOGIN] profesional_asignado normalization deferred "
                        f"status={status or 'unknown'} updated_rows={updated}"
                    )
            except Exception as exc:
                _log_capture(f"[LOGIN] profesional_asignado normalization failed: {exc}")
            finally:
                self._profesional_normalization_in_progress = False

        threading.Thread(target=_worker, daemon=True).start()
        return True

    def _handle_failed_login_attempt(self, *, message, silent=False, clear_saved=False):
        self._auto_login_in_progress = False
        if clear_saved:
            _clear_login_credentials()
            try:
                if self.remember_login_var is not None:
                    self.remember_login_var.set(False)
            except Exception:
                pass
            try:
                if self.auto_login_var is not None:
                    self.auto_login_var.set(False)
            except Exception:
                pass
            self._update_auto_login_toggle_state()
        if self.login_status is not None:
            try:
                self.login_status.config(text=message if silent else "")
            except tk.TclError:
                pass
        if silent:
            return False
        messagebox.showerror("Error", message)
        return False

    def _handle_login(self, silent=False, auto_login=False):
        username_input = self.login_user_entry.get().strip()
        username = _normalize_login_value(username_input)
        password = self.login_pass_entry.get()
        if not username or not password:
            message = "Ingresa usuario y contraseña."
            if auto_login:
                message = "No fue posible ingresar automáticamente. Ingresa tus credenciales."
            return self._handle_failed_login_attempt(message=message, silent=silent)
        saved_creds = _load_saved_login_credentials()
        cached_email = (
            saved_creds.get("resolved_email", "")
            if _normalize_login_value(saved_creds.get("username", "")) == username
            else ""
        )
        try:
            self.login_status.config(
                text=(
                    "Intentando ingresar automáticamente..."
                    if auto_login
                    else "Validando credenciales..."
                )
            )
            self.update_idletasks()
            auth_result = self._resolve_login_attempt(
                username,
                password,
                cached_email=cached_email,
            )
        except Exception as exc:
            if not _is_connectivity_exception(exc):
                return self._handle_failed_login_attempt(
                    message=_log_user_error("login", exc),
                    silent=silent,
                    clear_saved=bool(auto_login and _is_invalid_credentials_exception(exc)),
                )
            auth_result = {
                "user_row": None,
                "used_offline": False,
                "auth_exc": exc,
            }
        return self._complete_login_with_auth_result(
            username_input=username_input,
            username=username,
            password=password,
            auth_result=auth_result,
            silent=silent,
            auto_login=auto_login,
        )

    def _resolve_login_attempt(self, username, password, *, cached_email=""):
        used_offline = False
        auth_exc = None
        user_row = None
        try:
            user_row = self._authenticate_user(username, password, cached_email=cached_email)
        except Exception as exc:
            auth_exc = exc
            if not _is_connectivity_exception(exc):
                raise
        if not user_row:
            can_use_offline = bool(auth_exc and _is_connectivity_exception(auth_exc))
            if not can_use_offline:
                try:
                    can_use_offline = not _supabase_ping(timeout=3)
                except Exception:
                    can_use_offline = True
            if can_use_offline:
                user_row = self._authenticate_user_offline(username, password)
                if user_row:
                    used_offline = True
        return {
            "user_row": user_row,
            "used_offline": used_offline,
            "auth_exc": auth_exc,
        }

    def _complete_login_with_auth_result(
        self,
        *,
        username_input,
        username,
        password,
        auth_result,
        silent=False,
        auto_login=False,
    ):
        user_row = (auth_result or {}).get("user_row")
        used_offline = bool((auth_result or {}).get("used_offline"))
        auth_exc = (auth_result or {}).get("auth_exc")
        if used_offline and self.login_status is not None:
            try:
                self.login_status.config(text="Modo offline: sesión local")
            except tk.TclError:
                pass
        if not user_row:
            if auth_exc:
                return self._handle_failed_login_attempt(
                    message=(
                        "No fue posible ingresar automáticamente. Verifica tu conexión o inicia sesión manualmente."
                        if auto_login
                        else _log_user_error("login", auth_exc)
                    ),
                    silent=silent,
                )
            return self._handle_failed_login_attempt(
                message=(
                    "No fue posible ingresar automáticamente. Vuelve a iniciar sesión."
                    if auto_login
                    else "Usuario y contraseña incorrectos."
                ),
                silent=silent,
                clear_saved=bool(auto_login),
            )
        if not used_offline and self._must_force_password_change(user_row, password):
            changed = self._prompt_force_password_change(user_row, password)
            if not changed:
                self._auto_login_in_progress = False
                self.login_status.config(
                    text=(
                        "Debes iniciar sesión manualmente para cambiar la contraseña."
                        if auto_login
                        else ""
                    )
                )
                if not silent:
                    messagebox.showwarning(
                        "Cambio requerido",
                        "Debes cambiar la contraseña para continuar.",
                    )
                return False
            # Reload profile to keep local state aligned.
            try:
                refreshed = _supabase_rpc("get_my_profesional_profile", {})
                if isinstance(refreshed, dict):
                    user_row = refreshed
            except Exception:
                pass
        self._finalize_login_success(
            user_row=user_row,
            username_input=username_input,
            username=username,
            password=password,
        )
        return True

    def _authenticate_user(self, username, password, cached_email=""):
        username_norm = _normalize_login_value(username)
        if not username_norm:
            return None
        _clear_supabase_session()

        def _resolve_email_fresh():
            try:
                resolved = _supabase_rpc(
                    "resolve_login_email",
                    {"p_login": username_norm},
                    use_session=False,
                )
            except Exception as exc:
                raise RuntimeError(str(exc)) from exc
            if isinstance(resolved, str):
                return resolved.strip()
            if isinstance(resolved, dict):
                return str(
                    resolved.get("resolve_login_email")
                    or resolved.get("email")
                    or ""
                ).strip()
            return ""

        email = str(cached_email or "").strip()
        email_was_cached = bool(email)

        if not email:
            email = _resolve_email_fresh()
        if not email:
            return None

        try:
            _supabase_auth_password_login(email, password)
        except Exception as exc:
            # If login failed and we used a cached email, retry with a fresh resolve
            # in case the email changed since the cache was written.
            if email_was_cached and _is_invalid_credentials_exception(exc):
                _log_capture("[LOGIN] cached email rejected, re-resolving fresh")
                email = _resolve_email_fresh()
                if not email:
                    return None
                _supabase_auth_password_login(email, password)
                email_was_cached = False  # signal caller to update cache
            else:
                raise

        try:
            profile = _supabase_rpc("get_my_profesional_profile", {})
        except Exception as exc:
            if _is_profile_permission_exception(exc):
                _log_capture(
                    f"[LOGIN] profile authorization denied user={username_norm!r} reason={exc}"
                )
                raise RuntimeError(
                    "No fue posible validar tu perfil con los permisos actuales."
                ) from exc
            raise
        if isinstance(profile, dict) and profile.get("id"):
            profile["_auth_source"] = "jwt"
            profile["_resolved_email"] = email
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
        if _offline_auth_entry_is_expired(cached):
            try:
                users.pop(username_norm, None)
                _save_offline_auth_store(store)
            except Exception:
                pass
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
            if len(str(password or "").strip()) > MAX_PASSWORD_LENGTH:
                return
            pass_hash = _hash_password(str(password or "").strip())
        if not pass_hash:
            return
        store = _load_offline_auth_store()
        store["version"] = OFFLINE_AUTH_STORE_VERSION
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
            "cached_at": time.time(),
            "ttl_days": OFFLINE_AUTH_TTL_DAYS,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        _save_offline_auth_store(store)

    def _must_force_password_change(self, user_row, current_password):
        _ = current_password
        return bool(user_row.get("auth_password_temp"))

    def _validate_new_password(self, new_password, current_password):
        pwd = str(new_password or "")
        if len(pwd) > MAX_PASSWORD_LENGTH:
            return False, "La nueva contraseña supera la longitud máxima permitida."
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
        except Exception as exc:
            _log_capture(f"[USAGE] _usage_upsert_sync failed table={table} err={exc}")
            try:
                _supabase_enqueue_upsert(table, [row], on_conflict=on_conflict)
                _log_capture(f"[USAGE] enqueued as fallback table={table}")
                return True
            except Exception as q_exc:
                _log_capture(f"[USAGE] enqueue_fallback failed table={table} err={q_exc}")
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

    def _set_form_card_state(self, form_id, *, active):
        button = (self._form_action_buttons or {}).get(str(form_id or ""))
        if button is None:
            return
        try:
            button.config(
                text=("En progreso..." if active else "Abrir"),
                state=("disabled" if active else "normal"),
            )
        except tk.TclError:
            return

    def _register_open_form_window(self, form_id, window):
        if not form_id or window is None:
            return
        self._open_form_windows[str(form_id)] = window
        self._set_form_card_state(form_id, active=True)

    def _release_form_window(self, form_id, window=None):
        key = str(form_id or "")
        if not key:
            return
        current = (self._open_form_windows or {}).get(key)
        if window is not None and current is not None and current is not window:
            return
        self._open_form_windows.pop(key, None)
        self._set_form_card_state(key, active=False)

    def _find_form_completion_record_by_source_item_key(self, source_item_key):
        source_key = str(source_item_key or "").strip()
        if not source_key:
            return {}
        try:
            rows = _supabase_get_paged(
                "formatos_finalizados_il",
                {
                    "select": "registro_id",
                    "source_item_key": f"eq.{source_key}",
                },
                page_size=1,
                max_pages=1,
            )
        except Exception as exc:
            _log_capture(f"find_form_completion_record failed source_item_key={source_key} err={exc}")
            return {}
        return dict(rows[0]) if rows else {}

    def create_form_completion_record(
        self,
        form_name,
        company_name,
        output_path=None,
        *,
        source_item_key=None,
        payload_schema_version=1,
        payload_source="form_cache",
        payload_raw=None,
        payload_normalized=None,
        payload_generated_at=None,
        drive_file_id="",
        upload_status="pending",
        upload_error="",
        upload_attempted_at=None,
        uploaded_at=None,
    ):
        if not self._should_track_usage():
            return ""
        usuario_login = (self.current_user_profile.get("usuario_login") or self.current_user or "").strip()
        nombre_usuario = (self.current_user_profile.get("nombre_profesional") or self.current_user or "").strip()
        now_col = self._get_colombia_now()
        existing = self._find_form_completion_record_by_source_item_key(source_item_key)
        registro_id = str(existing.get("registro_id") or uuid.uuid4())
        row = {
            "registro_id": registro_id,
            "session_id": self.current_session_id,
            "usuario_login": usuario_login,
            "nombre_usuario": nombre_usuario,
            "nombre_formato": (form_name or "").strip(),
            "nombre_empresa": (company_name or "").strip(),
            "path_formato": str(output_path or "").strip(),
            "drive_file_id": str(drive_file_id or "").strip(),
            "finalizado_at_colombia": now_col.strftime("%Y-%m-%d %H:%M:%S"),
            "finalizado_at_iso": now_col.isoformat(),
            "upload_status": str(upload_status or "").strip() or "pending",
            "upload_error": str(upload_error or "").strip(),
            "upload_attempted_at": upload_attempted_at,
            "uploaded_at": uploaded_at,
            "source_item_key": str(source_item_key or "").strip() or None,
            "payload_schema_version": int(payload_schema_version or 1),
            "payload_source": str(payload_source or "").strip() or "form_cache",
            "payload_raw": payload_raw,
            "payload_normalized": payload_normalized,
            "payload_generated_at": payload_generated_at,
        }
        try:
            _store_completed_form_locally(row)
        except Exception as exc:
            _log_capture(f"store_completed_form_locally failed registro_id={registro_id} err={exc}")
        if output_path:
            _log_capture(
                f"create_form_completion_record form={form_name} company={company_name} output={output_path}"
            )
        try:
            saved = self._usage_upsert_sync("formatos_finalizados_il", row, on_conflict="registro_id")
            if not saved:
                _log_capture(
                    f"[USAGE] create_form_completion_record upsert returned False "
                    f"registro_id={registro_id} form={form_name}"
                )
        except Exception as exc:
            _log_capture(f"create_form_completion_record failed registro_id={registro_id} err={exc}")
        return registro_id

    def finalize_form_delivery(
        self,
        output_path,
        *,
        form_name,
        company_name,
        drive_job=None,
        completion_payload=None,
    ):
        completion_payload = dict(completion_payload or {})
        payload_normalized = completion_payload.get("payload_normalized")
        payload_raw = completion_payload.get("payload_raw")
        payload_source = (
            ((payload_normalized or {}).get("metadata") or {}).get("payload_source")
            or "form_cache"
        )
        payload_generated_at = (
            ((payload_normalized or {}).get("metadata") or {}).get("generated_at")
            or ((payload_raw or {}).get("metadata") or {}).get("generated_at")
        )
        job = {
            "registro_id": self.create_form_completion_record(
                form_name,
                company_name,
                output_path=output_path,
                source_item_key=completion_payload.get("source_item_key"),
                payload_schema_version=completion_payloads.PAYLOAD_SCHEMA_VERSION,
                payload_source=payload_source,
                payload_raw=payload_raw,
                payload_normalized=payload_normalized,
                payload_generated_at=payload_generated_at,
            ),
            "form_name": form_name,
            "company_name": company_name,
            "local_excel_path": output_path,
        }
        if isinstance(drive_job, dict):
            job.update(drive_job)
        if job.get("status") == "synced":
            remote_url = str(job.get("remote_url") or "").strip()
            drive_file_id = str(job.get("drive_file_id") or "").strip()
            uploaded_at = _get_colombia_now().isoformat()
            _update_form_completion_upload_status(
                job.get("registro_id"),
                upload_status="synced",
                upload_error="",
                upload_attempted_at=uploaded_at,
                uploaded_at=uploaded_at,
                path_formato=remote_url,
                drive_file_id=drive_file_id,
            )
            result = {
                "status": "synced",
                "error": "",
                "attempted_at": uploaded_at,
                "uploaded_at": uploaded_at,
                "output_path": output_path,
                "remote_url": remote_url,
                "drive_file_id": drive_file_id,
            }
        else:
            try:
                result = _perform_drive_upload_attempt(job)
            except Exception as exc:
                _log_capture(
                    f"finalize_form_delivery unexpected_drive_exception "
                    f"registro_id={job.get('registro_id')} err={exc}"
                )
                result = _build_drive_upload_result_from_exception(job, exc)
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

    def record_followup_completion(self, *, case_target, case_path, case_record=None, followup_index):
        if not self._should_track_usage():
            return ""
        base_payload = seguimientos.get_base_payload(case_target)
        followup_payload = seguimientos.get_followup_payload(case_target, followup_index)
        case_record_data = dict(case_record or {})
        output_path = str(case_path or case_record_data.get("local_path") or "").strip()
        remote_url = str(case_record_data.get("webViewLink") or "").strip()
        drive_file_id = str(case_record_data.get("id") or case_record_data.get("drive_file_id") or "").strip()
        completion_payload = completion_payloads.build_followup_completion_payload(
            case_ref=case_target,
            followup_index=followup_index,
            base_payload=base_payload,
            followup_payload=followup_payload,
            form_name=f"Seguimiento al Proceso de Inclusion Laboral #{int(followup_index)}",
            session_id=self.current_session_id,
            app_version=get_version(),
            extra_context={
                "payload_source": "seguimientos_sheet",
                "local_path": output_path,
                "remote_url": remote_url,
                "drive_file_id": drive_file_id,
                "case_record": case_record_data,
                "case_meta": seguimientos.get_case_meta(case_target),
            },
        )
        now_iso = self._get_colombia_now().isoformat()
        upload_status = "synced" if drive_file_id or remote_url else "pending"
        return self.create_form_completion_record(
            f"Seguimiento al Proceso de Inclusion Laboral #{int(followup_index)}",
            str(base_payload.get("nombre_empresa") or "").strip(),
            output_path=output_path or remote_url,
            source_item_key=completion_payload.get("source_item_key"),
            payload_schema_version=completion_payloads.PAYLOAD_SCHEMA_VERSION,
            payload_source="seguimientos_sheet",
            payload_raw=completion_payload.get("payload_raw"),
            payload_normalized=completion_payload.get("payload_normalized"),
            payload_generated_at=(
                ((completion_payload.get("payload_normalized") or {}).get("metadata") or {}).get("generated_at")
            ),
            drive_file_id=drive_file_id,
            upload_status=upload_status,
            upload_attempted_at=now_iso if upload_status == "synced" else None,
            uploaded_at=now_iso if upload_status == "synced" else None,
        )

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

    def _get_expected_update_installer_asset_name(self):
        try:
            _owner, _repo, _token, installer_asset, _hash_asset = _updater_repo_config()
            return str(installer_asset or "").strip()
        except Exception as exc:
            _log_capture(f"_get_expected_update_installer_asset_name error: {exc}")
            return ""

    def _open_release_page(self, version=None):
        try:
            url = get_release_page_url(version)
            webbrowser.open(url)
            _log_capture(f"_open_release_page: {url}")
            return True
        except Exception as exc:
            _log_capture(f"_open_release_page error: {exc}")
            messagebox.showerror("Actualización", f"No se pudo abrir el release: {exc}")
            return False

    def _update_assets_are_usable(self, assets):
        installer_asset = self._get_expected_update_installer_asset_name()
        if not installer_asset or not isinstance(assets, dict):
            return False
        return bool(str((assets or {}).get(installer_asset) or "").strip())

    def _normalize_update_snapshot(self, snapshot):
        raw = dict(snapshot or {})
        local_raw = raw.get("local_version")
        remote_raw = raw.get("remote_version")
        error_raw = raw.get("error")
        assets = dict(raw.get("assets") or {})
        local = str(local_raw or "0.0.0").strip() or "0.0.0"
        remote = str(remote_raw).strip() if remote_raw else ""
        error = str(error_raw).strip() if error_raw else ""
        update_available = is_update_available(local, remote or None)
        has_installer_asset = self._update_assets_are_usable(assets)
        can_start_update = bool(update_available and has_installer_asset)
        return {
            "local_version": local,
            "remote_version": remote or None,
            "assets": assets,
            "error": error or None,
            "update_available": update_available,
            "has_installer_asset": has_installer_asset,
            "can_start_update": can_start_update,
        }

    def _resolve_update_snapshot(self):
        local = get_version() or "0.0.0"
        _log_capture(f"_resolve_update_snapshot start local={local}")
        snapshot = {
            "local_version": local,
            "remote_version": None,
            "assets": {},
            "error": None,
        }
        try:
            remote, assets = get_latest_release_assets()
            snapshot["remote_version"] = remote
            snapshot["assets"] = dict(assets or {})
            _log_capture(
                "_resolve_update_snapshot success "
                f"remote={remote} assets={list(snapshot['assets'].keys())}"
            )
        except Exception as exc:
            snapshot["error"] = str(exc)
            _log_capture(f"_resolve_update_snapshot error: {exc}")
        normalized = self._normalize_update_snapshot(snapshot)
        _log_capture(
            "_resolve_update_snapshot normalized "
            f"remote={normalized['remote_version']} "
            f"update_available={normalized['update_available']} "
            f"can_start_update={normalized['can_start_update']}"
        )
        return normalized

    def _apply_update_snapshot(self, snapshot, *, prompt_startup=False):
        normalized = self._normalize_update_snapshot(snapshot)
        self._latest_update_snapshot = normalized
        self.set_version_info(normalized["local_version"], normalized["remote_version"])
        if prompt_startup:
            try:
                self.after(0, lambda snap=dict(normalized): self._maybe_prompt_startup_update(snap))
            except tk.TclError:
                pass
        return normalized

    def _confirm_and_start_update(self, snapshot, *, source):
        normalized = self._normalize_update_snapshot(snapshot)
        if not normalized["can_start_update"]:
            _log_capture(f"{source}: actualización omitida por flujo no elegible")
            return False
        confirm = messagebox.askyesno(
            "Actualización disponible",
            f"Hay una nueva versión disponible ({normalized['remote_version']}).\n¿Deseas actualizar ahora?",
        )
        if not confirm:
            _log_capture(f"{source}: usuario canceló actualización")
            return False
        _log_capture(f"{source}: usuario confirmó actualización")
        self._start_manual_update(normalized)
        return True

    def _maybe_prompt_startup_update(self, snapshot):
        normalized = self._normalize_update_snapshot(snapshot)
        if self._startup_update_prompt_shown:
            _log_capture("_maybe_prompt_startup_update omitido: prompt ya mostrado")
            return False
        if normalized["error"]:
            _log_capture("_maybe_prompt_startup_update omitido: error en snapshot")
            return False
        if not normalized["remote_version"]:
            _log_capture("_maybe_prompt_startup_update omitido: versión remota vacía")
            return False
        if not normalized["update_available"]:
            _log_capture("_maybe_prompt_startup_update omitido: sin actualización")
            return False
        if not normalized["has_installer_asset"]:
            _log_capture("_maybe_prompt_startup_update omitido: release sin instalador válido")
            return False
        self._startup_update_prompt_shown = True
        _log_capture(
            "_maybe_prompt_startup_update: actualización disponible "
            f"remote={normalized['remote_version']} local={normalized['local_version']}"
        )
        return self._confirm_and_start_update(normalized, source="_maybe_prompt_startup_update")

    def _handle_manual_update_snapshot(self, snapshot):
        normalized = self._apply_update_snapshot(snapshot, prompt_startup=False)
        if normalized["error"]:
            _log_capture(f"_handle_manual_update_snapshot error: {normalized['error']}")
            messagebox.showerror("Actualización", f"No se pudo verificar: {normalized['error']}")
            return False
        if not normalized["remote_version"]:
            _log_capture("_handle_manual_update_snapshot: remote vacío")
            messagebox.showerror("Actualización", "No se pudo obtener la versión remota.")
            return False
        if not normalized["update_available"]:
            _log_capture("_handle_manual_update_snapshot: sin actualización disponible")
            messagebox.showinfo("Actualización", "Ya estás usando la última versión.")
            return False
        if not normalized["has_installer_asset"]:
            installer_asset = self._get_expected_update_installer_asset_name() or "instalador"
            _log_capture(
                "_handle_manual_update_snapshot: release sin instalador esperado "
                f"asset={installer_asset}"
            )
            messagebox.showerror(
                "Actualización",
                "Hay una nueva versión disponible "
                f"({normalized['remote_version']}), pero el release no incluye el instalador "
                f"esperado ({installer_asset}).",
            )
            return False
        _log_capture(
            "_handle_manual_update_snapshot: actualización disponible "
            f"remote={normalized['remote_version']} local={normalized['local_version']}"
        )
        return self._confirm_and_start_update(normalized, source="_handle_manual_update_snapshot")

    def _refresh_version_info_async(self):
        if self._version_check_thread and self._version_check_thread.is_alive():
            _log_capture("_refresh_version_info_async omitido: hilo activo")
            return

        def _worker():
            snapshot = self._resolve_update_snapshot()
            try:
                self.after(0, lambda snap=snapshot: self._apply_update_snapshot(snap, prompt_startup=True))
            except tk.TclError:
                pass

        self._version_check_thread = threading.Thread(target=_worker, daemon=True)
        _log_capture("_refresh_version_info_async hilo iniciado")
        self._version_check_thread.start()

    def _open_update_page(self):
        _log_capture("_open_update_page: inicio de verificación manual")
        dialog = LoadingDialog(self, title="Verificando actualización")
        dialog.set_status("Consultando versión en GitHub...")
        dialog.set_progress(20)

        result = {"snapshot": None}

        def _worker():
            result["snapshot"] = self._resolve_update_snapshot()

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()

        def _check_done():
            if thread.is_alive():
                self.after(200, _check_done)
                return
            dialog.close()
            self._handle_manual_update_snapshot(result["snapshot"] or {})

        self.after(200, _check_done)

    def _start_manual_update(self, snapshot):
        normalized = self._normalize_update_snapshot(snapshot)
        assets = dict(normalized.get("assets") or {})
        _log_capture(
            "_start_manual_update: inicio "
            f"remote={normalized.get('remote_version') or '?'} assets={list(assets.keys())}"
        )
        dialog = LoadingDialog(self, title="Descargando instalador")
        dialog.set_status("Preparando descarga...")
        dialog.set_progress(5)
        result = {"error": None, "path": None}

        def _progress(message, value):
            self.after(0, lambda: dialog.set_status(message))
            self.after(0, lambda: dialog.set_progress(value))

        def _worker():
            try:
                path = download_installer(assets, progress_callback=_progress)
                result["path"] = path
                _log_capture(f"_start_manual_update: instalador descargado en {path}")
                self.after(0, lambda: dialog.set_status("Listo. Cerrando para instalar..."))
                self.after(0, lambda: dialog.set_progress(100))
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
            _log_capture("_start_manual_update resultado OK: iniciando instalación")
            self._show_restart_countdown(result["path"])

        self.after(300, _check_done)

    def _show_restart_countdown(self, installer_path, seconds=5):
        modal = tk.Toplevel(self)
        modal.title("Actualización")
        modal.configure(bg=COLOR_LIGHT_BG)
        modal.transient(self)
        modal.grab_set()
        modal.resizable(False, False)

        body = tk.Frame(modal, bg=COLOR_LIGHT_BG, padx=18, pady=14)
        body.pack(fill="both", expand=True)
        tk.Label(
            body,
            text="Descarga lista",
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 6))
        countdown = tk.Label(
            body,
            text=f"La aplicación se cerrará en {seconds} segundos para instalar...",
            font=("Arial", 10),
            fg=COLOR_TEAL,
            bg=COLOR_LIGHT_BG,
        )
        countdown.pack(anchor="w")

        modal.update_idletasks()
        w, h = 400, 140
        x = (modal.winfo_screenwidth() // 2) - (w // 2)
        y = (modal.winfo_screenheight() // 2) - (h // 2)
        modal.geometry(f"{w}x{h}+{x}+{y}")

        def _tick(remaining):
            if remaining <= 0:
                try:
                    modal.destroy()
                except Exception:
                    pass
                self.after(200, lambda: self._launch_installer_and_exit(installer_path))
                return
            countdown.config(text=f"La aplicación se cerrará en {remaining} segundos para instalar...")
            self.after(1000, lambda: _tick(remaining - 1))

        _tick(seconds)

    def _launch_installer_and_exit(self, installer_path):
        args = [
            str(installer_path),
            "/VERYSILENT",
            "/CURRENTUSER",
            "/SUPPRESSMSGBOXES",
            "/NORESTART",
        ]
        try:
            subprocess.Popen(
                args,
                creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP,
            )
            _log_capture(f"_launch_installer_and_exit: installer lanzado {installer_path}")
        except Exception as exc:
            _log_capture(f"_launch_installer_and_exit error al lanzar: {exc}")
        os._exit(0)

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
            text="Selecciona el formulario que necesitas diligenciar hoy",
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
            text="● Verificando...",
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
                badge_text, badge_color = _get_services_badge_state(
                    internet_state,
                    supabase_state,
                    drive_state,
                    total_pending=total_pending,
                    total_failed=total_failed,
                )
                self._net_status_label.config(text=badge_text, fg=badge_color)
            if self._sync_panel_btn:
                self._sync_panel_btn.config(text="Ver detalles")
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

        tk.Label(
            frame,
            text="Estado de servicios",
            font=("Arial", 11, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 6))

        internet_status_lbl = tk.Label(
            frame,
            text="",
            font=("Arial", 9),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
        )
        internet_status_lbl.pack(anchor="w")
        supabase_status_lbl = tk.Label(
            frame,
            text="",
            font=("Arial", 9),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
        )
        supabase_status_lbl.pack(anchor="w")
        drive_status_lbl = tk.Label(
            frame,
            text="",
            font=("Arial", 9),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
        )
        drive_status_lbl.pack(anchor="w")
        queue_summary_lbl = tk.Label(
            frame,
            text="",
            font=("Arial", 9),
            fg="#555555",
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
        )
        queue_summary_lbl.pack(anchor="w", pady=(2, 8))

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
            ui_feedback.set_semantic_label(
                internet_status_lbl,
                (
                    f"Internet: {'Online' if self._is_online else 'Offline'}"
                    f" | Último estado: {internet_state.get('status_text') or '-'}"
                ),
                state=("success" if self._is_online else "error"),
            )
            ui_feedback.set_semantic_label(
                supabase_status_lbl,
                f"Supabase: {'OK' if supabase_state.get('ok') else 'Falla'} | Detalle: {supabase_state.get('status_text') or '-'}",
                state=("success" if supabase_state.get("ok") else "error"),
            )
            ui_feedback.set_semantic_label(
                drive_status_lbl,
                f"Drive: {'OK' if drive_state.get('ok') else 'Falla'} | Detalle: {drive_state.get('status_text') or '-'}",
                state=("success" if drive_state.get("ok") else "error"),
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

            queue_summary_lbl.config(
                text=(
                    f"Colas: Supabase {len(supabase_pending_rows)} pendientes y {len(supabase_failed_rows)} fallidos | "
                    f"Drive {len(drive_pending_rows)} pendientes y {len(drive_failed_rows)} fallidos"
                ),
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
        try:
            profesionales = _supabase_get_paged(
                "profesionales",
                {"select": "nombre_profesional"},
                page_size=1000,
                max_pages=20,
            )
        except Exception as exc:
            _log_capture(f"[LOGIN] profesionales fetch failed during normalization: {exc}")
            return {"status": "deferred", "updated_rows": 0, "error": str(exc)}
        alias_map = {}
        for row in profesionales:
            nombre = (row.get("nombre_profesional") or "").strip()
            if not nombre:
                continue
            for alias in self._build_profesional_aliases(nombre):
                alias_map.setdefault(alias, nombre)
        if not alias_map:
            return {"status": "skipped", "updated_rows": 0, "error": ""}

        try:
            empresas = _supabase_get_paged(
                "empresas",
                {
                    "select": "id,profesional_asignado",
                    "profesional_asignado": "not.is.null",
                },
                page_size=1000,
                max_pages=50,
            )
        except Exception as exc:
            _log_capture(f"[LOGIN] empresas fetch failed during normalization: {exc}")
            return {"status": "deferred", "updated_rows": 0, "error": str(exc)}
        updates = []
        for row in empresas:
            current = (row.get("profesional_asignado") or "").strip()
            if not current:
                continue
            key = self._norm_match(current)
            target = alias_map.get(key)
            if target and target != current:
                updates.append({"id": row.get("id"), "profesional_asignado": target})
        if not updates:
            return {"status": "ok", "updated_rows": 0, "error": ""}
        try:
            result = _supabase_upsert_with_queue("empresas", updates, on_conflict="id")
        except Exception as exc:
            _log_capture(f"[LOGIN] empresas upsert failed during normalization: {exc}")
            return {"status": "deferred", "updated_rows": 0, "error": str(exc)}
        return {
            "status": str((result or {}).get("status") or "ok"),
            "updated_rows": len(updates),
            "error": str((result or {}).get("error") or ""),
        }

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
        _attach_autoexpand(comentarios_txt, 6, 20)
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
            drafts = []
        items = [item for item in drafts if isinstance(item, dict)]
        items.extend(_list_followup_local_drafts_for_user(user_login))
        items.sort(key=lambda item: str(item.get("updated_at") or ""), reverse=True)
        return items

    def _get_user_completed_forms(self):
        user_login = self._get_current_user_login()
        if not user_login:
            return []
        _sync_completed_forms_from_remote(user_login)
        data = _load_completed_forms_store()
        users = data.get("users", {})
        entries = users.get(user_login, [])
        if not isinstance(entries, list):
            return []
        reopenable = []
        for item in entries:
            if not isinstance(item, dict):
                continue
            if _resolve_completed_restore_form_meta(item) is None:
                continue
            reopenable.append(item)
        return reopenable

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

    def _persist_form_draft(
        self,
        window,
        *,
        allow_empty=False,
        silent=False,
        toast_text="",
        source="manual",
    ):
        form_id = getattr(window, "_form_id", "") or ""
        form_name = getattr(window, "_form_name", "") or form_id
        form_meta = _resolve_form_meta(form_id)
        if not _form_supports_drafts(form_meta):
            if not silent:
                messagebox.showinfo("Guardar", "Este formulario no tiene guardado disponible.")
            return False
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

        draft_source = str(source or "manual").strip().lower() or "manual"
        if draft_source == "autosave" and ui_section == "section_1":
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
        _update_window_save_status(window, now, section_id=ui_section, source="draft")
        self._refresh_drafts_badge()
        if toast_text:
            self.show_toast(toast_text)
        return True

    def _schedule_window_draft_autosave(self, window, delay_ms=250):
        if not window or not window.winfo_exists():
            return
        form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
        if form_id and not _form_supports_drafts(form_id):
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
                source="autosave",
            )

        try:
            window._draft_autosave_after_id = window.after(delay_ms, _run)
        except tk.TclError:
            window._draft_autosave_after_id = None

    def _install_form_autosave_bindings(self, window):
        form_id = getattr(window, "_form_id", "") or WINDOW_CLASS_FORM_ID_MAP.get(window.__class__.__name__, "")
        if form_id and not _form_supports_drafts(form_id):
            return
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
        if self._refresh_db_status_label:
            self._refresh_db_status_label.config(text="")

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
                    message = _log_user_error("database_refresh", err)
                    if self._refresh_db_status_label:
                        self._refresh_db_status_label.config(text=message, fg=COLOR_DANGER)
                    return
                self._companies_all = rows
                self._empresa_names_cache = sorted(
                    {(r.get("nombre_empresa") or "").strip() for r in rows if (r.get("nombre_empresa") or "").strip()},
                    key=str.lower,
                )
                self._render_companies()
                if self._refresh_db_status_label:
                    self._refresh_db_status_label.config(text="Base de datos actualizada ✓", fg=COLOR_SUCCESS)
                    self.after(3000, lambda: self._refresh_db_status_label.config(text=""))

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
            source="manual",
        )

    def _open_followup_local_draft_entry(self, draft):
        case_record = dict(draft.get("case_record") or {})
        case_path = str(draft.get("case_path") or "")
        case_target = case_record if case_record else case_path
        if not case_target:
            messagebox.showerror("Borradores", "El borrador de seguimientos no tiene un caso válido para restaurar.")
            return
        try:
            bootstrap = _load_followup_editor_bootstrap(case_target)
        except RuntimeError as exc:
            messagebox.showerror("Borradores", str(exc))
            return
        sheet_name = str(draft.get("sheet_name") or "").strip()
        if sheet_name:
            suggestion = dict((bootstrap or {}).get("suggestion") or {})
            suggestion["sheet"] = sheet_name
            suggestion["message"] = (
                f"Borrador local restaurado en "
                f"{_friendly_followup_sheet_title(sheet_name, (bootstrap or {}).get('workflow') or {})}."
            )
            bootstrap["suggestion"] = suggestion
        editor = SeguimientoEditorWindow(
            self,
            case_path=case_path,
            case_record=case_record,
            bootstrap=bootstrap,
        )
        _focus_window(editor)
        self.track_form_open("seguimientos", "Seguimientos")

    def _delete_followup_local_draft_entry(self, draft):
        draft_id = str(draft.get("draft_id") or "").strip()
        if not draft_id:
            return False
        deleted = _delete_followup_local_sheet_draft_by_id(
            draft_id,
            user_login=self._get_current_user_login(),
        )
        if deleted:
            self._refresh_drafts_badge()
        return deleted

    def _open_draft_entry(self, draft):
        if str(draft.get("draft_type") or "").strip() == "followup_local":
            self._open_followup_local_draft_entry(draft)
            return
        form_id = str(draft.get("form_id") or "")
        form_meta = next((item for item in get_forms() if item.get("id") == form_id), None)
        if not form_meta:
            messagebox.showerror("Borradores", "No se encontro el formulario en el HUB.")
            return
        if not _form_supports_drafts(form_meta):
            messagebox.showinfo(
                "Borradores",
                "Este formulario ya no admite borradores automaticos. Abre el caso desde el flujo principal.",
            )
            return
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
            if str(draft.get("draft_type") or "").strip() == "followup_local":
                if self._delete_followup_local_draft_entry(draft):
                    tree.delete(sel)
                    draft_by_iid.pop(sel, None)
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

        ttk.Button(actions, text="Cerrar", style="Secondary.TButton", command=modal.destroy).pack(side="left")
        ttk.Button(actions, text="Eliminar", style="DangerOutline.TButton", command=_delete_selected).pack(side="right", padx=(0, 8))
        ttk.Button(actions, text="Abrir", style="Primary.TButton", command=_open_selected).pack(side="right", padx=(0, 8))
        tree.bind("<Double-1>", lambda _e: _open_selected())
        modal.bind("<Escape>", lambda _event: modal.destroy())
        modal.bind("<Return>", lambda _event: _open_selected())

    def _open_completed_entry(self, completed_entry):
        if not isinstance(completed_entry, dict):
            messagebox.showerror("Terminados", "El registro seleccionado no es valido.")
            return
        form_meta = _resolve_completed_restore_form_meta(completed_entry)
        if not form_meta:
            messagebox.showerror(
                "Terminados",
                "Este formulario terminado ya no se puede reabrir desde el flujo normal.",
            )
            return
        payload_raw = _extract_completed_payload_raw(completed_entry)
        cache_snapshot = payload_raw.get("cache_snapshot")
        if not isinstance(cache_snapshot, dict) or not cache_snapshot:
            messagebox.showerror("Terminados", "El registro terminado no tiene cache para restaurar.")
            return
        self._pending_draft_restore = {
            "form_id": str(form_meta.get("id") or ""),
            "draft_id": "",
            "draft_session_key": uuid.uuid4().hex,
            "cache": copy.deepcopy(cache_snapshot),
            "ui_section": "section_1",
            "ui_snapshot": [],
        }
        window = self._open_form(form_meta)
        if not window:
            self._pending_draft_restore = None
            return
        window._draft_id = ""
        window._draft_session_key = uuid.uuid4().hex

    def _open_completed_window(self):
        completed = self._get_user_completed_forms()
        modal = tk.Toplevel(self)
        modal.title("Terminados - Últimos 30 días")
        modal.configure(bg=COLOR_LIGHT_BG)
        modal.geometry("920x420")
        modal.transient(self)
        modal.grab_set()

        frame = tk.Frame(modal, bg=COLOR_LIGHT_BG, padx=14, pady=12)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text="Formularios terminados",
            font=("Arial", 12, "bold"),
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 8))

        tk.Label(
            frame,
            text="Solo se conservan localmente por 30 dias y se reabren en el flujo normal con datos precargados.",
            font=("Arial", 9),
            fg="#555555",
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(0, 8))

        box = tk.Frame(frame, bg="white", bd=1, relief="solid")
        box.pack(fill="both", expand=True)
        yscroll = tk.Scrollbar(box, orient="vertical", width=SCROLLBAR_WIDTH)
        yscroll.pack(side="right", fill="y")

        tree = ttk.Treeview(
            box,
            columns=("form", "empresa", "finalizado", "estado"),
            show="headings",
            yscrollcommand=yscroll.set,
        )
        tree.heading("form", text="Formulario")
        tree.heading("empresa", text="Empresa")
        tree.heading("finalizado", text="Finalizado")
        tree.heading("estado", text="Estado")
        tree.column("form", width=260, anchor="w")
        tree.column("empresa", width=300, anchor="w")
        tree.column("finalizado", width=180, anchor="w")
        tree.column("estado", width=120, anchor="w")
        tree.pack(side="left", fill="both", expand=True)
        yscroll.config(command=tree.yview)

        entry_by_iid = {}
        for idx, item in enumerate(completed, start=1):
            iid = f"completed_{idx}"
            entry_by_iid[iid] = item
            tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    str(item.get("form_name") or item.get("form_id") or ""),
                    str(item.get("company_name") or "Sin empresa"),
                    str(item.get("finalizado_at_colombia") or item.get("finalizado_at_iso") or ""),
                    str(item.get("upload_status") or ""),
                ),
            )
        if not entry_by_iid:
            tree.insert("", "end", iid="__empty__", values=("-", "No hay formularios terminados recientes.", "-", "-"))

        actions = tk.Frame(frame, bg=COLOR_LIGHT_BG)
        actions.pack(fill="x", pady=(8, 0))

        def _open_selected():
            sel = tree.focus()
            if not sel or sel == "__empty__":
                return
            entry = entry_by_iid.get(sel)
            if not entry:
                return
            modal.destroy()
            self._open_completed_entry(entry)

        ttk.Button(actions, text="Cerrar", style="Secondary.TButton", command=modal.destroy).pack(side="left")
        ttk.Button(actions, text="Abrir", style="Primary.TButton", command=_open_selected).pack(side="right", padx=(0, 8))
        tree.bind("<Double-1>", lambda _e: _open_selected())
        modal.bind("<Escape>", lambda _event: modal.destroy())
        modal.bind("<Return>", lambda _event: _open_selected())

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
        right.grid_rowconfigure(3, weight=1)
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
            style="Secondary.TButton",
            command=self._refresh_database_cache,
        )
        self._refresh_db_btn.pack(side="right", padx=(0, 8))
        self._refresh_db_status_label = tk.Label(
            left_header,
            text="",
            font=("Arial", 9, "bold"),
            fg=COLOR_SUCCESS,
            bg=COLOR_LIGHT_BG,
        )
        self._refresh_db_status_label.pack(side="right", padx=(0, 8))
        self._completed_btn = ttk.Button(
            left_header,
            text="Terminados",
            style="Secondary.TButton",
            command=self._open_completed_window,
        )
        self._completed_btn.pack(side="right", padx=(0, 8))
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

        forms = [form for form in get_forms() if not bool(form.get("hidden"))]
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
                description_text = str(form.get("hub_description") or "").strip()
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
                title.grid(
                    row=0,
                    column=0,
                    sticky="ew",
                    padx=(12, 8),
                    pady=((10, 2) if description_text else 8),
                )
                if description_text:
                    description = tk.Label(
                        card,
                        text=description_text,
                        font=("Arial", 9),
                        bg="white",
                        fg="#666666",
                        anchor="w",
                        justify="left",
                        wraplength=460,
                    )
                    description.grid(row=1, column=0, sticky="ew", padx=(12, 8), pady=(0, 10))
                action = ttk.Button(
                    card,
                    text="Abrir",
                    style="Primary.TButton",
                    command=lambda f=form: self._open_form(f),
                )
                action.grid(
                    row=0,
                    column=1,
                    rowspan=(2 if description_text else 1),
                    sticky="e",
                    padx=12,
                    pady=8,
                )
                self._form_action_buttons[str(form.get("id") or "")] = action
                if bool((self._open_form_windows or {}).get(str(form.get("id") or ""))):
                    self._set_form_card_state(form.get("id"), active=True)
                widgets_to_bind = [card, title, action]
                if description_text:
                    widgets_to_bind.append(description)
                for _w in widgets_to_bind:
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
        tk.Label(
            right,
            text="Doble clic para editar la empresa seleccionada.",
            font=("Arial", 9),
            fg="#666666",
            bg=COLOR_LIGHT_BG,
        ).grid(row=1, column=0, sticky="w", pady=(0, 6))

        controls = tk.Frame(right, bg=COLOR_LIGHT_BG)
        controls.grid(row=2, column=0, sticky="ew", pady=(0, 8))
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
        box.grid(row=3, column=0, sticky="nsew")

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
            messagebox.showwarning("Empresas", _log_user_error("database_refresh", exc))
        self._empresa_names_cache = sorted(
            {(r.get("nombre_empresa") or "").strip() for r in self._companies_all if (r.get("nombre_empresa") or "").strip()},
            key=str.lower,
        )
        self._companies_search_var.trace_add("write", self._render_companies)
        sort_combo.bind("<<ComboboxSelected>>", self._render_companies)
        tree.bind("<Double-1>", self._on_company_double_click)
        self._render_companies()

    def _bind_form_runtime(self, window, form_meta):
        if not window or not form_meta:
            return
        form_id = str(form_meta.get("id") or "")
        form_name = str(form_meta.get("name") or form_id)
        supports_drafts = _form_supports_drafts(form_meta)
        window._form_id = form_id
        window._form_name = form_name
        window._form_meta = dict(form_meta or {})
        window._empresa_names_cache = getattr(self, "_empresa_names_cache", [])
        module = FORM_MODULE_MAP.get(form_id)
        if (
            supports_drafts
            and module
            and hasattr(module, "get_form_cache")
            and hasattr(module, "save_cache_to_file")
        ):
            window._save_draft_command = lambda w=window: self._save_current_form_draft(w)
        else:
            window._save_draft_command = None
        window._current_section = "section_1"
        window._draft_session_key = getattr(window, "_draft_session_key", "") or uuid.uuid4().hex
        window._draft_autosave_after_id = None
        window._draft_last_fingerprint = getattr(window, "_draft_last_fingerprint", "")
        window._skip_close_guard = False
        _ensure_form_save_status_label(window)
        if not getattr(window, "_usage_close_tracking_installed", False):
            window._usage_close_tracking_installed = True
            window._usage_finish_logged = False
            original_destroy = window.destroy

            def _track_usage_finish_once():
                if getattr(window, "_usage_finish_logged", False):
                    return
                window._usage_finish_logged = True
                _log_capture(f"[USAGE] form_window_closed form={form_id}")
                try:
                    self.track_form_finished(form_id)
                except Exception:
                    pass
                try:
                    self._release_form_window(form_id, window)
                except Exception:
                    pass

            def _tracked_destroy(*args, **kwargs):
                if not getattr(window, "_skip_close_guard", False):
                    if _guard_form_action(window, action_label="cerrar"):
                        return None
                _track_usage_finish_once()
                return original_destroy(*args, **kwargs)

            window._track_usage_finish_once = _track_usage_finish_once
            window._original_destroy = original_destroy
            window.destroy = _tracked_destroy
            try:
                window.protocol("WM_DELETE_WINDOW", window.destroy)
            except tk.TclError:
                pass
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
                        _ensure_wizard_runtime_widgets(window)
                    except Exception:
                        pass
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
                    _refresh_form_save_status(window)
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
        window._runtime_sections_ready = True
        pending_route_name = str(getattr(window, "_draft_restore_route_name", "") or "").strip()
        if pending_route_name:
            route = getattr(window, pending_route_name, None)
            if callable(route):
                try:
                    route()
                except Exception as exc:
                    _log_capture(
                        f"[DRAFT] restore_route_replay_failed form={form_id} route={pending_route_name} err={exc}"
                    )
            window._draft_restore_route_name = ""
        try:
            _ensure_wizard_runtime_widgets(window)
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
        _refresh_form_save_status(window)

    def _open_form(self, form_meta):
        form_id = str(form_meta.get("id") or "")
        if bool(form_meta.get("singleton_window")):
            current = (self._open_form_windows or {}).get(form_id)
            try:
                if current is not None and current.winfo_exists():
                    _focus_window(current)
                    return current
            except Exception:
                self._open_form_windows.pop(form_id, None)

        def _finalize_window(window):
            if window is None:
                return None
            self._bind_form_runtime(window, form_meta)
            if bool(form_meta.get("singleton_window")):
                self._register_open_form_window(form_id, window)
            _focus_window(window)
            self.track_form_open(form_meta["id"], form_meta["name"])
            return window

        if form_meta["id"] == "presentacion_programa":
            window = Section1Window(self)
            return _finalize_window(window)
        if form_meta["id"] == "evaluacion_accesibilidad":
            window = EvaluacionAccesibilidadWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "condiciones_vacante":
            window = CondicionesVacanteWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "condiciones_vacante_labs":
            window = CondicionesVacanteLabsWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "seleccion_incluyente":
            window = SeleccionIncluyenteWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "seleccion_incluyente_labs":
            window = SeleccionIncluyenteLabsWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "contratacion_incluyente":
            window = ContratacionIncluyenteWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "induccion_organizacional":
            window = InduccionOrganizacionalWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "induccion_operativa":
            window = InduccionOperativaWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "sensibilizacion":
            window = SensibilizacionWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "seguimientos":
            window = SeguimientosWindow(self)
            return _finalize_window(window)
        if form_meta["id"] == "interprete_lsc":
            window = LSCWindow(self)
            return _finalize_window(window)
        messagebox.showinfo("Formulario", f"Abrir formulario: {form_meta['name']}")
        return None

    def _open_labs_flow(self):
        if not _confirm_labs_experimental_warning(self):
            _log_labs("open_labs_flow cancelled_by_warning")
            return None
        selected_form_id = _select_labs_flow(self)
        if not selected_form_id:
            _log_labs("open_labs_flow cancelled_by_selector", level="WARN")
            return None
        _log_labs(f"open_labs_flow accepted form_id={selected_form_id}")
        return self._open_form(_resolve_form_meta(selected_form_id))

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

    def show_toast(self, text, duration_ms=None, *, level="info"):
        if duration_ms is None:
            duration_ms = _TOAST_DURATIONS.get(str(level or "info").lower(), 4000)
        self._ensure_toast()
        if self._toast_after_id is not None:
            self.after_cancel(self._toast_after_id)
            self._toast_after_id = None
        self._toast_label.config(text=text)
        self._toast_label.lift()
        self._toast_label.place(relx=1.0, rely=1.0, x=-24, y=-24, anchor="se")
        if duration_ms is not None:
            self._toast_after_id = self.after(duration_ms, self._hide_toast)

    def _toast_async(self, text, duration_ms=None, *, level="info"):
        self.after(0, lambda: self.show_toast(text, duration_ms, level=level))

        threading.Thread(target=_run, daemon=True).start()

# ── VENTANA: EvaluacionAccesibilidadWindow ───────────────────────────────────


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
                "section_2_1": self._show_section_2,
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
        if evaluacion_accesibilidad.cache_file_exists():
            _clear_local_resume_state(evaluacion_accesibilidad)
        return False


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
        _fn = getattr(self, '_pending_autosave', None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
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

    def _create_detail_text_widget(self, parent, *, width=80, min_lines=2, max_lines=10):
        detail = tk.Text(parent, width=width, height=min_lines, wrap="word")
        _attach_autoexpand(detail, min_lines, max_lines)
        return detail

    def _show_section_1(self):
        self._clear_section_container()
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._build_actions(content)
        _restore_section1_cached_state(self, evaluacion_accesibilidad)

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
                obs = tk.Text(row, width=80, height=2, wrap="word")
                obs.grid(row=2, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                _attach_autoexpand(obs, 2, 6)
                self.section2_1_fields[field_id]["observaciones"] = obs

            elif question["type"] == "texto":
                detail = self._create_detail_text_widget(row)
                detail.grid(row=1, column=0, columnspan=4, sticky="we", padx=8, pady=6)
                self.section2_1_fields[field_id]["texto"] = detail
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
        self._pending_autosave = lambda f=self.section2_1_fields: _autosave_section(evaluacion_accesibilidad, "section_2_1", lambda: self._collect_section_fields(f))
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
                obs = tk.Text(row, width=80, height=2, wrap="word")
                obs.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                _attach_autoexpand(obs, 2, 6)
                self.section2_2_fields[field_id]["observaciones"] = obs

            elif question["type"] == "texto":
                tk.Label(
                    row,
                    text="Detalle",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                detail = self._create_detail_text_widget(row)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                self.section2_2_fields[field_id]["texto"] = detail

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
                    obs = tk.Text(row, width=80, height=2, wrap="word")
                    obs.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                    _attach_autoexpand(obs, 2, 6)
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
        self._pending_autosave = lambda f=self.section2_2_fields: _autosave_section(evaluacion_accesibilidad, "section_2_2", lambda: self._collect_section_fields(f))
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
                obs = tk.Text(row, width=80, height=2, wrap="word")
                obs.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                _attach_autoexpand(obs, 2, 6)
                self.section2_3_fields[field_id]["observaciones"] = obs

            elif question["type"] == "texto":
                tk.Label(
                    row,
                    text="Detalle",
                    font=("Arial", 9, "bold"),
                    bg="white",
                ).grid(row=current_row, column=0, sticky="w", padx=8, pady=4)
                detail = self._create_detail_text_widget(row)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                self.section2_3_fields[field_id]["texto"] = detail

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

        self._prefill_section_fields("section_2_3", self.section2_3_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section2_3_fields: _autosave_section(evaluacion_accesibilidad, "section_2_3", lambda: self._collect_section_fields(f))
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
                obs = tk.Text(row, width=80, height=2, wrap="word")
                obs.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                _attach_autoexpand(obs, 2, 6)
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
                detail = self._create_detail_text_widget(row)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                self.section2_4_fields[field_id]["detalle"] = detail

        self._prefill_section_fields("section_2_4", self.section2_4_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section2_4_fields: _autosave_section(evaluacion_accesibilidad, "section_2_4", lambda: self._collect_section_fields(f))
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
                obs = tk.Text(row, width=80, height=2, wrap="word")
                obs.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                _attach_autoexpand(obs, 2, 6)
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
                detail = self._create_detail_text_widget(row)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                self.section2_5_fields[field_id]["detalle"] = detail

        self._prefill_section_fields("section_2_5", self.section2_5_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section2_5_fields: _autosave_section(evaluacion_accesibilidad, "section_2_5", lambda: self._collect_section_fields(f))
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
                detail = self._create_detail_text_widget(row)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                self.section2_6_fields[field_id]["detalle"] = detail

        self._prefill_section_fields("section_2_6", self.section2_6_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section2_6_fields: _autosave_section(evaluacion_accesibilidad, "section_2_6", lambda: self._collect_section_fields(f))
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
                obs = tk.Text(row, width=80, height=2, wrap="word")
                obs.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                _attach_autoexpand(obs, 2, 6)
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
                detail = self._create_detail_text_widget(row)
                detail.grid(row=current_row, column=1, columnspan=3, sticky="we", padx=4, pady=4)
                self.section3_fields[field_id]["detalle"] = detail

        self._prefill_section_fields("section_3", self.section3_fields)

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section3_fields: _autosave_section(evaluacion_accesibilidad, "section_3", lambda: self._collect_section_fields(f))
        ttk.Button(actions, text="Regresar", command=self._show_section_2_6).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_3).pack(side="right")
    def _confirm_section_2_5(self):
        payload = self._collect_section_fields(self.section2_5_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_2_6()

    def _confirm_section_2_6(self):
        payload = self._collect_section_fields(self.section2_6_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        self._pending_autosave = lambda: _autosave_section(
            evaluacion_accesibilidad,
            "section_4",
            self._collect_section4_payload,
        )
        ttk.Button(actions, text="Regresar", command=self._show_section_3).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_4).pack(side="right")
    def _confirm_section_3(self):
        payload = self._collect_section_fields(self.section3_fields)
        try:
            evaluacion_accesibilidad.confirm_section_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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


    def _collect_section4_payload(self):
        nivel = self.section4_level_var.get().strip()
        descripcion = evaluacion_accesibilidad.SECTION_4["descriptions"].get(nivel, "")
        return {
            "nivel_accesibilidad": nivel,
            "descripcion": descripcion,
        }


    def _confirm_section_4(self):
        payload = self._collect_section4_payload()
        try:
            evaluacion_accesibilidad.confirm_section_4(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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

            suggested_label = tk.Label(
                row,
                text="Ajustes sugeridos:",
                font=("Arial", 9, "bold"),
                bg="white",
            )
            suggested_label.grid(row=2, column=0, sticky="w", padx=8, pady=(6, 2))
            suggested_value = tk.Label(
                row,
                text=self._clean_text(item["ajustes"]),
                font=("Arial", 9),
                bg="white",
                fg="#333333",
                wraplength=760,
                justify="left",
            )
            suggested_value.grid(row=3, column=0, columnspan=3, sticky="w", padx=8)

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
                "_suggested_label": suggested_label,
                "_suggested_value": suggested_value,
            }

            def _toggle_suggested_widgets(_event=None, *, combo=aplica_combo, lbl=suggested_label, value_lbl=suggested_value):
                show_suggested = combo.get().strip() == "Aplica"
                if show_suggested:
                    lbl.grid()
                    value_lbl.grid()
                else:
                    lbl.grid_remove()
                    value_lbl.grid_remove()

            aplica_combo.bind("<<ComboboxSelected>>", _toggle_suggested_widgets, add="+")
            _toggle_suggested_widgets()

        self._prefill_section_fields("section_5", self.section5_fields)
        for widgets in self.section5_fields.values():
            combo = widgets.get("lista")
            suggested_label = widgets.get("_suggested_label")
            suggested_value = widgets.get("_suggested_value")
            if not combo or not suggested_label or not suggested_value:
                continue
            show_suggested = combo.get().strip() == "Aplica"
            if show_suggested:
                suggested_label.grid()
                suggested_value.grid()
            else:
                suggested_label.grid_remove()
                suggested_value.grid_remove()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section5_fields: _autosave_section(
            evaluacion_accesibilidad,
            "section_5",
            lambda: self._collect_evaluacion_section5_payload(f),
        )
        ttk.Button(actions, text="Regresar", command=self._show_section_4).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_5).pack(side="right")

    def _collect_evaluacion_section5_payload(self, fields=None):
        payload = self._collect_section_fields(fields or self.section5_fields)
        for item in evaluacion_accesibilidad.SECTION_5.get("items", []):
            field_id = item.get("id")
            if not field_id:
                continue
            payload[f"{field_id}_ajustes"] = (
                item.get("ajustes", "")
                if payload.get(field_id) == "Aplica"
                else "No aplica"
            )
        return payload

    def _confirm_section_5(self):
        payload = self._collect_evaluacion_section5_payload(self.section5_fields)
        try:
            evaluacion_accesibilidad.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            _attach_autoexpand(text_box, 4, 15)
            self.section6_fields[field["id"]] = {"texto": text_box}

        self._prefill_section_fields("section_6", self.section6_fields)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section6_fields: _autosave_section(evaluacion_accesibilidad, "section_6", lambda: self._collect_section_fields(f))
        ttk.Button(actions, text="Regresar", command=self._show_section_5).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_6).pack(side="right")

    def _confirm_section_6(self):
        payload = self._collect_section_fields(self.section6_fields)
        try:
            evaluacion_accesibilidad.confirm_section_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            _attach_autoexpand(text_box, 4, 15)
            self.section7_fields[field["id"]] = {"texto": text_box}

        self._prefill_section_fields("section_7", self.section7_fields)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section7_fields: _autosave_section(evaluacion_accesibilidad, "section_7", lambda: self._collect_section_fields(f))
        ttk.Button(actions, text="Regresar", command=self._show_section_6).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_7).pack(side="right")

    def _confirm_section_7(self):
        payload = self._collect_section_fields(self.section7_fields)
        try:
            evaluacion_accesibilidad.confirm_section_7(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        self._pending_autosave = lambda: _autosave_section(evaluacion_accesibilidad, "section_8", lambda: [{"nombre": n.get().strip(), "cargo": c.get().strip()} for n, c in self.section8_entries])
        ttk.Button(actions, text="Regresar", command=self._show_section_7).pack(side="left")
        ttk.Button(actions, text="📞 Solicitar Intérprete LSC", command=self._open_lsc_window).pack(side="left", padx=(8, 0))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        _queue_or_run_main_form_export(self, self._export_form)

    def _export_form(self):
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

            return _raise_finalize_stage(
                "preparando el acta",
                lambda: evaluacion_accesibilidad.export_to_excel(progress_callback=_on_progress),
            )

        _start_background_finalization(
            self,
            loading,
            form_name="Evaluacion Accesibilidad",
            company_name=company_name,
            form_id="evaluacion_accesibilidad",
            worker_fn=_worker,
        )

    def _open_lsc_window(self):
        ctx = _build_lsc_context(
            self,
            module=evaluacion_accesibilidad,
            source_form="evaluacion_accesibilidad",
        )
        _launch_linked_lsc_window(
            self,
            context=ctx,
            return_to_final_section=self._show_section_8,
            main_finish_action=self._confirm_section_8,
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
                if str(key).startswith("_"):
                    continue
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
                if str(key).startswith("_"):
                    continue
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_2_3()

    def _confirm_section_2_3(self):
        payload = self._collect_section_fields(self.section2_3_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_2_4()

    def _confirm_section_2_4(self):
        payload = self._collect_section_fields(self.section2_4_fields)
        try:
            evaluacion_accesibilidad.confirm_section_2_4(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_2_5()

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        modalidad_options = next(
            (
                list(field.get("options") or [])
                for field in evaluacion_accesibilidad.SECTION_1.get("fields", [])
                if field.get("id") == "modalidad"
            ),
            ["Virtual", "Presencial", "Mixto", "No aplica"],
        )
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
            modalidad_options=modalidad_options,
            modalidad_aliases={"Mixta": "Mixto"},
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
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(
            self,
            mode=mode,
            lookup=evaluacion_accesibilidad,
            button=target_button,
        )

    def _confirm_and_continue(self):
        _confirm_section1_and_continue(
            self,
            confirm_fn=evaluacion_accesibilidad.confirm_section_1,
            next_step=self._show_section_2,
        )

# ── VENTANA: CondicionesVacanteWindow ────────────────────────────────────────


class CondicionesVacanteWindow(tk.Toplevel, FormMousewheelMixin):
    FORM_META_ID = "condiciones_vacante"
    WINDOW_TITLE = "Condiciones de Vacante - Seccion 1"

    def __init__(self, parent):
        super().__init__(parent)
        self.title(self.WINDOW_TITLE)
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
        _fn = getattr(self, '_pending_autosave', None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
        for child in self.section_container.winfo_children():
            child.destroy()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            self.FORM_META_ID,
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
        if condiciones_vacante.cache_file_exists():
            _clear_local_resume_state(condiciones_vacante)
        return False

    def _build_condiciones_section2_voice_banner(self, parent):
        return None

    def _get_condiciones_voice_ui_namespace(self):
        return "Vacante"

    def _show_section_1(self):
        self._clear_section_container()
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        self._build_actions(content)
        _restore_section1_cached_state(self, condiciones_vacante)

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
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(
            self,
            mode=mode,
            lookup=condiciones_vacante,
            button=target_button,
        )

    def _confirm_and_continue(self):
        _confirm_section1_and_continue(
            self,
            confirm_fn=condiciones_vacante.confirm_section_1,
            next_step=self._show_section_2,
        )

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_2["title"])
        self.header_subtitle.config(text="Completa caracteristicas de la vacante.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_condiciones_section2_voice_banner(content)

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
                _attach_autoexpand(obs_text, 3, 12)
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
        self._pending_autosave = lambda f=self.section2_fields: _autosave_section(condiciones_vacante, "section_2", lambda: _collect_flat_fields(f))
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

    def _apply_condiciones_section2_updates(self, updates):
        for field_id, value in (updates or {}).items():
            if value in (None, ""):
                continue
            widget = self.section2_fields.get(field_id)
            if widget is None:
                continue
            if isinstance(widget, ttk.Combobox):
                widget.set(str(value))
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", str(value))
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(value))
        nivel = str((updates or {}).get("nivel_cargo") or "").strip()
        if nivel:
            self._populate_competencias(nivel)

    def _on_condiciones_section2_dictation(self):
        subsection_key = "section_2_vacancy"
        section_label = "2. Caracteristicas de la vacante"
        _log_labs(f"vacancy_dictation_open subsection={subsection_key}")
        dialog = LabsSection2VoiceDialog(
            self,
            subsection_key=subsection_key,
            section_label=section_label,
            candidate_index=1,
            form_id="condiciones_vacante",
            session_provider=lambda: _supabase_get_access_token(".env"),
            spec=get_vacancy_section2_spec(subsection_key),
            function_name=VACANCY_SECTION2_VOICE_FUNCTION,
            section_id="section_2",
            record_label="Vacante",
            ui_namespace=self._get_condiciones_voice_ui_namespace(),
        )
        response = dialog.show()
        if not response:
            _log_labs(
                f"vacancy_dictation_cancelled subsection={subsection_key}",
                level="WARN",
            )
            return

        extraction = response.get("extraction") or {}
        transcription = str(response.get("transcription") or "").strip()
        try:
            processed = postprocess_vacancy_section2_extraction(
                extraction,
                subsection_key=subsection_key,
                candidate_index=1,
            )
        except Exception as exc:
            _log_labs(
                f"vacancy_dictation_invalid_structured_response subsection={subsection_key} detail={exc}",
                level="ERROR",
            )
            messagebox.showerror(
                "Vacante",
                f"No fue posible interpretar la respuesta estructurada.\n\nDetalle: {exc}",
                parent=self,
            )
            return

        updates = processed.get("candidate") or {}
        preview_lines = summarize_vacancy_section2_updates(updates, subsection_key=subsection_key)
        warnings = []
        for item in list(response.get("warnings") or []) + list(processed.get("warnings") or []):
            text = str(item or "").strip()
            if text and text not in warnings:
                warnings.append(text)

        preview = LabsSection2PreviewDialog(
            self,
            section_label=section_label,
            candidate_index=1,
            transcription=transcription,
            preview_lines=preview_lines,
            warnings=warnings,
            record_label="Vacante",
            ui_namespace=self._get_condiciones_voice_ui_namespace(),
        )
        if not preview.show():
            _log_labs(
                f"vacancy_preview_cancelled subsection={subsection_key} fields={','.join(sorted(updates.keys()))}",
                level="WARN",
            )
            return
        self._apply_condiciones_section2_updates(updates)
        _log_labs(
            f"vacancy_preview_applied subsection={subsection_key} fields={','.join(sorted(updates.keys()))} "
            f"warnings={len(warnings)}"
        )

    def _build_condiciones_section2_1_voice_banner(self, parent):
        return None

    def _apply_condiciones_section2_1_updates(self, updates):
        for field_id, value in (updates or {}).items():
            if value in (None, ""):
                continue
            widget = self.section2_1_fields.get(field_id)
            if widget is None:
                continue
            if isinstance(widget, ttk.Combobox):
                widget.set(str(value))
            elif isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", str(value))
            elif isinstance(widget, tk.BooleanVar):
                continue
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(value))

    def _on_condiciones_section2_1_dictation(self):
        subsection_key = "section_2_1_schedule_experience"
        section_label = "2.1 Horarios, experiencia, funciones y herramientas"
        _log_labs(f"vacancy_dictation_open subsection={subsection_key}")
        dialog = LabsSection2VoiceDialog(
            self,
            subsection_key=subsection_key,
            section_label=section_label,
            candidate_index=1,
            form_id="condiciones_vacante",
            session_provider=lambda: _supabase_get_access_token(".env"),
            spec=get_vacancy_section2_spec(subsection_key),
            function_name=VACANCY_SECTION2_VOICE_FUNCTION,
            section_id="section_2_1",
            record_label="Vacante",
            ui_namespace=self._get_condiciones_voice_ui_namespace(),
        )
        response = dialog.show()
        if not response:
            _log_labs(
                f"vacancy_dictation_cancelled subsection={subsection_key}",
                level="WARN",
            )
            return

        extraction = response.get("extraction") or {}
        transcription = str(response.get("transcription") or "").strip()
        try:
            processed = postprocess_vacancy_section2_extraction(
                extraction,
                subsection_key=subsection_key,
                candidate_index=1,
            )
        except Exception as exc:
            _log_labs(
                f"vacancy_dictation_invalid_structured_response subsection={subsection_key} detail={exc}",
                level="ERROR",
            )
            messagebox.showerror(
                "Vacante",
                f"No fue posible interpretar la respuesta estructurada.\n\nDetalle: {exc}",
                parent=self,
            )
            return

        updates = processed.get("candidate") or {}
        preview_lines = summarize_vacancy_section2_updates(updates, subsection_key=subsection_key)
        warnings = []
        for item in list(response.get("warnings") or []) + list(processed.get("warnings") or []):
            text = str(item or "").strip()
            if text and text not in warnings:
                warnings.append(text)

        preview = LabsSection2PreviewDialog(
            self,
            section_label=section_label,
            candidate_index=1,
            transcription=transcription,
            preview_lines=preview_lines,
            warnings=warnings,
            record_label="Vacante",
            ui_namespace=self._get_condiciones_voice_ui_namespace(),
        )
        if not preview.show():
            _log_labs(
                f"vacancy_preview_cancelled subsection={subsection_key} fields={','.join(sorted(updates.keys()))}",
                level="WARN",
            )
            return
        self._apply_condiciones_section2_1_updates(updates)
        _log_labs(
            f"vacancy_preview_applied subsection={subsection_key} fields={','.join(sorted(updates.keys()))} "
            f"warnings={len(warnings)}"
        )

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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_2_1()

    def _show_section_2_1(self):
        self._clear_section_container()
        self.header_title.config(text=condiciones_vacante.SECTION_2_1["title"])
        self.header_subtitle.config(text="Completa formacion academica.")
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_condiciones_section2_1_voice_banner(content)

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
                _attach_autoexpand(widget, 3, 12)
            else:
                widget = tk.Entry(row, width=48)
                widget.grid(row=0, column=1, sticky="w", padx=8, pady=8)

            self.section2_1_fields[field["id"]] = widget

        self._prefill_section2_1_fields()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section2_1_fields: _autosave_section(condiciones_vacante, "section_2_1", lambda: _collect_flat_fields(f))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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

        for category in condiciones_vacante.SECTION_3["categories"]:
            category_frame = tk.Frame(content, bg=COLOR_LIGHT_BG)
            category_frame.pack(fill="x", pady=(8, 4))
            tk.Label(
                category_frame,
                text=category["title"],
                font=FONT_SECTION,
                fg=COLOR_PURPLE,
                bg=COLOR_LIGHT_BG,
            ).pack(anchor="w")
            self._build_condiciones_category_actions(
                category_frame,
                self.section3_fields,
                [field_id for field_id, _label in category["items"]],
            )

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
            _attach_autoexpand(obs_text, 3, 12)
            self.section3_fields[category["observaciones_id"]] = obs_text

        self._prefill_section3_fields()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section3_fields: _autosave_section(condiciones_vacante, "section_3", lambda: _collect_flat_fields(f))
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

    def _build_condiciones_category_actions(self, parent, fields, field_ids):
        actions = tk.Frame(parent, bg=COLOR_LIGHT_BG)
        actions.pack(anchor="w", pady=(4, 0))
        action_specs = [
            ("Todo alto", "Alto."),
            ("Todo medio", "Medio."),
            ("Todo bajo", "Bajo."),
            ("Todo no aplica", "No aplica"),
        ]
        for label, value in action_specs:
            ttk.Button(
                actions,
                text=label,
                command=lambda selected=value, ids=tuple(field_ids): self._set_condiciones_habilidades_nivel(
                    fields,
                    selected,
                    ids,
                ),
            ).pack(side="left", padx=(0, 8))

    def _set_section3_habilidades_nivel(self, value, field_ids=None):
        self._set_condiciones_habilidades_nivel(self.section3_fields, value, field_ids)

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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        self._pending_autosave = lambda f=self.section4_fields: _autosave_section(condiciones_vacante, "section_4", lambda: {field_id: widget.get().strip() for field_id, widget in f.items()})
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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

        for category in condiciones_vacante.SECTION_5["categories"]:
            category_frame = tk.Frame(content, bg=COLOR_LIGHT_BG)
            category_frame.pack(fill="x", pady=(8, 4))
            tk.Label(
                category_frame,
                text=category["title"],
                font=FONT_SECTION,
                fg=COLOR_PURPLE,
                bg=COLOR_LIGHT_BG,
            ).pack(anchor="w")
            self._build_condiciones_category_actions(
                category_frame,
                self.section5_fields,
                [item[0] for item in category["items"]],
            )

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
        _attach_autoexpand(obs_text, 4, 15)
        self.section5_fields[condiciones_vacante.SECTION_5["observaciones"]["id"]] = obs_text

        self._prefill_section5_fields()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda f=self.section5_fields: _autosave_section(condiciones_vacante, "section_5", lambda: _collect_flat_fields(f))
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

    def _set_section5_habilidades_nivel(self, value, field_ids=None):
        self._set_condiciones_habilidades_nivel(self.section5_fields, value, field_ids)

    def _set_condiciones_habilidades_nivel(self, fields, value, field_ids=None):
        allowed_ids = set(field_ids or []) if field_ids else None
        for field_id, widget in fields.items():
            if allowed_ids is not None and field_id not in allowed_ids:
                continue
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        self.section6_add_btn = None
        self.section6_remove_btn = None
        self.section6_container = tk.Frame(content, bg=COLOR_LIGHT_BG)
        self.section6_container.pack(fill="x")
        self.disability_options = condiciones_vacante.SECTION_6["options"]
        self.disability_descriptions = condiciones_vacante.get_disability_descriptions()
        self.section6_base_rows = condiciones_vacante.SECTION_6.get("base_rows", 4)
        for _ in range(self.section6_base_rows):
            self._add_disability_row()

        row_actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        row_actions.pack(anchor="w", pady=(6, 8))
        self.section6_add_btn = ttk.Button(
            row_actions,
            text="Agregar discapacidad",
            command=self._add_disability_row,
        )
        self.section6_add_btn.pack(side="left")
        self.section6_remove_btn = ttk.Button(
            row_actions,
            text="Eliminar última discapacidad",
            command=self._remove_last_disability_row,
        )
        self.section6_remove_btn.pack(side="left", padx=(8, 0))

        cached_rows = condiciones_vacante.get_form_cache().get("section_6", [])
        if cached_rows:
            for idx, entry in enumerate(cached_rows):
                if idx >= len(self.section6_rows):
                    self._add_disability_row()
                self._set_disability_row(self.section6_rows[idx], entry)
        self._update_section6_row_actions()

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda: _autosave_section(condiciones_vacante, "section_6", lambda: [{"discapacidad": r["combo"].get().strip(), "descripcion": r["descripcion"].get("1.0", tk.END).strip()} for r in self.section6_rows if r["combo"].get().strip() or r["descripcion"].get("1.0", tk.END).strip()])
        ttk.Button(actions, text="Regresar", command=self._show_section_5).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_6).pack(side="right")

    def _add_disability_row(self):
        row = tk.Frame(self.section6_container, bg="white", bd=1, relief="solid")
        row.pack(fill="x", pady=4)
        row.grid_columnconfigure(0, weight=2, minsize=360)
        row.grid_columnconfigure(1, weight=3)

        combo = ttk.Combobox(
            row,
            values=self.disability_options,
            state="readonly",
            width=48,
        )
        combo.grid(row=0, column=0, sticky="ew", padx=8, pady=6)

        descripcion = tk.Text(row, height=3, wrap="word", width=50, state="disabled")
        descripcion.grid(row=0, column=1, sticky="we", padx=8, pady=6)

        row_entry = {
            "frame": row,
            "combo": combo,
            "descripcion": descripcion,
        }
        self.section6_rows.append(row_entry)

        combo.bind(
            "<<ComboboxSelected>>",
            lambda _event, target=row_entry: self._update_disability_description(target),
        )
        self._update_section6_row_actions()

        return row_entry

    def _remove_last_disability_row(self):
        if len(self.section6_rows) <= getattr(self, "section6_base_rows", 4):
            return
        row_entry = self.section6_rows.pop()
        row_entry["frame"].destroy()
        self._update_section6_row_actions()

    def _update_section6_row_actions(self):
        button = getattr(self, "section6_remove_btn", None)
        if not button:
            return
        state = "normal" if len(self.section6_rows) > getattr(self, "section6_base_rows", 4) else "disabled"
        try:
            if not button.winfo_exists():
                self.section6_remove_btn = None
                return
            button.configure(state=state)
        except (tk.TclError, Exception):
            self.section6_remove_btn = None
            return

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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        _attach_autoexpand(self.section7_text, 8, 25)

        cached = condiciones_vacante.get_form_cache().get("section_7", {})
        cached_text = cached.get(condiciones_vacante.SECTION_7["field_id"])
        if cached_text:
            self.section7_text.delete("1.0", tk.END)
            self.section7_text.insert("1.0", cached_text)

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda: _autosave_section(condiciones_vacante, "section_7", lambda: {condiciones_vacante.SECTION_7["field_id"]: self.section7_text.get("1.0", tk.END).strip()})
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        self._pending_autosave = lambda: _autosave_section(
            condiciones_vacante,
            "section_8",
            lambda: self._get_section8_asistentes_values(),
        )
        ttk.Button(actions, text="Regresar", command=self._show_section_7).pack(side="left")
        ttk.Button(actions, text="📞 Solicitar Intérprete LSC", command=self._open_lsc_window).pack(side="left", padx=(8, 0))
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
        # Mínimo: al menos 1 asistente + 1 asesor agencia (última fila siempre)
        min_rows = max(condiciones_vacante.SECTION_8.get("rows", 3), 2)
        while len(rows) < min_rows:
            rows.append({"nombre": "", "cargo": ""})

        for nombre_entry, cargo_entry in self.section8_rows:
            nombre_entry.destroy()
            cargo_entry.destroy()
        self.section8_rows = []

        last_idx = len(rows) - 1
        for idx, entry in enumerate(rows):
            row_idx = idx + 1
            is_last = idx == last_idx
            if is_last:
                nombre_entry, cargo_entry = _create_asesor_agencia_inputs(
                    self.section8_table,
                    ENTRY_W_WIDE,
                    catalog=getattr(self, "_asesores_agencia_catalog", None),
                )
            elif idx == 0:
                nombre_entry, cargo_entry = _create_asistente_inputs(
                    self.section8_table,
                    ENTRY_W_WIDE,
                    use_catalog=True,
                    catalog=_get_asistentes_profesionales_catalog(),
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
            # Inserta fila vacía antes del último (asesor agencia), preservando sus datos
            rows.insert(len(rows) - 1, {"nombre": "", "cargo": ""})
        else:
            rows = [{"nombre": "", "cargo": ""}]
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        _queue_or_run_main_form_export(self, self._export_form)

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando acta...")
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
            form_id=getattr(self, "_form_id", self.FORM_META_ID),
            worker_fn=lambda: _raise_finalize_stage(
                "preparando el acta",
                condiciones_vacante.export_to_excel,
            ),
        )

    def _open_lsc_window(self):
        ctx = _build_lsc_context(
            self,
            module=condiciones_vacante,
            source_form="condiciones_vacante",
        )
        _launch_linked_lsc_window(
            self,
            context=ctx,
            return_to_final_section=self._show_section_8,
            main_finish_action=self._confirm_section_8,
        )


class CondicionesVacanteLabsWindow(CondicionesVacanteWindow):
    FORM_META_ID = "condiciones_vacante_labs"
    WINDOW_TITLE = "Condiciones de Vacante Labs - Seccion 1"

    def _build_condiciones_section2_voice_banner(self, parent):
        voice_box = tk.Frame(parent, bg="#FFF4E5", bd=1, relief="solid")
        voice_box.pack(fill="x", pady=(0, 12))
        tk.Label(
            voice_box,
            text=(
                "Dictado experimental para esta seccion. Usa un solo audio y responde solo lo que este confirmado. "
                "Salario, edad, pruebas y firma de contrato pueden decirse en lenguaje natural."
            ),
            bg="#FFF4E5",
            fg="#7A4100",
            justify="left",
            wraplength=760,
            padx=12,
            pady=10,
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(
            voice_box,
            text="Dictar seccion 2",
            command=self._on_condiciones_section2_dictation,
        ).pack(side="right", padx=10, pady=10)
        return voice_box

    def _get_condiciones_voice_ui_namespace(self):
        return "Vacante Labs"

    def _build_condiciones_section2_1_voice_banner(self, parent):
        voice_box = tk.Frame(parent, bg="#FFF4E5", bd=1, relief="solid")
        voice_box.pack(fill="x", pady=(0, 12))
        tk.Label(
            voice_box,
            text=(
                "Dictado experimental para esta subseccion. Aqui solo cubre horarios, experiencia, funciones y herramientas. "
                "Los niveles educativos y la formacion academica siguen siendo manuales."
            ),
            bg="#FFF4E5",
            fg="#7A4100",
            justify="left",
            wraplength=760,
            padx=12,
            pady=10,
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(
            voice_box,
            text="Dictar 2.1",
            command=self._on_condiciones_section2_1_dictation,
        ).pack(side="right", padx=10, pady=10)
        return voice_box


# ── VENTANA: SeleccionIncluyenteWindow ───────────────────────────────────────


class SeleccionIncluyenteWindow(tk.Toplevel, FormMousewheelMixin):
    FORM_META_ID = "seleccion_incluyente"
    FORM_MODULE = seleccion_incluyente
    WINDOW_TITLE = "Seleccion Incluyente - Seccion 1"

    def __init__(self, parent):
        super().__init__(parent)
        self.title(self.WINDOW_TITLE)
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._seleccion_module = self.FORM_MODULE
        self._empresa_lookup = self._seleccion_module

        self.company_data = None
        self.fields = {}
        self.cedula_options = []

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    def _maybe_resume_form(self):
        module = self._seleccion_module
        if _consume_pending_draft_restore(
            self,
            self.FORM_META_ID,
            module,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_5": self._show_section_5,
                "section_6": self._show_section_6,
            },
            self._show_section_1,
        ):
            return True
        if module.cache_file_exists():
            _clear_local_resume_state(module)
        return False

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
        _fn = getattr(self, '_pending_autosave', None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
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
            self.cedula_options = self._seleccion_module.get_usuarios_reca_cedulas()
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
        if len(normalized) > 10:
            widget.delete(0, tk.END)
            return
        if normalized and normalized != cedula:
            widget.delete(0, tk.END)
            widget.insert(0, normalized)
        try:
            data = self._seleccion_module.get_usuario_reca_by_cedula(normalized)
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
        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self.continue_btn = ttk.Button(
            actions,
            text="Continuar",
            command=self._confirm_and_continue,
            state="disabled",
        )
        self.continue_btn.pack(side="right")
        _restore_section1_cached_state(self, self._seleccion_module)

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
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(
            self,
            mode=mode,
            lookup=self._seleccion_module,
            button=target_button,
        )

    def _confirm_and_continue(self):
        _confirm_section1_and_continue(
            self,
            confirm_fn=self._seleccion_module.confirm_section_1,
            next_step=self._show_section_2,
        )

    def _label_for_field(self, field_id):
        return getattr(self, "_section1_labels", {}).get(field_id, field_id)

    def _prefill_section_1(self):
        cache = self._seleccion_module.get_form_cache().get("section_1", {})
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

    def _labs_voice_ui_enabled(self):
        return getattr(self, "FORM_META_ID", "") == "seleccion_incluyente_labs"

    def _build_selection_subsection_shell(self, parent, title, subsection_key):
        if not self._labs_voice_ui_enabled():
            frame = tk.LabelFrame(
                parent,
                text=title,
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            frame.pack(fill="x", pady=(0, 8))
            return frame, frame, None, None

        frame = tk.Frame(
            parent,
            bg="white",
            bd=1,
            relief="solid",
            highlightbackground="#D9D2E3",
            highlightthickness=1,
        )
        frame.pack(fill="x", pady=(0, 10))
        header = tk.Frame(frame, bg="#F1EAF8")
        header.pack(fill="x")
        title_label = tk.Label(
            header,
            text=title,
            font=FONT_LABEL,
            bg="#F1EAF8",
            fg="#222222",
            anchor="w",
        )
        title_label.pack(side="left", padx=10, pady=8)
        body = tk.Frame(frame, bg="white")
        body.pack(fill="x", padx=10, pady=10)
        return frame, body, header, title_label

    def _install_labs_subsection_button(self, header, title_label, subsection_key, label, candidate_fields=None):
        if not self._labs_voice_ui_enabled() or not header or not title_label:
            return None
        button = ttk.Button(
            header,
            text="Dictar",
            command=lambda key=subsection_key, section_label=label: self._on_labs_subsection_dictation(
                key,
                section_label,
                candidate_fields,
            ),
        )
        button.pack(side="left", padx=(10, 0), pady=6)
        return button

    def _attach_labs_text_dictation(self, text_widget, header, title_label, subsection_key):
        if not self._labs_voice_ui_enabled() or not isinstance(text_widget, tk.Text):
            return
        try:
            attach_dictation(
                text_widget,
                form_id=getattr(self, "_form_id", self.FORM_META_ID),
                field_id=f"{subsection_key}:voice",
                session_provider=lambda: _supabase_get_access_token(".env"),
                log_fn=_log_capture,
                show_controls=True,
                controls_parent=header,
                anchor_widget=title_label,
                placement="inline_right",
            )
        except Exception as exc:
            _log_capture(
                f"[DICTATION] attach_labs_subsection_failed form={getattr(self, '_form_id', self.FORM_META_ID)} "
                f"section={subsection_key} err={exc}"
            )

    def _get_labs_candidate_index(self, candidate_fields):
        if not candidate_fields:
            return 1
        try:
            return int(self.oferente_blocks.index(candidate_fields)) + 1
        except Exception:
            return 1

    def _apply_labs_candidate_updates(self, candidate_fields, updates, subsection_key):
        if not isinstance(candidate_fields, dict):
            return
        for field_id, value in (updates or {}).items():
            if value in (None, ""):
                continue
            widget = candidate_fields.get(field_id)
            if widget is None:
                continue
            if isinstance(widget, ttk.Combobox):
                widget.set(str(value))
                continue
            if isinstance(widget, tk.Text):
                widget.delete("1.0", tk.END)
                widget.insert("1.0", str(value))
                continue
            try:
                state = str(widget.cget("state"))
            except Exception:
                state = ""
            if state == "readonly":
                _set_readonly_entry_value(widget, str(value))
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(value))

        fecha_widget = candidate_fields.get("fecha_nacimiento")
        edad_widget = candidate_fields.get("edad")
        if fecha_widget is not None and edad_widget is not None and updates.get("fecha_nacimiento"):
            try:
                self._format_birthdate(None, fecha_widget, edad_widget)
            except Exception:
                pass

        if subsection_key == "section_3_desarrollo":
            widget = candidate_fields.get("desarrollo_actividad")
            sync_fn = getattr(self, "_labs_section2_sync_desarrollo_widgets", None)
            if callable(sync_fn) and isinstance(widget, tk.Text):
                sync_fn(widget)

    def _on_labs_subsection_dictation(self, subsection_key, section_label, candidate_fields=None):
        candidate_index = self._get_labs_candidate_index(candidate_fields)
        _log_labs(
            f"subsection_dictation_open subsection={subsection_key} candidate_index={candidate_index}"
        )
        dialog = LabsSection2VoiceDialog(
            self,
            subsection_key=subsection_key,
            section_label=section_label,
            candidate_index=candidate_index,
            form_id=getattr(self, "_form_id", self.FORM_META_ID),
            session_provider=lambda: _supabase_get_access_token(".env"),
        )
        response = dialog.show()
        if not response:
            _log_labs(
                f"subsection_dictation_cancelled subsection={subsection_key} candidate_index={candidate_index}",
                level="WARN",
            )
            return

        extraction = response.get("extraction") or {}
        transcription = str(response.get("transcription") or "").strip()
        try:
            processed = postprocess_selection_labs_extraction(
                extraction,
                subsection_key=subsection_key,
                candidate_index=candidate_index,
            )
        except Exception as exc:
            _log_labs(
                f"subsection_dictation_invalid_structured_response subsection={subsection_key} "
                f"candidate_index={candidate_index} detail={exc}",
                level="ERROR",
            )
            messagebox.showerror(
                "Labs",
                f"No fue posible interpretar la respuesta estructurada.\n\nDetalle: {exc}",
                parent=self,
            )
            return

        updates = processed.get("candidate") or {}
        preview_lines = summarize_selection_labs_updates(updates)
        warnings = []
        for item in list(response.get("warnings") or []) + list(processed.get("warnings") or []):
            text = str(item or "").strip()
            if text and text not in warnings:
                warnings.append(text)

        preview = LabsSection2PreviewDialog(
            self,
            section_label=section_label,
            candidate_index=candidate_index,
            transcription=transcription,
            preview_lines=preview_lines,
            warnings=warnings,
        )
        if not preview.show():
            _log_labs(
                f"subsection_preview_cancelled subsection={subsection_key} candidate_index={candidate_index} "
                f"fields={','.join(sorted(updates.keys()))}",
                level="WARN",
            )
            return
        self._apply_labs_candidate_updates(candidate_fields, updates, subsection_key)
        _log_labs(
            f"subsection_preview_applied subsection={subsection_key} candidate_index={candidate_index} "
            f"fields={','.join(sorted(updates.keys()))} warnings={len(warnings)}"
        )

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. DESARROLLO DE LA ACTIVIDAD")
        self.header_subtitle.config(
            text="Registra un único desarrollo de la actividad y luego completa el o los oferentes."
        )
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        self._load_cedula_options()
        content = _build_scrollable_content(section_frame, self)
        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        remove_btn = None

        self.oferente_blocks = []
        self.oferente_frames = []
        self.section2_shared_desarrollo_widget = None
        self.section2_shared_desarrollo_frame = None
        self.section2_shared_desarrollo_frame = None
        self.section2_shared_desarrollo_proxy_fields = None
        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)

        field_meta = {field["id"]: field for field in self._seleccion_module.SECTION_2["fields"]}

        def _create_widget(parent, field_id, width=30, text_height=4):
            meta = field_meta.get(field_id, {})
            if meta.get("type") == "lista":
                return ttk.Combobox(parent, values=meta.get("options", []), state="readonly", width=width)
            if meta.get("type") == "texto_largo":
                w = tk.Text(parent, width=width, height=text_height, wrap="word")
                _attach_autoexpand(w, text_height, 20)
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
                        self._apply_numeric_entry(widget, max_len=10)
                    if field_id == "certificado_porcentaje":
                        self._apply_decimal_entry(widget)
                    if field_id in {"telefono_oferente", "telefono_emergencia"}:
                        self._apply_numeric_entry(widget, max_len=10)
                    if field_id in {"nombre_oferente", "nombre_contacto_emergencia"}:
                        self._apply_name_entry(widget)
                fields[field_id] = widget
            _bind_prefixed_dropdown_fields(fields)
            return fields

        def _add_question_block(parent, title, field_ids, subitems=None, sync_binder=None):
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
            if callable(sync_binder):
                sync_binder(fields)
            else:
                _bind_prefixed_dropdown_fields(fields)
            return fields

        def _add_activity_block(parent, title, nivel_id, observacion_id, nota_id, subitems, sync_binder=None):
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
            if callable(sync_binder):
                sync_binder(fields)
            else:
                _bind_prefixed_dropdown_fields(fields)
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

        def _section2_header_title():
            return "2. DESARROLLO DE LA ACTIVIDAD"

        def _section2_header_subtitle():
            return "Registra un único desarrollo de la actividad y luego completa el o los oferentes."

        def _shared_desarrollo_title():
            return "2. DESARROLLO DE LA ACTIVIDAD"

        def _linked_section_title():
            return "3. DATOS DEL OFERENTE"

        def _get_shared_desarrollo_value():
            widget = getattr(self, "section2_shared_desarrollo_widget", None)
            if isinstance(widget, tk.Text):
                return widget.get("1.0", tk.END).strip()
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
            widget = getattr(self, "section2_shared_desarrollo_widget", None)
            if isinstance(widget, tk.Text) and widget is not source_widget:
                _set_text_widget_value(widget, shared_value)

        def _bind_shared_desarrollo(widget):
            if not isinstance(widget, tk.Text):
                return
            widget.bind(
                "<FocusOut>",
                lambda _event=None, w=widget: _sync_desarrollo_widgets(w),
                add="+",
            )

        self._labs_section2_sync_desarrollo_widgets = _sync_desarrollo_widgets

        def _refresh_shared_desarrollo_refs():
            for entry_fields in self.oferente_blocks:
                entry_fields["desarrollo_actividad"] = self.section2_shared_desarrollo_widget

        def _refresh_section_dictation_binding():
            section_name = str(getattr(self, "_current_section", "section_2") or "section_2").strip() or "section_2"
            if section_name != "section_2":
                return
            form_id = str(
                getattr(self, "_form_id", getattr(self, "FORM_META_ID", "seleccion_incluyente"))
                or getattr(self, "FORM_META_ID", "seleccion_incluyente")
            ).strip()
            if not form_id:
                return
            try:
                self.after_idle(
                    lambda w=self, fid=form_id, sec=section_name: _attach_dictation_for_section(w, fid, sec)
                )
            except Exception:
                pass

        def _create_shared_desarrollo_section(parent, *, after_widget=None, before_widget=None):
            shared_value = _get_shared_desarrollo_value()
            if self.section2_shared_desarrollo_frame is not None:
                try:
                    self.section2_shared_desarrollo_frame.destroy()
                except Exception:
                    pass
            (
                section3_frame,
                section3_body,
                section3_header,
                section3_title,
            ) = self._build_selection_subsection_shell(
                parent,
                _shared_desarrollo_title(),
                "section_3_desarrollo",
            )
            proxy_fields = {}
            self._install_labs_subsection_button(
                section3_header,
                section3_title,
                "section_3_desarrollo",
                "2. Desarrollo de la actividad",
                proxy_fields,
            )
            widget = _create_widget(
                section3_body,
                "desarrollo_actividad",
                width=80,
                text_height=6,
            )
            widget.pack(fill="x", padx=6, pady=6)
            _bind_shared_desarrollo(widget)
            if shared_value:
                widget.insert("1.0", shared_value)
            section3_frame.pack_forget()
            pack_kwargs = {"fill": "x"}
            if parent is content:
                pack_kwargs["pady"] = (0, 8)
                if before_widget is not None:
                    pack_kwargs["before"] = before_widget
            else:
                pack_kwargs["padx"] = 8
                pack_kwargs["pady"] = (0, 8)
                if after_widget is not None:
                    pack_kwargs["after"] = after_widget
            section3_frame.pack(**pack_kwargs)
            proxy_fields["desarrollo_actividad"] = widget
            self.section2_shared_desarrollo_widget = widget
            self.section2_shared_desarrollo_frame = section3_frame
            self.section2_shared_desarrollo_proxy_fields = proxy_fields
            _refresh_shared_desarrollo_refs()
            _refresh_section_dictation_binding()

        def _reposition_shared_desarrollo_section():
            before_widget = self.oferente_frames[0] if self.oferente_frames else actions
            _create_shared_desarrollo_section(content, before_widget=before_widget)

        def _refresh_section_titles():
            self.header_title.config(text=_section2_header_title())
            self.header_subtitle.config(text=_section2_header_subtitle())
            for idx, fields in enumerate(self.oferente_blocks, start=1):
                if idx - 1 < len(self.oferente_frames):
                    try:
                        self.oferente_frames[idx - 1].configure(text=f"Oferente {idx}")
                    except Exception:
                        pass
                section2_frame = getattr(self.oferente_frames[idx - 1], "_section2_frame", None)
                if section2_frame is not None:
                    section2_frame.configure(text=_linked_section_title())

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

            section2_frame, section2_body, section2_header, section2_title = self._build_selection_subsection_shell(
                block,
                _linked_section_title(),
                "section_2_fields",
            )
            block._section2_frame = section2_frame
            section2_body.grid_columnconfigure(1, weight=1)
            self._install_labs_subsection_button(
                section2_header,
                section2_title,
                "section_2_fields",
                "3. Datos del oferente",
                fields,
            )

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
            fields.update(_add_fields_grid(section2_body, section2_fields, columns=2))
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

            section41_frame, section41_body, section41_header, section41_title = self._build_selection_subsection_shell(
                block,
                "4.1 Condiciones medicas y de salud",
                "section_4_1_salud",
            )
            self._install_labs_subsection_button(
                section41_header,
                section41_title,
                "section_4_1_salud",
                "4.1 Condiciones médicas y de salud",
                fields,
            )

            fields.update(
                _add_question_block(
                    section41_body,
                    "¿Toma medicamentos?",
                    ["medicamentos_nivel_apoyo", "medicamentos_conocimiento", "medicamentos_horarios", "medicamentos_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section41_body,
                    "¿Presenta alguna alergia?",
                    ["alergias_nivel_apoyo", "alergias_tipo", "alergias_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section41_body,
                    "¿Tiene algún tipo de restricción médica?",
                    ["restriccion_nivel_apoyo", "restriccion_conocimiento", "restriccion_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section41_body,
                    "¿Asiste a controles médicos con especialista?",
                    ["controles_nivel_apoyo", "controles_asistencia", "controles_frecuencia", "controles_nota"],
                )
            )

            section42a_frame, section42a_body, section42a_header, section42a_title = self._build_selection_subsection_shell(
                block,
                "4.2A Habilidades basicas de la vida diaria",
                "section_4_2_a_habilidades",
            )
            self._install_labs_subsection_button(
                section42a_header,
                section42a_title,
                "section_4_2_a_habilidades",
                "4.2A Habilidades básicas de la vida diaria",
                fields,
            )

            fields.update(
                _add_question_block(
                    section42a_body,
                    "¿Se desplaza por la ciudad de manera independiente?",
                    ["desplazamiento_nivel_apoyo", "desplazamiento_modo", "desplazamiento_transporte", "desplazamiento_nota"],
                    sync_binder=lambda block_fields: _bind_prefixed_dropdown_subset(
                        block_fields,
                        ("desplazamiento_nivel_apoyo", "desplazamiento_modo"),
                    ),
                )
            )
            fields.update(
                _add_question_block(
                    section42a_body,
                    "¿Se le facilita ubicarse dentro de la ciudad?",
                    ["ubicacion_nivel_apoyo", "ubicacion_ciudad", "ubicacion_aplicaciones", "ubicacion_nota"],
                    sync_binder=lambda block_fields: _bind_prefixed_dropdown_subset(
                        block_fields,
                        ("ubicacion_nivel_apoyo", "ubicacion_ciudad"),
                    ),
                )
            )
            fields.update(
                _add_question_block(
                    section42a_body,
                    "¿Reconoce y maneja el dinero?",
                    ["dinero_nivel_apoyo", "dinero_reconocimiento", "dinero_manejo", "dinero_medios", "dinero_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42a_body,
                    "Presentacion personal",
                    ["presentacion_nivel_apoyo", "presentacion_personal", "presentacion_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42a_body,
                    "¿Conoce y maneja algún apoyo de comunicación escrita?",
                    ["comunicacion_escrita_nivel_apoyo", "comunicacion_escrita_apoyo", "comunicacion_escrita_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42a_body,
                    "¿Conoce y maneja algún apoyo de comunicación verbal?",
                    ["comunicacion_verbal_nivel_apoyo", "comunicacion_verbal_apoyo", "comunicacion_verbal_nota"],
                )
            )
            fields.update(
                _add_question_block(
                    section42a_body,
                    "¿A quién recurre al momento de tomar decisiones?",
                    ["decisiones_nivel_apoyo", "toma_decisiones", "toma_decisiones_nota"],
                )
            )

            section42b_frame, section42b_body, section42b_header, section42b_title = self._build_selection_subsection_shell(
                block,
                "4.2B Actividades, apoyos y discriminacion",
                "section_4_2_b_actividades",
            )
            self._install_labs_subsection_button(
                section42b_header,
                section42b_title,
                "section_4_2_b_actividades",
                "4.2B Actividades, apoyos y discriminación",
                fields,
            )
            fields.update(
                _add_activity_block(
                    section42b_body,
                    "¿Necesita apoyo en algunas de las siguientes actividades de la vida diaria?",
                    "aseo_nivel_apoyo",
                    "alimentacion",
                    "aseo_nota",
                    subitems=[
                        ("Criar y cuidado de ninos", "aseo_criar_apoyo", "Alimentacion", "aseo_alimentacion"),
                        ("Uso de los sistemas de comunicacion", "aseo_comunicacion_apoyo", "Movilidad funcional", "aseo_movilidad_funcional"),
                        ("Cuidado de las ayudas tecnicas personales", "aseo_ayudas_apoyo", "Higiene personal y aseo (Control de esfinter)", "aseo_higiene_aseo"),
                    ],
                    sync_binder=lambda block_fields: _bind_selection_activity_dropdown_fields(
                        block_fields,
                        primary_field_id="aseo_nivel_apoyo",
                        secondary_field_id="alimentacion",
                        dependent_field_ids=(
                            "aseo_criar_apoyo",
                            "aseo_comunicacion_apoyo",
                            "aseo_ayudas_apoyo",
                            "aseo_alimentacion",
                            "aseo_movilidad_funcional",
                            "aseo_higiene_aseo",
                        ),
                    ),
                )
            )
            fields.update(
                _add_activity_block(
                    section42b_body,
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
                    sync_binder=lambda block_fields: _bind_selection_activity_dropdown_fields(
                        block_fields,
                        primary_field_id="instrumentales_nivel_apoyo",
                        secondary_field_id="instrumentales_actividades",
                        dependent_field_ids=(
                            "instrumentales_criar_apoyo",
                            "instrumentales_comunicacion_apoyo",
                            "instrumentales_movilidad_apoyo",
                            "instrumentales_finanzas",
                            "instrumentales_cocina_limpieza",
                            "instrumentales_crear_hogar",
                            "instrumentales_salud_cuenta_apoyo",
                        ),
                    ),
                )
            )
            fields.update(
                _add_question_block(
                    section42b_body,
                    "¿Necesita apoyo durante actividades laborales?",
                    ["actividades_nivel_apoyo", "actividades_apoyo", "actividades_nota"],
                    subitems=[
                        ("Actividades de esparcimiento con familia", "actividades_esparcimiento_apoyo", "Psicologico en salud", "actividades_esparcimiento_cuenta_apoyo"),
                        ("Complementarios médicos", "actividades_complementarios_apoyo", "Actividades académicas de hijos", "actividades_complementarios_cuenta_apoyo"),
                        ("Subsidios economicos para estudio de hijos", None, "", "actividades_subsidios_cuenta_apoyo"),
                    ],
                    sync_binder=lambda block_fields: _bind_selection_activity_dropdown_fields(
                        block_fields,
                        primary_field_id="actividades_nivel_apoyo",
                        secondary_field_id="actividades_apoyo",
                        dependent_field_ids=(
                            "actividades_esparcimiento_apoyo",
                            "actividades_esparcimiento_cuenta_apoyo",
                            "actividades_complementarios_apoyo",
                            "actividades_complementarios_cuenta_apoyo",
                            "actividades_subsidios_cuenta_apoyo",
                        ),
                    ),
                )
            )
            fields.update(
                _add_question_block(
                    section42b_body,
                    "¿Ha sufrido o vivido discriminación?",
                    ["discriminacion_nivel_apoyo", "discriminacion", "discriminacion_nota"],
                    subitems=[
                        ("Violencia fisica", "discriminacion_violencia_apoyo", "Acoso laboral", "discriminacion_violencia_cuenta_apoyo"),
                        ("Vulneracion de derechos", "discriminacion_vulneracion_apoyo", "Violencia psicosocial", "discriminacion_vulneracion_cuenta_apoyo"),
                    ],
                    sync_binder=lambda block_fields: _bind_selection_activity_dropdown_fields(
                        block_fields,
                        primary_field_id="discriminacion_nivel_apoyo",
                        secondary_field_id="discriminacion",
                        dependent_field_ids=(
                            "discriminacion_violencia_apoyo",
                            "discriminacion_violencia_cuenta_apoyo",
                            "discriminacion_vulneracion_apoyo",
                            "discriminacion_vulneracion_cuenta_apoyo",
                        ),
                    ),
                )
            )
            self.oferente_blocks.append(fields)
            _update_remove_button_state()
            self.oferente_frames.append(block)
            _refresh_oferente_numbers()
            _refresh_section_titles()
            _reposition_shared_desarrollo_section()
            _update_remove_button_state()

        def _remove_oferente_block():
            if len(self.oferente_blocks) <= 1:
                return
            self.oferente_blocks.pop()
            frame = self.oferente_frames.pop()
            frame.destroy()
            _refresh_oferente_numbers()
            _refresh_section_titles()
            _reposition_shared_desarrollo_section()
            _update_remove_button_state()

        def _prefill_section_2():
            cache = self._seleccion_module.get_form_cache().get("section_2", [])
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
        _refresh_section_titles()

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
        shared_widget = getattr(self, "section2_shared_desarrollo_widget", None)
        shared_desarrollo = ""
        if isinstance(shared_widget, tk.Text):
            shared_desarrollo = shared_widget.get("1.0", tk.END).strip()
        payload = []
        for fields in self.oferente_blocks:
            entry = {}
            for key, widget in fields.items():
                if key == "desarrollo_actividad":
                    continue
                if isinstance(widget, ttk.Combobox):
                    entry[key] = widget.get().strip()
                elif isinstance(widget, tk.Text):
                    entry[key] = widget.get("1.0", tk.END).strip()
                else:
                    entry[key] = widget.get().strip()
            entry["desarrollo_actividad"] = shared_desarrollo
            payload.append(entry)
        try:
            self._seleccion_module.confirm_section_2(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        for idx, (template_key, label) in enumerate(self._seleccion_module.AJUSTES_ENTREVISTA_TEMPLATE_BUTTONS):
            btn = ttk.Button(
                template_actions,
                text=label,
                command=lambda key=template_key: self._insert_seleccion_section5_template(key),
            )
            btn.grid(row=idx // 3, column=idx % 3, sticky="w", padx=(0, 8), pady=(0, 8))
        ajustes = tk.Text(content, height=8, width=TEXT_WIDE, wrap="word")
        ajustes.pack(fill="x", padx=4, pady=(0, 10))
        _attach_autoexpand(ajustes, 8, 30)
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

        cache = self._seleccion_module.get_form_cache().get("section_5", {})
        if cache:
            ajustes.delete("1.0", tk.END)
            ajustes.insert("1.0", cache.get("ajustes_recomendaciones", ""))
            nota.delete(0, tk.END)
            nota.insert(0, cache.get("nota", ""))

        self._pending_autosave = lambda f=self.section5_fields: _autosave_section(self._seleccion_module, "section_5", lambda: _collect_flat_fields(f))
        _build_wizard_actions(
            content,
            back_command=self._show_section_2,
            primary_command=self._confirm_section_5,
        )

    def _insert_seleccion_section5_template(self, template_key):
        text_widget = self.section5_fields.get("ajustes_recomendaciones")
        if not text_widget:
            return
        template_text = self._seleccion_module.AJUSTES_ENTREVISTA_TEMPLATES.get(template_key, "").strip()
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
            self._seleccion_module.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_6()

    def _show_section_6(self):
        self._clear_section_container()
        self.header_title.config(text=self._seleccion_module.SECTION_6["title"])
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
        for idx in range(self._seleccion_module.SECTION_6["rows"]):
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

        cached_rows = self._seleccion_module.get_form_cache().get("section_6", [])
        while len(self.section6_rows) < len(cached_rows):
            _add_asistente_row()
        for idx, entry in enumerate(cached_rows):
            nombre_entry, cargo_entry = self.section6_rows[idx]
            nombre_entry.delete(0, tk.END)
            nombre_entry.insert(0, entry.get("nombre", ""))
            cargo_entry.delete(0, tk.END)
            cargo_entry.insert(0, entry.get("cargo", ""))
        self._pending_autosave = lambda: _autosave_section(
            self._seleccion_module,
            "section_6",
            lambda: _collect_asistente_rows(self.section6_rows),
        )

        _build_wizard_actions(
            content,
            back_command=self._show_section_5,
            primary_command=self._confirm_section_6,
            primary_text="Finalizar",
            left_buttons=[("📞 Solicitar Intérprete LSC", self._open_lsc_window)],
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
            self._seleccion_module.confirm_section_6(asistentes)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        _queue_or_run_main_form_export(self, self._export_form)

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando acta...")
        loading.set_progress(40)
        cache_snapshot = self._seleccion_module.get_form_cache()
        cache = cache_snapshot
        section_1 = cache.get("section_1", {})
        company_name = section_1.get("nombre_empresa")

        def _worker():
            output_path = _raise_finalize_stage(
                "preparando el acta",
                lambda: self._seleccion_module.export_to_excel(clear_cache=False),
            )
            _update_loading_async(
                loading,
                status="Guardando en Supabase...",
                progress=70,
            )
            _raise_finalize_stage(
                "guardando en Supabase",
                self._seleccion_module.sync_usuarios_reca,
            )
            return output_path

        _start_background_finalization(
            self,
            loading,
            form_name=getattr(self, "_form_name", self._seleccion_module.FORM_NAME),
            company_name=company_name,
            form_id=getattr(self, "_form_id", self.FORM_META_ID),
            worker_fn=_worker,
            post_delivery_fn=lambda: _clear_form_cache_safe(self._seleccion_module),
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

    def _open_lsc_window(self):
        cache = self._seleccion_module.get_form_cache()
        section_1 = cache.get("section_1", {})
        empresa = section_1 if section_1.get("nombre_empresa") else (
            self.company_data if isinstance(getattr(self, "company_data", None), dict) else None
        )
        raw = cache.get("section_2", [])
        oferentes = [
            {
                "nombre_oferente": (c.get("nombre_oferente") or "").strip(),
                "cedula": (c.get("cedula") or c.get("cedula_oferente") or "").strip(),
                "proceso": "Selección incluyente",
            }
            for c in (raw if isinstance(raw, list) else [])
            if c.get("nombre_oferente") or c.get("cedula")
        ]
        ctx = _build_lsc_context(
            self,
            module=self._seleccion_module,
            source_form="seleccion_incluyente",
            oferentes=oferentes,
        )
        _launch_linked_lsc_window(
            self,
            context=ctx,
            return_to_final_section=self._show_section_6,
            main_finish_action=self._confirm_section_6,
        )

    def _apply_decimal_entry(self, entry):
        _bind_decimal_entry(entry)

    def _apply_name_entry(self, entry):
        _bind_name_entry(entry)


class SeleccionIncluyenteLabsWindow(SeleccionIncluyenteWindow):
    FORM_META_ID = "seleccion_incluyente_labs"
    FORM_MODULE = seleccion_incluyente
    WINDOW_TITLE = "Seleccion Incluyente Labs - Seccion 1"

    def __init__(self, parent):
        self._disable_section_dictation_sections = {"section_2"}
        super().__init__(parent)


# ── VENTANA: ContratacionIncluyenteWindow ────────────────────────────────────


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
        _fn = getattr(self, '_pending_autosave', None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
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
        if contratacion_incluyente.cache_file_exists():
            _clear_local_resume_state(contratacion_incluyente)
        return False

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
        _restore_section1_cached_state(self, contratacion_incluyente)

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. DATOS DEL VINCULADO")
        self.header_subtitle.config(
            text="Completa la información del vinculado. Puedes agregar más vinculados."
        )
        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        self._load_cedula_options()
        content = _build_scrollable_content(section_frame, self)
        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        remove_btn = None

        self.oferente_blocks = []
        self.oferente_frames = []
        self.section2_shared_desarrollo_frame = None
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
            fields = _add_fields_grid(inner, fields_def, columns=2)
            _bind_prefixed_dropdown_fields(fields)
            return fields

        def _is_group_variant_ui():
            # Unified template always uses group layout
            return True

        def _section2_header_title():
            if _is_group_variant_ui():
                return "2. DESARROLLO DE LA ACTIVIDAD"
            return "2. DATOS DEL VINCULADO"

        def _section2_header_subtitle():
            if _is_group_variant_ui():
                return "Registra un único desarrollo de la actividad y luego completa los vinculados."
            return "Completa la información del vinculado. Puedes agregar más vinculados."

        def _shared_desarrollo_title():
            if _is_group_variant_ui():
                return "2. DESARROLLO DE LA ACTIVIDAD"
            return "4. DESARROLLO DE LA ACTIVIDAD"

        def _linked_section_title():
            if _is_group_variant_ui():
                return "3. DATOS DEL VINCULADO"
            return "2. DATOS DEL VINCULADO"

        def _additional_section_title():
            if _is_group_variant_ui():
                return "4. DATOS ADICIONALES"
            return "3. DATOS ADICIONALES"

        def _get_shared_desarrollo_value():
            widget = getattr(self, "section2_shared_desarrollo_widget", None)
            if isinstance(widget, tk.Text):
                return widget.get("1.0", tk.END).strip()
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

        def _refresh_section_dictation_binding():
            section_name = str(getattr(self, "_current_section", "section_2") or "section_2").strip() or "section_2"
            if section_name != "section_2":
                return
            form_id = str(getattr(self, "_form_id", "contratacion_incluyente") or "contratacion_incluyente").strip()
            if not form_id:
                return
            try:
                self.after_idle(
                    lambda w=self, fid=form_id, sec=section_name: _attach_dictation_for_section(w, fid, sec)
                )
            except Exception:
                pass

        def _create_shared_desarrollo_section(parent, *, after_widget=None, before_widget=None):
            shared_value = _get_shared_desarrollo_value()
            current_frame = getattr(self, "section2_shared_desarrollo_frame", None)
            if current_frame is not None:
                try:
                    current_frame.destroy()
                except Exception:
                    pass
            frame = tk.LabelFrame(
                parent,
                text=_shared_desarrollo_title(),
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            text_widget = tk.Text(frame, height=5, wrap="word")
            text_widget.pack(fill="x", padx=6, pady=6)
            _attach_autoexpand(text_widget, 5, 20)
            if shared_value:
                text_widget.insert("1.0", shared_value)
            pack_kwargs = {"fill": "x"}
            if parent is content:
                pack_kwargs["pady"] = (0, 8)
                if before_widget is not None:
                    pack_kwargs["before"] = before_widget
            else:
                pack_kwargs["padx"] = 8
                pack_kwargs["pady"] = (0, 8)
                if after_widget is not None:
                    pack_kwargs["after"] = after_widget
            frame.pack(**pack_kwargs)
            self.section2_shared_desarrollo_frame = frame
            self.section2_shared_desarrollo_widget = text_widget
            _refresh_section_dictation_binding()

        def _reposition_shared_desarrollo_section():
            if not self.oferente_frames:
                return
            if _is_group_variant_ui():
                _create_shared_desarrollo_section(content, before_widget=self.oferente_frames[0])
                return
            first_block = self.oferente_frames[0]
            anchor = getattr(first_block, "_section3_frame", None)
            _create_shared_desarrollo_section(first_block, after_widget=anchor)

        def _refresh_variant_dependent_options():
            valid_values = {
                "contrato_lee_observacion": contratacion_incluyente.get_section_2_field_options(
                    "contrato_lee_observacion",
                ),
            }
            for fields in self.oferente_blocks:
                for field_id, values in valid_values.items():
                    widget = fields.get(field_id)
                    if not isinstance(widget, ttk.Combobox):
                        continue
                    current = widget.get().strip()
                    widget.configure(values=values)
                    if current and current not in values:
                        widget.set("")

        def _refresh_section_titles():
            self.header_title.config(text=_section2_header_title())
            self.header_subtitle.config(text=_section2_header_subtitle())
            for frame in self.oferente_frames:
                section2_frame = getattr(frame, "_section2_frame", None)
                if section2_frame is not None:
                    section2_frame.configure(text=_linked_section_title())
                section3_frame = getattr(frame, "_section3_frame", None)
                if section3_frame is not None:
                    section3_frame.configure(text=_additional_section_title())

        def _refresh_layout():
            _refresh_oferente_numbers()
            _refresh_variant_dependent_options()
            _refresh_section_titles()
            _reposition_shared_desarrollo_section()
            _update_remove_button_state()

        def _add_oferente_block():
            idx = len(self.oferente_blocks) + 1
            block = tk.Frame(content, bg="white", bd=1, relief="solid")
            pack_kwargs = {"fill": "x", "pady": 8}
            if actions.winfo_manager():
                pack_kwargs["before"] = actions
            block.pack(**pack_kwargs)
            self.oferente_frames.append(block)

            header = tk.Label(
                block,
                text=f"Vinculado {idx}",
                font=FONT_LABEL,
                bg="white",
                fg="#222222",
            )
            header.pack(anchor="w", padx=10, pady=(8, 4))

            fields = {}

            section2_frame = tk.LabelFrame(
                block,
                text=_linked_section_title(),
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section2_frame.pack(fill="x", padx=8, pady=(0, 8))
            block._section2_frame = section2_frame

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
                            "width": 34,
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
                text=_additional_section_title(),
                bg="white",
                fg="#222222",
                font=FONT_LABEL,
                padx=8,
                pady=6,
            )
            section3_frame.pack(fill="x", padx=8, pady=(0, 8))
            block._section3_frame = section3_frame
            fields.update(
                _add_fields_grid(
                    section3_frame,
                    [
                        {
                            "id": "tipo_contrato",
                            "label": "Tipo de contrato",
                            "options": contratacion_incluyente.TIPO_CONTRATO_FIRMADO_OPTIONS,
                            "width": 42,
                        },
                        {"id": "fecha_fin", "label": "Fecha de fin", "width": 14},
                    ],
                    columns=2,
                )
            )

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
                            "options": contratacion_incluyente.get_section_2_field_options(
                                "contrato_lee_observacion",
                            ),
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
                            "options": contratacion_incluyente.CONTRATO_TIPO_CONTRATO_OPTIONS,
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
            _refresh_layout()

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
            _refresh_layout()

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
            _refresh_layout()

        _prefill_section_2()

        _pack_actions(actions)
        ttk.Button(actions, text="Agregar vinculado", command=_add_oferente_block).pack(
            side="left"
        )
        remove_btn = ttk.Button(
            actions, text="Eliminar ultimo vinculado", command=_remove_oferente_block
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
        if True:  # unified template always uses group layout
            self.header_title.config(text="5. AJUSTES RAZONABLES Y RECOMENDACIONES")
        else:
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
        _attach_autoexpand(ajustes, 6, 25)
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_7()

    def _show_section_7(self):
        self._clear_section_container()
        if True:  # unified template always uses group layout
            self.header_title.config(text="6. ASISTENTES")
        else:
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
        self._pending_autosave = lambda: _autosave_section(
            contratacion_incluyente,
            "section_7",
            lambda: _collect_asistente_rows(self.section7_rows),
        )

        _build_wizard_actions(
            content,
            back_command=self._show_section_6,
            primary_command=self._confirm_section_7,
            primary_text="Finalizar",
            left_buttons=[("📞 Solicitar Intérprete LSC", self._open_lsc_window), ("Agregar asistente", _add_asistente_row)],
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        _queue_or_run_main_form_export(self, self._export_form)

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando acta...")
        loading.set_progress(40)
        cache = contratacion_incluyente.get_form_cache()
        section_1 = cache.get("section_1", {})
        company_name = section_1.get("nombre_empresa")

        def _worker():
            output_path = _raise_finalize_stage(
                "preparando el acta",
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
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(
            self,
            mode=mode,
            lookup=contratacion_incluyente,
            button=target_button,
        )

    def _open_lsc_window(self):
        cache = contratacion_incluyente.get_form_cache()
        section_1 = cache.get("section_1", {})
        empresa = section_1 if section_1.get("nombre_empresa") else (
            self.company_data if isinstance(getattr(self, "company_data", None), dict) else None
        )
        raw = cache.get("section_2", [])
        oferentes = [
            {
                "nombre_oferente": (c.get("nombre_oferente") or "").strip(),
                "cedula": (c.get("cedula") or c.get("cedula_oferente") or "").strip(),
                "proceso": "Contratación incluyente",
            }
            for c in (raw if isinstance(raw, list) else [])
            if c.get("nombre_oferente") or c.get("cedula")
        ]
        ctx = _build_lsc_context(
            self,
            module=contratacion_incluyente,
            source_form="contratacion_incluyente",
            oferentes=oferentes,
        )
        _launch_linked_lsc_window(
            self,
            context=ctx,
            return_to_final_section=self._show_section_7,
            main_finish_action=self._confirm_section_7,
        )

    def _confirm_and_continue(self):
        _confirm_section1_and_continue(
            self,
            confirm_fn=contratacion_incluyente.confirm_section_1,
            next_step=self._show_section_2,
        )


# ── VENTANA: InduccionOrganizacionalWindow ───────────────────────────────────


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
        _fn = getattr(self, '_pending_autosave', None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
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
        if induccion_organizacional.cache_file_exists():
            _clear_local_resume_state(induccion_organizacional)
        return False

    def _show_section_1(self):
        self._clear_section_container()
        self.header_title.config(text="1. DATOS GENERALES")
        self.header_subtitle.config(text="Busca empresa por NIT y confirma datos.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)
        self._build_search(content)
        self._build_groups(content)
        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._close_to_hub).pack(side="left")
        self.continue_btn = ttk.Button(actions, text="Continuar", command=self._confirm_and_continue)
        self.continue_btn.pack(side="right")
        _restore_section1_cached_state(self, induccion_organizacional)

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

        def _set_section3_visto(value, item_ids):
            allowed_ids = set(item_ids or [])
            for item_id, widgets in self.section3_fields.items():
                if item_id not in allowed_ids:
                    continue
                visto_widget = widgets.get("visto")
                if isinstance(visto_widget, ttk.Combobox):
                    visto_widget.set(value)
            autosave_fn = getattr(self, "_pending_autosave", None)
            if callable(autosave_fn):
                autosave_fn()

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

            subsection_item_ids = [item["id"] for item in subsection["items"]]
            subsection_actions = tk.Frame(section_box, bg=COLOR_LIGHT_BG)
            subsection_actions.grid(row=0, column=0, columnspan=5, sticky="w", padx=4, pady=(0, 8))
            for label, value in (
                ("Todo si", "Si"),
                ("Todo no", "No"),
                ("Todo no aplica", "No aplica"),
            ):
                ttk.Button(
                    subsection_actions,
                    text=label,
                    command=lambda selected=value, ids=subsection_item_ids: _set_section3_visto(selected, ids),
                ).pack(side="left", padx=(0, 8))

            tk.Label(section_box, text="Tematica", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=1, column=0, sticky="w", padx=4, pady=(0, 6)
            )
            tk.Label(section_box, text="Visto", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=1, column=1, sticky="w", padx=4, pady=(0, 6)
            )
            tk.Label(section_box, text="Responsable", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=1, column=2, sticky="w", padx=4, pady=(0, 6)
            )
            tk.Label(
                section_box, text="Medio de socializacion", bg=COLOR_LIGHT_BG, font=FONT_LABEL
            ).grid(row=1, column=3, sticky="w", padx=4, pady=(0, 6))
            tk.Label(section_box, text="Descripción", bg=COLOR_LIGHT_BG, font=FONT_LABEL).grid(
                row=1, column=4, sticky="w", padx=4, pady=(0, 6)
            )

            for idx, item in enumerate(subsection["items"], start=2):
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
        self._pending_autosave = lambda f=self.section3_fields: _autosave_section(induccion_organizacional, "section_3", lambda: _collect_flat_fields(f))
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
            _attach_autoexpand(texto, 8, 25)

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
        _attach_autoexpand(self.section5_text, 10, 30)

        cache = induccion_organizacional.get_form_cache().get("section_5", {})
        if cache.get("observaciones"):
            self.section5_text.insert("1.0", cache.get("observaciones", ""))

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda: _autosave_section(induccion_organizacional, "section_5", lambda: {"observaciones": self.section5_text.get("1.0", tk.END).strip()})
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
        self._pending_autosave = lambda: _autosave_section(
            induccion_organizacional,
            "section_6",
            lambda: _collect_asistente_rows(self.section6_rows),
        )
        ttk.Button(actions, text="Regresar", command=self._show_section_5).pack(side="left")
        ttk.Button(actions, text="📞 Solicitar Intérprete LSC", command=self._open_lsc_window).pack(side="left", padx=(8, 0))
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
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(
            self,
            mode=mode,
            lookup=induccion_organizacional,
            button=target_button,
        )
        return
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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

    def _open_lsc_window(self):
        cache = induccion_organizacional.get_form_cache()
        section_1 = cache.get("section_1", {})
        empresa = section_1 if section_1.get("nombre_empresa") else (
            self.company_data if isinstance(getattr(self, "company_data", None), dict) else None
        )
        raw = cache.get("section_2", [])
        oferentes = [
            {
                "nombre_oferente": (c.get("nombre_oferente") or "").strip(),
                "cedula": (c.get("cedula") or c.get("cedula_oferente") or "").strip(),
                "proceso": "Inducción organizacional",
            }
            for c in (raw if isinstance(raw, list) else [])
            if c.get("nombre_oferente") or c.get("cedula")
        ]
        ctx = _build_lsc_context(
            self,
            module=induccion_organizacional,
            source_form="induccion_organizacional",
            oferentes=oferentes,
        )
        _launch_linked_lsc_window(
            self,
            context=ctx,
            return_to_final_section=self._show_section_6,
            main_finish_action=self._confirm_section_6,
        )

    def _confirm_and_continue(self):
        _confirm_section1_and_continue(
            self,
            confirm_fn=induccion_organizacional.confirm_section_1,
            next_step=self._show_section_2,
        )

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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_5()

    def _confirm_section_5(self):
        payload = {
            "observaciones": self.section5_text.get("1.0", tk.END).strip(),
        }
        try:
            induccion_organizacional.confirm_section_5(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        _queue_or_run_main_form_export(self, self._export_form)

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando acta...")
        loading.set_progress(35)

        cache_snapshot = induccion_organizacional.get_form_cache()
        section_1 = cache_snapshot.get("section_1", {})
        company_name = section_1.get("nombre_empresa")
        def _worker():
            output_path = _raise_finalize_stage(
                "preparando el acta",
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


# ── VENTANA: InduccionOperativaWindow ────────────────────────────────────────


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
        _fn = getattr(self, '_pending_autosave', None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
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
        if induccion_operativa.cache_file_exists():
            _clear_local_resume_state(induccion_operativa)
        return False

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
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(
            self,
            mode=mode,
            lookup=induccion_operativa,
            button=target_button,
        )
        return
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._close_to_hub).pack(side="left")
        self.continue_btn = ttk.Button(actions, text="Continuar", command=self._confirm_and_continue)
        self.continue_btn.pack(side="right")
        _restore_section1_cached_state(self, induccion_operativa)

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

        bulk_actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        bulk_actions.pack(anchor="w", padx=FORM_PADX, pady=(8, 4))
        for label, value in (
            ("Todo si", "Si"),
            ("Todo no", "No"),
            ("Todo no aplica", "No aplica"),
        ):
            ttk.Button(
                bulk_actions,
                text=label,
                command=lambda selected=value: self._set_section3_ejecucion(selected),
            ).pack(side="left", padx=(0, 8))

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
        self._pending_autosave = lambda f=self.section3_fields: _autosave_section(induccion_operativa, "section_3", lambda: _collect_flat_fields(f))
        ttk.Button(actions, text="Regresar", command=self._show_section_2).pack(side="left")
        ttk.Button(actions, text="Continuar", command=self._confirm_section_3).pack(side="right")

    def _set_section3_ejecucion(self, value, item_ids=None):
        allowed_ids = set(item_ids or []) if item_ids else None
        for item_id, widgets in self.section3_fields.items():
            if allowed_ids is not None and item_id not in allowed_ids:
                continue
            ejecucion_widget = widgets.get("ejecucion")
            if isinstance(ejecucion_widget, ttk.Combobox):
                ejecucion_widget.set(value)
        autosave_fn = getattr(self, "_pending_autosave", None)
        if callable(autosave_fn):
            autosave_fn()

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
                _bind_prefixed_dropdown_fields(self.section4_item_widgets[item["id"]])
                sync_fn = getattr(nivel, "_nivel_apoyo_observacion_sync", None)
                if callable(sync_fn):
                    sync_fn()

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
        self._pending_autosave = lambda f=self.section5_fields: _autosave_section(induccion_operativa, "section_5", lambda: _collect_flat_fields(f))
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
        _attach_autoexpand(self.section6_text, 8, 30)

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
        _attach_autoexpand(self.section8_text, 8, 30)

        cached = induccion_operativa.get_form_cache().get("section_8", {})
        if cached.get("observaciones_recomendaciones"):
            self.section8_text.insert("1.0", cached.get("observaciones_recomendaciones", ""))

        actions = tk.Frame(section_frame, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        self._pending_autosave = lambda: _autosave_section(induccion_operativa, "section_8", lambda: {"observaciones_recomendaciones": self.section8_text.get("1.0", tk.END).strip()})
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
        self._pending_autosave = lambda: _autosave_section(
            induccion_operativa,
            "section_9",
            lambda: _collect_asistente_rows(self.section9_rows),
        )
        ttk.Button(actions, text="Regresar", command=self._show_section_8).pack(side="left")
        ttk.Button(actions, text="📞 Solicitar Intérprete LSC", command=self._open_lsc_window).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Agregar asistente", command=_add_row).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Eliminar ultimo", command=_remove_last).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Finalizar", command=self._confirm_section_9).pack(side="right")

    def _open_lsc_window(self):
        cache = induccion_operativa.get_form_cache()
        section_1 = cache.get("section_1", {})
        empresa = section_1 if section_1.get("nombre_empresa") else (
            self.company_data if isinstance(getattr(self, "company_data", None), dict) else None
        )
        raw = cache.get("section_2", [])
        oferentes = [
            {
                "nombre_oferente": (c.get("nombre_oferente") or "").strip(),
                "cedula": (c.get("cedula") or c.get("cedula_oferente") or "").strip(),
                "proceso": "Inducción operativa",
            }
            for c in (raw if isinstance(raw, list) else [])
            if c.get("nombre_oferente") or c.get("cedula")
        ]
        ctx = _build_lsc_context(
            self,
            module=induccion_operativa,
            source_form="induccion_operativa",
            oferentes=oferentes,
        )
        _launch_linked_lsc_window(
            self,
            context=ctx,
            return_to_final_section=self._show_section_9,
            main_finish_action=self._confirm_section_9,
        )

    def _confirm_and_continue(self):
        _confirm_section1_and_continue(
            self,
            confirm_fn=induccion_operativa.confirm_section_1,
            next_step=self._show_section_2,
        )

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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_6()

    def _confirm_section_6(self):
        payload = {
            "ajustes_requeridos": self.section6_text.get("1.0", tk.END).strip(),
        }
        try:
            induccion_operativa.confirm_section_6(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_7()

    def _confirm_section_7(self):
        payload = {
            "fecha_primer_seguimiento": self.section7_date.get().strip(),
        }
        try:
            induccion_operativa.confirm_section_7(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_8()

    def _confirm_section_8(self):
        payload = {
            "observaciones_recomendaciones": self.section8_text.get("1.0", tk.END).strip(),
        }
        try:
            induccion_operativa.confirm_section_8(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        _queue_or_run_main_form_export(self, self._export_form)

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando acta...")
        loading.set_progress(35)

        cache_snapshot = induccion_operativa.get_form_cache()
        section_1 = cache_snapshot.get("section_1", {})
        company_name = section_1.get("nombre_empresa")
        def _worker():
            output_path = _raise_finalize_stage(
                "preparando el acta",
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


# ── VENTANA: SensibilizacionWindow ───────────────────────────────────────────


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
        _fn = getattr(self, '_pending_autosave', None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
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
        if sensibilizacion.cache_file_exists():
            _clear_local_resume_state(sensibilizacion)
        return False

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
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(
            self,
            mode=mode,
            lookup=sensibilizacion,
            button=target_button,
        )
        return
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._close_to_hub).pack(side="left")
        self.continue_btn = ttk.Button(actions, text="Continuar", command=self._confirm_section_1)
        self.continue_btn.pack(side="right")
        _restore_section1_cached_state(self, sensibilizacion)

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
        _attach_autoexpand(self.section3_text, 8, 30)
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
        self._pending_autosave = lambda: _autosave_section(
            sensibilizacion,
            "section_5",
            lambda: _collect_asistente_rows(self.section5_rows),
        )
        ttk.Button(actions, text="Regresar", command=self._show_section_4).pack(side="left")
        ttk.Button(actions, text="📞 Solicitar Intérprete LSC", command=self._open_lsc_window).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Agregar asistente", command=_add_row).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Eliminar ultimo", command=_remove_last).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Finalizar", command=self._confirm_section_5).pack(side="right")

    def _open_lsc_window(self):
        cache = sensibilizacion.get_form_cache()
        section_1 = cache.get("section_1", {})
        empresa = section_1 if section_1.get("nombre_empresa") else (
            self.company_data if isinstance(getattr(self, "company_data", None), dict) else None
        )
        raw = cache.get("section_2", [])
        oferentes = [
            {
                "nombre_oferente": (c.get("nombre_oferente") or "").strip(),
                "cedula": (c.get("cedula") or c.get("cedula_oferente") or "").strip(),
                "proceso": "Sensibilización",
            }
            for c in (raw if isinstance(raw, list) else [])
            if c.get("nombre_oferente") or c.get("cedula")
        ]
        ctx = _build_lsc_context(
            self,
            module=sensibilizacion,
            source_form="sensibilizacion",
            oferentes=oferentes,
        )
        _launch_linked_lsc_window(
            self,
            context=ctx,
            return_to_final_section=self._show_section_5,
            main_finish_action=self._confirm_section_5,
        )

    def _confirm_section_1(self):
        _confirm_section1_and_continue(
            self,
            confirm_fn=sensibilizacion.confirm_section_1,
            next_step=self._show_section_2,
        )

    def _confirm_section_2(self):
        try:
            sensibilizacion.confirm_section_2({})
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_3()

    def _confirm_section_3(self):
        payload = {"observaciones": self.section3_text.get("1.0", tk.END).strip()}
        try:
            sensibilizacion.confirm_section_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_4()

    def _confirm_section_4(self):
        try:
            sensibilizacion.confirm_section_4({})
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
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
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        _queue_or_run_main_form_export(self, self._export_form)

    def _export_form(self):
        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando acta...")
        loading.set_progress(35)

        cache_snapshot = sensibilizacion.get_form_cache()
        section_1 = cache_snapshot.get("section_1", {})
        company_name = section_1.get("nombre_empresa")
        def _worker():
            output_path = _raise_finalize_stage(
                "preparando el acta",
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


# ── VENTANA: SeguimientosWindow ──────────────────────────────────────────────


class SeguimientosWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Seguimientos - Inicio de Caso")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self.user_row = None
        self.linked_company = {}
        self.case_record = None
        self.case_path = None
        self.cedula_options = []
        self._filtered_cedulas = []

        self.status_var = tk.StringVar(value="Ingresa la cédula y busca el vinculado.")
        self.user_name_var = tk.StringVar(value="")
        self.user_phone_var = tk.StringVar(value="")
        self.user_role_var = tk.StringVar(value="")
        self.user_discapacidad_var = tk.StringVar(value="")
        self.path_var = tk.StringVar(value="")
        self.company_var = tk.StringVar(value="")
        self.followups_var = tk.StringVar(value="")
        self.suggestion_var = tk.StringVar(value="")
        self.profesional_var = tk.StringVar(value="")
        self.case_type_var = tk.StringVar(value="")
        self.max_followups_var = tk.StringVar(value="")
        self.next_step_var = tk.StringVar(value="")
        self.compensar_var = tk.StringVar(value="")
        self.company_nit_var = tk.StringVar(value="")
        self.company_name_search_var = tk.StringVar(value="")
        self._case_summary_cache = None
        self._case_suggestion_cache = None
        self._case_cache_key = ""
        self._flow_stage_model = []
        self._buscar_vinculado_busy = False
        self._cedula_load_busy = False
        self._case_bootstrap_busy = False
        self._cedula_lookup_hint_id = "seguimientos_cedula_lookup"
        self._company_lookup_hint_id = "seguimientos_company_lookup"
        self._compensar_hint_id = "seguimientos_compensar"
        self.status_label_widget = None
        self.flow_stage_container = None
        self._pending_case_bootstrap = None

        self._build_header()
        self._build_body()
        self._refresh_case_flow_visuals()
        self._set_cedula_lookup_hint("Cargando cédulas disponibles...", state="loading")
        self._load_cedulas()

    def _get_case_target(self):
        if (
            self.case_record
            and str((self.case_record or {}).get("source") or "").strip() == "drive"
            and str((self.case_record or {}).get("mime_type") or "").strip()
            == seguimientos.GOOGLE_SHEETS_MIME
        ):
            return self.case_record
        return None

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
                "Identifica el vinculado, confirma la empresa y continúa siempre "
                "desde la etapa sugerida del caso."
            ),
            font=FONT_SUBTITLE,
            fg="#333333",
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w", pady=(2, 0))
        ttk.Button(
            header,
            text="📞 Solicitar Intérprete LSC",
            command=self._open_lsc_window,
        ).pack(anchor="e", pady=(6, 0))

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
        ui_feedback.bind_placeholder(self.cedula_combo, "Ej: 12345678")
        self.cedula_combo.bind("<KeyRelease>", self._filter_cedulas)

        self.buscar_vinculado_btn = ttk.Button(search, text="Buscar", style="Primary.TButton", command=self._buscar_vinculado)
        self.buscar_vinculado_btn.grid(
            row=0, column=2, sticky="w", padx=(12, 0), pady=4
        )
        self.reload_cedulas_btn = ttk.Button(
            search,
            text="Recargar lista",
            style="Secondary.TButton",
            command=self._reload_cedulas,
        )
        self.reload_cedulas_btn.grid(row=0, column=3, sticky="w", padx=(8, 0), pady=4)
        self.cedula_hint_label = tk.Label(search, text="", font=("Arial", 9), fg="#6B6B6B", bg=COLOR_LIGHT_BG)
        self.cedula_hint_label.grid(row=1, column=1, columnspan=3, sticky="w", pady=(2, 0))
        ui_feedback.register_group_hint(self, self._cedula_lookup_hint_id, self.cedula_hint_label)

        tk.Label(
            search,
            text="¿Empresa afiliada a Compensar?",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.compensar_combo = ttk.Combobox(
            search,
            textvariable=self.compensar_var,
            values=["", "Si (Compensar)", "No (Otro)"],
            state="disabled",
            width=30,
        )
        self.compensar_combo.grid(row=2, column=1, sticky="w", pady=4)
        self.compensar_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_compensar_choice_changed())
        compensar_hint = tk.Label(search, text="", font=("Arial", 9), fg="#6B6B6B", bg=COLOR_LIGHT_BG)
        compensar_hint.grid(row=3, column=1, columnspan=2, sticky="w", pady=(2, 0))
        ui_feedback.register_group_hint(self, self._compensar_hint_id, compensar_hint)

        company_lookup = tk.LabelFrame(
            container,
            text="Confirma la empresa del caso",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        company_lookup.pack(fill="x", pady=(12, 0))
        company_hint = tk.Label(company_lookup, text="", font=("Arial", 9), fg="#6B6B6B", bg=COLOR_LIGHT_BG)
        company_hint.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        ui_feedback.register_group_hint(self, self._company_lookup_hint_id, company_hint)

        tk.Label(company_lookup, text="NIT empresa:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=4
        )
        self.company_nit_entry = ttk.Combobox(
            company_lookup, textvariable=self.company_nit_var, width=28, state="disabled", values=()
        )
        self.company_nit_entry.grid(row=1, column=1, sticky="w", pady=4)
        ui_feedback.bind_placeholder(self.company_nit_entry, "Ej: 900123456-7")
        self.company_nit_entry.bind("<KeyRelease>", self._update_company_nit_suggestions)
        self.company_nit_entry.bind("<Button-1>", self._update_company_nit_suggestions)
        self.company_nit_entry.bind("<FocusIn>", self._update_company_nit_suggestions)
        self.company_nit_entry.bind("<<ComboboxSelected>>", self._buscar_empresa_manual_por_nit)
        self.company_nit_entry.bind("<Return>", self._buscar_empresa_manual_por_nit)
        self.company_nit_btn = ttk.Button(
            company_lookup,
            text="Buscar por NIT",
            style="Secondary.TButton",
            command=self._buscar_empresa_manual_por_nit,
            state="disabled",
        )
        self.company_nit_btn.grid(row=1, column=2, sticky="w", padx=(12, 0), pady=4)

        tk.Label(company_lookup, text="Nombre empresa:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=4
        )
        self.company_name_entry = ttk.Combobox(
            company_lookup, textvariable=self.company_name_search_var, width=62, state="disabled", values=()
        )
        self.company_name_entry.grid(row=2, column=1, sticky="w", pady=4)
        ui_feedback.bind_placeholder(self.company_name_entry, "Escribe al menos 2 letras")
        self.company_name_entry.bind("<KeyRelease>", self._update_company_name_suggestions)
        self.company_name_entry.bind("<Button-1>", self._update_company_name_suggestions)
        self.company_name_entry.bind("<FocusIn>", self._update_company_name_suggestions)
        self.company_name_entry.bind("<<ComboboxSelected>>", self._buscar_empresa_manual_por_nombre)
        self.company_name_entry.bind("<Return>", self._buscar_empresa_manual_por_nombre)
        self.company_name_btn = ttk.Button(
            company_lookup,
            text="Buscar por nombre",
            style="Secondary.TButton",
            command=self._buscar_empresa_manual_por_nombre,
            state="disabled",
        )
        self.company_name_btn.grid(row=2, column=2, sticky="w", padx=(12, 0), pady=4)

        context = tk.LabelFrame(
            container,
            text="Contexto del caso",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        context.pack(fill="x", pady=(12, 0))
        context.grid_columnconfigure(1, weight=1)

        self._add_info_row(context, 0, "Vinculado:", self.user_name_var)
        self._add_info_row(context, 1, "Teléfono:", self.user_phone_var)
        self._add_info_row(context, 2, "Cargo:", self.user_role_var)
        self._add_info_row(context, 3, "Discapacidad:", self.user_discapacidad_var)
        self._add_info_row(context, 4, "Empresa:", self.company_var)
        self._add_info_row(context, 5, "Profesional RECA:", self.profesional_var)
        self._add_info_row(context, 6, "Tipo de empresa:", self.case_type_var)
        self._add_info_row(context, 7, "Seguimientos visibles:", self.max_followups_var)
        self._add_info_row(context, 8, "Próximo paso:", self.next_step_var)
        self._add_info_row(context, 9, "Historial completado:", self.followups_var)

        status = tk.LabelFrame(
            container,
            text="Ruta del seguimiento",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        status.pack(fill="x", pady=(12, 0))
        status.grid_columnconfigure(0, weight=1)

        self._add_info_row(status, 0, "Caso:", self.path_var)
        self._add_info_row(status, 1, "Resumen operativo:", self.suggestion_var)
        self.status_label_widget = tk.Label(
            status,
            textvariable=self.status_var,
            font=("Arial", 10),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
            wraplength=860,
        )
        self.status_label_widget.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 10))
        self.flow_stage_container = tk.Frame(status, bg=COLOR_LIGHT_BG)
        self.flow_stage_container.grid(row=3, column=0, columnspan=2, sticky="ew")

        actions = tk.Frame(container, bg=COLOR_LIGHT_BG)
        _pack_actions(actions, pad_y=(14, FORM_PADY), pad_x=False)

        ttk.Button(actions, text="Regresar", style="Secondary.TButton", command=self._close_to_hub).pack(side="left")
        self.create_btn = ttk.Button(
            actions,
            text="Continuar donde voy",
            style="Primary.TButton",
            command=self._continue_case_flow,
            state="disabled",
        )
        self.create_btn.pack(side="right")
        self.open_btn = ttk.Button(
            actions,
            text="Abrir en Drive",
            style="Secondary.TButton",
            command=self._abrir_archivo,
            state="disabled",
        )
        self.open_btn.pack(side="right", padx=(0, 8))
        ui_feedback.set_group_hint(self, self._company_lookup_hint_id, "Primero busca al vinculado arriba.", state="hint")
        ui_feedback.set_group_hint(self, self._compensar_hint_id, "Primero busca al vinculado arriba.", state="hint")

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

    def _get_case_type_label(self):
        choice = str(self.compensar_var.get() or "").strip()
        if choice.startswith("Si"):
            return "Compensar"
        if choice.startswith("No"):
            return "Otro"
        caja = str((self.linked_company or {}).get("caja_compensacion") or "").strip()
        if caja:
            return caja
        return "Por confirmar"

    def _update_case_context(self, *, suggestion=None, summary=None):
        suggestion = dict(suggestion or self._case_suggestion_cache or {})
        summary = dict(summary or self._case_summary_cache or {})
        self.profesional_var.set(
            str(
                summary.get("profesional_asignado")
                or (self.linked_company or {}).get("profesional_asignado")
                or ""
            ).strip()
            or "(Sin asignar)"
        )
        self.case_type_var.set(self._get_case_type_label())
        max_followups = (
            suggestion.get("max_seguimientos")
            or (self.case_record or {}).get("max_seguimientos")
            or (self._case_summary_cache or {}).get("max_seguimientos")
            or 3
        )
        self.max_followups_var.set(str(max_followups))
        next_title = _friendly_followup_sheet_title(
            suggestion.get("sheet"),
            self._case_summary_cache or {},
            base_sheet_name=(self.case_record or {}).get("base_sheet_name") or seguimientos.SHEET_BASE,
        )
        next_message = str(suggestion.get("message") or "").strip()
        next_text = next_title
        if next_message and next_message != next_title:
            next_text = f"{next_title}. {next_message}" if next_title else next_message
        self.next_step_var.set(next_text or "Busca primero el vinculado.")

    def _render_flow_stage_cards(self):
        container = self.flow_stage_container
        if container is None:
            return
        for child in container.winfo_children():
            child.destroy()
        stages = list(self._flow_stage_model or [])
        if not stages:
            return
        for index, stage in enumerate(stages):
            palette = _get_followup_stage_palette(
                stage.get("status"),
                is_suggested=bool(stage.get("is_suggested")),
            )
            card = tk.Frame(
                container,
                bg=palette["bg"],
                highlightbackground=palette["accent"],
                highlightthickness=2,
                bd=0,
            )
            card.grid(
                row=index // 3,
                column=index % 3,
                sticky="nsew",
                padx=(0 if index % 3 == 0 else 8, 0),
                pady=(0 if index < 3 else 8, 0),
            )
            container.grid_columnconfigure(index % 3, weight=1)
            inner = tk.Frame(card, bg=palette["bg"], padx=12, pady=10)
            inner.pack(fill="both", expand=True)
            tk.Label(
                inner,
                text=str(stage.get("title") or ""),
                font=FONT_LABEL,
                fg=palette["accent"],
                bg=palette["bg"],
                anchor="w",
                justify="left",
            ).pack(anchor="w")
            tk.Label(
                inner,
                text=_format_followup_stage_status(
                    stage.get("status"),
                    coverage_percent=stage.get("coverage_percent"),
                    is_suggested=bool(stage.get("is_suggested")),
                ),
                font=("Arial", 9, "bold"),
                fg=palette["muted"],
                bg=palette["bg"],
                anchor="w",
                justify="left",
            ).pack(anchor="w", pady=(4, 2))
            tk.Label(
                inner,
                text=str(stage.get("helper_text") or ""),
                font=("Arial", 9),
                fg="#333333",
                bg=palette["bg"],
                anchor="w",
                justify="left",
                wraplength=250,
            ).pack(anchor="w")

    def _refresh_case_flow_visuals(self, *, suggestion=None, summary=None):
        suggestion = dict(suggestion or self._case_suggestion_cache or {})
        summary = dict(summary or self._case_summary_cache or {})
        workflow = {}
        if self.case_record:
            workflow = {
                "base_sheet_name": summary.get("base_sheet_name") or seguimientos.SHEET_BASE,
                "max_seguimientos": summary.get("max_seguimientos") or suggestion.get("max_seguimientos") or (self.case_record or {}).get("max_seguimientos") or 3,
                "stage_model": list(summary.get("sheet_progress") or []),
                "sheet_progress": list(summary.get("sheet_progress") or []),
                "suggested_sheet": suggestion.get("sheet") or "",
            }
        self._flow_stage_model = _build_followup_window_flow_model(
            user_row=self.user_row,
            linked_company=self.linked_company,
            compensar_choice=self.compensar_var.get(),
            case_record=self.case_record,
            workflow=workflow,
            summary=summary,
            suggestion=suggestion,
        )
        self._update_case_context(suggestion=suggestion, summary=summary)
        self._render_flow_stage_cards()

    def _set_cedula_lookup_hint(self, text, *, state="hint"):
        ui_feedback.set_group_hint(self, self._cedula_lookup_hint_id, text, state=state)

    def _set_cedula_search_enabled(self, enabled):
        combo_state = "normal" if enabled else "disabled"
        button_state = "normal" if enabled else "disabled"
        try:
            self.cedula_combo.config(state=combo_state)
        except Exception:
            pass
        try:
            self.buscar_vinculado_btn.config(state=button_state)
        except Exception:
            pass

    def _format_cedula_load_error(self, exc):
        text = str(exc or "").strip().lower()
        if any(token in text for token in ("http 401", "http 403", "unauthorized", "permission denied", "42501")):
            return (
                "No fue posible cargar las cédulas porque la sesión o los permisos de la base de datos "
                "no están disponibles. Reintenta la carga; si persiste, cierra sesión e ingresa nuevamente."
            )
        if _is_connectivity_exception(exc) or any(
            token in text
            for token in ("timeout", "timed out", "connection", "network", "dns", "failed to resolve")
        ):
            return "No fue posible cargar las cédulas por un problema de conexión con la base de datos. Reintenta la carga."
        return _log_user_error("linked_user_search", exc)

    def _apply_cedula_options(self, options, *, show_success_hint=False):
        self.cedula_options = list(options or [])
        self._filtered_cedulas = list(self.cedula_options)
        self.cedula_combo["values"] = self._filtered_cedulas[:50]
        if self.cedula_options:
            self._set_cedula_search_enabled(True)
            self._set_cedula_lookup_hint(
                "Lista de cédulas recargada. Puedes escribir para filtrar." if show_success_hint else "",
                state="success" if show_success_hint else "hint",
            )
            if not self.user_row:
                self.status_var.set("Ingresa la cédula y busca el vinculado.")
                self._refresh_case_flow_visuals()
            return True
        message = "No hay cédulas disponibles en usuarios_reca. Reintenta la carga después de sincronizar la base."
        self._set_cedula_search_enabled(False)
        self._set_cedula_lookup_hint(message, state="warning")
        self.status_var.set(message)
        self._refresh_case_flow_visuals()
        return False

    def _handle_cedula_load_error(self, exc):
        message = self._format_cedula_load_error(exc)
        if self.cedula_options:
            message = f"{message} Se conserva la última lista cargada."
            self._set_cedula_search_enabled(True)
            self.cedula_combo["values"] = self._filtered_cedulas[:50]
        else:
            self.cedula_options = []
            self._filtered_cedulas = []
            self.cedula_combo.set("")
            self.cedula_combo["values"] = ()
            self._set_cedula_search_enabled(False)
        self._set_cedula_lookup_hint(message, state="error")
        self.status_var.set(message)
        self._refresh_case_flow_visuals()
        return False

    def _set_company_lookup_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        try:
            self.company_nit_entry.config(state=state)
            self.company_name_entry.config(state=state)
            self.company_nit_btn.config(state=state)
            self.company_name_btn.config(state=state)
        except Exception:
            pass
        hint_text = ""
        if not enabled and not self.user_row:
            hint_text = "Primero busca al vinculado arriba."
        ui_feedback.set_group_hint(
            self,
            self._company_lookup_hint_id,
            hint_text,
            state="hint",
        )

    def _set_compensar_choice_enabled(self, enabled):
        try:
            self.compensar_combo.config(state="readonly" if enabled else "disabled")
        except Exception:
            pass
        hint_text = ""
        if not enabled:
            hint_text = (
                "Confirma la empresa del caso arriba."
                if self.user_row
                else "Primero busca al vinculado arriba."
            )
        ui_feedback.set_group_hint(
            self,
            self._compensar_hint_id,
            hint_text,
            state="hint",
        )

    def _resolve_compensar_choice(self, company):
        caja = str((company or {}).get("caja_compensacion") or "").strip().lower()
        if not caja:
            return ""
        return "Si (Compensar)" if "compensar" in caja else "No (Otro)"

    def _current_case_cache_key(self, record=None, path=None):
        record = record or self.case_record or {}
        path = path or self.case_path or ""
        file_id = str((record or {}).get("file_id") or "").strip()
        if file_id:
            return f"file:{file_id}"
        return f"path:{str(path or '').strip()}"

    def _compensar_choice_is_resolved(self, value=None):
        choice = str(value if value is not None else self.compensar_var.get()).strip()
        return choice.startswith("Si") or choice.startswith("No")

    def _build_case_user_row_snapshot(self, *, row=None, linked_company=None):
        user_row = dict(row or self.user_row or {})
        company = dict(linked_company or self.linked_company or {})
        if company.get("nit_empresa"):
            user_row["empresa_nit"] = str(company.get("nit_empresa") or "").strip()
        if company.get("nombre_empresa"):
            user_row["empresa_nombre"] = str(company.get("nombre_empresa") or "").strip()
        return user_row

    def _clear_case_state(self):
        self.case_record = None
        self.case_path = None
        self._case_suggestion_cache = None
        self._case_summary_cache = None
        self._case_cache_key = ""
        self.path_var.set("(Sin Google Sheet aún)")
        self.open_btn.config(state="disabled")
        self.create_btn.config(state="normal" if self.user_row else "disabled")
        self._refresh_case_flow_visuals()

    def _set_pending_case_bootstrap(self, *, row, normalized_cedula, linked_company=None):
        self._pending_case_bootstrap = {
            "row": dict(row or {}),
            "cedula": str(normalized_cedula or "").strip(),
            "linked_company": dict(linked_company or {}),
        }
        self._clear_case_state()
        self.suggestion_var.set("Se creará el Google Sheet para continuar con la ficha inicial del proceso.")
        self.followups_var.set("Sin hojas completadas")
        self.next_step_var.set("Confirma la empresa para crear el caso.")
        self._refresh_case_flow_visuals()

    def _finalize_case_bootstrap_result(self, data):
        result = dict((data or {}).get("result") or {})
        self.case_record = result.get("record") or {}
        self.case_path = None
        self._pending_case_bootstrap = None
        self.path_var.set((self.case_record or {}).get("webViewLink") or "")
        self.open_btn.config(state="normal" if self.case_record else "disabled")
        self.create_btn.config(state="normal" if self.case_record else "disabled")
        self._case_suggestion_cache = data.get("suggestion")
        self._case_summary_cache = data.get("summary")
        self._case_cache_key = self._current_case_cache_key(self.case_record, self.case_path)

        state_error = data.get("state_error")
        if state_error is not None:
            self._case_suggestion_cache = None
            self._case_summary_cache = None
            self.suggestion_var.set("No fue posible leer la etapa sugerida.")
            self.followups_var.set("")
            self.status_var.set(_format_followup_case_state_error(state_error))
            self._refresh_case_flow_visuals()
            return

        self._apply_summary_result(
            suggestion=data.get("suggestion"),
            summary=data.get("summary"),
        )
        created = bool(result.get("created"))
        if created:
            self.status_var.set(
                "La cédula existe, pero no tenía caso en seguimientos; se creó carpeta y Google Sheet."
            )
        else:
            self.status_var.set("La persona ya existe en seguimientos.")

    def _maybe_start_pending_case_bootstrap(self):
        pending = dict(self._pending_case_bootstrap or {})
        if not pending or bool(getattr(self, "_case_bootstrap_busy", False)):
            return False
        if not self._compensar_choice_is_resolved():
            return False
        row = dict(pending.get("row") or {})
        normalized_cedula = str(pending.get("cedula") or "").strip()
        linked_company = dict(self.linked_company or pending.get("linked_company") or {})
        user_row = self._build_case_user_row_snapshot(row=row, linked_company=linked_company)
        is_compensar = self.compensar_var.get().startswith("Si")

        def _worker(progress):
            progress("Creando el Google Sheet del caso...", 28)
            result = seguimientos.ensure_case_record(
                normalized_cedula,
                user_row,
                is_compensar=is_compensar,
            )
            record = result.get("record") or {}
            progress("Leyendo el estado actual del caso...", 74)
            state_snapshot = _read_followup_case_state(record)
            progress("Organizando la información del caso...", 100)
            return {
                "result": result,
                "suggestion": state_snapshot.get("suggestion"),
                "summary": state_snapshot.get("summary"),
                "state_error": state_snapshot.get("error"),
            }

        def _on_success(data):
            self._finalize_case_bootstrap_result(data)

        return self._run_loading_job(
            title="Creando caso",
            initial_status="Creando carpeta y Google Sheet del caso...",
            worker=_worker,
            on_success=_on_success,
            on_error_context="followup_case",
            busy_attr="_case_bootstrap_busy",
            busy_widgets=[self.cedula_combo, getattr(self, "buscar_vinculado_btn", None), getattr(self, "create_btn", None)],
        )

    def _on_compensar_choice_changed(self):
        if self._pending_case_bootstrap and self._compensar_choice_is_resolved():
            self._maybe_start_pending_case_bootstrap()

    def _update_company_name_suggestions(self, _event=None):
        combo = self.company_name_entry
        if not combo:
            return
        prefix = self.company_name_search_var.get().strip()
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

    def _update_company_nit_suggestions(self, _event=None):
        combo = self.company_nit_entry
        if not combo:
            return
        prefix = "".join(self.company_nit_var.get().split())
        if len(prefix) < 2:
            combo["values"] = ()
            return
        try:
            rows = seguimientos.get_empresas_by_nit_prefix(prefix, limit=12)
        except Exception:
            rows = []
        values = []
        seen = set()
        for row in rows:
            nit = "".join(str(row.get("nit_empresa") or "").split())
            name = str(row.get("nombre_empresa") or "").strip()
            if not nit:
                continue
            label = f"{nit} - {name}" if name else nit
            if nit in seen:
                continue
            seen.add(nit)
            values.append(label)
        combo["values"] = values

    def _apply_linked_company_to_ui(self):
        linked_name = str((self.linked_company or {}).get("nombre_empresa") or "").strip()
        linked_nit = str((self.linked_company or {}).get("nit_empresa") or "").strip()
        if linked_nit:
            self.company_nit_var.set(linked_nit)
        if linked_name:
            self.company_name_search_var.set(linked_name)
        if linked_name and linked_nit:
            self.company_var.set(f"{linked_name} ({linked_nit})")
        elif linked_name:
            self.company_var.set(linked_name)
        else:
            self.company_var.set("(Sin empresa cargada)")
        self.profesional_var.set(
            str((self.linked_company or {}).get("profesional_asignado") or "").strip() or "(Sin asignar)"
        )
        compensar_choice = self._resolve_compensar_choice(self.linked_company)
        self.compensar_var.set(compensar_choice)
        self.case_type_var.set(self._get_case_type_label())
        self._set_company_lookup_enabled(not bool(linked_name or linked_nit))
        self._set_compensar_choice_enabled(bool(linked_name or linked_nit) and not compensar_choice)
        self._refresh_case_flow_visuals()
        if self._pending_case_bootstrap and self._compensar_choice_is_resolved(compensar_choice):
            self._maybe_start_pending_case_bootstrap()

    def _buscar_empresa_manual_por_nit(self, _event=None):
        raw_nit = ui_feedback.get_widget_value(self.company_nit_entry)
        nit = raw_nit.split(" - ", 1)[0].strip()
        if not nit:
            _show_inline_feedback(self, "Ingresa el NIT para buscar la empresa.", state="error")
            return

        def _worker():
            return seguimientos.get_empresa_by_nit(nit)

        def _on_success(company):
            if not company:
                self.compensar_var.set("")
                self._set_compensar_choice_enabled(True)
                _show_inline_feedback(self, "No se encontró empresa para ese NIT.", state="warning")
                return
            self.linked_company = company
            self._apply_linked_company_to_ui()
            _show_inline_feedback(self, "Empresa encontrada por NIT.", state="success")
            self._maybe_start_pending_case_bootstrap()

        def _on_error(exc):
            _show_inline_feedback(self, _log_user_error("followup_case", exc), state="error")

        _run_async_ui_task(
            self,
            busy_attr="_manual_company_lookup_busy",
            widgets=[self.company_nit_entry, self.company_name_entry, self.company_nit_btn, self.company_name_btn],
            loading_button=self.company_nit_btn,
            loading_button_text="Buscando...",
            status_label=self.status_label_widget,
            loading_text="Buscando empresa...",
            loading_state="loading",
            worker=_worker,
            on_success=_on_success,
            on_error=_on_error,
        )

    def _buscar_empresa_manual_por_nombre(self, _event=None):
        name = ui_feedback.get_widget_value(self.company_name_entry)
        if not name:
            _show_inline_feedback(self, "Ingresa el nombre para buscar la empresa.", state="error")
            return
        options = list(self.company_name_entry.cget("values") or [])
        if options:
            exact = next((v for v in options if str(v).strip().casefold() == name.casefold()), None)
            if exact and exact != name:
                self.company_name_search_var.set(str(exact).strip())
                name = str(exact).strip()

        def _worker():
            return seguimientos.get_empresa_by_nombre(name)

        def _on_success(company):
            if not company:
                self.compensar_var.set("")
                self._set_compensar_choice_enabled(True)
                _show_inline_feedback(self, "No se encontró empresa con ese nombre.", state="warning")
                return
            self.linked_company = company
            self._apply_linked_company_to_ui()
            _show_inline_feedback(self, "Empresa encontrada por nombre.", state="success")
            self._maybe_start_pending_case_bootstrap()

        def _on_error(exc):
            _show_inline_feedback(self, _log_user_error("followup_case", exc), state="error")

        _run_async_ui_task(
            self,
            busy_attr="_manual_company_lookup_busy",
            widgets=[self.company_nit_entry, self.company_name_entry, self.company_nit_btn, self.company_name_btn],
            loading_button=self.company_name_btn,
            loading_button_text="Buscando...",
            status_label=self.status_label_widget,
            loading_text="Buscando empresa...",
            loading_state="loading",
            worker=_worker,
            on_success=_on_success,
            on_error=_on_error,
        )

    def _run_loading_job(
        self,
        *,
        title,
        initial_status,
        worker,
        on_success,
        on_error_context="ui_error",
        busy_attr=None,
        busy_widgets=None,
    ):
        if busy_attr and bool(getattr(self, busy_attr, False)):
            return False
        if busy_attr:
            setattr(self, busy_attr, True)
        snapshots = _capture_widget_snapshots(busy_widgets or [])
        for widget in list(busy_widgets or []):
            _disable_widget(widget)
        _set_window_busy_cursor(self, True)
        dialog = LoadingDialog(self, title=title)
        dialog.set_status(initial_status)
        dialog.set_progress(8)
        result = {"value": None, "error": None}

        def _progress(status=None, progress=None):
            _update_loading_async(dialog, status=status, progress=progress)

        def _worker():
            try:
                result["value"] = worker(_progress)
            except Exception as exc:
                result["error"] = exc

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()

        def _check_done():
            if thread.is_alive():
                self.after(180, _check_done)
                return
            dialog.close()
            _restore_widget_snapshots(snapshots)
            _set_window_busy_cursor(self, False)
            if busy_attr:
                setattr(self, busy_attr, False)
            if result["error"] is not None:
                _show_inline_feedback(self, _log_user_error(on_error_context, result["error"]), state="error")
                return False
            on_success(result["value"])
            return True

        self.after(180, _check_done)
        return True

    def _apply_summary_result(self, *, suggestion=None, summary=None, missing_case_status=None):
        if not suggestion and not summary:
            self._refresh_suggestion()
            return
        if not summary:
            summary = {}
        if not suggestion:
            suggestion = {}
        sheet = _friendly_followup_sheet_title(
            suggestion.get("sheet"),
            {"stage_model": list(summary.get("sheet_progress") or []), "base_sheet_name": summary.get("base_sheet_name") or seguimientos.SHEET_BASE},
        )
        msg = suggestion.get("message") or ""
        self.suggestion_var.set(f"{sheet} - {msg}" if sheet or msg else "")
        company_text = summary.get("empresa") or ""
        if not company_text:
            linked_name = str((self.linked_company or {}).get("nombre_empresa") or "").strip()
            linked_nit = str((self.linked_company or {}).get("nit_empresa") or "").strip()
            if linked_name and linked_nit:
                company_text = f"{linked_name} ({linked_nit})"
            else:
                company_text = linked_name
        self.company_var.set(company_text or "(Sin empresa cargada)")
        completed = list(summary.get("completed_sheets") or [])
        if completed:
            self.followups_var.set(", ".join(completed))
        else:
            self.followups_var.set("Sin hojas completadas")
        self._refresh_case_flow_visuals(suggestion=suggestion, summary=summary)
        if missing_case_status:
            self.status_var.set(missing_case_status)
        else:
            self.status_var.set("La persona ya existe en seguimientos.")

    def _load_cedulas(self):
        try:
            options = seguimientos.get_usuarios_reca_cedulas()
        except Exception as exc:
            return self._handle_cedula_load_error(exc)
        return self._apply_cedula_options(options)

    def _reload_cedulas(self):
        self._set_cedula_lookup_hint("Recargando cédulas disponibles...", state="loading")

        def _worker():
            return seguimientos.get_usuarios_reca_cedulas()

        def _on_success(options):
            self._apply_cedula_options(options, show_success_hint=True)

        def _on_error(exc):
            self._handle_cedula_load_error(exc)

        _run_async_ui_task(
            self,
            busy_attr="_cedula_load_busy",
            widgets=[self.cedula_combo, self.buscar_vinculado_btn, self.reload_cedulas_btn],
            loading_button=self.reload_cedulas_btn,
            loading_button_text="Recargando...",
            status_label=self.status_label_widget,
            loading_text="Recargando cédulas disponibles...",
            loading_state="loading",
            worker=_worker,
            on_success=_on_success,
            on_error=_on_error,
        )

    def _filter_cedulas(self, _event=None):
        raw = self.cedula_combo.get().strip()
        if not raw:
            self._filtered_cedulas = list(self.cedula_options)
        else:
            self._filtered_cedulas = [c for c in self.cedula_options if raw in str(c)]
        self.cedula_combo["values"] = self._filtered_cedulas[:50]

    def _buscar_vinculado(self):
        cedula = ui_feedback.get_widget_value(self.cedula_combo)
        if not cedula:
            _show_inline_feedback(self, "Ingresa una cédula para buscar el caso.", state="error")
            return
        normalized = re.sub(r"\D+", "", cedula)

        def _worker(progress):
            progress("Buscando el vinculado en usuarios RECA...", 15)
            row = seguimientos.get_usuario_reca_by_cedula(cedula)
            if not row:
                raise RuntimeError("No se encontró la cédula en usuarios_reca.")
            progress("Resolviendo la empresa asociada desde contratación...", 35)
            try:
                linked_company = seguimientos.get_linked_company_for_user(row)
            except Exception:
                linked_company = {}
            progress("Buscando el caso...", 60)
            case_record = seguimientos.find_case_record(normalized, row.get("nombre_usuario"))
            suggestion = None
            summary = None
            state_error = None
            if case_record:
                progress("Leyendo el estado actual del caso...", 82)
                state_snapshot = _read_followup_case_state(case_record)
                suggestion = state_snapshot.get("suggestion")
                summary = state_snapshot.get("summary")
                state_error = state_snapshot.get("error")
            progress("Organizando la información del caso...", 100)
            return {
                "row": row,
                "normalized_cedula": normalized,
                "linked_company": linked_company,
                "case_record": case_record,
                "suggestion": suggestion,
                "summary": summary,
                "state_error": state_error,
            }

        def _on_success(data):
            row = data.get("row") or {}
            self.user_row = row
            self.linked_company = data.get("linked_company") or {}
            self.user_name_var.set(str(row.get("nombre_usuario") or ""))
            self.user_phone_var.set(str(row.get("telefono_oferente") or ""))
            self.user_role_var.set(str(row.get("cargo_oferente") or ""))
            discapacidad = row.get("discapacidad_detalle") or row.get("discapacidad_usuario") or ""
            self.user_discapacidad_var.set(str(discapacidad))
            self._apply_linked_company_to_ui()
            if not (self.linked_company or {}).get("nombre_empresa") and not (self.linked_company or {}).get("nit_empresa"):
                self.compensar_var.set("")
                self._set_compensar_choice_enabled(True)
            self.case_record = data.get("case_record")
            self.case_path = None
            self._case_suggestion_cache = data.get("suggestion")
            self._case_summary_cache = data.get("summary")
            self._case_cache_key = self._current_case_cache_key(self.case_record, self.case_path)
            self.path_var.set(
                ((self.case_record or {}).get("webViewLink") or "")
                if self.case_record
                else "(Sin Google Sheet aún)"
            )
            self.create_btn.config(state="normal")
            self.open_btn.config(state="normal" if self.case_record else "disabled")
            if not self._get_case_target():
                self._set_pending_case_bootstrap(
                    row=row,
                    normalized_cedula=data.get("normalized_cedula"),
                    linked_company=self.linked_company,
                )
                status_text = (
                    "La cédula existe, pero no tiene caso en seguimientos. "
                    "Confirma si la empresa es Compensar para crear carpeta y Google Sheet."
                )
                self.status_var.set(status_text)
                self._refresh_case_flow_visuals()
                if self._compensar_choice_is_resolved():
                    self._maybe_start_pending_case_bootstrap()
                return
            state_error = data.get("state_error")
            if state_error is not None:
                self._case_suggestion_cache = None
                self._case_summary_cache = None
                self.suggestion_var.set("No fue posible leer la etapa sugerida.")
                self.followups_var.set("")
                self.status_var.set(_format_followup_case_state_error(state_error))
                self._refresh_case_flow_visuals()
                return
            self._apply_summary_result(
                suggestion=data.get("suggestion"),
                summary=data.get("summary"),
            )

        self._run_loading_job(
            title="Buscando vinculado",
            initial_status="Preparando la búsqueda del caso...",
            worker=_worker,
            on_success=_on_success,
            on_error_context="linked_user_search",
            busy_attr="_buscar_vinculado_busy",
            busy_widgets=[self.cedula_combo, getattr(self, "buscar_vinculado_btn", None)],
        )

    def _refresh_suggestion(self):
        case_target = self._get_case_target()
        if not case_target:
            self.suggestion_var.set("Se creará el caso para iniciar en la ficha inicial del proceso.")
            self._apply_linked_company_to_ui()
            self.followups_var.set("Sin hojas completadas")
            self.status_var.set(
                "La cédula existe, pero no tiene caso en seguimientos. "
                "Confirma si la empresa es Compensar para crear carpeta y Google Sheet."
            )
            self._refresh_case_flow_visuals()
            return
        state_snapshot = _read_followup_case_state(case_target)
        suggestion = state_snapshot.get("suggestion")
        summary = state_snapshot.get("summary")
        state_error = state_snapshot.get("error")
        if state_error is not None:
            self.suggestion_var.set("No fue posible leer la etapa sugerida.")
            self._apply_linked_company_to_ui()
            self.followups_var.set("")
            self.status_var.set(_format_followup_case_state_error(state_error))
            self._refresh_case_flow_visuals()
            return
        sheet = _friendly_followup_sheet_title(
            suggestion.get("sheet"),
            {"stage_model": list(summary.get("sheet_progress") or []), "base_sheet_name": summary.get("base_sheet_name") or seguimientos.SHEET_BASE},
        )
        msg = suggestion.get("message") or ""
        self.suggestion_var.set(f"{sheet} - {msg}")
        company_text = summary.get("empresa") or ""
        if company_text:
            self.company_var.set(company_text)
        else:
            self._apply_linked_company_to_ui()
        completed = summary.get("completed_sheets") or []
        if completed:
            self.followups_var.set(", ".join(completed))
        else:
            self.followups_var.set("Sin hojas completadas")
        self.status_var.set("La persona ya existe en seguimientos.")
        self._refresh_case_flow_visuals(suggestion=suggestion, summary=summary)

    def _continue_case_flow(self):
        if not self._get_case_target():
            if self._pending_case_bootstrap and self._compensar_choice_is_resolved():
                self._maybe_start_pending_case_bootstrap()
                return
            _show_inline_feedback(
                self,
                "Todavía no existe el Google Sheet del caso. Confirma la empresa y si es Compensar para continuar.",
                state="warning",
            )
            return
        self._open_editor()

    def _crear_o_actualizar_caso(self):
        if not self.user_row:
            messagebox.showerror("Error", "Primero busca y selecciona una cédula válida.")
            return False
        raw = self.cedula_combo.get().strip()
        cedula = re.sub(r"\D+", "", raw)
        if not cedula:
            messagebox.showerror("Error", "Cédula inválida.")
            return False
        is_compensar = self.compensar_var.get().startswith("Si")
        user_row = dict(self.user_row or {})
        if self.linked_company:
            if self.linked_company.get("nit_empresa"):
                user_row["empresa_nit"] = str(self.linked_company.get("nit_empresa") or "").strip()
            if self.linked_company.get("nombre_empresa"):
                user_row["empresa_nombre"] = str(self.linked_company.get("nombre_empresa") or "").strip()
        try:
            result = seguimientos.ensure_case_record(cedula, user_row, is_compensar=is_compensar)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return False
        self.case_record = result.get("record") or {}
        self.case_path = self.case_record.get("local_path")
        self.path_var.set(self.case_record.get("webViewLink") or self.case_path or "")
        self.open_btn.config(state="normal" if self.case_record else "disabled")
        created = bool(result.get("created"))
        max_seg = result.get("max_seguimientos")
        action = "creado" if created else "actualizado"
        self.status_var.set(
            f"Caso {action} correctamente. Límite de seguimientos: {max_seg}."
        )
        self._refresh_suggestion()
        return True

    def _abrir_preparar_caso(self):
        if not self.user_row:
            messagebox.showerror("Error", "Primero busca y selecciona una cédula válida.")
            return
        raw = self.cedula_combo.get().strip()
        cedula = re.sub(r"\D+", "", raw)
        if not cedula:
            messagebox.showerror("Error", "Cédula inválida.")
            return
        is_compensar = self.compensar_var.get().startswith("Si")
        user_row = dict(self.user_row or {})
        if self.linked_company:
            if self.linked_company.get("nit_empresa"):
                user_row["empresa_nit"] = str(self.linked_company.get("nit_empresa") or "").strip()
            if self.linked_company.get("nombre_empresa"):
                user_row["empresa_nombre"] = str(self.linked_company.get("nombre_empresa") or "").strip()

        def _worker(progress):
            progress("Preparando el caso...", 20)
            result = seguimientos.ensure_case_record(cedula, user_row, is_compensar=is_compensar)
            record = result.get("record") or {}
            case_path = record.get("local_path")
            case_target = (
                record
                if str((record or {}).get("source") or "").strip() == "drive"
                and str((record or {}).get("mime_type") or "").strip() == seguimientos.GOOGLE_SHEETS_MIME
                else case_path
            )
            cache_key = self._current_case_cache_key(record, case_path)
            if (not bool(result.get("created"))) and cache_key and cache_key == self._case_cache_key and self._case_suggestion_cache and self._case_summary_cache:
                progress("Usando el estado del caso ya leído...", 75)
                suggestion = self._case_suggestion_cache
                summary = self._case_summary_cache
                state_error = None
            else:
                progress("Leyendo la etapa sugerida para continuar...", 75)
                state_snapshot = _read_followup_case_state(case_target)
                suggestion = state_snapshot.get("suggestion")
                summary = state_snapshot.get("summary")
                state_error = state_snapshot.get("error")
            progress("Abriendo el editor del caso...", 100)
            return {
                "result": result,
                "suggestion": suggestion,
                "summary": summary,
                "cache_key": cache_key,
                "state_error": state_error,
            }

        def _on_success(data):
            result = data.get("result") or {}
            self.case_record = result.get("record") or {}
            self.case_path = self.case_record.get("local_path")
            self.path_var.set(self.case_record.get("webViewLink") or self.case_path or "")
            self.open_btn.config(state="normal" if self.case_record else "disabled")
            self._case_suggestion_cache = data.get("suggestion")
            self._case_summary_cache = data.get("summary")
            self._case_cache_key = str(data.get("cache_key") or self._current_case_cache_key(self.case_record, self.case_path))
            created = bool(result.get("created"))
            max_seg = result.get("max_seguimientos")
            action = "creado" if created else "actualizado"
            state_error = data.get("state_error")
            if state_error is not None:
                self._case_suggestion_cache = None
                self._case_summary_cache = None
                self.suggestion_var.set("No fue posible leer la etapa sugerida.")
                self.followups_var.set("")
                self.status_var.set(_format_followup_case_state_error(state_error))
                if self._get_case_target():
                    self._open_editor()
                return
            self._apply_summary_result(
                suggestion=data.get("suggestion"),
                summary=data.get("summary"),
            )
            self.status_var.set(f"Caso {action} correctamente. Límite de seguimientos: {max_seg}.")
            if self._get_case_target():
                self._open_editor()

        self._run_loading_job(
            title="Preparando caso",
            initial_status="Validando la cédula y preparando el caso...",
            worker=_worker,
            on_success=_on_success,
            on_error_context="followup_case",
        )

    def _abrir_archivo(self):
        if not self.case_record:
            messagebox.showerror("Error", "No hay archivo para abrir.")
            return
        try:
            open_target = str((self.case_record or {}).get("webViewLink") or "").strip()
            if not open_target:
                raise RuntimeError("No se encontró el enlace del Google Sheet del caso.")
            _open_url_prefer_chrome(open_target)
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo abrir el archivo.\n{exc}")

    def _open_editor(self):
        case_target = self._get_case_target()
        if not case_target:
            messagebox.showerror("Error", "No hay archivo de seguimiento para diligenciar.")
            return
        case_record_snapshot = dict(self.case_record or {})

        def _worker(progress):
            case_record = dict(case_record_snapshot or {})
            progress("Preparando el editor del caso...", 18)
            if not case_record:
                raise RuntimeError("No hay archivo de seguimiento para diligenciar.")
            progress("Leyendo la estructura del caso...", 82)
            bootstrap = _load_followup_editor_bootstrap(case_record)
            progress("Abriendo el editor...", 100)
            return {
                "case_record": case_record,
                "case_path": str(case_record.get("local_path") or self.case_path or "").strip() or None,
                "bootstrap": bootstrap,
            }

        def _on_success(data):
            self.case_record = data.get("case_record") or {}
            self.case_path = data.get("case_path")
            self.path_var.set((self.case_record or {}).get("webViewLink") or self.case_path or "")
            self.open_btn.config(state="normal" if self.case_record else "disabled")
            try:
                editor = SeguimientoEditorWindow(
                    self,
                    self.case_path,
                    case_record=self.case_record,
                    bootstrap=data.get("bootstrap"),
                )
            except Exception as exc:
                messagebox.showerror("Error", _log_user_error("ui_error", exc))
                return
            _focus_window(editor)

        self._run_loading_job(
            title="Abriendo seguimiento",
            initial_status="Preparando el editor del caso...",
            worker=_worker,
            on_success=_on_success,
            on_error_context="followup_case",
        )

    def _close_to_hub(self):
        _return_to_hub(self)
        self.destroy()

    def _open_lsc_window(self):
        empresa = (
            self.linked_company
            if isinstance(self.linked_company, dict) and self.linked_company.get("nombre_empresa")
            else None
        )
        oferentes = []
        if self.user_row and isinstance(self.user_row, dict):
            nombre = (
                self.user_row.get("nombre_usuario")
                or self.user_row.get("nombre_completo")
                or ""
            ).strip()
            cedula = (
                self.user_row.get("cedula_usuario")
                or self.user_row.get("cedula")
                or ""
            ).strip()
            if nombre or cedula:
                oferentes.append(
                    {
                        "nombre_oferente": nombre,
                        "cedula": cedula,
                        "proceso": "Seguimiento",
                    }
                )
        ctx = {"empresa": empresa, "oferentes": oferentes, "source_form": "seguimientos"} if empresa else None
        LSCWindow(self, context=ctx)


class _SeguimientoDateField(tk.Frame):
    _DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y")

    def __init__(
        self,
        master,
        *,
        host,
        textvariable=None,
        width=14,
        allow_text_values=False,
        quick_actions=None,
    ):
        super().__init__(master, bg=COLOR_LIGHT_BG)
        self.host = host
        self.var = textvariable or tk.StringVar(master=self, value="")
        self.always_disabled = False
        self.allow_text_values = bool(allow_text_values)
        self.quick_actions = list(quick_actions or [])

        self.entry = tk.Entry(self, textvariable=self.var, width=width, state="readonly")
        self.entry.pack(side="left")
        self.host._always_readonly_widgets.add(self.entry)

        self.select_button = ttk.Button(self, text="Seleccionar", width=11, command=self.open_picker)
        self.select_button.pack(side="left", padx=(6, 0))
        self.clear_button = ttk.Button(self, text="Limpiar", width=8, command=self.clear)
        self.clear_button.pack(side="left", padx=(4, 0))
        self.quick_buttons = []
        for label, value in self.quick_actions:
            button = ttk.Button(
                self,
                text=str(label or ""),
                width=max(10, len(str(label or "")) + 2),
                command=lambda resolved=value: self._set_quick_value(resolved),
            )
            button.pack(side="left", padx=(4, 0))
            self.quick_buttons.append(button)

    def _parse_date(self, value):
        text = str(value or "").strip()
        if not text:
            return None
        for fmt in self._DATE_FORMATS:
            try:
                return datetime.strptime(text, fmt).date()
            except Exception:
                continue
        return None

    def _center_dialog(self, dialog):
        dialog.update_idletasks()
        parent = self.winfo_toplevel()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width() or parent.winfo_reqwidth()
        parent_h = parent.winfo_height() or parent.winfo_reqheight()
        dialog_w = dialog.winfo_width() or dialog.winfo_reqwidth()
        dialog_h = dialog.winfo_height() or dialog.winfo_reqheight()
        x = parent_x + max(0, (parent_w - dialog_w) // 2)
        y = parent_y + max(0, (parent_h - dialog_h) // 2)
        dialog.geometry(f"+{x}+{y}")

    def set_enabled(self, editable):
        enabled = bool(editable) and not self.always_disabled
        entry_state = "normal" if (enabled and self.allow_text_values) else ("readonly" if enabled else "disabled")
        self.entry.config(state=entry_state)
        self.select_button.config(state="normal" if enabled else "disabled")
        self.clear_button.config(state="normal" if enabled else "disabled")
        for button in list(self.quick_buttons or []):
            button.config(state="normal" if enabled else "disabled")

    def set(self, value):
        if self.allow_text_values:
            self.var.set(str(value or ""))
            return
        normalized = self._parse_date(value)
        self.var.set(normalized.strftime("%Y-%m-%d") if normalized else "")

    def get(self):
        return str(self.var.get() or "")

    def clear(self):
        self.var.set("")

    def _set_quick_value(self, value):
        self.var.set(str(value or ""))

    def open_picker(self):
        if str(self.select_button.cget("state") or "") == "disabled":
            return
        initial_date = self._parse_date(self.var.get()) or date.today()
        dialog = tk.Toplevel(self)
        dialog.title("Seleccionar fecha")
        dialog.configure(bg=COLOR_LIGHT_BG)
        dialog.resizable(False, False)
        dialog.transient(self.winfo_toplevel())

        result = {"value": None}
        calendar = Calendar(
            dialog,
            selectmode="day",
            year=initial_date.year,
            month=initial_date.month,
            day=initial_date.day,
            date_pattern="yyyy-mm-dd",
        )
        calendar.pack(padx=12, pady=(12, 8))

        actions = tk.Frame(dialog, bg=COLOR_LIGHT_BG)
        actions.pack(fill="x", padx=12, pady=(0, 12))

        def _accept(_event=None):
            try:
                chosen = calendar.selection_get()
            except Exception:
                chosen = None
            result["value"] = chosen
            dialog.destroy()

        def _cancel(_event=None):
            dialog.destroy()

        ttk.Button(actions, text="Cancelar", command=_cancel).pack(side="right")
        ttk.Button(actions, text="Aceptar", command=_accept).pack(side="right", padx=(0, 8))

        dialog.bind("<Return>", _accept)
        dialog.bind("<Escape>", _cancel)
        self._center_dialog(dialog)
        dialog.grab_set()
        calendar.focus_set()
        dialog.wait_window()

        if result["value"] is not None:
            self.var.set(result["value"].strftime("%Y-%m-%d"))

# ── VENTANA: SeguimientoEditorWindow — editor de un seguimiento individual ───


FOLLOWUP_LOCAL_BASE_EDITABLE_FIELDS = (
    "fecha_visita",
    "modalidad",
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
    "fecha_firma_contrato",
    "fecha_inicio_contrato",
    "fecha_fin_contrato",
    "apoyos_ajustes",
    "funciones_1_5",
    "funciones_6_10",
)


def _merge_followup_local_payload(base_payload, draft_payload, *, save_kind):
    merged = copy.deepcopy(base_payload or {})
    local_payload = dict(draft_payload or {})
    if str(save_kind or "").strip() == "base":
        for field_id in FOLLOWUP_LOCAL_BASE_EDITABLE_FIELDS:
            if field_id in local_payload:
                merged[field_id] = copy.deepcopy(local_payload.get(field_id))
        return merged
    if str(save_kind or "").strip() == "followup":
        for key, value in local_payload.items():
            merged[key] = copy.deepcopy(value)
    return merged


class SeguimientoEditorWindow(tk.Toplevel, FormMousewheelMixin):
    def __init__(self, parent, case_path, case_record=None, bootstrap=None):
        super().__init__(parent)
        self.owner = parent
        self.case_path = case_path
        self.case_record = case_record or {}
        self.case_target = (
            self.case_record
            if (
                str((self.case_record or {}).get("source") or "").strip() == "drive"
                and str((self.case_record or {}).get("mime_type") or "").strip()
                == seguimientos.GOOGLE_SHEETS_MIME
            )
            else self.case_path
        )
        self.title("Seguimientos - Diligenciamiento")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1200x820")
        _maximize_window(self)

        bootstrap = bootstrap or _load_followup_editor_bootstrap(self.case_target)
        self.meta = dict((bootstrap or {}).get("meta") or {})
        self.workflow = dict((bootstrap or {}).get("workflow") or {})
        suggestion = dict(
            (bootstrap or {}).get("suggestion")
            or _build_followup_suggestion_from_workflow(self.workflow)
        )
        self.max_seg = int(
            self.workflow.get("max_seguimientos") or self.meta.get("max_seguimientos") or 3
        )
        self.base_sheet_name = str(
            self.workflow.get("base_sheet_name")
            or self.meta.get("base_sheet_name")
            or seguimientos.SHEET_BASE
        )
        self.sheet_options = list(self.workflow.get("visible_sheets") or [self.base_sheet_name])
        self.sheet_stage_model = []
        self.sheet_display_var = tk.StringVar(value="")
        self._sheet_title_by_name = {}
        self._sheet_name_by_title = {}
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
        self.base_company_widgets = {}
        self.base_field_widgets = {}
        self.base_date_widgets = {}
        self.base_date_na_button = None
        self.base_modalidad_widget = None
        self._always_readonly_widgets = set()
        self.company_name_combo = None

        self.follow_vars = {}
        self.follow_widgets = {}
        self.follow_text = {}
        self.follow_item_obs = []
        self.follow_item_auto = []
        self.follow_item_auto_widgets = []
        self.follow_item_emp = []
        self.follow_item_emp_widgets = []
        self.follow_emp_eval = []
        self.follow_emp_eval_widgets = []
        self.follow_emp_obs = []
        self.follow_asistentes = []
        self.follow_date_widget = None
        self.current_followup_index = None
        self.save_button = None
        self.continue_stage_button = None
        self.stage_overview_container = None
        self._sheet_remote_payload = {}
        self._overwrite_fields = []
        self._rendered_sheet_name = ""
        self._sheet_autosave_after_id = None
        self._sheet_autosave_trace_tokens = []
        self._sheet_autosave_suspend = False
        self._sheet_autosave_busy = False
        self._sheet_autosave_pending_request = None
        self._sheet_autosave_last_fingerprint = ""
        self._sheet_autosave_debounce_ms = 1500
        self._base_saved_in_session = False
        self._saved_followup_indices_in_session = []
        self._last_saved_followup_index = None

        self._build_header()
        self._build_controls()
        self._build_scroller()
        self.protocol("WM_DELETE_WINDOW", self._close_editor)
        self._refresh_sheet_selector_model()
        self._render_selected_sheet()

    def _get_hub_window(self):
        if isinstance(self.owner, SeguimientosWindow) and isinstance(self.owner.master, HubWindow):
            return self.owner.master
        if isinstance(self.master, HubWindow):
            return self.master
        return None

    def _build_local_draft_metadata(self):
        hub = self._get_hub_window()
        user_login = ""
        if hub and hasattr(hub, "_get_current_user_login"):
            try:
                user_login = str(hub._get_current_user_login() or "").strip().lower()
            except Exception:
                user_login = ""
        company_name = ""
        if isinstance(self.base_vars, dict) and self.base_vars.get("nombre_empresa") is not None:
            try:
                company_name = str(self.base_vars["nombre_empresa"].get() or "").strip()
            except Exception:
                company_name = ""
        if not company_name:
            company_name = str(
                (self.case_record or {}).get("folder_name")
                or (self.case_record or {}).get("file_name")
                or (self.case_record or {}).get("cedula")
                or self.case_path
                or "Seguimientos"
            ).strip()
        return {
            "user_login": user_login,
            "case_record": copy.deepcopy(self.case_record or {}),
            "case_path": str(self.case_path or ""),
            "company_name": company_name,
            "case_label": company_name,
            "folder_name": str((self.case_record or {}).get("folder_name") or ""),
        }

    def _refresh_sheet_selector_model(self):
        self.sheet_stage_model = list(self.workflow.get("stage_model") or self.workflow.get("sheet_progress") or [])
        self._sheet_title_by_name = {}
        self._sheet_name_by_title = {}
        display_values = []
        for sheet_name in list(self.sheet_options or []):
            title = _friendly_followup_sheet_title(
                sheet_name,
                {"stage_model": self.sheet_stage_model, "base_sheet_name": self.base_sheet_name},
                base_sheet_name=self.base_sheet_name,
            )
            self._sheet_title_by_name[sheet_name] = title
            self._sheet_name_by_title[title] = sheet_name
            display_values.append(title)
        if getattr(self, "sheet_combo", None) is not None:
            self.sheet_combo["values"] = display_values
        current_title = self._sheet_title_by_name.get(self.sheet_var.get(), self.sheet_var.get())
        self.sheet_display_var.set(current_title)

    def _on_sheet_combo_selected(self):
        selected_title = str(self.sheet_display_var.get() or "").strip()
        target_sheet = self._sheet_name_by_title.get(selected_title) or self.sheet_var.get()
        if target_sheet != self.sheet_var.get():
            self.sheet_var.set(target_sheet)
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
            text=f"Caso: {((self.case_record or {}).get('webViewLink') or self.case_path or '')}",
            font=("Arial", 10),
            bg=COLOR_LIGHT_BG,
            fg="#333333",
            wraplength=1120,
            justify="left",
        ).pack(anchor="w")

    def _build_controls(self):
        controls = tk.Frame(self, bg=COLOR_LIGHT_BG)
        controls.pack(fill="x", padx=FORM_PADX, pady=(0, 8))

        tk.Label(controls, text="Etapa:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
            side="left", padx=(0, 6)
        )
        self.sheet_combo = ttk.Combobox(
            controls,
            textvariable=self.sheet_display_var,
            values=(),
            state="readonly",
            width=45,
        )
        self.sheet_combo.pack(side="left")
        self.sheet_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_sheet_combo_selected())

        self.save_button = ttk.Button(controls, text="Guardar etapa", command=self._save_current_sheet)
        self.save_button.pack(side="right")
        self.continue_stage_button = ttk.Button(
            controls,
            text="Ir a etapa sugerida",
            style="Primary.TButton",
            command=self._go_to_suggested_stage,
        )
        self.continue_stage_button.pack(side="right", padx=(0, 8))
        ttk.Button(controls, text="Cerrar", command=self._close_editor).pack(side="right", padx=(0, 8))
        ttk.Button(controls, text="Abrir en Drive", command=self._open_excel).pack(
            side="right", padx=(0, 8)
        )

        self.status_label_widget = tk.Label(
            self,
            textvariable=self.status_var,
            font=("Arial", 10),
            fg="#333333",
            bg=COLOR_LIGHT_BG,
            anchor="w",
            justify="left",
            wraplength=1140,
        )
        self.status_label_widget.pack(fill="x", padx=FORM_PADX, pady=(0, 8))
        self.stage_overview_container = tk.Frame(self, bg=COLOR_LIGHT_BG)
        self.stage_overview_container.pack(fill="x", padx=FORM_PADX, pady=(0, 10))

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
        self._bind_mousewheel(self.canvas, self)

    def _refresh_scroller_layout(self):
        try:
            self.update_idletasks()
        except Exception:
            pass
        try:
            bbox = self.canvas.bbox("all")
            if bbox:
                self.canvas.configure(scrollregion=bbox)
        except Exception:
            pass

    def _normalize_overwrite_value(self, value):
        return str(value or "").strip()

    def _bind_overwrite_refresh_widget(self, widget):
        if widget is None or getattr(widget, "_followup_overwrite_bound", False):
            return

        def _refresh(_event=None):
            self._refresh_overwrite_highlights()

        for sequence in ("<KeyRelease>", "<FocusOut>", "<<ComboboxSelected>>", "<<Modified>>"):
            try:
                widget.bind(sequence, _refresh, add="+")
            except Exception:
                continue
        widget._followup_overwrite_bound = True

    def _set_widget_overwrite_state(self, widget, changed):
        if widget is None:
            return
        try:
            klass = str(widget.winfo_class() or "")
        except Exception:
            klass = ""
        if klass in {"TEntry", "TCombobox"}:
            original_style = getattr(widget, "_overwrite_original_style", None)
            if original_style is None:
                try:
                    original_style = str(widget.cget("style") or "")
                except Exception:
                    original_style = ""
                widget._overwrite_original_style = original_style
            target_style = f"Changed.{klass}" if changed else original_style
            try:
                widget.configure(style=target_style)
            except Exception:
                pass
            return
        if klass in {"Entry", "Text", "DateEntry"}:
            original_bg = getattr(widget, "_overwrite_original_bg", None)
            if original_bg is None:
                try:
                    original_bg = widget.cget("bg")
                except Exception:
                    original_bg = "white"
                widget._overwrite_original_bg = original_bg
            try:
                state = str(widget.cget("state") or "")
            except Exception:
                state = ""
            if klass == "Entry" and state == "readonly":
                original_readonly_bg = getattr(widget, "_overwrite_original_readonly_bg", None)
                if original_readonly_bg is None:
                    try:
                        original_readonly_bg = widget.cget("readonlybackground")
                    except Exception:
                        original_readonly_bg = original_bg
                    widget._overwrite_original_readonly_bg = original_readonly_bg
                try:
                    widget.configure(
                        readonlybackground=COLOR_FIELD_WARNING_BG if changed else original_readonly_bg
                    )
                except Exception:
                    pass
                return
            try:
                widget.configure(bg=COLOR_FIELD_WARNING_BG if changed else original_bg)
            except Exception:
                pass

    def _reset_overwrite_tracking(self, remote_payload=None):
        self._sheet_remote_payload = copy.deepcopy(remote_payload or {})
        self._overwrite_fields = []

    def _register_overwrite_field(self, *, label, getter, widgets=None, original_value=""):
        entry = {
            "label": str(label or "").strip(),
            "getter": getter,
            "widgets": list(widgets or []),
            "original_value": self._normalize_overwrite_value(original_value),
        }
        self._overwrite_fields.append(entry)
        for widget in list(entry["widgets"]):
            self._bind_overwrite_refresh_widget(widget)

    def _get_overwrite_changes(self):
        changes = []
        for entry in list(self._overwrite_fields or []):
            getter = entry.get("getter")
            try:
                current_value = getter() if callable(getter) else ""
            except Exception:
                current_value = ""
            current = self._normalize_overwrite_value(current_value)
            original = self._normalize_overwrite_value(entry.get("original_value"))
            changed = bool(original) and current != original
            for widget in list(entry.get("widgets") or []):
                self._set_widget_overwrite_state(widget, changed)
            if changed:
                changes.append(
                    {
                        "label": str(entry.get("label") or "").strip(),
                        "original": original,
                        "current": current,
                    }
                )
        return changes

    def _refresh_overwrite_highlights(self):
        self._get_overwrite_changes()

    def _build_overwrite_warning_message(self, changes):
        labels = [str((item or {}).get("label") or "").strip() for item in list(changes or []) if str((item or {}).get("label") or "").strip()]
        if not labels:
            return ""
        preview = labels[:8]
        details = "\n".join(f"- {label}" for label in preview)
        remaining = len(labels) - len(preview)
        if remaining > 0:
            details += f"\n- y {remaining} campo(s) más"
        return (
            "Se van a sobreescribir datos que ya estaban diligenciados en esta etapa:\n\n"
            f"{details}\n\n"
            "Los campos marcados en amarillo cambiarán su valor actual.\n"
            "¿Quieres guardar estos cambios?"
        )

    def _confirm_overwrite_changes(self):
        changes = self._get_overwrite_changes()
        if not changes:
            return True
        labels = [str(item.get("label") or "").strip() for item in changes if str(item.get("label") or "").strip()]
        preview = ", ".join(labels[:4])
        if len(labels) > 4:
            preview = f"{preview} y {len(labels) - 4} campo(s) más"
        _show_inline_feedback(
            self,
            f"Vas a sobreescribir datos ya diligenciados: {preview}. Confirma si quieres guardarlos.",
            state="warning",
        )
        return bool(
            messagebox.askyesno(
                "Confirmar sobreescritura",
                self._build_overwrite_warning_message(changes),
                parent=self,
            )
        )

    def _go_to_suggested_stage(self):
        suggested_sheet = str((self.workflow or {}).get("suggested_sheet") or "").strip()
        if not suggested_sheet or suggested_sheet == self.sheet_var.get():
            return
        if suggested_sheet == seguimientos.SHEET_FINAL:
            self.status_var.set("El resultado final es de solo lectura.")
        self.sheet_var.set(suggested_sheet)
        self._refresh_sheet_selector_model()
        self._render_selected_sheet()

    def _refresh_continue_stage_button(self):
        button = self.continue_stage_button
        if button is None:
            return
        suggested_sheet = str((self.workflow or {}).get("suggested_sheet") or "").strip()
        current_sheet = str(self.sheet_var.get() or "").strip()
        if not suggested_sheet or suggested_sheet == current_sheet:
            button.pack_forget()
            return
        suggested_title = _friendly_followup_sheet_title(
            suggested_sheet,
            {"stage_model": self.sheet_stage_model, "base_sheet_name": self.base_sheet_name},
            base_sheet_name=self.base_sheet_name,
        )
        button.configure(text=f"Continuar a {suggested_title}")
        button.pack(side="right", padx=(0, 8))

    def _render_stage_overview(self):
        container = self.stage_overview_container
        if container is None:
            return
        for child in container.winfo_children():
            child.destroy()
        stage_model = list(self.sheet_stage_model or [])
        if not stage_model:
            return

        current_title = _friendly_followup_sheet_title(
            (self.workflow or {}).get("suggested_sheet"),
            {"stage_model": stage_model, "base_sheet_name": self.base_sheet_name},
            base_sheet_name=self.base_sheet_name,
        ) or "Pendiente"
        history_titles = [
            str((entry or {}).get("title") or "")
            for entry in stage_model
            if str((entry or {}).get("stage_id") or "").startswith("followup_")
            and str((entry or {}).get("status") or "") == "completed"
        ]
        cards = [
            {
                "title": "Ficha inicial del proceso",
                "status": next(
                    (
                        dict(entry)
                        for entry in stage_model
                        if str((entry or {}).get("stage_id") or "") == "base_process"
                    ),
                    {},
                ),
            },
            {
                "title": "Seguimiento actual",
                "status": {
                    "status": next(
                        (
                            str((entry or {}).get("status") or "")
                            for entry in stage_model
                            if str((entry or {}).get("sheet_name") or "").strip()
                            == str((self.workflow or {}).get("suggested_sheet") or "").strip()
                        ),
                        "pending",
                    ),
                    "coverage_percent": next(
                        (
                            int((entry or {}).get("coverage_percent") or 0)
                            for entry in stage_model
                            if str((entry or {}).get("sheet_name") or "").strip()
                            == str((self.workflow or {}).get("suggested_sheet") or "").strip()
                        ),
                        0,
                    ),
                    "is_suggested": True,
                    "helper_text": current_title,
                },
            },
            {
                "title": "Historial de seguimientos",
                "status": {
                    "status": "completed" if history_titles else "pending",
                    "coverage_percent": 100 if history_titles else 0,
                    "is_suggested": False,
                    "helper_text": ", ".join(history_titles) if history_titles else "Aún no hay seguimientos completos.",
                },
            },
            {
                "title": "Resultado final",
                "status": next(
                    (
                        dict(entry)
                        for entry in stage_model
                        if str((entry or {}).get("stage_id") or "") == "final_result"
                    ),
                    {},
                ),
            },
        ]

        for index, item in enumerate(cards):
            status_entry = dict(item.get("status") or {})
            palette = _get_followup_stage_palette(
                status_entry.get("status"),
                is_suggested=bool(status_entry.get("is_suggested")),
            )
            card = tk.Frame(
                container,
                bg=palette["bg"],
                highlightbackground=palette["accent"],
                highlightthickness=2,
                bd=0,
            )
            card.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else 8, 0))
            container.grid_columnconfigure(index, weight=1)
            inner = tk.Frame(card, bg=palette["bg"], padx=12, pady=10)
            inner.pack(fill="both", expand=True)
            tk.Label(
                inner,
                text=item.get("title") or "",
                font=FONT_LABEL,
                fg=palette["accent"],
                bg=palette["bg"],
                anchor="w",
                justify="left",
            ).pack(anchor="w")
            tk.Label(
                inner,
                text=_format_followup_stage_status(
                    status_entry.get("status"),
                    coverage_percent=status_entry.get("coverage_percent"),
                    is_suggested=bool(status_entry.get("is_suggested")),
                ),
                font=("Arial", 9, "bold"),
                fg=palette["muted"],
                bg=palette["bg"],
                anchor="w",
                justify="left",
            ).pack(anchor="w", pady=(4, 2))
            tk.Label(
                inner,
                text=str(status_entry.get("helper_text") or ""),
                font=("Arial", 9),
                fg="#333333",
                bg=palette["bg"],
                anchor="w",
                justify="left",
                wraplength=240,
            ).pack(anchor="w")

    def _clear_content(self):
        for child in self.content_frame.winfo_children():
            child.destroy()

    def _close_editor(self):
        try:
            self._schedule_sheet_autosave(delay_ms=0)
            self._flush_pending_sheet_autosave()
        except Exception as exc:
            _log_capture(f"followup_editor_local_draft_close_failed case={self.case_target!r} err={exc}")
        self.destroy()

    def _cancel_sheet_autosave_timer(self):
        after_id = getattr(self, "_sheet_autosave_after_id", None)
        if not after_id:
            return
        try:
            self.after_cancel(after_id)
        except Exception:
            pass
        self._sheet_autosave_after_id = None

    def _clear_sheet_autosave_traces(self):
        for var, token in list(getattr(self, "_sheet_autosave_trace_tokens", []) or []):
            try:
                var.trace_remove("write", token)
            except Exception:
                continue
        self._sheet_autosave_trace_tokens = []

    def _bind_sheet_autosave_var(self, var):
        return

    def _bind_sheet_autosave_widget(self, widget, *, include_focus_out=True):
        if widget is None or getattr(widget, "_seguimiento_autosave_bound", False):
            return
        if include_focus_out:
            try:
                widget.bind(
                    "<FocusOut>",
                    lambda _event=None, w=self: w._schedule_sheet_autosave(),
                    add="+",
                )
            except Exception:
                pass
        widget._seguimiento_autosave_bound = True

    def _bind_sheet_autosave_date_field(self, field):
        if field is None or getattr(field, "_seguimiento_autosave_bound", False):
            return
        self._bind_sheet_autosave_widget(field.entry)
        self._bind_sheet_autosave_widget(field.select_button)
        self._bind_sheet_autosave_widget(field.clear_button)
        for button in list(getattr(field, "quick_buttons", []) or []):
            self._bind_sheet_autosave_widget(button, include_focus_out=False)

        original_clear = field.clear
        original_open_picker = field.open_picker
        original_quick_set = field._set_quick_value

        def _clear_and_schedule():
            original_clear()
            self._schedule_sheet_autosave()
            self._refresh_overwrite_highlights()

        def _open_picker_and_schedule():
            before = field.get()
            original_open_picker()
            if field.get() != before:
                self._schedule_sheet_autosave()
                self._refresh_overwrite_highlights()

        def _set_quick_value_and_schedule(value):
            before = field.get()
            original_quick_set(value)
            if field.get() != before:
                self._schedule_sheet_autosave()
                self._refresh_overwrite_highlights()

        field.clear = _clear_and_schedule
        field.open_picker = _open_picker_and_schedule
        field._set_quick_value = _set_quick_value_and_schedule
        try:
            field.select_button.configure(command=field.open_picker)
        except Exception:
            pass
        try:
            field.clear_button.configure(command=field.clear)
        except Exception:
            pass
        for button, (_label, value) in zip(list(getattr(field, "quick_buttons", []) or []), list(getattr(field, "quick_actions", []) or [])):
            try:
                button.configure(command=lambda resolved=value, current_field=field: current_field._set_quick_value(resolved))
            except Exception:
                continue
        field._seguimiento_autosave_bound = True

    def _install_sheet_autosave_bindings(self):
        self._clear_sheet_autosave_traces()
        for _path, widget in _iter_widget_paths(self.content_frame):
            if isinstance(widget, (tk.Entry, tk.Text, ttk.Combobox)):
                self._bind_sheet_autosave_widget(widget)
        for field in list(self.base_date_widgets.values()) + list(self.base_dates_1) + list(self.base_dates_2):
            self._bind_sheet_autosave_date_field(field)
        if self.follow_date_widget is not None:
            self._bind_sheet_autosave_date_field(self.follow_date_widget)

    def _build_sheet_save_request(self, *, sheet_name=None, validate_base=False):
        sheet = str(sheet_name or self._rendered_sheet_name or self.sheet_var.get() or "").strip()
        if not sheet or sheet == seguimientos.SHEET_FINAL:
            return None
        if sheet != self.base_sheet_name and not self._is_sheet_editable(sheet):
            return None
        if sheet == self.base_sheet_name:
            payload = self._collect_base_payload()
            if validate_base and not self._validate_base_payload(payload):
                return None
            request = {
                "sheet": sheet,
                "save_kind": "base",
                "payload": payload,
                "followup_index": None,
            }
        elif sheet.startswith(seguimientos.SHEET_PREFIX):
            match = re.search(r"(\d+)$", sheet)
            if not match:
                return None
            idx = int(match.group(1))
            request = {
                "sheet": sheet,
                "save_kind": "followup",
                "payload": self._collect_followup_payload(idx),
                "followup_index": idx,
            }
        else:
            return None
        request["fingerprint"] = self._fingerprint_sheet_request(request)
        return request

    def _fingerprint_sheet_request(self, request):
        if not isinstance(request, dict):
            return ""
        payload = {
            "sheet": str(request.get("sheet") or "").strip(),
            "save_kind": str(request.get("save_kind") or "").strip(),
            "followup_index": request.get("followup_index"),
            "payload": request.get("payload"),
        }
        try:
            encoded = json.dumps(
                payload,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
        except Exception:
            return ""
        return hashlib.sha256(encoded).hexdigest()

    def _load_local_sheet_draft_payload(self, sheet_name, *, save_kind):
        draft = _get_followup_local_sheet_draft(self.case_target, sheet_name)
        if not isinstance(draft, dict):
            return None
        if str(draft.get("save_kind") or "").strip() != str(save_kind or "").strip():
            return None
        payload = draft.get("payload")
        if not isinstance(payload, dict):
            return None
        return copy.deepcopy(payload)

    def _schedule_sheet_autosave(self, delay_ms=None):
        if self._sheet_autosave_suspend:
            return False
        try:
            request = self._build_sheet_save_request()
        except Exception:
            return False
        if not request:
            return False
        fingerprint = str(request.get("fingerprint") or "")
        if fingerprint and fingerprint == str(self._sheet_autosave_last_fingerprint or ""):
            self._sheet_autosave_pending_request = None
            self._cancel_sheet_autosave_timer()
            return False
        self._sheet_autosave_pending_request = request
        self._cancel_sheet_autosave_timer()

        def _run():
            self._sheet_autosave_after_id = None
            self._flush_pending_sheet_autosave()

        try:
            resolved_delay = self._sheet_autosave_debounce_ms if delay_ms is None else int(delay_ms or 0)
        except Exception:
            resolved_delay = self._sheet_autosave_debounce_ms
        try:
            if int(resolved_delay or 0) <= 0:
                self._sheet_autosave_after_id = self.after_idle(_run)
            else:
                self._sheet_autosave_after_id = self.after(int(resolved_delay), _run)
        except Exception:
            self._sheet_autosave_after_id = None
            return False
        return True

    def _flush_pending_sheet_autosave(self):
        self._cancel_sheet_autosave_timer()
        request = dict(self._sheet_autosave_pending_request or {})
        if not request:
            return False
        fingerprint = str(request.get("fingerprint") or "")
        if fingerprint and fingerprint == str(self._sheet_autosave_last_fingerprint or ""):
            self._sheet_autosave_pending_request = None
            return False

        sheet_label = str(request.get("sheet") or "").strip()
        stage_title = self._sheet_title_by_name.get(sheet_label, sheet_label)
        try:
            saved = _save_followup_local_sheet_draft(
                self.case_target,
                request,
                metadata=self._build_local_draft_metadata(),
            )
        except Exception as exc:
            self.status_var.set(f"No se pudo guardar el borrador local de {stage_title}.")
            _show_inline_feedback(self, _log_user_error("save_sheet", exc), state="warning")
            return False
        if not saved:
            return False
        self._sheet_autosave_last_fingerprint = fingerprint
        current_fingerprint = str((self._sheet_autosave_pending_request or {}).get("fingerprint") or "")
        if current_fingerprint == fingerprint:
            self._sheet_autosave_pending_request = None
        self.status_var.set(f"Borrador local guardado en {stage_title}.")
        hub = self._get_hub_window()
        if hub and hasattr(hub, "_refresh_drafts_badge"):
            try:
                hub._refresh_drafts_badge()
            except Exception:
                pass
        return True

    def _execute_sheet_save_worker(self, request, *, trigger, progress=None):
        request = dict(request or {})
        sheet = str(request.get("sheet") or "").strip()
        save_kind = str(request.get("save_kind") or "").strip()
        payload = dict(request.get("payload") or {})
        idx = request.get("followup_index")

        def _progress(status, percent):
            if callable(progress):
                progress(status, percent)

        _progress("Preparando los datos de la hoja...", 12)
        if save_kind == "base":
            _progress("Guardando la ficha inicial...", 42)
            try:
                seguimientos.save_base_payload(self.case_target, payload)
            except PermissionError as exc:
                raise RuntimeError(
                    "No se pudo guardar porque el Excel está abierto en otra aplicación."
                ) from exc
        elif save_kind == "followup":
            _progress(f"Guardando seguimiento {idx}...", 42)
            try:
                seguimientos.save_followup_payload(self.case_target, idx, payload)
            except PermissionError as exc:
                raise RuntimeError(
                    "No se pudo guardar porque el Excel está abierto en otra aplicación."
                ) from exc
        else:
            raise RuntimeError("No se pudo identificar la hoja que se intenta guardar.")

        sync_warning = ""
        if str(trigger or "").strip().lower() == "manual":
            _progress("Sincronizando el caso con Drive...", 72)
            if self.case_record and self.case_path:
                try:
                    seguimientos.sync_case_record_from_local(self.case_record, self.case_path)
                except Exception as exc:
                    sync_warning = str(exc)
            if save_kind == "followup":
                _progress("Actualizando el estado del seguimiento...", 88)
                hub = self._get_hub_window()
                if hub:
                    try:
                        hub.record_followup_completion(
                            case_target=self.case_target,
                            case_path=self.case_path,
                            case_record=self.case_record,
                            followup_index=idx,
                        )
                    except Exception as exc:
                        _log_capture(
                            f"record_followup_completion failed case={self.case_target!r} idx={idx} err={exc}"
                        )
            _progress("Actualizando la pantalla...", 96)

        return {
            "sheet": sheet,
            "sync_warning": sync_warning,
            "fingerprint": str(request.get("fingerprint") or ""),
        }

    def _add_labeled_entry(self, parent, row, label, var, width=40, readonly=False):
        tk.Label(parent, text=label, font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=3
        )
        entry = tk.Entry(parent, textvariable=var, width=width)
        if readonly:
            entry.configure(state="readonly")
            self._always_readonly_widgets.add(entry)
        entry.grid(row=row, column=1, sticky="w", pady=3)
        return entry

    def _add_labeled_date(self, parent, row, label, var, width=20, **field_kwargs):
        tk.Label(parent, text=label, font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=3
        )
        date_field = _SeguimientoDateField(
            parent,
            host=self,
            textvariable=var,
            width=width,
            **field_kwargs,
        )
        date_field.grid(row=row, column=1, sticky="w", pady=3)
        date_field.set(var.get())
        return date_field

    def _add_dictation_subsection(
        self,
        parent,
        *,
        title,
        store,
        key,
        field_id,
        initial_value="",
        height=5,
        bottom_pad=8,
    ):
        section = tk.Frame(parent, bg=COLOR_LIGHT_BG, bd=1, relief="groove")
        section.pack(fill="x", pady=(0, bottom_pad))
        header = tk.Frame(section, bg=COLOR_LIGHT_BG)
        header.pack(fill="x", padx=10, pady=(8, 4))
        title_label = tk.Label(header, text=title, font=FONT_LABEL, bg=COLOR_LIGHT_BG)
        title_label.pack(side="left", anchor="w")
        text_widget = tk.Text(section, height=height, wrap="word")
        text_widget.pack(fill="x", padx=10, pady=(0, 10))
        _attach_autoexpand(text_widget, height, 30)
        attach_dictation(
            text_widget,
            form_id="seguimientos",
            field_id=field_id,
            session_provider=lambda: _supabase_get_access_token(".env"),
            log_fn=_log_capture,
            controls_parent=header,
            anchor_widget=title_label,
            placement="inline_right",
        )
        text_widget.insert("1.0", str(initial_value or ""))
        store[key] = text_widget
        return text_widget

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
        previous_sheet = str(self._rendered_sheet_name or "").strip()
        target_sheet = str(self.sheet_var.get() or "").strip()
        if previous_sheet and previous_sheet != target_sheet:
            self._schedule_sheet_autosave(delay_ms=0)
            self._flush_pending_sheet_autosave()
        self._refresh_sheet_selector_model()
        self._sheet_autosave_suspend = True
        self._clear_sheet_autosave_traces()
        self._clear_content()
        sheet = target_sheet
        if sheet == self.base_sheet_name:
            self._render_sheet_base()
        elif sheet.startswith(seguimientos.SHEET_PREFIX):
            match = re.search(r"(\d+)$", sheet)
            if not match:
                self.status_var.set("No se pudo identificar la etapa de seguimiento.")
                return
            self._render_sheet_followup(int(match.group(1)))
        else:
            self._render_sheet_final()
        self._apply_sheet_access_rules()
        self._rendered_sheet_name = sheet
        self.sheet_display_var.set(self._sheet_title_by_name.get(sheet, sheet))
        self._render_stage_overview()
        self._refresh_continue_stage_button()
        self._install_sheet_autosave_bindings()
        self._refresh_overwrite_highlights()
        self._sheet_autosave_pending_request = None
        try:
            current_request = self._build_sheet_save_request(sheet_name=sheet, validate_base=False)
        except Exception:
            current_request = None
        self._sheet_autosave_last_fingerprint = self._fingerprint_sheet_request(current_request)
        self._sheet_autosave_suspend = False
        self._refresh_scroller_layout()
        try:
            self.after_idle(self._refresh_scroller_layout)
        except Exception:
            pass
        self.canvas.yview_moveto(0)

    def _refresh_workflow_state(self, preferred_sheet=None):
        try:
            bootstrap = _load_followup_editor_bootstrap(self.case_target)
        except RuntimeError as exc:
            self.status_var.set(str(exc))
            return
        self.meta = dict((bootstrap or {}).get("meta") or {})
        self.workflow = dict((bootstrap or {}).get("workflow") or {})
        suggestion = dict(
            (bootstrap or {}).get("suggestion")
            or _build_followup_suggestion_from_workflow(self.workflow)
        )
        self.max_seg = int(
            self.workflow.get("max_seguimientos") or self.meta.get("max_seguimientos") or 3
        )
        self.base_sheet_name = str(
            self.workflow.get("base_sheet_name")
            or self.meta.get("base_sheet_name")
            or seguimientos.SHEET_BASE
        )
        self.sheet_options = list(self.workflow.get("visible_sheets") or [self.base_sheet_name])
        target_sheet = preferred_sheet or self.sheet_var.get().strip()
        if target_sheet not in self.sheet_options:
            target_sheet = (
                suggestion.get("sheet")
                or (self.sheet_options[0] if self.sheet_options else self.base_sheet_name)
            )
        self.sheet_var.set(target_sheet)
        self._refresh_sheet_selector_model()

    def _run_loading_job(
        self,
        *,
        title,
        initial_status,
        worker,
        on_success,
        on_error_context="ui_error",
        busy_attr=None,
        busy_widgets=None,
    ):
        if busy_attr and bool(getattr(self, busy_attr, False)):
            return False
        if busy_attr:
            setattr(self, busy_attr, True)
        snapshots = _capture_widget_snapshots(busy_widgets or [])
        for widget in list(busy_widgets or []):
            _disable_widget(widget)
        _set_window_busy_cursor(self, True)
        dialog = LoadingDialog(self, title=title)
        dialog.set_status(initial_status)
        dialog.set_progress(8)
        result = {"value": None, "error": None}

        def _progress(status=None, progress=None):
            _update_loading_async(dialog, status=status, progress=progress)

        def _worker():
            try:
                result["value"] = worker(_progress)
            except Exception as exc:
                result["error"] = exc

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()

        def _check_done():
            if thread.is_alive():
                self.after(180, _check_done)
                return
            dialog.close()
            _restore_widget_snapshots(snapshots)
            _set_window_busy_cursor(self, False)
            if busy_attr:
                setattr(self, busy_attr, False)
            if result["error"] is not None:
                _show_inline_feedback(self, _log_user_error(on_error_context, result["error"]), state="error")
                return
            on_success(result["value"])

        self.after(180, _check_done)
        return True

    def _is_sheet_editable(self, sheet_name):
        current = str(sheet_name or "").strip()
        if not current or current == seguimientos.SHEET_FINAL:
            return False
        return current in list(self.sheet_options or [])

    def _set_single_widget_edit_state(self, widget, editable):
        if isinstance(widget, _SeguimientoDateField):
            widget.set_enabled(editable)
            return
        try:
            klass = str(widget.winfo_class() or "")
        except Exception:
            klass = ""
        try:
            if klass == "Text":
                widget.config(state="normal" if editable else "disabled")
            elif klass in {"Entry", "TEntry", "DateEntry"}:
                if widget in self._always_readonly_widgets:
                    widget.config(state="readonly" if editable else "disabled")
                else:
                    widget.config(state="normal" if editable else "disabled")
            elif klass == "TCombobox":
                widget.config(state="readonly" if editable else "disabled")
            elif klass in {"Button", "TButton"}:
                widget.config(state="normal" if editable else "disabled")
        except Exception:
            pass

    def _get_base_followup_date_widget(self, followup_number):
        try:
            idx = int(followup_number or 0)
        except Exception:
            return None
        if 1 <= idx <= 3 and len(self.base_dates_1) >= idx:
            return self.base_dates_1[idx - 1]
        if 4 <= idx <= 6:
            pos = idx - 4
            if len(self.base_dates_2) > pos:
                return self.base_dates_2[pos]
        return None

    def _apply_base_followup_date_states(self, active_followup):
        try:
            active_idx = int(active_followup or 0)
        except Exception:
            active_idx = 0
        for idx, widget in enumerate(list(self.base_dates_1) + list(self.base_dates_2), start=1):
            self._set_single_widget_edit_state(widget, bool(active_idx > 0 and idx == active_idx))

    def _set_widget_edit_state(self, widget, editable):
        if isinstance(widget, _SeguimientoDateField):
            widget.set_enabled(editable)
            return
        try:
            klass = str(widget.winfo_class() or "")
        except Exception:
            klass = ""
        try:
            if klass == "Text":
                widget.config(state="normal" if editable else "disabled")
            elif klass in {"Entry", "TEntry", "DateEntry"}:
                if widget in self._always_readonly_widgets:
                    widget.config(state="readonly" if editable else "disabled")
                else:
                    widget.config(state="normal" if editable else "disabled")
            elif klass == "TCombobox":
                widget.config(state="readonly" if editable else "disabled")
            elif klass in {"Button", "TButton"}:
                widget.config(state="normal" if editable else "disabled")
        except Exception:
            pass
        for child in widget.winfo_children():
            self._set_widget_edit_state(child, editable)

    def _apply_sheet_access_rules(self):
        editable = self._is_sheet_editable(self.sheet_var.get())
        self._set_widget_edit_state(self.content_frame, editable)
        current_sheet = self.sheet_var.get().strip()
        suggested_sheet = str((self.workflow or {}).get("suggested_sheet") or "").strip()
        current_title = self._sheet_title_by_name.get(current_sheet, current_sheet)
        suggested_title = self._sheet_title_by_name.get(suggested_sheet, suggested_sheet)

        if current_sheet == str(self.base_sheet_name or "").strip():
            self._apply_base_followup_date_states(0)
        if self.save_button is not None:
            self.save_button.config(state="normal" if editable else "disabled")
        if not editable:
            if current_sheet == seguimientos.SHEET_FINAL:
                self.status_var.set("Resultado final es de solo lectura.")
            return
        if current_sheet == str(self.base_sheet_name or "").strip():
            if suggested_sheet == current_sheet:
                self.status_var.set(
                    str((self.workflow or {}).get("message") or "Editando la ficha inicial del proceso.")
                )
            else:
                self.status_var.set(
                    "Editando la ficha inicial del proceso. "
                    f"Las fechas del historial se alimentan desde cada seguimiento. "
                    f"Etapa sugerida actual: {suggested_title}."
                )
            return
        if current_sheet != suggested_sheet and suggested_sheet:
            self.status_var.set(
                f"Editando {current_title}. Etapa sugerida actual: {suggested_title}. "
                "Puedes continuar aquí si necesitas corregir información."
            )
            return
        if (self.workflow or {}).get("message"):
            self.status_var.set(str((self.workflow or {}).get("message")))

    def _render_sheet_base(self):
        remote_payload = seguimientos.get_base_payload(self.case_target)
        payload = copy.deepcopy(remote_payload)
        local_payload = self._load_local_sheet_draft_payload(self.base_sheet_name, save_kind="base")
        if local_payload:
            payload = _merge_followup_local_payload(payload, local_payload, save_kind="base")
        self.base_company_widgets = {}
        self.base_field_widgets = {}
        self.base_date_widgets = {}
        self.base_date_na_button = None
        self.base_modalidad_widget = None
        self.company_name_combo = None
        self.follow_date_widget = None
        self._reset_overwrite_tracking(remote_payload)
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
                "caja_compensacion",
                "profesional_asignado",
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
                "fecha_firma_contrato",
                "fecha_inicio_contrato",
                "fecha_fin_contrato",
            ]
        }

        visit = tk.LabelFrame(
            self.content_frame,
            text="Datos de visita",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        visit.pack(fill="x", pady=(0, 10))
        visit.grid_columnconfigure(1, weight=1)

        row = 0
        self.base_date_widgets["fecha_visita"] = self._add_labeled_date(
            visit,
            row,
            "Fecha de visita:",
            self.base_vars["fecha_visita"],
            width=18,
        )
        self._register_overwrite_field(
            label="Fecha de visita",
            getter=lambda field=self.base_date_widgets["fecha_visita"]: field.get().strip(),
            widgets=[self.base_date_widgets["fecha_visita"].entry],
            original_value=remote_payload.get("fecha_visita"),
        )
        row += 1
        tk.Label(visit, text="Modalidad:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=row, column=0, sticky="w", padx=(0, 8), pady=3
        )
        mod_combo = ttk.Combobox(
            visit,
            textvariable=self.base_vars["modalidad"],
            values=seguimientos.MODALIDAD_OPTIONS,
            state="readonly",
            width=26,
        )
        mod_combo.grid(row=row, column=1, sticky="w", pady=3)
        self.base_modalidad_widget = mod_combo
        self.base_field_widgets["modalidad"] = mod_combo
        self._register_overwrite_field(
            label="Modalidad",
            getter=lambda var=self.base_vars["modalidad"]: var.get().strip(),
            widgets=[mod_combo],
            original_value=remote_payload.get("modalidad"),
        )

        company = tk.LabelFrame(
            self.content_frame,
            text="Empresa",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        company.pack(fill="x", pady=(0, 10))
        company.grid_columnconfigure(1, weight=1)

        row = 0
        for field_id, label, width in [
            ("nombre_empresa", "Nombre empresa:", 72),
            ("nit_empresa", "NIT:", 25),
            ("ciudad_empresa", "Ciudad/Municipio:", 40),
            ("direccion_empresa", "Dirección:", 70),
            ("correo_1", "Correo:", 60),
            ("telefono_empresa", "Teléfonos:", 40),
            ("contacto_empresa", "Contacto empresa:", 45),
            ("cargo", "Cargo empresa:", 45),
            ("asesor", "Asesor:", 40),
            ("sede_empresa", "Sede Compensar:", 30),
            ("caja_compensacion", "Caja de compensación:", 30),
            ("profesional_asignado", "Profesional asignado RECA:", 40),
        ]:
            widget = self._add_labeled_entry(
                company,
                row,
                label,
                self.base_vars[field_id],
                width=width,
                readonly=True,
            )
            self.base_company_widgets[field_id] = widget
            self.base_field_widgets[field_id] = widget
            self._register_overwrite_field(
                label=label.rstrip(":"),
                getter=lambda var=self.base_vars[field_id]: var.get().strip(),
                widgets=[widget],
                original_value=remote_payload.get(field_id),
            )
            row += 1

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
            widget = self._add_labeled_entry(vinc, r, f"{label}:", self.base_vars[key], width=50)
            self.base_field_widgets[key] = widget
            self._register_overwrite_field(
                label=label,
                getter=lambda var=self.base_vars[key]: var.get().strip(),
                widgets=[widget],
                original_value=remote_payload.get(key),
            )
            r += 1
        self.base_date_widgets["fecha_firma_contrato"] = self._add_labeled_date(
            vinc,
            r,
            "Fecha firma contrato:",
            self.base_vars["fecha_firma_contrato"],
            width=18,
        )
        self._register_overwrite_field(
            label="Fecha firma contrato",
            getter=lambda field=self.base_date_widgets["fecha_firma_contrato"]: field.get().strip(),
            widgets=[self.base_date_widgets["fecha_firma_contrato"].entry],
            original_value=remote_payload.get("fecha_firma_contrato"),
        )
        r += 1
        self.base_date_widgets["fecha_inicio_contrato"] = self._add_labeled_date(
            vinc,
            r,
            "Fecha inicio contrato:",
            self.base_vars["fecha_inicio_contrato"],
            width=18,
        )
        self._register_overwrite_field(
            label="Fecha inicio contrato",
            getter=lambda field=self.base_date_widgets["fecha_inicio_contrato"]: field.get().strip(),
            widgets=[self.base_date_widgets["fecha_inicio_contrato"].entry],
            original_value=remote_payload.get("fecha_inicio_contrato"),
        )
        r += 1
        self.base_date_widgets["fecha_fin_contrato"] = self._add_labeled_date(
            vinc,
            r,
            "Fecha fin contrato:",
            self.base_vars["fecha_fin_contrato"],
            width=18,
            allow_text_values=True,
            quick_actions=[("No aplica", "No aplica")],
        )
        self.base_date_na_button = next(
            iter(getattr(self.base_date_widgets["fecha_fin_contrato"], "quick_buttons", []) or []),
            None,
        )
        self._register_overwrite_field(
            label="Fecha fin contrato",
            getter=lambda field=self.base_date_widgets["fecha_fin_contrato"]: field.get().strip(),
            widgets=[
                self.base_date_widgets["fecha_fin_contrato"].entry,
                *list(getattr(self.base_date_widgets["fecha_fin_contrato"], "quick_buttons", []) or []),
            ],
            original_value=remote_payload.get("fecha_fin_contrato"),
        )
        r += 1

        support = tk.LabelFrame(
            self.content_frame,
            text="Funciones y apoyos",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        support.pack(fill="x", pady=(0, 10))
        tk.Label(
            support,
            text="Apoyos y/o ajustes razonables",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).pack(anchor="w")
        self._add_dictation_subsection(
            support,
            title="Apoyos y/o ajustes razonables requeridos:",
            store=self.base_text,
            key="apoyos_ajustes",
            field_id="base:apoyos_ajustes",
            initial_value=payload.get("apoyos_ajustes", ""),
            height=4,
            bottom_pad=10,
        )
        self._register_overwrite_field(
            label="Apoyos y/o ajustes razonables",
            getter=lambda widget=self.base_text["apoyos_ajustes"]: widget.get("1.0", tk.END).strip(),
            widgets=[self.base_text["apoyos_ajustes"]],
            original_value=remote_payload.get("apoyos_ajustes"),
        )

        funcs = tk.Frame(support, bg=COLOR_LIGHT_BG)
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
            self._register_overwrite_field(
                label=f"Función {i + 1}",
                getter=lambda entry=e1: entry.get().strip(),
                widgets=[e1],
                original_value=(remote_payload.get("funciones_1_5") or [""] * 5)[i] if i < len(remote_payload.get("funciones_1_5") or []) else "",
            )

            e2 = tk.Entry(funcs, width=60)
            e2.grid(row=i + 1, column=1, sticky="w", padx=(18, 0), pady=2)
            if i < len(vals_2):
                e2.insert(0, vals_2[i] or "")
            self.base_func_entries_2.append(e2)
            self._register_overwrite_field(
                label=f"Función {i + 6}",
                getter=lambda entry=e2: entry.get().strip(),
                widgets=[e2],
                original_value=(remote_payload.get("funciones_6_10") or [""] * 5)[i] if i < len(remote_payload.get("funciones_6_10") or []) else "",
            )

        dates = tk.Frame(support, bg=COLOR_LIGHT_BG)
        dates.pack(fill="x")
        tk.Label(
            dates,
            text="Línea de tiempo de seguimientos",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        self.base_dates_1 = []
        self.base_dates_2 = []
        d1 = payload.get("seguimiento_fechas_1_3") or []
        d2 = payload.get("seguimiento_fechas_4_6") or []
        stage_lookup = {
            str((entry or {}).get("stage_id") or ""): dict(entry or {})
            for entry in list(self.sheet_stage_model or [])
        }
        for idx in range(1, self.max_seg + 1):
            row_idx = ((idx - 1) // 3) + 1
            col_idx = (idx - 1) % 3
            entry = dict(stage_lookup.get(f"followup_{idx}") or {})
            palette = _get_followup_stage_palette(
                entry.get("status"),
                is_suggested=bool(entry.get("is_suggested")),
            )
            card = tk.Frame(
                dates,
                bg=palette["bg"],
                highlightbackground=palette["accent"],
                highlightthickness=2,
                bd=0,
                padx=10,
                pady=8,
            )
            card.grid(row=row_idx, column=col_idx, sticky="nsew", padx=(0 if col_idx == 0 else 8, 0), pady=6)
            dates.grid_columnconfigure(col_idx, weight=1)
            tk.Label(
                card,
                text=f"Seguimiento {idx}",
                font=FONT_LABEL,
                fg=palette["accent"],
                bg=palette["bg"],
                anchor="w",
            ).pack(anchor="w")
            date_value = d1[idx - 1] if idx <= 3 and idx - 1 < len(d1) else ""
            if idx > 3:
                date_pos = idx - 4
                date_value = d2[date_pos] if date_pos < len(d2) else ""
            field_var = tk.StringVar(value=date_value or "")
            field = _SeguimientoDateField(card, host=self, textvariable=field_var, width=12)
            field.pack(anchor="w", pady=(6, 4))
            field.set(date_value or "")
            field.always_disabled = True
            field.set_enabled(False)
            if idx <= 3:
                self.base_dates_1.append(field)
            else:
                self.base_dates_2.append(field)
            tk.Label(
                card,
                text=_format_followup_stage_status(
                    entry.get("status") or "pending",
                    coverage_percent=entry.get("coverage_percent"),
                    is_suggested=bool(entry.get("is_suggested")),
                ),
                font=("Arial", 9, "bold"),
                fg=palette["muted"],
                bg=palette["bg"],
                anchor="w",
                justify="left",
                wraplength=220,
            ).pack(anchor="w")

        self.status_var.set("Editando la ficha inicial del proceso.")
        self._refresh_overwrite_highlights()

    def _build_followup_eval_actions(self, parent, specs, *, row, columnspan):
        actions = tk.LabelFrame(
            parent,
            text="Acciones rápidas de diligenciamiento",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=10,
            pady=8,
        )
        actions.grid(row=row, column=0, columnspan=columnspan, sticky="w", pady=(0, 10))
        for row_index, (label, group_key) in enumerate(specs):
            action_row = tk.Frame(actions, bg=COLOR_LIGHT_BG)
            action_row.pack(fill="x", pady=(0, 6 if row_index < len(specs) - 1 else 0))
            tk.Label(
                action_row,
                text=label,
                font=FONT_LABEL,
                bg=COLOR_LIGHT_BG,
                width=22,
                anchor="w",
            ).pack(side="left", padx=(0, 8))
            buttons = tk.Frame(action_row, bg=COLOR_LIGHT_BG)
            buttons.pack(side="left", fill="x", expand=True)
            for idx, option in enumerate(seguimientos.EVAL_OPTIONS):
                ttk.Button(
                    buttons,
                    text=option,
                    width=18,
                    command=lambda selected=option, current_group=group_key: self._set_followup_eval_group_values(
                        current_group,
                        selected,
                    ),
                ).grid(
                    row=idx // 3,
                    column=idx % 3,
                    sticky="w",
                    padx=(0, 8),
                    pady=(0, 4),
                )

    def _set_followup_eval_group_values(self, group_key, value):
        group_map = {
            "auto": self.follow_item_auto,
            "item_empresa": self.follow_item_emp,
            "empresa_eval": self.follow_emp_eval,
        }
        fields = list(group_map.get(str(group_key or "").strip()) or [])
        for var in fields:
            try:
                var.set(str(value or ""))
            except Exception:
                continue
        status_map = {
            "auto": "autoevaluación",
            "item_empresa": "evaluación de la empresa",
            "empresa_eval": "evaluación empresarial",
        }
        group_label = status_map.get(str(group_key or "").strip(), "evaluación")
        self.status_var.set(f"Se aplicó '{value}' a toda la {group_label}.")
        self._refresh_overwrite_highlights()
        self._schedule_sheet_autosave()

    def _render_sheet_followup(self, idx):
        self.current_followup_index = idx
        remote_payload = seguimientos.get_followup_payload(self.case_target, idx)
        payload = copy.deepcopy(remote_payload)
        local_payload = self._load_local_sheet_draft_payload(
            f"{seguimientos.SHEET_PREFIX}{idx}",
            save_kind="followup",
        )
        if local_payload:
            payload = _merge_followup_local_payload(payload, local_payload, save_kind="followup")
        self.follow_vars = {
            "modalidad": tk.StringVar(value=str(payload.get("modalidad") or "")),
            "seguimiento_numero": tk.StringVar(value=str(idx)),
            "fecha_seguimiento": tk.StringVar(value=str(payload.get("fecha_seguimiento") or "")),
            "tipo_apoyo": tk.StringVar(value=str(payload.get("tipo_apoyo") or "")),
        }
        self.follow_widgets = {}
        self.follow_date_widget = None
        self.follow_item_auto_widgets = []
        self.follow_item_emp_widgets = []
        self.follow_emp_eval_widgets = []
        self._reset_overwrite_tracking(remote_payload)
        top = tk.LabelFrame(
            self.content_frame,
            text="Datos del seguimiento",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        top.pack(fill="x", pady=(0, 10))

        tk.Label(top, text="Modalidad:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=0, sticky="w"
        )
        modalidad_combo = ttk.Combobox(
            top,
            textvariable=self.follow_vars["modalidad"],
            values=seguimientos.MODALIDAD_OPTIONS,
            state="readonly",
            width=24,
        )
        modalidad_combo.grid(row=0, column=1, sticky="w")
        self.follow_widgets["modalidad"] = modalidad_combo
        self._register_overwrite_field(
            label="Modalidad",
            getter=lambda var=self.follow_vars["modalidad"]: var.get().strip(),
            widgets=[modalidad_combo],
            original_value=remote_payload.get("modalidad"),
        )
        tk.Label(top, text="Seguimiento #:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=0, column=2, sticky="w", padx=(24, 6)
        )
        tk.Label(
            top,
            textvariable=self.follow_vars["seguimiento_numero"],
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
            width=8,
            anchor="w",
        ).grid(row=0, column=3, sticky="w")
        self.follow_date_widget = self._add_labeled_date(
            top,
            1,
            "Fecha seguimiento:",
            self.follow_vars["fecha_seguimiento"],
            width=18,
        )
        self._register_overwrite_field(
            label="Fecha seguimiento",
            getter=lambda field=self.follow_date_widget: field.get().strip(),
            widgets=[self.follow_date_widget.entry],
            original_value=remote_payload.get("fecha_seguimiento"),
        )
        tk.Label(top, text="Tipo de apoyo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=2, sticky="w", padx=(24, 6)
        )
        tipo_apoyo_combo = ttk.Combobox(
            top,
            textvariable=self.follow_vars["tipo_apoyo"],
            values=seguimientos.TIPO_APOYO_OPTIONS,
            state="readonly",
            width=34,
        )
        tipo_apoyo_combo.grid(row=1, column=3, sticky="w")
        self.follow_widgets["tipo_apoyo"] = tipo_apoyo_combo
        self._register_overwrite_field(
            label="Tipo de apoyo",
            getter=lambda var=self.follow_vars["tipo_apoyo"]: var.get().strip(),
            widgets=[tipo_apoyo_combo],
            original_value=remote_payload.get("tipo_apoyo"),
        )
        if idx > 1:
            ttk.Button(
                top,
                text="Copiar datos del seguimiento anterior",
                command=self._copy_previous_followup_values,
            ).grid(row=0, column=4, rowspan=2, sticky="w", padx=(24, 0))

        items = tk.LabelFrame(
            self.content_frame,
            text="Desempeño del vinculado",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        items.pack(fill="x", pady=(0, 10))
        self._build_followup_eval_actions(
            items,
            [
                ("Aplicar a autoevaluación:", "auto"),
                ("Aplicar a eval. empresa:", "item_empresa"),
            ],
            row=0,
            columnspan=4,
        )
        tk.Label(items, text="Ítem", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(row=1, column=0, sticky="w")
        tk.Label(items, text="Observación", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=1, sticky="w", padx=(8, 0)
        )
        tk.Label(items, text="Autoevaluación", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=2, sticky="w", padx=(8, 0)
        )
        tk.Label(items, text="Eval. empresa", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=3, sticky="w", padx=(8, 0)
        )

        self.follow_item_obs = []
        self.follow_item_auto = []
        self.follow_item_auto_widgets = []
        self.follow_item_emp = []
        self.follow_item_emp_widgets = []
        labels = payload.get("item_labels") or []
        obs_vals = payload.get("item_observaciones") or []
        auto_vals = payload.get("item_autoevaluacion") or []
        emp_vals = payload.get("item_eval_empresa") or []
        for i, label in enumerate(labels):
            tk.Label(items, text=label, bg=COLOR_LIGHT_BG, anchor="w", justify="left").grid(
                row=i + 2, column=0, sticky="w", pady=2
            )
            e_obs = tk.Entry(items, width=40)
            e_obs.grid(row=i + 2, column=1, sticky="w", padx=(8, 0), pady=2)
            if i < len(obs_vals):
                e_obs.insert(0, obs_vals[i] or "")
            self.follow_item_obs.append(e_obs)
            self._register_overwrite_field(
                label=f"Observación del vinculado: {label}",
                getter=lambda entry=e_obs: entry.get().strip(),
                widgets=[e_obs],
                original_value=(remote_payload.get("item_observaciones") or [])[i] if i < len(remote_payload.get("item_observaciones") or []) else "",
            )

            v_auto = tk.StringVar(value=auto_vals[i] if i < len(auto_vals) else "")
            c_auto = ttk.Combobox(
                items, textvariable=v_auto, values=seguimientos.EVAL_OPTIONS, state="readonly", width=20
            )
            c_auto.grid(row=i + 2, column=2, sticky="w", padx=(8, 0), pady=2)
            self.follow_item_auto.append(v_auto)
            self.follow_item_auto_widgets.append(c_auto)
            self._register_overwrite_field(
                label=f"Autoevaluación: {label}",
                getter=lambda var=v_auto: var.get().strip(),
                widgets=[c_auto],
                original_value=(remote_payload.get("item_autoevaluacion") or [])[i] if i < len(remote_payload.get("item_autoevaluacion") or []) else "",
            )

            v_emp = tk.StringVar(value=emp_vals[i] if i < len(emp_vals) else "")
            c_emp = ttk.Combobox(
                items, textvariable=v_emp, values=seguimientos.EVAL_OPTIONS, state="readonly", width=20
            )
            c_emp.grid(row=i + 2, column=3, sticky="w", padx=(8, 0), pady=2)
            self.follow_item_emp.append(v_emp)
            self.follow_item_emp_widgets.append(c_emp)
            self._register_overwrite_field(
                label=f"Evaluación de empresa: {label}",
                getter=lambda var=v_emp: var.get().strip(),
                widgets=[c_emp],
                original_value=(remote_payload.get("item_eval_empresa") or [])[i] if i < len(remote_payload.get("item_eval_empresa") or []) else "",
            )

        middle = tk.LabelFrame(
            self.content_frame,
            text="Evaluación de la empresa",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        middle.pack(fill="x", pady=(0, 10))
        self._build_followup_eval_actions(
            middle,
            [("Aplicar a evaluación empresarial:", "empresa_eval")],
            row=0,
            columnspan=3,
        )

        tk.Label(middle, text="Ítem", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(row=1, column=0, sticky="w", pady=(8, 2))
        tk.Label(middle, text="Evaluación", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=1, sticky="w", padx=(8, 0), pady=(8, 2)
        )
        tk.Label(middle, text="Observación", font=FONT_LABEL, bg=COLOR_LIGHT_BG).grid(
            row=1, column=2, sticky="w", padx=(8, 0), pady=(8, 2)
        )

        self.follow_emp_eval = []
        self.follow_emp_eval_widgets = []
        self.follow_emp_obs = []
        emp_labels = payload.get("empresa_item_labels") or []
        emp_eval_vals = payload.get("empresa_eval") or []
        emp_obs_vals = payload.get("empresa_observacion") or []
        for i, label in enumerate(emp_labels):
            tk.Label(middle, text=label, bg=COLOR_LIGHT_BG, anchor="w", justify="left").grid(
                row=i + 2, column=0, sticky="w", pady=2
            )
            v = tk.StringVar(value=emp_eval_vals[i] if i < len(emp_eval_vals) else "")
            emp_combo = ttk.Combobox(
                middle, textvariable=v, values=seguimientos.EVAL_OPTIONS, state="readonly", width=20
            )
            emp_combo.grid(row=i + 2, column=1, sticky="w", padx=(8, 0), pady=2)
            self.follow_emp_eval.append(v)
            self.follow_emp_eval_widgets.append(emp_combo)
            self._register_overwrite_field(
                label=f"Evaluación empresarial: {label}",
                getter=lambda var=v: var.get().strip(),
                widgets=[emp_combo],
                original_value=(remote_payload.get("empresa_eval") or [])[i] if i < len(remote_payload.get("empresa_eval") or []) else "",
            )
            e = tk.Entry(middle, width=45)
            e.grid(row=i + 2, column=2, sticky="w", padx=(8, 0), pady=2)
            if i < len(emp_obs_vals):
                e.insert(0, emp_obs_vals[i] or "")
            self.follow_emp_obs.append(e)
            self._register_overwrite_field(
                label=f"Observación empresarial: {label}",
                getter=lambda entry=e: entry.get().strip(),
                widgets=[e],
                original_value=(remote_payload.get("empresa_observacion") or [])[i] if i < len(remote_payload.get("empresa_observacion") or []) else "",
            )

        txt = tk.LabelFrame(
            self.content_frame,
            text="Situación encontrada y estrategias",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        txt.pack(fill="x", pady=(0, 10))
        self._add_dictation_subsection(
            txt,
            title="Situación encontrada:",
            store=self.follow_text,
            key="situacion_encontrada",
            field_id=f"followup_{idx}:situacion_encontrada",
            initial_value=payload.get("situacion_encontrada") or "",
            height=5,
        )
        self._register_overwrite_field(
            label="Situación encontrada",
            getter=lambda widget=self.follow_text["situacion_encontrada"]: widget.get("1.0", tk.END).strip(),
            widgets=[self.follow_text["situacion_encontrada"]],
            original_value=remote_payload.get("situacion_encontrada"),
        )
        self._add_dictation_subsection(
            txt,
            title="Estrategias:",
            store=self.follow_text,
            key="estrategias_ajustes",
            field_id=f"followup_{idx}:estrategias_ajustes",
            initial_value=payload.get("estrategias_ajustes") or "",
            height=5,
            bottom_pad=0,
        )
        self._register_overwrite_field(
            label="Estrategias y ajustes",
            getter=lambda widget=self.follow_text["estrategias_ajustes"]: widget.get("1.0", tk.END).strip(),
            widgets=[self.follow_text["estrategias_ajustes"]],
            original_value=remote_payload.get("estrategias_ajustes"),
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
            original_asistentes = list(remote_payload.get("asistentes") or [])
            original_entry = original_asistentes[i] if i < len(original_asistentes) else {}
            self._register_overwrite_field(
                label=f"Asistente {i + 1} nombre",
                getter=lambda widget=e_name: _get_input_value(widget),
                widgets=[e_name],
                original_value=(original_entry or {}).get("nombre"),
            )
            self._register_overwrite_field(
                label=f"Asistente {i + 1} cargo",
                getter=lambda widget=e_cargo: _get_input_value(widget),
                widgets=[e_cargo],
                original_value=(original_entry or {}).get("cargo"),
            )

        self.status_var.set(f"Editando Seguimiento {idx}.")
        self._refresh_overwrite_highlights()

    def _set_entry_value(self, entry, value):
        entry.delete(0, tk.END)
        entry.insert(0, str(value or ""))

    def _copy_previous_followup_values(self):
        idx = int(self.current_followup_index or 0)
        if idx <= 1:
            _show_inline_feedback(
                self,
                "El seguimiento 1 no tiene un seguimiento anterior para copiar.",
                state="warning",
            )
            return
        try:
            previous = seguimientos.get_followup_payload(self.case_target, idx - 1)
        except Exception as exc:
            _show_inline_feedback(self, _log_user_error("followup_case", exc), state="error")
            return
        local_previous = self._load_local_sheet_draft_payload(
            f"{seguimientos.SHEET_PREFIX}{idx - 1}",
            save_kind="followup",
        )
        if local_previous:
            previous = _merge_followup_local_payload(previous, local_previous, save_kind="followup")

        self.follow_vars["modalidad"].set(str(previous.get("modalidad") or ""))
        self.follow_vars["tipo_apoyo"].set(str(previous.get("tipo_apoyo") or ""))

        prev_item_obs = previous.get("item_observaciones") or []
        prev_item_auto = previous.get("item_autoevaluacion") or []
        prev_item_emp = previous.get("item_eval_empresa") or []
        prev_emp_eval = previous.get("empresa_eval") or []
        prev_emp_obs = previous.get("empresa_observacion") or []

        for i, entry in enumerate(self.follow_item_obs):
            self._set_entry_value(entry, prev_item_obs[i] if i < len(prev_item_obs) else "")
        for i, var in enumerate(self.follow_item_auto):
            var.set(str(prev_item_auto[i] if i < len(prev_item_auto) else ""))
        for i, var in enumerate(self.follow_item_emp):
            var.set(str(prev_item_emp[i] if i < len(prev_item_emp) else ""))
        for i, var in enumerate(self.follow_emp_eval):
            var.set(str(prev_emp_eval[i] if i < len(prev_emp_eval) else ""))
        for i, entry in enumerate(self.follow_emp_obs):
            self._set_entry_value(entry, prev_emp_obs[i] if i < len(prev_emp_obs) else "")
        prev_asistentes = previous.get("asistentes") or []
        for i, (name_widget, cargo_widget) in enumerate(self.follow_asistentes):
            current_asistente = prev_asistentes[i] if i < len(prev_asistentes) else {}
            _set_input_value(name_widget, current_asistente.get("nombre") or "")
            _set_input_value(cargo_widget, current_asistente.get("cargo") or "")

        self.status_var.set(
            f"Se copiaron los datos del seguimiento {idx - 1}. "
            "No se copiaron la fecha, situación encontrada ni estrategias."
        )
        self._refresh_overwrite_highlights()
        self._schedule_sheet_autosave()

    def _render_sheet_final(self):
        self._reset_overwrite_tracking({})
        card = tk.LabelFrame(
            self.content_frame,
            text="Resultado final",
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            padx=12,
            pady=10,
        )
        card.pack(fill="x")
        tk.Label(
            card,
            text=(
                "Este bloque consolida automáticamente el resultado del caso.\n"
                "No se diligencia manualmente: se actualiza al guardar la ficha inicial y los seguimientos."
            ),
            bg=COLOR_LIGHT_BG,
            justify="left",
            anchor="w",
            font=("Arial", 10),
        ).pack(fill="x")
        self.status_var.set("Resultado final es de solo lectura.")

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
        self._refresh_overwrite_highlights()
        self._schedule_sheet_autosave()

    def _buscar_empresa_por_nit(self):
        nit = self.base_vars.get("nit_empresa").get().strip() if self.base_vars.get("nit_empresa") else ""
        if not nit:
            _show_inline_feedback(self, "Ingresa NIT para buscar empresa.", state="error")
            return
        try:
            company = seguimientos.get_empresa_by_nit(nit)
        except Exception as exc:
            _show_inline_feedback(self, _log_user_error("followup_case", exc), state="error")
            return
        if not company:
            _show_inline_feedback(self, "No se encontró empresa para ese NIT.", state="warning")
            return
        self._apply_company_data(company)
        _show_inline_feedback(self, "Empresa encontrada por NIT.", state="success")

    def _buscar_empresa_por_nombre(self):
        name = (
            self.base_vars.get("nombre_empresa").get().strip()
            if self.base_vars.get("nombre_empresa")
            else ""
        )
        if not name:
            _show_inline_feedback(self, "Ingresa nombre de empresa para buscar.", state="error")
            return
        try:
            company = seguimientos.get_empresa_by_nombre(name)
        except Exception as exc:
            _show_inline_feedback(self, _log_user_error("followup_case", exc), state="error")
            return
        if not company:
            _show_inline_feedback(self, "No se encontró empresa con ese nombre.", state="warning")
            return
        self._apply_company_data(company)
        _show_inline_feedback(self, "Empresa encontrada por nombre.", state="success")

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
            "caja_compensacion": "caja_compensacion",
            "profesional_asignado": "profesional_asignado",
        }
        for field, key in mapping.items():
            if field in self.base_vars:
                self.base_vars[field].set(str(company.get(key) or ""))
        self.status_var.set("Datos de empresa cargados desde Supabase.")
        self._refresh_overwrite_highlights()
        self._schedule_sheet_autosave()

    def _collect_base_payload(self):
        payload = {k: v.get().strip() for k, v in self.base_vars.items()}
        payload["apoyos_ajustes"] = self.base_text["apoyos_ajustes"].get("1.0", tk.END).strip()
        payload["funciones_1_5"] = [e.get().strip() for e in self.base_func_entries_1]
        payload["funciones_6_10"] = [e.get().strip() for e in self.base_func_entries_2]
        payload["seguimiento_fechas_1_3"] = [e.get().strip() for e in self.base_dates_1]
        payload["seguimiento_fechas_4_6"] = [e.get().strip() for e in self.base_dates_2]
        return payload

    def _validate_base_payload(self, payload, *, focus_invalid=True):
        date_widget = self.base_date_widgets.get("fecha_visita")
        date_entry = getattr(date_widget, "entry", None)
        if date_entry is not None:
            try:
                if date_entry.winfo_exists():
                    ui_feedback.register_field(self, "fecha_visita", date_entry)
            except Exception:
                pass
        if self.base_modalidad_widget is not None:
            try:
                if self.base_modalidad_widget.winfo_exists():
                    ui_feedback.register_field(self, "modalidad", self.base_modalidad_widget)
            except Exception:
                pass
        ui_feedback.clear_field_errors(self, ["fecha_visita", "modalidad"])
        missing = []
        if not str((payload or {}).get("fecha_visita") or "").strip():
            missing.append("Fecha visita")
            ui_feedback.set_field_error(self, "fecha_visita", "Completa la fecha de visita.")
        if not str((payload or {}).get("modalidad") or "").strip():
            missing.append("Modalidad")
            ui_feedback.set_field_error(self, "modalidad", "Selecciona la modalidad.")
        if not missing:
            return True
        _show_inline_feedback(
            self,
            "Completa fecha de visita y modalidad antes de guardar la ficha inicial.",
            state="error",
        )
        if focus_invalid:
            if "Fecha visita" in missing:
                widget = self.base_date_widgets.get("fecha_visita")
                if widget is not None:
                    try:
                        if widget.select_button.winfo_exists():
                            widget.select_button.focus_set()
                    except Exception:
                        pass
            elif self.base_modalidad_widget is not None:
                try:
                    if self.base_modalidad_widget.winfo_exists():
                        self.base_modalidad_widget.focus_set()
                except Exception:
                    pass
        return False

    def _collect_followup_payload(self, idx):
        payload = {
            "modalidad": self.follow_vars["modalidad"].get().strip(),
            "seguimiento_numero": str(idx),
            "fecha_seguimiento": self.follow_vars["fecha_seguimiento"].get().strip(),
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

    def _sheet_title(self, sheet_name):
        return _friendly_followup_sheet_title(
            sheet_name,
            {"stage_model": self.sheet_stage_model, "base_sheet_name": self.base_sheet_name},
            base_sheet_name=self.base_sheet_name,
        )

    def _sheet_sort_key(self, sheet_name):
        current = str(sheet_name or "").strip()
        if current == str(self.base_sheet_name or "").strip():
            return (0, 0)
        match = re.search(r"(\d+)$", current)
        if match:
            return (1, int(match.group(1)))
        return (2, current.casefold())

    def _mark_session_sheet_saved(self, request):
        request = dict(request or {})
        save_kind = str(request.get("save_kind") or "").strip()
        if save_kind == "base":
            self._base_saved_in_session = True
            return
        if save_kind != "followup":
            return
        try:
            idx = int(request.get("followup_index") or 0)
        except Exception:
            return
        if idx <= 0:
            return
        self._saved_followup_indices_in_session = [
            value for value in list(self._saved_followup_indices_in_session or []) if value != idx
        ]
        self._saved_followup_indices_in_session.append(idx)
        self._last_saved_followup_index = idx

    def _validate_sheet_request(self, request, *, focus_invalid=False):
        request = dict(request or {})
        save_kind = str(request.get("save_kind") or "").strip()
        if save_kind != "base":
            return True
        return self._validate_base_payload(request.get("payload") or {}, focus_invalid=focus_invalid)

    def _build_pending_sheet_save_requests(self, current_sheet):
        requests_by_sheet = {}
        local_drafts = _list_followup_local_sheet_drafts(self.case_target)
        for draft in list(local_drafts.values()):
            if not isinstance(draft, dict):
                continue
            sheet_name = str(draft.get("sheet") or "").strip()
            if not sheet_name or sheet_name == seguimientos.SHEET_FINAL:
                continue
            payload = draft.get("payload")
            if not isinstance(payload, dict):
                continue
            requests_by_sheet[sheet_name] = {
                "sheet": sheet_name,
                "save_kind": str(draft.get("save_kind") or "").strip(),
                "followup_index": draft.get("followup_index"),
                "payload": copy.deepcopy(payload),
                "fingerprint": str(draft.get("fingerprint") or ""),
                "_was_dirty": True,
            }

        current_request = self._build_sheet_save_request(sheet_name=current_sheet, validate_base=False)
        if current_request:
            current_fingerprint = str(current_request.get("fingerprint") or "")
            current_request["_was_dirty"] = bool(
                current_sheet in requests_by_sheet
                or current_fingerprint != str(self._sheet_autosave_last_fingerprint or "")
            )
            requests_by_sheet[current_sheet] = current_request

        requests = []
        for sheet_name in sorted(requests_by_sheet.keys(), key=self._sheet_sort_key):
            request = dict(requests_by_sheet.get(sheet_name) or {})
            if not request:
                continue
            if not self._validate_sheet_request(
                request,
                focus_invalid=(sheet_name == str(current_sheet or "").strip()),
            ):
                raise RuntimeError(
                    f"Completa los campos obligatorios antes de guardar {self._sheet_title(sheet_name)}."
                )
            requests.append(request)
        return requests

    def _ask_generate_followup_pdf(self, *, has_followups):
        message = (
            "Las etapas con cambios quedaron guardadas en Google Sheets.\n\n"
            "¿Deseas generar el PDF de cierre ahora?"
        )
        if has_followups:
            message = (
                "Las etapas con cambios quedaron guardadas en Google Sheets.\n\n"
                "¿Deseas generar el PDF de cierre ahora? "
                "Luego podrás elegir el seguimiento que irá junto con la ficha inicial."
            )
        return bool(messagebox.askyesno("Generar PDF", message, parent=self))

    def _resolve_default_pdf_followup_index(self, candidates, saved_followups_in_batch):
        candidate_indices = {
            int(item.get("followup_index") or 0)
            for item in list(candidates or [])
            if int(item.get("followup_index") or 0) > 0
        }
        preferred = []
        if saved_followups_in_batch:
            preferred.append(int(saved_followups_in_batch[-1]))
        if self._last_saved_followup_index:
            preferred.append(int(self._last_saved_followup_index))
        preferred.extend(reversed(list(self._saved_followup_indices_in_session or [])))
        for idx in preferred:
            if idx in candidate_indices:
                return idx
        ordered_candidates = sorted(candidate_indices)
        return ordered_candidates[-1] if ordered_candidates else None

    def _prompt_pdf_followup_choice(self, candidates, *, default_index=None):
        items = [dict(item or {}) for item in list(candidates or [])]
        if not items:
            return None

        options = []
        option_by_label = {}
        for item in items:
            idx = int(item.get("followup_index") or 0)
            if idx <= 0:
                continue
            title = str(item.get("title") or f"Seguimiento {idx}").strip()
            fecha = str(item.get("fecha_seguimiento") or "").strip()
            label = f"{title} ({fecha})" if fecha else title
            options.append((label, idx))
            option_by_label[label] = idx
        if not options:
            return None

        default_label = options[-1][0]
        for label, idx in options:
            if idx == int(default_index or 0):
                default_label = label
                break

        dialog = tk.Toplevel(self)
        dialog.title("Seleccionar seguimiento")
        dialog.configure(bg=COLOR_LIGHT_BG)
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        result = {"value": None}
        selected_var = tk.StringVar(value=default_label)

        body = tk.Frame(dialog, bg=COLOR_LIGHT_BG, padx=24, pady=20)
        body.pack(fill="both", expand=True)

        tk.Label(
            body,
            text="Elige el seguimiento que irá en el PDF.",
            font=FONT_SECTION,
            fg=COLOR_PURPLE,
            bg=COLOR_LIGHT_BG,
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            body,
            text="La ficha inicial del proceso siempre se incluirá en el mismo PDF.",
            font=FONT_LABEL,
            bg=COLOR_LIGHT_BG,
            fg="#333333",
            justify="left",
            wraplength=420,
            anchor="w",
        ).pack(fill="x", pady=(8, 10))

        combo = ttk.Combobox(
            body,
            textvariable=selected_var,
            values=[label for label, _idx in options],
            state="readonly",
            width=48,
        )
        combo.pack(fill="x")

        actions = tk.Frame(body, bg=COLOR_LIGHT_BG)
        actions.pack(fill="x", pady=(18, 0))

        def _accept():
            result["value"] = option_by_label.get(str(selected_var.get() or "").strip())
            dialog.destroy()

        def _cancel():
            result["value"] = None
            dialog.destroy()

        ttk.Button(actions, text="Cancelar", command=_cancel).pack(side="right")
        ttk.Button(actions, text="Generar PDF", style="Primary.TButton", command=_accept).pack(
            side="right",
            padx=(0, 8),
        )

        dialog.protocol("WM_DELETE_WINDOW", _cancel)
        dialog.update_idletasks()
        width = max(dialog.winfo_reqwidth(), 470)
        height = dialog.winfo_reqheight()
        x = self.winfo_rootx() + max(0, (self.winfo_width() - width) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - height) // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.lift()
        try:
            combo.focus_set()
        except Exception:
            pass
        dialog.wait_window()
        return result["value"]

    def _start_followup_pdf_export(self, *, followup_index=None):
        def _worker(progress):
            progress("Preparando la exportación PDF...", 18)
            export_bundle = seguimientos.build_pdf_export_bundle(
                self.case_target,
                followup_index=followup_index,
            )
            progress("Encolando el PDF en Drive...", 70)
            pdf_folder_id = drive_upload._get_pdf_folder_id()
            job_id = _enqueue_pdf_export_job(
                sheet_file_id=str((self.case_record or {}).get("file_id") or "").strip(),
                tipo_acta=str(export_bundle.get("tipo_acta") or "").strip(),
                fecha_servicio=str(export_bundle.get("fecha_servicio") or "").strip(),
                acta_metadata=export_bundle.get("acta_metadata") or {},
                extra_name=export_bundle.get("extra_name"),
                pdf_folder_id=pdf_folder_id,
                company_name=str(
                    ((export_bundle.get("acta_metadata") or {}).get("nombre_empresa"))
                    or (self.case_record or {}).get("folder_name")
                    or ""
                ).strip(),
                selected_sheet_names=export_bundle.get("selected_sheet_names") or [],
                temp_parent_folder_id=str(export_bundle.get("temp_parent_folder_id") or "").strip(),
            )
            return {
                "job_id": job_id,
                "pdf_folder_id": pdf_folder_id,
                "company_name": str(
                    ((export_bundle.get("acta_metadata") or {}).get("nombre_empresa"))
                    or (self.case_record or {}).get("folder_name")
                    or ""
                ).strip(),
            }

        def _on_success(result):
            self.status_var.set("El PDF quedó en cola para generarse en Drive.")
            dialog_parent = self.owner if getattr(self, "owner", None) is not None else self
            try:
                if dialog_parent is not self and not dialog_parent.winfo_exists():
                    dialog_parent = self
            except Exception:
                dialog_parent = self
            pdf_folder_id = str((result or {}).get("pdf_folder_id") or "").strip()
            pdf_folder_url = (
                f"https://drive.google.com/drive/folders/{pdf_folder_id}"
                if pdf_folder_id
                else None
            )
            sheet_url = str((self.case_record or {}).get("webViewLink") or self.case_path or "").strip()
            company_name = str((result or {}).get("company_name") or "").strip()
            if self.winfo_exists():
                try:
                    self._refresh_workflow_state(preferred_sheet=self.sheet_var.get())
                    self._render_selected_sheet()
                except Exception:
                    pass
            close_before_dialog = dialog_parent is not self
            if close_before_dialog and self.winfo_exists():
                try:
                    self.destroy()
                except tk.TclError:
                    pass
            _show_acta_published_dialog(
                dialog_parent,
                sheet_url=sheet_url,
                company_name=company_name,
                pdf_folder_url=pdf_folder_url,
                dialog_title="PDF de seguimiento",
                header_text="PDF de seguimiento en proceso",
                body_text=(
                    f"La información del caso{' de ' + company_name if company_name else ''} "
                    "quedó guardada en Google Sheets."
                ),
                pdf_status_text=(
                    "El PDF se está generando con la ficha inicial y el seguimiento seleccionado. "
                    "Estará disponible en la carpeta de Drive en unos segundos."
                ),
            )
            if (not close_before_dialog) and self.winfo_exists():
                try:
                    self.destroy()
                except tk.TclError:
                    pass

        self._run_loading_job(
            title="Generando PDF",
            initial_status="Preparando la exportación...",
            worker=_worker,
            on_success=_on_success,
            on_error_context="pdf_export",
        )

    def _save_current_sheet(self):
        sheet = self.sheet_var.get()
        previous_base_completed = bool((self.workflow or {}).get("base_completed"))
        if sheet == seguimientos.SHEET_FINAL:
            _show_inline_feedback(self, "Esta etapa es de solo lectura.", state="warning")
            try:
                self.sheet_combo.focus_set()
            except Exception:
                pass
            return
        if not self._is_sheet_editable(sheet):
            _show_inline_feedback(
                self,
                "Esta etapa no está disponible para edición en este momento.",
                state="warning",
            )
            try:
                self.sheet_combo.focus_set()
            except Exception:
                pass
            return
        self._cancel_sheet_autosave_timer()
        self._sheet_autosave_pending_request = None
        try:
            requests = self._build_pending_sheet_save_requests(sheet)
        except RuntimeError as exc:
            if "Completa los campos obligatorios antes de guardar" in str(exc or ""):
                return
            _show_inline_feedback(self, _log_user_error("save_sheet", exc), state="error")
            return
        except Exception as exc:
            _show_inline_feedback(self, _log_user_error("save_sheet", exc), state="error")
            return
        if not requests:
            self.status_var.set("No hay cambios pendientes para guardar en este caso.")
            _show_inline_feedback(self, "No hay cambios pendientes para guardar.", state="warning")
            return
        if not self._confirm_overwrite_changes():
            self.status_var.set("Guardado cancelado. Revisa los campos resaltados en amarillo.")
            return

        def _worker(progress):
            total = max(1, len(requests))
            saved_requests = []
            sync_warnings = []
            for position, request in enumerate(requests, start=1):
                stage_title = self._sheet_title(request.get("sheet"))
                start_percent = int(((position - 1) * 100) / total)
                end_percent = int((position * 100) / total)

                def _stage_progress(status, percent):
                    try:
                        percent_value = int(percent or 0)
                    except Exception:
                        percent_value = 0
                    percent_value = max(0, min(100, percent_value))
                    mapped = start_percent + int(((end_percent - start_percent) * percent_value) / 100)
                    progress(f"{status} ({position}/{total})", max(6, min(mapped, 98)))

                try:
                    result = self._execute_sheet_save_worker(
                        request,
                        trigger="manual",
                        progress=_stage_progress,
                    )
                except Exception as exc:
                    raise RuntimeError(f"No se pudo guardar {stage_title}.\n{exc}") from exc

                try:
                    _delete_followup_local_sheet_draft(self.case_target, request.get("sheet"))
                except Exception as exc:
                    _log_capture(
                        "followup_editor_local_draft_clear_failed "
                        f"case={self.case_target!r} sheet={request.get('sheet')!r} err={exc}"
                    )

                saved_entry = dict(request)
                saved_entry["sync_warning"] = str((result or {}).get("sync_warning") or "").strip()
                saved_entry["fingerprint"] = str((result or {}).get("fingerprint") or request.get("fingerprint") or "")
                saved_requests.append(saved_entry)
                if saved_entry["sync_warning"]:
                    sync_warnings.append(saved_entry["sync_warning"])

            return {
                "saved_requests": saved_requests,
                "sync_warnings": sync_warnings,
            }

        def _on_success(result):
            saved_requests = list((result or {}).get("saved_requests") or [])
            sync_warnings = [str(item or "").strip() for item in list((result or {}).get("sync_warnings") or []) if str(item or "").strip()]
            self._sheet_autosave_pending_request = None
            saved_titles = []
            saved_followups_in_batch = []
            current_sheet_saved = False
            for request in saved_requests:
                sheet_name = str(request.get("sheet") or "").strip()
                if not sheet_name:
                    continue
                saved_titles.append(self._sheet_title(sheet_name))
                self._mark_session_sheet_saved(request)
                if sheet_name == str(sheet or "").strip():
                    self._sheet_autosave_last_fingerprint = str(request.get("fingerprint") or "")
                    current_sheet_saved = True
                if str(request.get("save_kind") or "").strip() == "followup":
                    try:
                        idx = int(request.get("followup_index") or 0)
                    except Exception:
                        idx = 0
                    if idx > 0:
                        saved_followups_in_batch.append(idx)
            if not current_sheet_saved:
                self._sheet_autosave_last_fingerprint = str(self._sheet_autosave_last_fingerprint or "")
            hub = self._get_hub_window()
            if hub and hasattr(hub, "_refresh_drafts_badge"):
                try:
                    hub._refresh_drafts_badge()
                except Exception:
                    pass
            if sync_warnings:
                for warning in sync_warnings:
                    _log_capture(f"[UI] context=sync_case_record err={warning}")
                _show_inline_feedback(
                    self,
                    "Las etapas se guardaron, pero quedó pendiente la sincronización del caso.",
                    state="warning",
                )
            if not sync_warnings:
                joined_titles = ", ".join(saved_titles[:3])
                if len(saved_titles) > 3:
                    joined_titles = f"{joined_titles} y {len(saved_titles) - 3} etapa(s) más"
                self.status_var.set(f"Etapas guardadas: {joined_titles}.")
            if isinstance(self.owner, SeguimientosWindow):
                self.owner.case_record = self.case_record
                self.owner.case_path = self.case_path
                self.owner.path_var.set(
                    (self.case_record or {}).get("webViewLink") or self.case_path
                )
                self.owner._refresh_suggestion()
            self._refresh_workflow_state(preferred_sheet=sheet)
            self._render_selected_sheet()
            if (
                sheet == str(self.base_sheet_name or "").strip()
                and not previous_base_completed
                and bool((self.workflow or {}).get("base_completed"))
            ):
                followup_1 = f"{seguimientos.SHEET_PREFIX}1"
                if followup_1 in list(self.sheet_options or []):
                    followup_title = self._sheet_title_by_name.get(followup_1, "Seguimiento 1")
                    self.status_var.set(
                        f"Ficha inicial completa. {followup_title} quedó listo para continuar."
                    )

            had_dirty_changes = any(bool(request.get("_was_dirty")) for request in saved_requests)
            if not had_dirty_changes:
                return
            candidates = seguimientos.list_pdf_followup_candidates(self.case_target)
            if not self._ask_generate_followup_pdf(has_followups=bool(candidates)):
                return

            followup_index = None
            if candidates:
                default_index = self._resolve_default_pdf_followup_index(
                    candidates,
                    saved_followups_in_batch,
                )
                followup_index = self._prompt_pdf_followup_choice(
                    candidates,
                    default_index=default_index,
                )
                if followup_index is None:
                    self.status_var.set("Las etapas se guardaron. La generación del PDF fue cancelada.")
                    return

            self._start_followup_pdf_export(followup_index=followup_index)

        self._run_loading_job(
            title="Guardando etapa",
            initial_status="Preparando el guardado...",
            worker=_worker,
            on_success=_on_success,
            on_error_context="save_sheet",
        )

    def _open_excel(self):
        try:
            if self.case_record and str((self.case_record or {}).get("webViewLink") or "").strip():
                _open_url_prefer_chrome(str(self.case_record.get("webViewLink")))
            else:
                raise RuntimeError("No hay enlace del Google Sheet para abrir.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo abrir el archivo.\n{exc}")


# ── VENTANA: LSCWindow ───────────────────────────────────────────────────────


class LSCWindow(tk.Toplevel, FormMousewheelMixin):
    """Ventana para el Servicio de Interpretación LSC.

    Soporta dos rutas de apertura:
      Ruta A (Hub): context=None → búsqueda normal de empresa, oferentes vacíos.
      Ruta B (desde proceso): context={empresa, oferentes, source_form} →
          empresa y oferentes pre-cargados desde el formulario activo.
    """

    def __init__(
        self,
        parent,
        context=None,
        *,
        linked_mode=False,
        parent_form=None,
        on_linked_export_started=None,
        on_linked_export_finished=None,
    ):
        super().__init__(parent)
        self.title("Servicio de Interpretación LSC")
        self.configure(bg=COLOR_LIGHT_BG)
        self.geometry("1000x700")
        _maximize_window(self)

        self._empresa_lookup = interprete_lsc
        self.company_data = None
        self.fields = {}
        self._context = context or interprete_lsc.consume_pending_context()
        self._linked_mode = bool(linked_mode)
        self._parent_form = parent_form
        self._on_linked_export_started = on_linked_export_started
        self._on_linked_export_finished = on_linked_export_finished

        # Si viene con contexto (Ruta B), pre-cargar empresa en cache
        if self._context.get("empresa"):
            self.company_data = dict(self._context["empresa"])
            interprete_lsc.SECTION_1_CACHE.update(self.company_data)

        self._build_header()
        self._build_section_container()
        if self._maybe_resume_form():
            return
        self._show_section_1()

    # ── Header ───────────────────────────────────────────────────────────────

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
        _fn = getattr(self, "_pending_autosave", None)
        if callable(_fn):
            try:
                _fn()
            except Exception:
                pass
            self._pending_autosave = None
        for child in self.section_container.winfo_children():
            child.destroy()

    def _maybe_resume_form(self):
        if _consume_pending_draft_restore(
            self,
            "interprete_lsc",
            interprete_lsc,
            {
                "section_1": self._show_section_1,
                "section_2": self._show_section_2,
                "section_3": self._show_section_3,
                "section_4": self._show_section_4,
            },
            self._show_section_1,
        ):
            return True
        if interprete_lsc.cache_file_exists():
            _clear_local_resume_state(interprete_lsc)
        return False

    # ── Helpers de sección 1 (empresa) ───────────────────────────────────────

    def _build_search(self, parent):
        _section1_build_search(self, parent)

    def _build_groups(self, parent):
        container = tk.Frame(parent, bg=COLOR_LIGHT_BG)
        container.pack(fill="both", expand=True)
        labels = {
            "nombre_empresa": "Nombre de la empresa",
            "ciudad_empresa": "Ciudad/Municipio",
            "direccion_empresa": "Dirección",
            "contacto_empresa": "Contacto en la empresa",
            "cargo": "Cargo",
        }
        self._section1_labels = labels

        group_label = tk.Label(
            container,
            text="Información de Empresa",
            bg=COLOR_GROUP_EMPRESA,
            fg=COLOR_PURPLE,
            font=FONT_LABEL,
            anchor="w",
        )
        group_label.pack(fill="x", padx=FORM_PADX, pady=(0, 0))

        box = tk.Frame(container, bg="#EAF5ED", bd=1, relief="solid")
        box.pack(fill="x", padx=FORM_PADX, pady=(0, FORM_PADY))

        for row_idx, field_id in enumerate(
            ["nombre_empresa", "ciudad_empresa", "direccion_empresa", "contacto_empresa", "cargo"]
        ):
            tk.Label(
                box,
                text=labels.get(field_id, field_id),
                font=FONT_LABEL,
                bg="#EAF5ED",
                anchor="w",
            ).grid(row=row_idx, column=0, sticky="w", padx=20, pady=8)
            entry = tk.Entry(box, width=ENTRY_W_XL)
            entry.grid(row=row_idx, column=1, sticky="w", padx=(12, 20), pady=8)
            entry.configure(state="readonly")
            self.fields[field_id] = entry

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
        target_button = self.search_name_btn if mode == "nombre" else self.search_nit_btn
        _run_section1_company_search(
            self,
            mode=mode,
            lookup=interprete_lsc,
            button=target_button,
        )

    def _prefill_section_1(self):
        cache = interprete_lsc.get_form_cache().get("section_1", {})
        if not cache:
            return
        self.company_data = cache
        self.fields["nit_empresa"].delete(0, tk.END)
        self.fields["nit_empresa"].insert(0, cache.get("nit_empresa", ""))
        fecha_val = cache.get("fecha_visita")
        if fecha_val and "fecha_visita" in self.fields:
            try:
                self.fields["fecha_visita"].set_date(fecha_val)
            except Exception:
                pass
        if "modalidad_interprete" in self.fields:
            self.fields["modalidad_interprete"].set(cache.get("modalidad_interprete", ""))
        if "modalidad_profesional_reca" in self.fields:
            self.fields["modalidad_profesional_reca"].set(cache.get("modalidad_profesional_reca", ""))
        for key in interprete_lsc.SECTION_1_SUPABASE_MAP.keys():
            self._set_readonly_value(key, cache.get(key, ""))

    def _prefill_section_1_from_context(self):
        context = self._context if isinstance(self._context, dict) else {}
        if not context:
            return False

        restored = False
        empresa = context.get("empresa") if isinstance(context.get("empresa"), dict) else {}
        if empresa:
            self.company_data = dict(empresa)
            nit_widget = self.fields.get("nit_empresa")
            if nit_widget is not None:
                _set_input_value(nit_widget, empresa.get("nit_empresa", ""))
            search_widget = self.fields.get("nombre_busqueda")
            if search_widget is not None:
                _set_input_value(search_widget, empresa.get("nombre_empresa", ""))
            for key in interprete_lsc.SECTION_1_SUPABASE_MAP.keys():
                self._set_readonly_value(key, empresa.get(key, ""))
            restored = True

        fecha_val = context.get("fecha_visita")
        if fecha_val not in (None, "") and "fecha_visita" in self.fields:
            try:
                self.fields["fecha_visita"].set_date(fecha_val)
            except Exception:
                _set_input_value(self.fields.get("fecha_visita"), fecha_val)
            restored = True

        if restored:
            _refresh_section1_continue_button(self)
        return restored

    # ── Sección 1: Empresa ────────────────────────────────────────────────────

    def _show_section_1(self):
        self._clear_section_container()
        self.header_title.config(text="1. DATOS DE LA EMPRESA")
        self.header_subtitle.config(text="Busca empresa por NIT y confirma datos del servicio.")

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        self._build_search(content)
        self._build_groups(content)

        # Campos de input adicionales: fecha + modalidades
        extra = tk.Frame(content, bg=COLOR_LIGHT_BG)
        extra.pack(fill="x", padx=FORM_PADX, pady=(8, 4))

        row1 = tk.Frame(extra, bg=COLOR_LIGHT_BG)
        row1.pack(fill="x", pady=4)
        tk.Label(row1, text="Fecha del servicio:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
            side="left", padx=(0, 8)
        )
        fecha_entry = DateEntry(row1, width=ENTRY_W_MED, date_pattern="yyyy-mm-dd")
        fecha_entry.pack(side="left")
        self.fields["fecha_visita"] = fecha_entry
        fecha_error = _build_inline_error_label(extra, wraplength=260)
        fecha_error.pack(anchor="w", padx=(160, 0), pady=(0, 4))
        ui_feedback.register_field(self, "fecha_visita", fecha_entry, error_label=fecha_error)

        row2 = tk.Frame(extra, bg=COLOR_LIGHT_BG)
        row2.pack(fill="x", pady=4)
        tk.Label(row2, text="Modalidad intérprete:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
            side="left", padx=(0, 8)
        )
        mod_interp = ttk.Combobox(
            row2,
            values=["Presencial", "Virtual", "Mixta"],
            state="readonly",
            width=20,
        )
        mod_interp.pack(side="left", padx=(0, 20))
        mod_interp.set("")
        self.fields["modalidad_interprete"] = mod_interp

        tk.Label(row2, text="Modalidad profesional RECA:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
            side="left", padx=(0, 8)
        )
        mod_prof = ttk.Combobox(
            row2,
            values=["Presencial", "Virtual", "No aplica"],
            state="readonly",
            width=20,
        )
        mod_prof.pack(side="left")
        mod_prof.set("")
        self.fields["modalidad_profesional_reca"] = mod_prof
        modalidad_interp_error = _build_inline_error_label(extra, wraplength=260)
        modalidad_interp_error.pack(anchor="w", padx=(160, 0), pady=(0, 2))
        modalidad_prof_error = _build_inline_error_label(extra, wraplength=260)
        modalidad_prof_error.pack(anchor="w", padx=(160, 0), pady=(0, 6))
        ui_feedback.register_field(
            self,
            "modalidad_interprete",
            mod_interp,
            error_label=modalidad_interp_error,
        )
        ui_feedback.register_field(
            self,
            "modalidad_profesional_reca",
            mod_prof,
            error_label=modalidad_prof_error,
        )

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._close_to_hub).pack(side="left")
        self.continue_btn = ttk.Button(actions, text="Continuar", command=self._confirm_section_1)
        self.continue_btn.pack(side="right")

        restored = _restore_section1_cached_state(self, interprete_lsc)
        if not restored:
            self._prefill_section_1_from_context()

    # ── Sección 2: Oferentes / Vinculados ────────────────────────────────────

    def _show_section_2(self):
        self._clear_section_container()
        self.header_title.config(text="2. OFERENTES / VINCULADOS")
        self.header_subtitle.config(
            text=f"Registra los candidatos acompañados (máx. {interprete_lsc.MAX_OFERENTES})."
        )

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        # Encabezado de tabla
        hdr = tk.Frame(content, bg=COLOR_LIGHT_BG)
        hdr.pack(fill="x", padx=FORM_PADX, pady=(8, 2))
        for col, w in [("No.", 4), ("Nombre completo", 38), ("Cédula", 16), ("Proceso / Observaciones", 38)]:
            tk.Label(hdr, text=col, font=FONT_LABEL, bg=COLOR_LIGHT_BG, width=w, anchor="w").pack(
                side="left", padx=2
            )

        rows_frame = tk.Frame(content, bg=COLOR_LIGHT_BG)
        rows_frame.pack(fill="x", padx=FORM_PADX)
        self._oferente_rows = []

        def _add_oferente_row(nombre="", cedula="", proceso=""):
            idx = len(self._oferente_rows) + 1
            if idx > interprete_lsc.MAX_OFERENTES:
                return
            row = tk.Frame(rows_frame, bg=COLOR_LIGHT_BG)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=str(idx), bg=COLOR_LIGHT_BG, width=4, anchor="w").pack(side="left", padx=2)
            e_nombre = tk.Entry(row, width=38)
            e_nombre.pack(side="left", padx=2)
            e_cedula = tk.Entry(row, width=16)
            e_cedula.pack(side="left", padx=2)
            e_proceso = tk.Entry(row, width=38)
            e_proceso.pack(side="left", padx=2)
            if nombre:
                e_nombre.insert(0, nombre)
            if cedula:
                e_cedula.insert(0, cedula)
            if proceso:
                e_proceso.insert(0, proceso)
            self._oferente_rows.append((row, e_nombre, e_cedula, e_proceso))

        def _remove_last_oferente():
            if len(self._oferente_rows) <= 1:
                return
            row, *_ = self._oferente_rows.pop()
            row.destroy()

        # Pre-cargar desde cache o contexto
        cached = interprete_lsc.get_form_cache().get("section_2", [])
        if not cached and self._context.get("oferentes"):
            cached = self._context["oferentes"]
        if cached:
            for of in cached:
                _add_oferente_row(
                    nombre=of.get("nombre_oferente") or of.get("nombre", ""),
                    cedula=of.get("cedula") or of.get("cedula_oferente", ""),
                    proceso=of.get("proceso") or of.get("proceso_observaciones", ""),
                )
        else:
            _add_oferente_row()

        self._pending_autosave = lambda: _autosave_section(
            interprete_lsc,
            "section_2",
            lambda: [
                {
                    "nombre_oferente": _normalize_person_name(r[1].get()),
                    "cedula": r[2].get().strip(),
                    "proceso": r[3].get().strip(),
                }
                for r in self._oferente_rows
                if r[1].get().strip() or r[2].get().strip()
            ],
        )

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_1).pack(side="left")
        ttk.Button(actions, text="+ Agregar", command=_add_oferente_row).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(actions, text="- Eliminar último", command=_remove_last_oferente).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(actions, text="Continuar", command=self._confirm_section_2).pack(side="right")

    # ── Sección 3: Intérpretes ────────────────────────────────────────────────

    def _show_section_3(self):
        self._clear_section_container()
        self.header_title.config(text="3. INTÉRPRETES")
        self.header_subtitle.config(
            text=f"Registra los intérpretes LSC (máx. {interprete_lsc.MAX_INTERPRETES}). "
            "Los tiempos totales se calculan automáticamente. Acepta horas como 9 30 am, 9:30 pm o 14:30."
        )

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        interp_frame = tk.Frame(content, bg=COLOR_LIGHT_BG)
        interp_frame.pack(fill="x", padx=FORM_PADX, pady=(8, 4))
        self._interprete_rows = []
        interpretes_catalog = _get_interpretes_catalog()

        def _recalc_sumatoria(*_):
            """Recalcula sumatoria cada vez que cambia un campo de hora."""
            try:
                sab = _sabana_var.get()
                hs = float(_sabana_horas_var.get() or 1.0)
                interps = []
                for _, e_nom, e_ini, e_fin, lbl_tot in self._interprete_rows:
                    ini = e_ini.get().strip()
                    fin = e_fin.get().strip()
                    tot = interprete_lsc.calc_total_tiempo(ini, fin)
                    lbl_tot.config(text=tot or "—")
                    if tot:
                        interps.append({"total_tiempo": tot})
                sumatoria = interprete_lsc.calc_sumatoria(interps, sab, hs)
                _sumatoria_var.set(sumatoria)
            except Exception:
                pass

        def _add_interprete_row(nombre="", hora_ini="", hora_fin="", total=""):
            idx = len(self._interprete_rows) + 1
            if idx > interprete_lsc.MAX_INTERPRETES:
                return
            row = tk.Frame(interp_frame, bg=COLOR_LIGHT_BG)
            row.pack(fill="x", pady=4)
            tk.Label(row, text=f"Intérprete {idx}:", font=FONT_LABEL, bg=COLOR_LIGHT_BG, width=12).pack(
                side="left"
            )
            e_nom = _create_interprete_name_input(row, 28, catalog=interpretes_catalog)
            e_nom.pack(side="left", padx=(0, 8))
            if nombre:
                e_nom.set(_normalize_person_name(nombre))

            tk.Label(row, text="Hora inicial:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 4)
            )
            e_ini = tk.Entry(row, width=8)
            e_ini.pack(side="left", padx=(0, 8))
            if hora_ini:
                e_ini.insert(0, interprete_lsc.normalize_time_value(hora_ini) or hora_ini)

            tk.Label(row, text="Hora final:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 4)
            )
            e_fin = tk.Entry(row, width=8)
            e_fin.pack(side="left", padx=(0, 8))
            if hora_fin:
                e_fin.insert(0, interprete_lsc.normalize_time_value(hora_fin) or hora_fin)

            tk.Label(row, text="Total:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 4)
            )
            lbl_tot = tk.Label(row, text=total or "—", font=FONT_LABEL, bg=COLOR_LIGHT_BG, width=7)
            lbl_tot.pack(side="left")

            _bind_lsc_time_entry(e_ini, on_change=_recalc_sumatoria)
            _bind_lsc_time_entry(e_fin, on_change=_recalc_sumatoria)
            self._interprete_rows.append((row, e_nom, e_ini, e_fin, lbl_tot))

        def _remove_last_interprete():
            if len(self._interprete_rows) <= 1:
                return
            row, *_ = self._interprete_rows.pop()
            row.destroy()
            _recalc_sumatoria()

        # Pre-cargar desde cache
        cached_s3 = interprete_lsc.get_form_cache().get("section_3") or {}
        cached_interps = cached_s3.get("interpretes", []) if isinstance(cached_s3, dict) else []
        if cached_interps:
            for it in cached_interps:
                _add_interprete_row(
                    nombre=it.get("nombre", ""),
                    hora_ini=it.get("hora_inicial", ""),
                    hora_fin=it.get("hora_final", ""),
                    total=it.get("total_tiempo", ""),
                )
        else:
            _add_interprete_row()

        # ── Sabana ────────────────────────────────────────────────────────────
        sep = tk.Frame(content, bg="#CCCCCC", height=1)
        sep.pack(fill="x", padx=FORM_PADX, pady=(10, 6))

        sabana_frame = tk.Frame(content, bg=COLOR_LIGHT_BG)
        sabana_frame.pack(fill="x", padx=FORM_PADX, pady=(0, 4))

        _sabana_var = tk.BooleanVar(value=False)
        _sabana_horas_var = tk.StringVar(value="1.0")
        _sumatoria_var = tk.StringVar(value="—")

        cached_sab = cached_s3.get("sabana", {}) if isinstance(cached_s3, dict) else {}
        if cached_sab.get("activo"):
            _sabana_var.set(True)
            _sabana_horas_var.set(str(cached_sab.get("horas", 1.0)))
        if isinstance(cached_s3, dict) and cached_s3.get("sumatoria_horas"):
            _sumatoria_var.set(cached_s3["sumatoria_horas"])

        def _toggle_sabana_entry(*_):
            if _sabana_var.get():
                _sabana_horas_entry.configure(state="normal")
            else:
                _sabana_horas_entry.configure(state="disabled")
            _recalc_sumatoria()

        sab_chk = tk.Checkbutton(
            sabana_frame,
            text="Servicio realizado en Sabana (sumar horas adicionales):",
            variable=_sabana_var,
            bg=COLOR_LIGHT_BG,
            font=FONT_LABEL,
            command=_toggle_sabana_entry,
        )
        sab_chk.pack(side="left", padx=(0, 8))
        _sabana_horas_entry = tk.Entry(sabana_frame, textvariable=_sabana_horas_var, width=8)
        _sabana_horas_entry.pack(side="left", padx=(0, 4))
        _sabana_horas_entry.configure(state="disabled" if not _sabana_var.get() else "normal")
        tk.Label(sabana_frame, text="horas", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(side="left")
        _sabana_horas_entry.bind("<FocusOut>", _recalc_sumatoria)

        # ── Sumatoria ─────────────────────────────────────────────────────────
        sum_frame = tk.Frame(content, bg=COLOR_LIGHT_BG)
        sum_frame.pack(fill="x", padx=FORM_PADX, pady=(4, 8))
        tk.Label(sum_frame, text="Sumatoria total de horas:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
            side="left", padx=(0, 8)
        )
        tk.Label(sum_frame, textvariable=_sumatoria_var, font=FONT_TITLE, fg=COLOR_PURPLE, bg=COLOR_LIGHT_BG).pack(
            side="left"
        )
        ttk.Button(sum_frame, text="↺ Recalcular", command=_recalc_sumatoria).pack(
            side="left", padx=(12, 0)
        )

        _recalc_sumatoria()

        self._pending_autosave = lambda: _autosave_section(
            interprete_lsc,
            "section_3",
            lambda: self._collect_section_3(
                _sabana_var, _sabana_horas_var, _sumatoria_var
            ),
        )

        self._s3_sabana_var = _sabana_var
        self._s3_sabana_horas_var = _sabana_horas_var
        self._s3_sumatoria_var = _sumatoria_var

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_2).pack(side="left")
        ttk.Button(
            actions,
            text="+ Agregar intérprete",
            command=lambda: [_add_interprete_row(), _recalc_sumatoria()],
        ).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="- Eliminar último", command=_remove_last_interprete).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(actions, text="Continuar", command=self._confirm_section_3).pack(side="right")

    def _collect_section_3(self, sabana_var, sabana_horas_var, sumatoria_var):
        interpretes = []
        for _, e_nom, e_ini, e_fin, lbl_tot in self._interprete_rows:
            nombre = _normalize_person_name(e_nom.get())
            hora_ini = interprete_lsc.normalize_time_value(e_ini.get())
            hora_fin = interprete_lsc.normalize_time_value(e_fin.get())
            total = interprete_lsc.calc_total_tiempo(hora_ini, hora_fin)
            if nombre or hora_ini or hora_fin:
                interpretes.append({
                    "nombre": nombre,
                    "hora_inicial": hora_ini,
                    "hora_final": hora_fin,
                    "total_tiempo": total,
                })
        return {
            "interpretes": interpretes,
            "sabana": {
                "activo": bool(sabana_var.get()),
                "horas": float(sabana_horas_var.get() or 1.0),
            },
            "sumatoria_horas": sumatoria_var.get(),
        }

    # ── Sección 4: Asistentes ─────────────────────────────────────────────────

    def _show_section_4(self):
        self._clear_section_container()
        self.header_title.config(text="4. ASISTENTES")
        self.header_subtitle.config(
            text=f"Registra los asistentes al servicio (máx. {interprete_lsc.MAX_ASISTENTES})."
        )

        section_frame = tk.Frame(self.section_container, bg=COLOR_LIGHT_BG)
        section_frame.pack(fill="both", expand=True)
        content = _build_scrollable_content(section_frame, self)

        asist_content = tk.Frame(content, bg=COLOR_LIGHT_BG)
        asist_content.pack(fill="x", padx=FORM_PADX, pady=(8, 8))
        self._asistente_rows = []
        catalog = _get_asistentes_profesionales_catalog()

        def _add_asistente_row(nombre="", cargo=""):
            if len(self._asistente_rows) >= interprete_lsc.MAX_ASISTENTES:
                return
            use_catalog = len(self._asistente_rows) == 0
            row = tk.Frame(asist_content, bg=COLOR_LIGHT_BG)
            row.pack(fill="x", pady=4)
            tk.Label(row, text="Nombre completo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 6)
            )
            nombre_entry, cargo_entry = _create_asistente_inputs(
                row,
                50,
                use_catalog=use_catalog,
                catalog=catalog,
            )
            nombre_entry.pack(side="left", padx=(0, 12))
            tk.Label(row, text="Cargo:", font=FONT_LABEL, bg=COLOR_LIGHT_BG).pack(
                side="left", padx=(0, 6)
            )
            cargo_entry.pack(side="left")
            if not use_catalog:
                _bind_name_entry(nombre_entry)
            if nombre:
                _set_input_value(nombre_entry, _normalize_person_name(nombre))
            if cargo:
                _set_input_value(cargo_entry, cargo)
            self._asistente_rows.append((row, nombre_entry, cargo_entry))

        def _remove_last_asistente():
            if len(self._asistente_rows) <= 1:
                return
            row, *_ = self._asistente_rows.pop()
            row.destroy()

        cached_asist = interprete_lsc.get_form_cache().get("section_4", [])
        if cached_asist:
            for a in cached_asist:
                _add_asistente_row(a.get("nombre", ""), a.get("cargo", ""))
        else:
            for _ in range(2):
                _add_asistente_row()

        self._pending_autosave = lambda: _autosave_section(
            interprete_lsc,
            "section_4",
            lambda: _collect_asistente_rows(self._asistente_rows),
        )

        actions = tk.Frame(content, bg=COLOR_LIGHT_BG)
        _pack_actions(actions)
        ttk.Button(actions, text="Regresar", command=self._show_section_3).pack(side="left")
        ttk.Button(actions, text="Agregar asistente", command=_add_asistente_row).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(actions, text="Eliminar último", command=_remove_last_asistente).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(actions, text="Finalizar", command=self._confirm_section_4).pack(side="right")

    # ── Confirmaciones ────────────────────────────────────────────────────────

    def _confirm_section_1(self):
        ui_feedback.clear_field_errors(
            self,
            ["nit_empresa", "nombre_busqueda", "fecha_visita", "modalidad_interprete", "modalidad_profesional_reca"],
        )
        _clear_inline_feedback(self)
        if not getattr(self, "company_data", None):
            _show_inline_feedback(self, "Busca una empresa antes de continuar.", state="error")
            ui_feedback.focus_first_invalid_field(self, ["nit_empresa", "nombre_busqueda"])
            return

        fecha_visita = _get_required_fecha_visita(self)
        modalidad_interprete = str(_get_input_value(self.fields.get("modalidad_interprete")) or "").strip()
        modalidad_profesional = str(_get_input_value(self.fields.get("modalidad_profesional_reca")) or "").strip()

        has_error = False
        if not modalidad_interprete:
            ui_feedback.set_field_error(
                self,
                "modalidad_interprete",
                "Selecciona la modalidad del intérprete.",
            )
            has_error = True
        else:
            ui_feedback.clear_field_error(self, "modalidad_interprete")

        if not modalidad_profesional:
            ui_feedback.set_field_error(
                self,
                "modalidad_profesional_reca",
                "Selecciona la modalidad del profesional RECA.",
            )
            has_error = True
        else:
            ui_feedback.clear_field_error(self, "modalidad_profesional_reca")

        if not fecha_visita or has_error:
            _show_inline_feedback(
                self,
                "Completa la fecha del servicio y las dos modalidades para continuar.",
                state="error",
            )
            ui_feedback.focus_first_invalid_field(
                self,
                ["fecha_visita", "modalidad_interprete", "modalidad_profesional_reca"],
            )
            return

        user_inputs = {
            "fecha_visita": fecha_visita,
            "modalidad_interprete": modalidad_interprete,
            "modalidad_profesional_reca": modalidad_profesional,
        }
        try:
            interprete_lsc.confirm_section_1(self.company_data, user_inputs)
        except Exception as exc:
            _show_inline_feedback(self, _log_user_error("section_confirm", exc), state="error")
            return
        _clear_inline_feedback(self)
        self._show_section_2()

    def _confirm_section_2(self):
        payload = []
        for _row, e_nom, e_ced, e_proc in self._oferente_rows:
            nombre = _normalize_person_name(e_nom.get())
            cedula = e_ced.get().strip()
            proceso = e_proc.get().strip()
            if nombre or cedula:
                payload.append({
                    "nombre_oferente": nombre,
                    "cedula": cedula,
                    "proceso": proceso,
                })
        if not payload:
            messagebox.showerror("Error", "Registra al menos un oferente/vinculado.")
            return
        if len(payload) > interprete_lsc.MAX_OFERENTES:
            messagebox.showerror(
                "Error", f"Máximo {interprete_lsc.MAX_OFERENTES} oferentes permitidos."
            )
            return
        try:
            interprete_lsc.confirm_section_2(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_3()

    def _confirm_section_3(self):
        if not hasattr(self, "_s3_sabana_var"):
            messagebox.showerror("Error", "Error interno: recarga la sección.")
            return
        payload = self._collect_section_3(
            self._s3_sabana_var,
            self._s3_sabana_horas_var,
            self._s3_sumatoria_var,
        )
        interpretes = payload.get("interpretes", [])
        if not interpretes:
            messagebox.showerror("Error", "Registra al menos un intérprete.")
            return
        if len(interpretes) > interprete_lsc.MAX_INTERPRETES:
            messagebox.showerror(
                "Error", f"Máximo {interprete_lsc.MAX_INTERPRETES} intérpretes permitidos."
            )
            return
        try:
            interprete_lsc.confirm_section_3(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._show_section_4()

    def _confirm_section_4(self):
        payload = _collect_asistente_rows(self._asistente_rows)
        try:
            interprete_lsc.confirm_section_4(payload)
        except Exception as exc:
            messagebox.showerror("Error", _log_user_error("ui_error", exc))
            return
        self._export_form()

    # ── Exportación ───────────────────────────────────────────────────────────

    def _export_form(self):
        cache_snapshot = interprete_lsc.get_form_cache()
        company_name = cache_snapshot.get("section_1", {}).get("nombre_empresa")

        def _worker():
            return _raise_finalize_stage(
                "preparando el acta LSC",
                lambda: interprete_lsc.export_to_excel(clear_cache=False),
            )

        if self._linked_mode and self._parent_form is not None:
            started = _start_background_finalization(
                self,
                None,
                form_name="Servicio de Interpretación LSC",
                company_name=company_name,
                form_id="interprete_lsc",
                worker_fn=_worker,
                post_delivery_fn=lambda: _clear_form_cache_safe(interprete_lsc),
                close_window_on_success=False,
                return_to_hub_on_success=False,
                show_completion_ui=False,
                show_error_dialog=False,
                on_success=lambda result: self._on_linked_export_success(result),
                on_error=lambda exc, message: self._on_linked_export_error(message),
            )
            if not started:
                return
            try:
                self.grab_release()
            except tk.TclError:
                pass
            self.withdraw()
            _focus_window(self._parent_form)
            if callable(self._on_linked_export_started):
                self._on_linked_export_started()
            return

        loading = LoadingDialog(self, title="Guardando")
        loading.set_status("Preparando acta LSC...")
        loading.set_progress(35)

        _start_background_finalization(
            self,
            loading,
            form_name="Servicio de Interpretación LSC",
            company_name=company_name,
            form_id="interprete_lsc",
            worker_fn=_worker,
            post_delivery_fn=lambda: _clear_form_cache_safe(interprete_lsc),
        )

    def _on_linked_export_success(self, result):
        try:
            if callable(self._on_linked_export_finished):
                self._on_linked_export_finished(status="success", result=result, error_message="")
        finally:
            try:
                self._skip_close_guard = True
                self.destroy()
            except tk.TclError:
                pass

    def _on_linked_export_error(self, message):
        try:
            if callable(self._on_linked_export_finished):
                self._on_linked_export_finished(status="failed", result=None, error_message=message)
        finally:
            try:
                self._skip_close_guard = True
                self.destroy()
            except tk.TclError:
                pass

    def _close_to_hub(self):
        _return_to_hub(self)
        self.destroy()

if __name__ == "__main__":
    try:
        _ensure_roaming_service_account_file()
    except Exception:
        pass
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






