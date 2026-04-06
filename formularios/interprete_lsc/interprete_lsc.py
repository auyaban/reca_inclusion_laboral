"""
formularios/interprete_lsc/interprete_lsc.py
Formulario: "SERVICIO DE INTERPRETACIÓN LENGUA DE SEÑAS COLOMBIANA LSC"

Responsabilidades:
  - Mapeo de campos a celdas del spreadsheet maestro LSC (pestaña "Maestro")
  - Secciones 1–4: datos empresa, oferentes/vinculados, intérpretes, asistentes
  - Cálculo dinámico de offsets de filas para oferentes (máx 10) e intérpretes (máx 5)
  - Inserción de filas extra vía row_insertions en publish_sheet_from_template

Estructura del template "Maestro":
  Filas 1–9:   Header + Sección 1 (empresa) — siempre fijas
  Fila  10:    Label "DATOS DE LOS OFERENTES/VINCULADOS"
  Fila  11:    Header tabla oferentes
  Filas 12–18: 7 slots de oferentes (TEMPLATE_OFERENTE_SLOTS)
  Fila  19:    1 slot de intérprete (TEMPLATE_INTERPRETE_SLOTS)
  Fila  20:    Sabana
  Fila  21:    Sumatoria
  Filas 22–23: Observaciones (texto fijo, no escrito por la app)
  Fila  24:    Label "ASISTENTES"
  Filas 25–26: 2 slots de asistentes

Entry points para app.py:
  confirm_section_1(company_data, user_inputs)
  confirm_section_2(payload)   → lista de oferentes
  confirm_section_3(payload)   → dict con interpretes, sabana, sumatoria_horas
  confirm_section_4(payload)   → lista de asistentes
  validate_before_finalize()   → lista ValidationIssue
  export_to_excel()            → escribe en Sheets y sube a Drive
  register_form()              → metadata para HubWindow
  set_pending_context(ctx)     → inyecta contexto antes de abrir la ventana (Ruta B)
  consume_pending_context()    → consume y limpia el contexto pendiente

Depende de: google_sheets_client, formularios/common, formularios/finalize_validation,
            formularios/evaluacion_programa (para búsqueda de empresa)
"""

import json
import os
import re
import time
from datetime import datetime, timedelta

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _get_local_app_cache_dir,
    _sanitize_filename,
    build_sheet_updates,
)
from formularios.finalize_validation import (
    ValidationIssue,
    field_pairs,
    raise_validation_error,
    require_value,
)
from logging_utils import log_excel_event


# ── Constantes de formulario ─────────────────────────────────────────────────

FORM_ID = "interprete_lsc"
FORM_NAME = "Servicio de Interpretación LSC"
SHEET_TAB_NAME = "Maestro"

# ID del spreadsheet LSC (cargado desde config.json, con fallback hardcodeado)
_LSC_SPREADSHEET_ID_FALLBACK = "1WLAoc5lKHEoH3dkR1aQv6UYpEw97b9iNc2k43hCKrmk"

# Estructura del template "Maestro"
TEMPLATE_FIRST_OFERENTE_ROW = 12   # primera fila de datos de oferente
TEMPLATE_OFERENTE_SLOTS = 7        # filas 12–18 pre-formateadas
TEMPLATE_FIRST_INTERPRETE_ROW = 19 # fila del intérprete en template vacío
TEMPLATE_INTERPRETE_SLOTS = 1      # sólo 1 fila de intérprete en template
TEMPLATE_SABANA_OFFSET = 1         # filas después del último intérprete
TEMPLATE_SUMATORIA_OFFSET = 2
TEMPLATE_OBS_OFFSET = 3            # ocupa 2 filas (22-23)
TEMPLATE_SECTION3_OFFSET = 5       # "3. ASISTENTES"
TEMPLATE_FIRST_ASISTENTE_OFFSET = 6

MAX_OFERENTES = 10
MAX_INTERPRETES = 5
MAX_ASISTENTES = 2


# ── Caches en memoria ────────────────────────────────────────────────────────

FORM_CACHE: dict = {}
SECTION_1_CACHE: dict = {}
_PENDING_CONTEXT: dict = {}   # contexto inyectado desde otra ventana (Ruta B)


# ── Definición de secciones ──────────────────────────────────────────────────

SECTION_1 = {
    "title": "1. DATOS DE LA EMPRESA",
    "nit_lookup_field": "nit_empresa",
    "fields": [
        {
            "id": "nombre_empresa",
            "label": "Nombre de la empresa",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "ciudad_empresa",
            "label": "Ciudad/Municipio",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "direccion_empresa",
            "label": "Dirección",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "nit_empresa",
            "label": "NIT",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "contacto_empresa",
            "label": "Contacto en la empresa",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "cargo",
            "label": "Cargo",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        # Campos de input del asesor
        {"id": "fecha_visita", "label": "Fecha del servicio", "source": "input"},
        {
            "id": "modalidad_interprete",
            "label": "Modalidad intérprete",
            "source": "input",
            "options": ["Presencial", "Virtual", "Mixta"],
        },
        {
            "id": "modalidad_profesional_reca",
            "label": "Modalidad profesional RECA",
            "source": "input",
            "options": ["Presencial", "Virtual", "No aplica"],
        },
    ],
}

SECTION_2 = {"title": "2. DATOS DE LOS OFERENTES / VINCULADOS"}
SECTION_3 = {"title": "3. INTÉRPRETES"}
SECTION_4 = {"title": "4. ASISTENTES"}

SECTION_1_SUPABASE_MAP = evaluacion_accesibilidad.SECTION_1_SUPABASE_MAP.copy()

# Mapeo de celdas para la sección 1 (filas fijas del template)
EXCEL_MAPPING = {
    "section_1": {
        "fecha_visita":            "D6",
        "ciudad_empresa":          "N6",
        "nombre_empresa":          "D7",
        "direccion_empresa":       "N7",
        "contacto_empresa":        "D8",
        "cargo":                   "N8",
        "modalidad_interprete":    "I9",
        "modalidad_profesional_reca": "P9",
    }
}


# ── Registro de formulario ───────────────────────────────────────────────────

def register_form():
    return {
        "id": FORM_ID,
        "name": FORM_NAME,
        "module": __name__,
        "hub_description": (
            "Documenta el servicio de interpretación en Lengua de Señas Colombiana (LSC), "
            "intérpretes, oferentes acompañados y asistentes."
        ),
        "singleton_window": True,
    }


# ── Cache en disco ───────────────────────────────────────────────────────────

def _get_cache_path():
    return os.path.join(_get_local_app_cache_dir(), f"{FORM_ID}.json")


def cache_file_exists():
    return os.path.exists(_get_cache_path())


def save_cache_to_file():
    payload = {
        "form_id": FORM_ID,
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "data": FORM_CACHE,
    }
    with open(_get_cache_path(), "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def load_cache_from_file():
    path = _get_cache_path()
    if not os.path.exists(path):
        return False
    with open(path, "r", encoding="utf-8") as handle:
        payload = json.load(handle) or {}
    data = payload.get("data") or {}
    FORM_CACHE.clear()
    FORM_CACHE.update(data)
    section_1 = data.get("section_1") or {}
    SECTION_1_CACHE.clear()
    SECTION_1_CACHE.update(section_1)
    return True


def clear_cache_file():
    path = _get_cache_path()
    if os.path.exists(path):
        os.remove(path)


def clear_form_cache():
    FORM_CACHE.clear()
    SECTION_1_CACHE.clear()


def set_section_cache(section_id, payload):
    if not section_id:
        raise ValueError("section_id requerido")
    FORM_CACHE[section_id] = payload if payload is not None else {}


def get_form_cache():
    return dict(FORM_CACHE)


# ── Contexto pendiente (Ruta B: inyección desde otra ventana) ────────────────

def set_pending_context(context: dict):
    """Llamado por otras ventanas antes de abrir LSCWindow con contexto."""
    _PENDING_CONTEXT.clear()
    if context:
        _PENDING_CONTEXT.update(context)


def consume_pending_context() -> dict:
    """Llamado una sola vez por LSCWindow.__init__ para obtener y limpiar el contexto."""
    ctx = dict(_PENDING_CONTEXT)
    _PENDING_CONTEXT.clear()
    return ctx


# ── Búsqueda de empresa (delega a evaluacion_accesibilidad) ─────────────────

def get_empresa_by_nit(nit, env_path=".env"):
    return evaluacion_accesibilidad.get_empresa_by_nit(nit, env_path=env_path)


def get_empresa_by_nombre(nombre, env_path=".env"):
    return evaluacion_accesibilidad.get_empresa_by_nombre(nombre, env_path=env_path)


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=50):
    return evaluacion_accesibilidad.get_empresas_by_nombre_prefix(
        prefix, env_path=env_path, limit=limit
    )


# ── Confirmación de secciones ────────────────────────────────────────────────

def confirm_section_1(company_data, user_inputs):
    if not company_data:
        raise ValueError("No hay datos de empresa para confirmar.")
    payload = {}
    for field in SECTION_1["fields"]:
        field_id = field["id"]
        if field["source"] == "input":
            payload[field_id] = user_inputs.get(field_id)
        else:
            payload[field_id] = company_data.get(field_id)
    SECTION_1_CACHE.update(payload)
    set_section_cache("section_1", payload)
    FORM_CACHE["_last_section"] = "section_1"
    save_cache_to_file()
    return payload


def confirm_section_2(payload):
    """payload: lista de dicts con nombre_oferente, cedula, proceso."""
    if payload is None:
        payload = []
    set_section_cache("section_2", payload)
    FORM_CACHE["_last_section"] = "section_2"
    save_cache_to_file()
    return payload


def confirm_section_3(payload):
    """payload: dict con keys 'interpretes' (lista), 'sabana' (dict), 'sumatoria_horas' (str)."""
    if payload is None:
        raise ValueError("section_3 requerida")
    set_section_cache("section_3", payload)
    FORM_CACHE["sabana"] = payload.get("sabana", {})
    FORM_CACHE["sumatoria_horas"] = payload.get("sumatoria_horas", "")
    FORM_CACHE["_last_section"] = "section_3"
    save_cache_to_file()
    return payload


def confirm_section_4(payload):
    """payload: lista de dicts con nombre, cargo."""
    if payload is None:
        payload = []
    set_section_cache("section_4", payload)
    FORM_CACHE["_last_section"] = "section_4"
    save_cache_to_file()
    return payload


# ── Validación pre-finalización ──────────────────────────────────────────────

def validate_before_finalize(cache=None):
    cache_data = FORM_CACHE if cache is None else (cache or {})
    issues = []

    section_1 = cache_data.get("section_1", {})
    require_value(issues, "section_1", section_1, "fecha_visita", "Fecha del servicio")
    require_value(issues, "section_1", section_1, "modalidad_interprete", "Modalidad del intérprete")
    require_value(
        issues,
        "section_1",
        section_1,
        "modalidad_profesional_reca",
        "Modalidad del profesional RECA",
    )

    oferentes = cache_data.get("section_2", [])
    if not isinstance(oferentes, list) or not oferentes:
        issues.append(
            ValidationIssue("section_2", "oferentes", "Debe haber al menos un oferente/vinculado.")
        )

    section_3 = cache_data.get("section_3", {})
    interpretes = section_3.get("interpretes", []) if isinstance(section_3, dict) else []
    if not isinstance(interpretes, list) or not interpretes:
        issues.append(
            ValidationIssue("section_3", "interpretes", "Debe haber al menos un intérprete.")
        )

    return issues


# ── Helpers de tiempo ────────────────────────────────────────────────────────

def calc_total_tiempo(hora_ini: str, hora_fin: str) -> str:
    """Calcula H:MM a partir de dos strings de hora (HH:MM)."""
    try:
        hora_ini_norm = normalize_time_value(hora_ini)
        hora_fin_norm = normalize_time_value(hora_fin)
        if not hora_ini_norm or not hora_fin_norm:
            return ""
        t1 = datetime.strptime(hora_ini_norm, "%H:%M")
        t2 = datetime.strptime(hora_fin_norm, "%H:%M")
        delta = t2 - t1
        if delta.total_seconds() < 0:
            delta += timedelta(hours=24)
        total_mins = int(delta.total_seconds() / 60)
        return f"{total_mins // 60}:{total_mins % 60:02d}"
    except Exception:
        return ""


def normalize_time_value(value: str) -> str:
    """Normaliza horas como '9 30 am', '930 pm' o '21:30' a HH:MM."""
    raw = str(value or "").strip().lower().replace(".", "")
    raw = re.sub(r"\s+", " ", raw)
    if not raw:
        return ""

    suffix = ""
    match = re.search(r"(am|pm|a|p)$", raw)
    if match:
        suffix = match.group(1)
        raw = raw[: match.start()].strip()

    raw = raw.replace(" ", ":")
    raw = re.sub(r"[^0-9:]", "", raw)
    if not raw:
        return ""

    hour = None
    minute = None
    if ":" in raw:
        parts = [part for part in raw.split(":") if part]
        if not parts:
            return ""
        if len(parts) == 1:
            digits = re.sub(r"\D+", "", parts[0])
            if len(digits) <= 2:
                hour = int(digits)
                minute = 0
            elif len(digits) == 3:
                hour = int(digits[0])
                minute = int(digits[1:])
            else:
                hour = int(digits[:-2])
                minute = int(digits[-2:])
        else:
            hour = int(parts[0])
            minute = int(parts[1])
    else:
        digits = re.sub(r"\D+", "", raw)
        if len(digits) <= 2:
            hour = int(digits)
            minute = 0
        elif len(digits) == 3:
            hour = int(digits[0])
            minute = int(digits[1:])
        else:
            hour = int(digits[:-2])
            minute = int(digits[-2:])

    if hour is None or minute is None or minute < 0 or minute > 59:
        return ""

    if suffix:
        if hour < 1 or hour > 12:
            return ""
        if suffix.startswith("p") and hour != 12:
            hour += 12
        elif suffix.startswith("a") and hour == 12:
            hour = 0
    elif hour < 0 or hour > 23:
        return ""

    return f"{hour:02d}:{minute:02d}"


def calc_sumatoria(interpretes: list, sabana_activo: bool, sabana_horas: float) -> str:
    """Suma todos los tiempos de intérpretes más las horas de Sabana si aplica."""
    total_mins = 0
    for interp in interpretes:
        tt = (interp.get("total_tiempo") or "").strip()
        if tt and ":" in tt:
            try:
                h, m = tt.split(":")
                total_mins += int(h) * 60 + int(m)
            except Exception:
                pass
    if sabana_activo:
        total_mins += int(float(sabana_horas or 1.0) * 60)
    return f"{total_mins // 60}:{total_mins % 60:02d}"


def format_sabana_value(activo: bool, horas: float) -> str:
    """Formatea el valor de la celda Sabana para el Sheet."""
    if not activo:
        return "No aplica"
    h = float(horas or 1.0)
    horas_int = int(h)
    mins = int(round((h - horas_int) * 60))
    return f"{horas_int}:{mins:02d} Hora"


# ── Cálculo dinámico de offsets de filas ─────────────────────────────────────

def _calculate_offsets(num_oferentes: int, num_interpretes: int) -> dict:
    """
    Calcula las filas reales (1-indexed) de cada sección dado el número
    de oferentes e intérpretes que se van a escribir.

    Con template vacío (7 oferentes, 1 intérprete):
      first_interprete = 19, sabana = 20, sumatoria = 21,
      obs = 22, section3 = 24, first_asistente = 25
    """
    num_oferentes = max(1, min(num_oferentes, MAX_OFERENTES))
    num_interpretes = max(1, min(num_interpretes, MAX_INTERPRETES))

    extra_oferentes = max(0, num_oferentes - TEMPLATE_OFERENTE_SLOTS)
    first_interprete = TEMPLATE_FIRST_INTERPRETE_ROW + extra_oferentes

    sabana_row   = first_interprete + num_interpretes
    sumatoria_row = sabana_row + 1
    obs_row       = sumatoria_row + 1   # ocupa filas obs_row y obs_row+1
    section3_row  = obs_row + 2
    first_asistente_row = section3_row + 1

    return {
        "first_interprete":   first_interprete,
        "sabana":             sabana_row,
        "sumatoria":          sumatoria_row,
        "obs":                obs_row,
        "section3":           section3_row,
        "first_asistente":    first_asistente_row,
    }


def _build_row_insertions(num_oferentes: int, num_interpretes: int) -> list:
    """
    Genera la lista de row_insertions para publish_sheet_from_template.
    Las inserciones se procesan en orden; la segunda ajusta su posición
    después de que la primera ya habrá corrido las filas.
    """
    insertions = []

    extra_oferentes = max(0, num_oferentes - TEMPLATE_OFERENTE_SLOTS)
    if extra_oferentes > 0:
        # Insertar filas ANTES de la fila de intérprete (row 19),
        # copiando formato de la última fila de oferente (row 18)
        insertions.append({
            "sheet_name":   SHEET_TAB_NAME,
            "start_row":    TEMPLATE_FIRST_OFERENTE_ROW,
            "base_rows":    TEMPLATE_OFERENTE_SLOTS,
            "total_rows":   num_oferentes,
            "insert_at_row": TEMPLATE_FIRST_INTERPRETE_ROW,
            "template_row": TEMPLATE_FIRST_INTERPRETE_ROW - 1,  # fila 18
        })

    extra_interpretes = max(0, num_interpretes - TEMPLATE_INTERPRETE_SLOTS)
    if extra_interpretes > 0:
        # Después de insertar extra_oferentes, el intérprete está en:
        first_interp_shifted = TEMPLATE_FIRST_INTERPRETE_ROW + extra_oferentes
        # Insertar filas DESPUÉS del primer intérprete, copiando su formato
        insertions.append({
            "sheet_name":   SHEET_TAB_NAME,
            "start_row":    first_interp_shifted,
            "base_rows":    TEMPLATE_INTERPRETE_SLOTS,
            "total_rows":   num_interpretes,
            "insert_at_row": first_interp_shifted + TEMPLATE_INTERPRETE_SLOTS,
            "template_row": first_interp_shifted,
        })

    return insertions


# ── Construcción de escrituras por sección ────────────────────────────────────

def _q(cell: str) -> str:
    """Prefija la celda con el nombre de la pestaña para la API de Sheets."""
    return f"'{SHEET_TAB_NAME}'!{cell}"


def _build_section_1_writes(payload: dict) -> list:
    return build_sheet_updates(
        SHEET_TAB_NAME,
        EXCEL_MAPPING["section_1"],
        payload or {},
    )


def _build_oferentes_writes(oferentes: list) -> list:
    writes = []
    for i, of in enumerate(oferentes[:MAX_OFERENTES]):
        row = TEMPLATE_FIRST_OFERENTE_ROW + i
        nombre = (of.get("nombre_oferente") or of.get("nombre") or "").strip()
        cedula = (of.get("cedula") or of.get("cedula_oferente") or "").strip()
        proceso = (of.get("proceso") or of.get("proceso_observaciones") or "").strip()
        writes += [
            {"range": _q(f"A{row}"), "value": str(i + 1)},
            {"range": _q(f"B{row}"), "value": nombre},
            {"range": _q(f"F{row}"), "value": cedula},
            {"range": _q(f"J{row}"), "value": proceso},
        ]
    return writes


def _build_interpretes_writes(interpretes: list, offsets: dict) -> list:
    writes = []
    for i, interp in enumerate(interpretes[:MAX_INTERPRETES]):
        row = offsets["first_interprete"] + i
        writes += [
            {"range": _q(f"D{row}"), "value": (interp.get("nombre") or "").strip()},
            {"range": _q(f"J{row}"), "value": (interp.get("hora_inicial") or "").strip()},
            {"range": _q(f"M{row}"), "value": (interp.get("hora_final") or "").strip()},
            {"range": _q(f"Q{row}"), "value": (interp.get("total_tiempo") or "").strip()},
        ]
    return writes


def _build_sabana_sumatoria_writes(section_3: dict, offsets: dict) -> list:
    sabana = section_3.get("sabana") or {}
    activo = bool(sabana.get("activo", False))
    horas = float(sabana.get("horas") or 1.0)
    sabana_val = format_sabana_value(activo, horas)
    sumatoria_val = section_3.get("sumatoria_horas") or ""
    return [
        {"range": _q(f"Q{offsets['sabana']}"),    "value": sabana_val},
        {"range": _q(f"Q{offsets['sumatoria']}"), "value": sumatoria_val},
    ]


def _build_asistentes_writes(asistentes: list, offsets: dict) -> list:
    writes = []
    for i, asist in enumerate(asistentes[:MAX_ASISTENTES]):
        row = offsets["first_asistente"] + i
        nombre = (asist.get("nombre") or "").strip()
        cargo = (asist.get("cargo") or "").strip()
        if nombre:
            writes.append({"range": _q(f"C{row}"), "value": nombre})
        if cargo:
            writes.append({"range": _q(f"K{row}"), "value": cargo})
    return writes


# ── Obtener ID del spreadsheet LSC ───────────────────────────────────────────

def _get_lsc_spreadsheet_id() -> str:
    """Lee el ID del spreadsheet LSC de config.json, con fallback a constante."""
    try:
        from formularios.common import _load_json_config
        cfg = _load_json_config("config.json") or {}
        sid = (cfg.get("google_sheets_lsc_template_id") or "").strip()
        if sid:
            return sid
    except Exception:
        pass
    return _LSC_SPREADSHEET_ID_FALLBACK


# ── Exportación a Google Sheets ──────────────────────────────────────────────

def export_to_excel(clear_cache=True):
    if not FORM_CACHE.get("section_1") and cache_file_exists():
        load_cache_from_file()
    raise_validation_error(validate_before_finalize())

    from drive_upload import publish_sheet_from_template

    section_1 = FORM_CACHE.get("section_1", {}) or {}
    empresa_nombre = section_1.get("nombre_empresa") or SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    fecha_servicio = str(section_1.get("fecha_visita") or "").strip()
    base_name = _sanitize_filename(
        f"Servicio Interpretacion LSC - {fecha_servicio}" if fecha_servicio else "Servicio Interpretacion LSC"
    )

    section_3 = FORM_CACHE.get("section_3") or {}
    oferentes  = FORM_CACHE.get("section_2") or []
    interpretes = section_3.get("interpretes") if isinstance(section_3, dict) else []
    interpretes = interpretes or []
    asistentes = FORM_CACHE.get("section_4") or []

    num_oferentes  = max(1, len(oferentes))
    num_interpretes = max(1, len(interpretes))
    offsets = _calculate_offsets(num_oferentes, num_interpretes)
    row_insertions = _build_row_insertions(num_oferentes, num_interpretes) or None

    writes = []
    writes.extend(_build_section_1_writes(FORM_CACHE.get("section_1", {})))
    writes.extend(_build_oferentes_writes(oferentes))
    writes.extend(_build_interpretes_writes(interpretes, offsets))
    writes.extend(_build_sabana_sumatoria_writes(section_3, offsets))
    writes.extend(_build_asistentes_writes(asistentes, offsets))

    try:
        log_excel_event(
            f"LSC export start empresa={empresa_nombre!r} "
            f"oferentes={len(oferentes)} interpretes={len(interpretes)}"
        )
    except Exception:
        pass

    result = publish_sheet_from_template(
        template_id=_get_lsc_spreadsheet_id(),
        sheet_writes=writes,
        base_name=base_name,
        folder_name=_sanitize_filename(empresa_nombre),
        reuse_existing=False,
        row_insertions=row_insertions,
    )

    asistentes_export = [
        {
            "nombre": str((item or {}).get("nombre") or "").strip(),
            "cargo": str((item or {}).get("cargo") or "").strip(),
        }
        for item in (asistentes if isinstance(asistentes, list) else [])
        if isinstance(item, dict) and str((item or {}).get("nombre") or "").strip()
    ]
    participantes = [
        {
            "nombre": str((item or {}).get("nombre_oferente") or (item or {}).get("nombre") or "").strip(),
            "cedula": str((item or {}).get("cedula") or (item or {}).get("cedula_oferente") or "").strip(),
            "proceso": str((item or {}).get("proceso") or "").strip(),
        }
        for item in (oferentes if isinstance(oferentes, list) else [])
        if isinstance(item, dict)
        and (
            str((item or {}).get("nombre_oferente") or (item or {}).get("nombre") or "").strip()
            or str((item or {}).get("cedula") or (item or {}).get("cedula_oferente") or "").strip()
        )
    ]
    sabana = section_3.get("sabana") if isinstance(section_3, dict) else {}
    acta_metadata = {
        "tipo_acta": "interprete_lsc",
        "nit_empresa": str(section_1.get("nit_empresa") or "").strip(),
        "nombre_empresa": str(empresa_nombre or "").strip(),
        "fecha_servicio": fecha_servicio,
        "nombre_profesional": str(
            section_1.get("profesional_asignado") or section_1.get("asesor") or section_1.get("contacto_empresa") or ""
        ).strip(),
        "modalidad_servicio": str(section_1.get("modalidad_profesional_reca") or "").strip(),
        "modalidad_interprete": str(section_1.get("modalidad_interprete") or "").strip(),
        "sumatoria_horas": str(section_3.get("sumatoria_horas") or "").strip(),
        "sabana": {
            "activo": bool((sabana or {}).get("activo", False)),
            "horas": float((sabana or {}).get("horas") or 1.0),
        },
        "asistentes": asistentes_export,
        "participantes": participantes,
        "interpretes": [
            {
                "nombre": str((item or {}).get("nombre") or "").strip(),
                "hora_inicial": str((item or {}).get("hora_inicial") or "").strip(),
                "hora_final": str((item or {}).get("hora_final") or "").strip(),
                "total_tiempo": str((item or {}).get("total_tiempo") or "").strip(),
            }
            for item in (interpretes if isinstance(interpretes, list) else [])
            if isinstance(item, dict) and str((item or {}).get("nombre") or "").strip()
        ],
    }

    if clear_cache:
        clear_cache_file()
        clear_form_cache()

    return {
        "output_path": result.get("webViewLink", ""),
        "drive_file_id": result.get("file_id", ""),
        "already_in_drive": True,
        "tipo_acta": "interprete_lsc",
        "fecha_servicio": fecha_servicio,
        "acta_metadata": acta_metadata,
    }
