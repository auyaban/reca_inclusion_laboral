import json
import os
import time

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _normalize_text,
    _sanitize_filename,
    build_sheet_updates,
)
from logging_utils import log_excel_event


FORM_ID = "sensibilizacion"
FORM_NAME = "Sensibilizacion"
SHEET_NAME = "8. SENSIBILIZACIÓN"

FORM_CACHE = {}
SECTION_1_CACHE = {}

SECTION_1 = {
    "title": "1. DATOS DE LA EMPRESA",
    "nit_lookup_field": "nit_empresa",
    "fields": [
        {"id": "fecha_visita", "label": "Fecha de la visita", "source": "input"},
        {
            "id": "modalidad",
            "label": "Modalidad",
            "source": "input",
            "options": ["Presencial", "Virtual", "Mixta", "No aplica"],
        },
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
            "label": "Dirección de la empresa",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {"id": "nit_empresa", "label": "Número de NIT", "source": "input"},
        {
            "id": "correo_1",
            "label": "Correo electrónico",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "telefono_empresa",
            "label": "Teléfonos",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "contacto_empresa",
            "label": "Persona que atiende la visita",
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
        {
            "id": "asesor",
            "label": "Asesor",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "sede_empresa",
            "label": "Sede Compensar",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
    ],
}

SECTION_2 = {"title": "2. PRESENTACION DE LOS TEMAS DE LA SENSIBILIZACION"}
SECTION_3 = {"title": "3. OBSERVACIONES"}
SECTION_4 = {"title": "4. REGISTRO FOTOGRAFICO"}
SECTION_5 = {"title": "5. ASISTENTES", "rows": 4}

SECTION_1_SUPABASE_MAP = evaluacion_accesibilidad.SECTION_1_SUPABASE_MAP.copy()

EXCEL_MAPPING = {
    "section_1": {
        "fecha_visita": "D7",
        "modalidad": "N7",
        "nombre_empresa": "D8",
        "ciudad_empresa": "N8",
        "direccion_empresa": "D9",
        "nit_empresa": "N9",
        "correo_1": "D10",
        "telefono_empresa": "N10",
        "contacto_empresa": "D11",
        "cargo": "N11",
        "asesor": "D12",
        "sede_empresa": "N12",
    }
}


def register_form():
    return {"id": FORM_ID, "name": FORM_NAME, "module": __name__}


def _get_cache_dir():
    base = os.getenv("LOCALAPPDATA")
    if not base:
        userprofile = os.getenv("USERPROFILE")
        if userprofile:
            base = os.path.join(userprofile, "AppData", "Local")
    if not base:
        base = os.getcwd()
    cache_dir = os.path.join(base, "RECA", "cache")
    os.makedirs(cache_dir, exist_ok=True)
    return cache_dir


def _get_cache_path():
    return os.path.join(_get_cache_dir(), f"{FORM_ID}.json")


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


def get_empresa_by_nit(nit, env_path=".env"):
    return evaluacion_accesibilidad.get_empresa_by_nit(nit, env_path=env_path)


def get_empresa_by_nombre(nombre, env_path=".env"):
    return evaluacion_accesibilidad.get_empresa_by_nombre(nombre, env_path=env_path)


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=50):
    return evaluacion_accesibilidad.get_empresas_by_nombre_prefix(
        prefix, env_path=env_path, limit=limit
    )


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


def confirm_section_2(payload=None):
    set_section_cache("section_2", payload or {})
    FORM_CACHE["_last_section"] = "section_2"
    save_cache_to_file()
    return payload or {}


def confirm_section_3(payload):
    if payload is None:
        raise ValueError("section_3 requerida")
    set_section_cache("section_3", payload)
    FORM_CACHE["_last_section"] = "section_3"
    save_cache_to_file()
    return payload


def confirm_section_4(payload=None):
    set_section_cache("section_4", payload or {})
    FORM_CACHE["_last_section"] = "section_4"
    save_cache_to_file()
    return payload or {}


def confirm_section_5(payload):
    if payload is None:
        raise ValueError("section_5 requerida")
    set_section_cache("section_5", payload)
    FORM_CACHE["_last_section"] = "section_5"
    save_cache_to_file()
    return payload


def _get_log_dir():
    output_path = FORM_CACHE.get("_output_path")
    if output_path:
        base_dir = os.path.dirname(output_path)
    else:
        base_dir = os.getcwd()
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    return log_dir


def _log_excel(message):
    try:
        log_excel_event(message)
    except Exception:
        return


SECTION_3_TITLE_ROW = 25
SECTION_3_OBSERVACIONES_ROW = 26
SECTION_5_TITLE_ROW = 31
SECTION_5_START_ROW = 32
SECTION_5_NOMBRE_COL = "C"
SECTION_5_CARGO_COL = "K"


def ws_write(ws, cell, value):
    try:
        ws[cell] = value
    except Exception:
        return


def _build_section_1_writes(payload):
    if not payload:
        payload = SECTION_1_CACHE
    return build_sheet_updates(SHEET_NAME, EXCEL_MAPPING.get("section_1", {}), payload or {})


def _build_section_3_writes(payload):
    if not payload:
        return []
    texto = (payload.get("observaciones") or "").strip()
    if not texto:
        return []
    return [{"range": f"'{SHEET_NAME}'!A{SECTION_3_OBSERVACIONES_ROW}", "value": texto}]


def _build_section_5_writes(payload):
    if not payload:
        return []
    writes = []
    for idx, entry in enumerate(payload):
        row = SECTION_5_START_ROW + idx
        nombre = (entry.get("nombre") or "").strip()
        cargo = (entry.get("cargo") or "").strip()
        if nombre:
            writes.append({"range": f"'{SHEET_NAME}'!{SECTION_5_NOMBRE_COL}{row}", "value": nombre})
        if cargo:
            writes.append({"range": f"'{SHEET_NAME}'!{SECTION_5_CARGO_COL}{row}", "value": cargo})
    return writes


def _build_section_5_row_insertions(payload):
    if not payload:
        return []
    base_rows = int(SECTION_5.get("rows", 4) or 4)
    total_rows = len(payload)
    if total_rows <= base_rows:
        return []
    return [
        {
            "sheet_name": SHEET_NAME,
            "start_row": SECTION_5_START_ROW,
            "base_rows": base_rows,
            "total_rows": total_rows,
        }
    ]


def _write_section_3(ws, payload):
    for write in _build_section_3_writes(payload):
        cell = str(write.get("range") or "").rsplit("!", 1)[-1].replace("'", "")
        ws_write(ws, cell, write.get("value", ""))


def _write_section_5(ws, payload):
    for write in _build_section_5_writes(payload):
        cell = str(write.get("range") or "").rsplit("!", 1)[-1].replace("'", "")
        ws_write(ws, cell, write.get("value", ""))


def export_to_excel(clear_cache=True):
    if not FORM_CACHE.get("section_1") and cache_file_exists():
        load_cache_from_file()

    from google_sheets_client import get_master_template_id
    from drive_upload import publish_sheet_from_template

    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    base_name = _sanitize_filename(empresa_nombre)

    writes = []
    writes.extend(_build_section_1_writes(FORM_CACHE.get("section_1", {})))
    writes.extend(_build_section_3_writes(FORM_CACHE.get("section_3", {})))
    writes.extend(_build_section_5_writes(FORM_CACHE.get("section_5", [])))
    row_insertions = _build_section_5_row_insertions(FORM_CACHE.get("section_5", []))

    result = publish_sheet_from_template(
        template_id=get_master_template_id(),
        sheet_writes=writes,
        base_name=base_name,
        folder_name=_sanitize_filename(empresa_nombre),
        row_insertions=row_insertions or None,
    )

    if clear_cache:
        clear_cache_file()
        clear_form_cache()

    return {
        "output_path": result.get("webViewLink", ""),
        "drive_file_id": result.get("file_id", ""),
        "already_in_drive": True,
    }
