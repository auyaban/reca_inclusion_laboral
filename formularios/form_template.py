import json
import os
import shutil
import time

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _get_desktop_dir,
    _normalize_text,
    _sanitize_filename,
    _supabase_get,
)


# Copy this file into a new folder and rename constants/functions.
FORM_ID = "nuevo_formulario"
FORM_NAME = "Nuevo Formulario"
SHEET_NAME = "NOMBRE DE HOJA EN EXCEL"

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
            "label": "Contacto de la empresa",
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
            "id": "caja_compensacion",
            "label": "Empresa afiliada a Caja de Compensación",
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
        {
            "id": "asesor",
            "label": "Asesor",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "profesional_asignado",
            "label": "Profesional asignado RECA",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
    ],
}

SECTION_1_SUPABASE_MAP = evaluacion_accesibilidad.SECTION_1_SUPABASE_MAP.copy()


# Define your sections and mappings.
# Example:
# SECTION_2 = {...}
# EXCEL_MAPPING = {
#     "section_1": {"fecha_visita": "F7", ...},
#     "section_2": {...},
# }


def register_form():
    return {"id": FORM_ID, "name": FORM_NAME}


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
    if not nit:
        return None
    nit = "".join(str(nit).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {"select": select_cols, "nit_empresa": f"eq.{nit}", "limit": 1}
    data = _supabase_get("empresas", params, env_path=env_path)
    return data[0] if data else None


def get_empresa_by_nombre(nombre, env_path=".env"):
    if not nombre:
        return None
    nombre = " ".join(str(nombre).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {"select": select_cols, "nombre_empresa": f"ilike.{nombre}", "limit": 2}
    data = _supabase_get("empresas", params, env_path=env_path)
    if not data:
        return None
    if len(data) > 1:
        raise ValueError("Hay más de una empresa con ese nombre. Usa el NIT.")
    return data[0]


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=10):
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


def _find_template_path(template_keyword):
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    templates_dir = os.path.join(base_dir, "templates")
    if not os.path.isdir(templates_dir):
        raise FileNotFoundError("No existe la carpeta templates.")
    for name in os.listdir(templates_dir):
        if name.startswith("~$"):
            continue
        normalized = _normalize_text(name).replace("_", "")
        if template_keyword in normalized and normalized.endswith(".xlsx"):
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template del formulario.")


def _ensure_output_path(template_keyword):
    output_path = FORM_CACHE.get("_output_path")
    if output_path and os.path.exists(output_path):
        return output_path
    template_path = _find_template_path(template_keyword)
    desktop = _get_desktop_dir()
    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    safe_company = _sanitize_filename(empresa_nombre)
    output_dir = os.path.join(desktop, "Formatos Inclusion Laboral", safe_company)
    os.makedirs(output_dir, exist_ok=True)
    output_name = f"{FORM_NAME} - {safe_company}.xlsx"
    output_path = os.path.join(output_dir, output_name)
    if not os.path.exists(output_path):
        shutil.copy2(template_path, output_path)
    FORM_CACHE["_output_path"] = output_path
    return output_path


def export_to_excel(clear_cache=True):
    # Implement writing logic for each mapped section.
    # Keep this signature to be compatible with app.py flow.
    output_path = _ensure_output_path(template_keyword="nuevoformulario")
    if clear_cache:
        clear_form_cache()
        clear_cache_file()
    return output_path

