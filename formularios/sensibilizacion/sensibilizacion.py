import json
import os
import shutil
import time

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import _get_desktop_dir, _normalize_text, _sanitize_filename


FORM_ID = "sensibilizacion"
FORM_NAME = "Sensibilizacion"
SHEET_NAME = "8. SENSIBILIZACION"

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


def _find_template_path():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    templates_dir = os.path.join(base_dir, "templates")
    if not os.path.isdir(templates_dir):
        raise FileNotFoundError("No existe la carpeta templates.")
    for name in os.listdir(templates_dir):
        if name.startswith("~$"):
            continue
        normalized = _normalize_text(name).replace("_", "")
        if "sensibilizacion" in normalized and normalized.endswith(".xlsx"):
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template de sensibilizacion.")


def _ensure_output_path():
    template_path = _find_template_path()
    desktop = _get_desktop_dir()
    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    safe_company = _sanitize_filename(empresa_nombre) or "Empresa"
    output_dir = os.path.join(desktop, "Formatos Inclusion Laboral", safe_company)
    os.makedirs(output_dir, exist_ok=True)
    process_name = "Sensibilizacion"
    output_name = f"{process_name} - {safe_company}.xlsx"
    output_path = os.path.join(output_dir, output_name)
    shutil.copy2(template_path, output_path)
    FORM_CACHE["_output_path"] = output_path
    return output_path


def _get_sheet_by_name(workbook):
    target = _normalize_text(SHEET_NAME).replace(" ", "")
    for ws in workbook.Worksheets:
        name_norm = _normalize_text(ws.Name).replace(" ", "")
        if name_norm == target:
            return ws
    raise KeyError(f"No existe la hoja {SHEET_NAME}.")


def _find_row_by_text(ws, text):
    cell = ws.Columns("A").Find(What=text, LookAt=1)
    if cell is not None:
        return cell.Row
    cell = ws.Columns("A").Find(What=text, LookAt=2)
    if cell is not None:
        return cell.Row
    target = _normalize_text(text)
    used = ws.UsedRange
    start_row = used.Row
    end_row = used.Row + used.Rows.Count - 1
    for row in range(start_row, end_row + 1):
        value = ws.Cells(row, 1).Value
        if not value:
            continue
        value_norm = _normalize_text(str(value))
        if value_norm == target or target in value_norm:
            return row
    raise ValueError(f"No se encontró el texto '{text}' en la columna A.")


def _write_section_1(ws, payload):
    if not payload:
        payload = SECTION_1_CACHE
    if not payload:
        return
    mapping = EXCEL_MAPPING.get("section_1", {})
    for key, cell in mapping.items():
        if key in payload:
            ws.Range(cell).Value = payload.get(key)


def _write_section_3(ws, payload):
    if not payload:
        return
    anchor = _find_row_by_text(ws, "3. OBSERVACIONES")
    texto = (payload.get("observaciones") or "").strip()
    if texto:
        ws.Range(f"A{anchor + 1}").Value = texto


def _write_section_5(ws, payload):
    if not payload:
        return
    title_row = _find_row_by_text(ws, "5. ASISTENTES")
    start_row = title_row + 1
    base_rows = SECTION_5["rows"]
    total = len(payload)

    if total > base_rows:
        insert_at = start_row + base_rows
        template_row = start_row + base_rows - 1
        for _ in range(total - base_rows):
            ws.Rows(insert_at).Insert()
            ws.Rows(template_row).Copy(ws.Rows(insert_at))
            ws.Rows(insert_at).RowHeight = ws.Rows(template_row).RowHeight
            insert_at += 1

    for idx, entry in enumerate(payload):
        row = start_row + idx
        nombre = (entry.get("nombre") or "").strip()
        cargo = (entry.get("cargo") or "").strip()
        if nombre:
            ws.Range(f"C{row}").Value = nombre
        if cargo:
            ws.Range(f"K{row}").Value = cargo


def export_to_excel(clear_cache=True):
    output_path = _ensure_output_path()
    if not FORM_CACHE.get("section_1") and cache_file_exists():
        load_cache_from_file()
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("pywin32 no esta instalado. Instala con pip install pywin32.") from exc
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(output_path)
        ws = _get_sheet_by_name(wb)
        _write_section_1(ws, FORM_CACHE.get("section_1", {}))
        _write_section_3(ws, FORM_CACHE.get("section_3", {}))
        _write_section_5(ws, FORM_CACHE.get("section_5", []))
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
    if clear_cache:
        clear_cache_file()
        clear_form_cache()
    return output_path


