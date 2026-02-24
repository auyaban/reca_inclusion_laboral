import os
import json
import time
import re
import shutil

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _get_desktop_dir,
    _normalize_cedula,
    _normalize_text,
    _parse_date_value,
    _sanitize_filename,
    _supabase_get,
    _supabase_upsert,
)


FORM_ID = "contratacion_incluyente"
FORM_NAME = "Contratacion Incluyente"
SHEET_NAME = "5. PROCESO DE CONTRATACION INCL"

FORM_CACHE = {}
SECTION_1_CACHE = {}

DISCAPACIDAD_OPTIONS = [
    "Discapacidad visual perdida total de la vision",
    "Discapacidad visual baja vision",
    "Discapacidad auditiva",
    "Discapacidad auditiva hipoacusia",
    "Trastorno de espectro autista",
    "Discapacidad intelectual",
    "Discapacidad fisica",
    "Discapacidad fisica usuario en silla de ruedas",
    "Discapacidad psicosocial",
    "Discapacidad multiple",
    "No aplica",
]

_DISCAPACIDAD_CATEGORIA_MAP = {
    "discapacidad visual perdida total de la vision": "Visual",
    "discapacidad visual baja vision": "Visual",
    "discapacidad auditiva": "Auditiva",
    "discapacidad auditiva hipoacusia": "Auditiva",
    "trastorno de espectro autista": "Intelectual",
    "discapacidad intelectual": "Intelectual",
    "discapacidad fisica": "Física",
    "discapacidad fisica usuario en silla de ruedas": "Física",
    "discapacidad psicosocial": "Psicosocial",
    "discapacidad multiple": "Múltiple",
    "no aplica": None,
}

LGTBIQ_OPTIONS = ["Si", "No", "No aplica", "Prefiere no responder"]
GRUPO_ETNICO_OPTIONS = ["Si", "No", "No aplica", "Prefiere no responder"]
GRUPO_ETNICO_CUAL_OPTIONS = [
    "Afrocolombiano / Negro / Raizal / Palenquero",
    "Indigena",
    "Gitano (ROM)",
    "NARP",
    "Otro",
    "No aplica",
]
CERTIFICADO_DISCAPACIDAD_OPTIONS = ["Si", "No", "No aplica"]
TIPO_CONTRATO_OPTIONS = [
    "Termino fijo",
    "Termino indefinido",
    "Obra o labor",
    "Prestacion de servicios",
    "Termino indefinido con clausula presuntiva",
    "Nombramiento",
    "Contrato de aprendizaje",
    "Nombramiento provisional",
]
NIVEL_APOYO_OPTIONS = [
    "0. No requiere apoyo.",
    "1. Nivel de apoyo Bajo.",
    "2. Nivel de apoyo medio.",
    "3. Nivel de apoyo alto.",
    "No aplica.",
]
OBS_LECTURA_CONTRATO_OPTIONS = [
    "Se acompana con recordatorio sobre horarios de toma.",
    "Se acompana con recordatorio sobre cantidades de medicamentos.",
    "Se acompana en la administracion del medicamento.",
    "No aplica.",
]
OBS_COMPRENDE_CONTRATO_OPTIONS = [
    "Comprende los medicamentos y su manejo.",
    "No comprende los medicamentos y su manejo.",
    "No aplica.",
]
OBS_TIPO_CONTRATO_OPTIONS = [
    "El vinculado tiene claras las caracteristicas del tipo de contrato.",
    "El vinculado NO tiene claras las caracteristicas del tipo de contrato.",
    "No aplica.",
]
JORNADA_LABORAL_OPTIONS = ["Tiempo completo", "Medio tiempo", "Otro", "No aplica"]
CLAUSULAS_CONTRATO_OPTIONS = [
    "El contrato cuenta con clausulas de confidencialidad.",
    "El contrato cuenta con clausulas de no competencia.",
    "El contrato cuenta con clausulas de permanencia minima.",
    "No aplica.",
]
OBS_CONDICIONES_SALARIALES_OPTIONS = [
    "Se aclaran las condiciones salariales.",
    "No se aclaran las condiciones salariales.",
    "No aplica.",
]
FRECUENCIA_PAGO_OPTIONS = ["Semanal", "Quincenal", "Mensual", "Otro", "No aplica"]
FORMA_PAGO_OPTIONS = ["Transferencia bancaria", "Efectivo", "Cheque", "Otro", "No aplica"]
OBS_PRESTACIONES_OPTIONS = [
    "Conoce los beneficios prestacionales.",
    "No conoce los beneficios prestacionales.",
    "No aplica.",
]
OBS_CONDUCTO_REGULAR_OPTIONS = [
    "Comprende el conducto regular.",
    "No comprende el conducto regular.",
    "No aplica.",
]
OBS_DESCARGOS_OPTIONS = [
    "Comprende el proceso de descargos.",
    "No comprende el proceso de descargos.",
    "No aplica.",
]
OBS_TRAMITES_OPTIONS = [
    "Comprende los tramites administrativos.",
    "No comprende los tramites administrativos.",
    "No aplica.",
]
OBS_PERMISOS_OPTIONS = [
    "Comprende el proceso de permisos.",
    "No comprende el proceso de permisos.",
    "No aplica.",
]
OBS_CAUSALES_OPTIONS = [
    "Conoce las causales de terminacion de contrato.",
    "No conoce las causales de terminacion de contrato.",
    "No aplica.",
]
OBS_RUTAS_OPTIONS = [
    "Conoce las rutas de atencion y denuncia.",
    "No conoce las rutas de atencion y denuncia.",
    "No aplica.",
]

EVALUADOR_NOMBRES = [
    "Sandra Milena Pachon Rojas",
    "Sara Zambrano",
    "Alejandra Perez",
    "Lenny Lugo",
    "Angie Diaz",
    "Adriana Viveros",
    "Janeth Camargo",
    "Gabriela Rubiano Isaza",
    "Andres Montes",
    "Sara Sanchez",
    "Catalina Salazar",
]

EVALUADOR_CARGOS = [
    "Coordinadora de inclusion laboral",
    "Coordinacion de inclusion laboral",
    "Gestora de inclusion laboral",
    "Profesional de apoyo de inclusion laboral",
    "Gestor de inclusion laboral",
    "Lider Empleo Inclusivo",
    "Gestora de proyectos y desarrollo",
    "Profesional de inclusion laboral",
    "Directora Fundacion Reca",
]

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

EXCEL_MAPPING = {
    "section_1": {
        "fecha_visita": "D7",
        "modalidad": "L7",
        "nombre_empresa": "D8",
        "ciudad_empresa": "L8",
        "direccion_empresa": "D9",
        "nit_empresa": "L9",
        "correo_1": "D10",
        "telefono_empresa": "L10",
        "contacto_empresa": "D11",
        "cargo": "L11",
        "caja_compensacion": "D12",
        "sede_empresa": "L12",
        "asesor": "D13",
        "profesional_asignado": "L13",
    },
    "section_7": {
        "start_row": 74,
        "rows": 3,
        "nombre_col": "C",
        "cargo_col": "K",
    },
}

SECTION_2_ANCHOR = "2. DATOS DEL VINCULADO"
SECTION_6_ANCHOR = "6. AJUSTES RAZONABLES / RECOMENDACIONES AL PROCESO DE CONTRATACION"
SECTION_2_TEMPLATE_ANCHOR_ROW = 14
SECTION_2_LAST_COLUMN = "R"

SECTION_2_CELL_MAP = {
    "numero": ("A", 18),
    "nombre_oferente": ("C", 18),
    "cedula": ("H", 18),
    "certificado_porcentaje": ("K", 18),
    "discapacidad": ("L", 18),
    "telefono_oferente": ("O", 18),
    "genero": ("C", 19),
    "correo_oferente": ("G", 19),
    "fecha_nacimiento": ("M", 19),
    "edad": ("Q", 19),
    "lgtbiq": ("E", 20),
    "grupo_etnico": ("L", 20),
    "grupo_etnico_cual": ("O", 20),
    "cargo_oferente": ("C", 21),
    "contacto_emergencia": ("I", 21),
    "parentesco": ("M", 21),
    "telefono_emergencia": ("Q", 21),
    "certificado_discapacidad": ("F", 22),
    "lugar_firma_contrato": ("L", 22),
    "fecha_firma_contrato": ("Q", 22),
    "tipo_contrato": ("G", 24),
    "fecha_fin": ("N", 24),
    "desarrollo_actividad": ("A", 26),
    "contrato_lee_nivel_apoyo": ("G", 30),
    "contrato_lee_observacion": ("L", 30),
    "contrato_lee_nota": ("M", 31),
    "contrato_comprendido_nivel_apoyo": ("G", 32),
    "contrato_comprendido_observacion": ("L", 32),
    "contrato_comprendido_nota": ("M", 33),
    "contrato_tipo_nivel_apoyo": ("G", 34),
    "contrato_tipo_observacion": ("L", 34),
    "contrato_tipo_contrato": ("L", 35),
    "contrato_jornada": ("L", 36),
    "contrato_clausulas": ("L", 37),
    "contrato_tipo_nota": ("M", 38),
    "condiciones_salariales_nivel_apoyo": ("G", 39),
    "condiciones_salariales_observacion": ("L", 39),
    "condiciones_salariales_frecuencia_pago": ("L", 40),
    "condiciones_salariales_forma_pago": ("L", 41),
    "condiciones_salariales_nota": ("M", 42),
    "prestaciones_cesantias_nivel_apoyo": ("G", 45),
    "prestaciones_cesantias_observacion": ("L", 45),
    "prestaciones_cesantias_nota": ("M", 46),
    "prestaciones_auxilio_transporte_nivel_apoyo": ("G", 47),
    "prestaciones_auxilio_transporte_observacion": ("L", 47),
    "prestaciones_auxilio_transporte_nota": ("M", 48),
    "prestaciones_prima_nivel_apoyo": ("G", 49),
    "prestaciones_prima_observacion": ("L", 49),
    "prestaciones_prima_nota": ("M", 50),
    "prestaciones_seguridad_social_nivel_apoyo": ("G", 51),
    "prestaciones_seguridad_social_observacion": ("L", 51),
    "prestaciones_seguridad_social_nota": ("M", 52),
    "prestaciones_vacaciones_nivel_apoyo": ("G", 53),
    "prestaciones_vacaciones_observacion": ("L", 53),
    "prestaciones_vacaciones_nota": ("M", 54),
    "prestaciones_auxilios_beneficios_nivel_apoyo": ("G", 55),
    "prestaciones_auxilios_beneficios_observacion": ("L", 55),
    "prestaciones_auxilios_beneficios_nota": ("M", 56),
    "conducto_regular_nivel_apoyo": ("G", 59),
    "conducto_regular_observacion": ("L", 59),
    "descargos_observacion": ("L", 60),
    "tramites_observacion": ("L", 61),
    "permisos_observacion": ("L", 62),
    "conducto_regular_nota": ("M", 63),
    "causales_fin_nivel_apoyo": ("G", 64),
    "causales_fin_observacion": ("L", 64),
    "causales_fin_nota": ("M", 65),
    "rutas_atencion_nivel_apoyo": ("G", 66),
    "rutas_atencion_observacion": ("L", 66),
    "rutas_atencion_nota": ("M", 67),
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
    return os.path.join(_get_cache_dir(), "contratacion_incluyente.json")


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
    if payload is None:
        payload = {}
    FORM_CACHE[section_id] = payload


def get_form_cache():
    return dict(FORM_CACHE)




def _infer_discapacidad_categoria(value):
    if not value:
        return None
    normalized = _normalize_text(value)
    if "no aplica" in normalized:
        return None
    if "multiple" in normalized:
        return "Múltiple"
    if "visual" in normalized:
        return "Visual"
    if "auditiva" in normalized or "hipoacusia" in normalized:
        return "Auditiva"
    if "fisica" in normalized:
        return "Física"
    if "psicosocial" in normalized:
        return "Psicosocial"
    if "intelectual" in normalized or "autismo" in normalized or "autista" in normalized:
        return "Intelectual"
    return _DISCAPACIDAD_CATEGORIA_MAP.get(normalized)




def get_usuarios_reca_cedulas(env_path=".env"):
    params = {
        "select": "cedula_usuario",
        "cedula_usuario": "not.is.null",
        "order": "cedula_usuario.asc",
    }
    data = _supabase_get("usuarios_reca", params, env_path=env_path)
    return [row.get("cedula_usuario") for row in data if row.get("cedula_usuario")]


def get_usuario_reca_by_cedula(cedula, env_path=".env"):
    normalized = _normalize_cedula(cedula)
    if not normalized:
        return None
    select_cols = ",".join(
        [
            "cedula_usuario",
            "nombre_usuario",
            "genero_usuario",
            "discapacidad_usuario",
            "discapacidad_detalle",
            "certificado_porcentaje",
            "telefono_oferente",
            "fecha_nacimiento",
            "cargo_oferente",
            "contacto_emergencia",
            "parentesco",
            "telefono_emergencia",
            "correo_oferente",
            "lgtbiq",
            "grupo_etnico",
            "grupo_etnico_cual",
            "certificado_discapacidad",
            "lugar_firma_contrato",
            "fecha_firma_contrato",
            "tipo_contrato",
            "fecha_fin",
        ]
    )
    params = {
        "select": select_cols,
        "cedula_usuario": f"eq.{normalized}",
        "limit": 1,
    }
    data = _supabase_get("usuarios_reca", params, env_path=env_path)
    return data[0] if data else None


def _find_template_path():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    templates_dir = os.path.join(base_dir, "templates")
    if not os.path.isdir(templates_dir):
        raise FileNotFoundError("No existe la carpeta templates.")
    for name in os.listdir(templates_dir):
        if name.startswith("~$"):
            continue
        normalized = _normalize_text(name).replace("_", "")
        if "contratacion" in normalized and "incluyente" in normalized and normalized.endswith(".xlsx"):
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template de contratacion incluyente.")


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
        log_dir = _get_log_dir()
        log_path = os.path.join(log_dir, "excel_log.txt")
        reset_log = False
        if os.path.exists(log_path):
            try:
                if os.path.getsize(log_path) >= 5 * 1024 * 1024:
                    reset_log = True
            except OSError:
                reset_log = True
        if reset_log:
            with open(log_path, "w", encoding="utf-8") as log_file:
                log_file.write("")
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] {message}\n")
    except OSError:
        return


def _ensure_output_path():
    template_path = _find_template_path()
    desktop = _get_desktop_dir()
    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    safe_company = _sanitize_filename(empresa_nombre)
    if not safe_company:
        safe_company = "Empresa"
    output_dir = os.path.join(desktop, "Formatos Inclusion Laboral", safe_company)
    os.makedirs(output_dir, exist_ok=True)
    process_name = "Proceso de Contratacion Incluyente"
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
    try:
        return workbook.Worksheets(SHEET_NAME)
    except Exception as exc:
        raise KeyError(f"No existe la hoja {SHEET_NAME}.") from exc


def get_empresa_by_nit(nit, env_path=".env"):
    return evaluacion_accesibilidad.get_empresa_by_nit(nit, env_path=env_path)


def get_empresa_by_nombre(nombre, env_path=".env"):
    return evaluacion_accesibilidad.get_empresa_by_nombre(nombre, env_path=env_path)


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=10):
    return evaluacion_accesibilidad.get_empresas_by_nombre_prefix(prefix, env_path=env_path, limit=limit)


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
    if payload is None:
        raise ValueError("section_2 requerida")
    set_section_cache("section_2", payload)
    FORM_CACHE["_last_section"] = "section_2"
    save_cache_to_file()
    return payload


def confirm_section_6(payload):
    if payload is None:
        raise ValueError("section_6 requerida")
    set_section_cache("section_6", payload)
    FORM_CACHE["_last_section"] = "section_6"
    save_cache_to_file()
    return payload


def confirm_section_7(payload):
    if payload is None:
        raise ValueError("section_7 requerida")
    set_section_cache("section_7", payload)
    FORM_CACHE["_last_section"] = "section_7"
    save_cache_to_file()
    return payload


def sync_usuarios_reca(env_path=".env"):
    data = FORM_CACHE.get("section_2")
    if not data and cache_file_exists():
        load_cache_from_file()
        data = FORM_CACHE.get("section_2")
    if not data:
        return 0

    rows = []
    for entry in data:
        cedula = _normalize_cedula(entry.get("cedula"))
        if not cedula:
            continue
        discapacidad_detalle = (entry.get("discapacidad") or "").strip()
        discapacidad_usuario = _infer_discapacidad_categoria(discapacidad_detalle)
        row = {
            "cedula_usuario": cedula,
            "nombre_usuario": (entry.get("nombre_oferente") or "").strip(),
            "genero_usuario": (entry.get("genero") or "").strip(),
            "discapacidad_usuario": discapacidad_usuario,
            "discapacidad_detalle": discapacidad_detalle or None,
            "certificado_porcentaje": (entry.get("certificado_porcentaje") or "").strip(),
            "telefono_oferente": (entry.get("telefono_oferente") or "").strip(),
            "fecha_nacimiento": _parse_date_value(entry.get("fecha_nacimiento")),
            "cargo_oferente": (entry.get("cargo_oferente") or "").strip(),
            "contacto_emergencia": (entry.get("contacto_emergencia") or "").strip(),
            "parentesco": (entry.get("parentesco") or "").strip(),
            "telefono_emergencia": (entry.get("telefono_emergencia") or "").strip(),
            "correo_oferente": (entry.get("correo_oferente") or "").strip(),
            "lgtbiq": (entry.get("lgtbiq") or "").strip(),
            "grupo_etnico": (entry.get("grupo_etnico") or "").strip(),
            "grupo_etnico_cual": (entry.get("grupo_etnico_cual") or "").strip(),
            "certificado_discapacidad": (entry.get("certificado_discapacidad") or "").strip(),
            "lugar_firma_contrato": (entry.get("lugar_firma_contrato") or "").strip(),
            "fecha_firma_contrato": _parse_date_value(entry.get("fecha_firma_contrato")),
            "tipo_contrato": (entry.get("tipo_contrato") or "").strip(),
            "fecha_fin": _parse_date_value(entry.get("fecha_fin")) or (entry.get("fecha_fin") or "").strip(),
            "resultado_certificado": (entry.get("resultado_certificado") or "").strip(),
            "pendiente_otros_oferentes": (entry.get("pendiente_otros_oferentes") or "").strip(),
            "cuenta_pension": (entry.get("cuenta_pension") or "").strip(),
            "tipo_pension": (entry.get("tipo_pension") or "").strip(),
        }
        cleaned = {k: v for k, v in row.items() if v not in ("", None)}
        rows.append(cleaned)
    if rows:
        _supabase_upsert("usuarios_reca", rows, env_path=env_path, on_conflict="cedula_usuario")
        cedulas = [row.get("cedula_usuario") for row in rows if row.get("cedula_usuario")]
        preview = ", ".join(cedulas[:10])
        extra = "" if len(cedulas) <= 10 else f" (+{len(cedulas) - 10} mas)"
        _log_excel(f"SUPABASE usuarios_reca upsert count={len(rows)} cedulas={preview}{extra}")
    return len(rows)


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
        if value_norm == target:
            return row
    for row in range(start_row, end_row + 1):
        value = ws.Cells(row, 1).Value
        if not value:
            continue
        value_norm = _normalize_text(str(value))
        if target in value_norm:
            if target.startswith("2.") or target.startswith("6."):
                if value_norm.startswith(target):
                    return row
            else:
                return row
    raise ValueError(f"No se encontró el texto '{text}' en la columna A.")


def _insert_person_block(ws, start_row, block_height, insert_at):
    start_end = start_row + block_height - 1
    dest_end = insert_at + block_height - 1
    source = ws.Range(f"A{start_row}:{SECTION_2_LAST_COLUMN}{start_end}")
    dest = ws.Range(f"A{insert_at}:{SECTION_2_LAST_COLUMN}{dest_end}")
    source.Copy()
    dest.Insert(Shift=-4121)
    for row_offset in range(block_height):
        ws.Rows(insert_at + row_offset).RowHeight = ws.Rows(start_row + row_offset).RowHeight
    ws.Application.CutCopyMode = False


def _write_section_2(ws, oferentes):
    if not oferentes:
        return
    total_oferentes = len(oferentes)
    if 2 <= total_oferentes <= 4:
        ws.Range("F1").Value = "PROCESO CONTRATACION INCLUYENTE GRUPAL - 2 A 4 OFERENTES"
    elif 5 <= total_oferentes <= 7:
        ws.Range("F1").Value = "PROCESO CONTRATACION INCLUYENTE GRUPAL - 5 A 7 VINCULADOS"
    elif total_oferentes >= 8:
        ws.Range("F1").Value = "PROCESO CONTRATACION INCLUYENTE GRUPAL - MAS DE 8 VINCULADOS"
    start_row = _find_row_by_text(ws, SECTION_2_ANCHOR)
    next_row = _find_row_by_text(ws, SECTION_6_ANCHOR)
    block_height = next_row - start_row
    if block_height <= 0:
        raise ValueError("Anclas de seccion 2 invalidas.")
    _log_excel(
        f"SECTION section=section_2 start_row={start_row} next_row={next_row} block_height={block_height} total={len(oferentes)}"
    )
    for idx in range(1, len(oferentes)):
        insert_at = start_row + (block_height * idx)
        _insert_person_block(ws, start_row, block_height, insert_at)
        _log_excel(f"INSERT section=section_2 rows={block_height} at={insert_at}")

    for idx, entry in enumerate(oferentes):
        base_row = start_row + (block_height * idx)
        for field_id, (col, row) in SECTION_2_CELL_MAP.items():
            offset = row - SECTION_2_TEMPLATE_ANCHOR_ROW
            target_row = base_row + offset
            value = entry.get(field_id, "")
            if value == "":
                continue
            _log_excel(
                f"WRITE section=section_2 cell={col}{target_row} key={field_id} value={value!r}"
            )
            ws.Range(f"{col}{target_row}").Value = value


def _write_section_6(ws, payload):
    if not payload:
        return
    anchor_row = _find_row_by_text(ws, SECTION_6_ANCHOR)
    ajustes_row = anchor_row + 1
    ajustes_value = payload.get("ajustes_recomendaciones", "")
    _log_excel(
        f"WRITE section=section_6 cell=A{ajustes_row} key=ajustes_recomendaciones value={ajustes_value!r}"
    )
    ws.Range(f"A{ajustes_row}").Value = ajustes_value


def _write_section_7(ws, payload):
    if not payload:
        return
    mapping = EXCEL_MAPPING.get("section_7", {})
    title_row = _find_row_by_text(ws, "7. ASISTENTES")
    start_row = title_row + 1
    base_rows = mapping.get("rows", 3)
    nombre_col = mapping.get("nombre_col", "C")
    cargo_col = mapping.get("cargo_col", "K")
    total = len(payload)
    if total > base_rows:
        insert_at = start_row + base_rows
        template_row = start_row + base_rows - 1
        for _ in range(total - base_rows):
            ws.Rows(insert_at).Insert()
            ws.Rows(template_row).Copy(ws.Rows(insert_at))
            insert_at += 1
            _log_excel(
                f"INSERT section=section_7 rows=1 at={insert_at - 1}"
            )
    for idx, entry in enumerate(payload):
        row = start_row + idx
        nombre = entry.get("nombre", "")
        cargo = entry.get("cargo", "")
        _log_excel(
            f"WRITE section=section_7 cell={nombre_col}{row} key=nombre value={nombre!r}"
        )
        _log_excel(
            f"WRITE section=section_7 cell={cargo_col}{row} key=cargo value={cargo!r}"
        )
        ws.Range(f"{nombre_col}{row}").Value = nombre
        ws.Range(f"{cargo_col}{row}").Value = cargo


def _write_section_1(ws, payload):
    if not payload:
        payload = SECTION_1_CACHE
    if not payload:
        try:
            if load_cache_from_file():
                payload = FORM_CACHE.get("section_1", {}) or SECTION_1_CACHE
        except Exception:
            payload = payload or {}
    mapping = EXCEL_MAPPING.get("section_1", {})
    for key, cell in mapping.items():
        if key in payload:
            value = payload.get(key)
            ws.Range(cell).Value = value
            _log_excel(
                f"WRITE section=section_1 cell={cell} key={key} value={value!r}"
            )


def export_to_excel(clear_cache=True):
    output_path = _ensure_output_path()
    if not FORM_CACHE.get("section_1") and cache_file_exists():
        load_cache_from_file()
    _log_excel(f"START export_all output={output_path}")
    try:
        import win32com.client as win32
    except ImportError as exc:
        _log_excel("ERROR export_all error=pywin32_not_installed")
        raise RuntimeError("pywin32 no esta instalado. Instala con pip install pywin32.") from exc
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(output_path)
        ws = _get_sheet_by_name(wb)
        _write_section_1(ws, FORM_CACHE.get("section_1", {}))
        _write_section_2(ws, FORM_CACHE.get("section_2", []))
        _write_section_6(ws, FORM_CACHE.get("section_6", {}))
        _write_section_7(ws, FORM_CACHE.get("section_7", []))
        wb.Save()
        _log_excel("SUCCESS export_all")
    except Exception as exc:
        _log_excel(f"ERROR export_all error={exc!r}")
        raise
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
    if clear_cache:
        clear_cache_file()
        clear_form_cache()
    return output_path

