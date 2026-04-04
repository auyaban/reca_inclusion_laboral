import os
import json
import time
from datetime import datetime
from functools import lru_cache

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _get_local_app_cache_dir,
    _normalize_cedula,
    _normalize_decimal_value,
    _normalize_text,
    _parse_date_value,
    _coerce_excel_decimal_value,
    _sanitize_filename,
    _supabase_get,
    _supabase_upsert_with_queue,
)
from formularios.finalize_validation import (
    append_missing_issue,
    field_pairs,
    humanize_field_id,
    is_meaningful,
    raise_validation_error,
    require_value,
    validate_dynamic_rows,
)
from logging_utils import log_excel_event


FORM_ID = "contratacion_incluyente"
FORM_NAME = "Contratacion Incluyente"
SHEET_NAME = "5. CONTRATACIÓN INCLUYENTE"

TEMPLATE_VARIANT_INDIVIDUAL = "individual"
TEMPLATE_VARIANT_GROUP_2_PLUS = "group_2_plus"

# Vinculado block geometry (unified format — one sheet for individual & group)
VINCULADO_BLOCK_HEIGHT = 52            # rows per vinculado block (rows 16-67 for first)
VINCULADO_FIRST_BLOCK_START_ROW = 16
VINCULADO_SECOND_BLOCK_START_ROW = VINCULADO_FIRST_BLOCK_START_ROW + VINCULADO_BLOCK_HEIGHT
DESARROLLO_ACTIVIDAD_CELL = "A15"      # shared across all vinculados
GROUP_EXPORT_TITLE_CELL = "F1"
SECTION_2_LAST_COLUMN = "Q"

# Base row positions for 1 vinculado (shift by (N-1)*BLOCK_HEIGHT for N vinculados)
SECTION_6_BASE_AJUSTES_ROW = 70        # ajustes text row
SECTION_7_BASE_START_ROW = 76          # first asistente data row
SECTION_7_NOMBRE_COL = "C"
SECTION_7_CARGO_COL = "K"
SECTION_6_TITLE_ROW_BY_TEMPLATE = {
    TEMPLATE_VARIANT_INDIVIDUAL: 69,
    TEMPLATE_VARIANT_GROUP_2_PLUS: 69,
}
SECTION_7_TITLE_ROW_BY_TEMPLATE = {
    TEMPLATE_VARIANT_INDIVIDUAL: 75,
    TEMPLATE_VARIANT_GROUP_2_PLUS: 75,
}


def ws_write(ws, cell, value):
    try:
        ws[cell] = value
    except Exception:
        return

FORM_CACHE = {}
SECTION_1_CACHE = {}

DISCAPACIDAD_OPTIONS = [
    "Discapacidad visual pérdida total de la visión",
    "Discapacidad visual baja visión",
    "Discapacidad auditiva",
    "Discapacidad auditiva hipoacusia",
    "Trastorno de espectro autista",
    "Discapacidad intelectual",
    "Discapacidad física",
    "Discapacidad física usuario en silla de ruedas",
    "Discapacidad psicosocial",
    "Discapacidad múltiple",
    "No aplica",
]

GENERO_OPTIONS = ["Binario", "No binario", "Otro"]

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
    "Afocolombiano",
    "Afrodescendiente",
    "Rom o Gitano",
    "Indígena",
    "Palenquero de San Basilio",
    "Otro",
    "No aplica",
    "Mulato",
    "Autorreconocimiento",
    "Pueblo Indígena",
    "Negro",
    "Raizal del Archipiélago de San Andrés Y Providencia",
]
CERTIFICADO_DISCAPACIDAD_OPTIONS = ["Si", "No", "No aplica"]
TIPO_CONTRATO_FIRMADO_OPTIONS = [
    "Contrato por obra o labor",
    "Contrato de trabajo a término fijo",
    "Contrato de trabajo a término indefinido",
    "Contrato de aprendizaje",
    "Contrato temporal",
    "Contrato a término indefinido con orden clausulada",
    "Contrato a término fijo a un año",
    "Contrato a término fijo a seis meses",
    "Contrato por prestación de servicios",
]
TIPO_CONTRATO_OPTIONS = TIPO_CONTRATO_FIRMADO_OPTIONS
CONTRATO_TIPO_CONTRATO_OPTIONS = [
    "Contrato a término indefinido.",
    "Contrato a término fijo.",
    "Contrato por obra o labor.",
]
NIVEL_APOYO_OPTIONS = [
    "0. No requiere apoyo.",
    "1. Nivel de apoyo Bajo.",
    "2. Nivel de apoyo medio.",
    "3. Nivel de apoyo alto.",
    "No aplica.",
]
OBS_LECTURA_CONTRATO_OPTIONS = [
    "1. Se acompaña en la lectura del contrato.",
    "2. Se apoya en la lectura del contrato.",
    "3. Cuando requiere un apoyo adicional al del gestor (lector de pantalla, intérprete LSC u otro).",
    "No aplica.",
    "0. No requiere apoyo.",
]
OBS_LECTURA_CONTRATO_OPTIONS_GROUP = [
    "1. Se acompaña en la lectura del contrato.",
    "2. Se apoya en la lectura del contrato.",
    "3. Cuando requiere un apoyo adicional al del gestor (lector de pantalla, intérprete LSC u otro).",
    "No aplica.",
]
OBS_LECTURA_CONTRATO_OPTIONS_INDIVIDUAL = [
    "1. Se acompaña en la lectura del contrato.",
    "2. Se apoya en la lectura del contrato.",
    "3. Cuando requiere un apoyo adicional al del gestor (lector de pantalla, intérprete LSC u otro).",
    "No aplica.",
    "0. No requiere apoyo.",
]
OBS_COMPRENDE_CONTRATO_OPTIONS = [
    "1. Comprende la información, pero no se familiariza con las características del contrato.",
    "2. Explicación de algunas cláusulas del contrato.",
    "3. Explicación total del contrato.",
    "0. Comprende con claridad el contrato.",
]
OBS_TIPO_CONTRATO_OPTIONS = [
    "1. El vinculado reconoce el tipo de contrato, pero no comprende sus condiciones.",
    "2. El vinculado requiere aclaración de algunas de las condiciones del contrato.",
    "3. El vinculado no conoce ninguna de las condiciones del tipo de contrato a firmar.",
    "0. El vinculado tiene claras las condiciones del tipo de contrato a firmar.",
]
JORNADA_LABORAL_OPTIONS = ["Tiempo Completo.", "Medio Tiempo.", "Por horas."]
CLAUSULAS_CONTRATO_OPTIONS = [
    "Cláusula de confidencialidad.",
    "Cláusulas adicionales.",
]
OBS_CONDICIONES_SALARIALES_OPTIONS = [
    "1. Se aclaran las condiciones salariales asignadas al cargo.",
    "2. Se explica de manera parcial las condiciones salariales asignadas al cargo.",
    "3. Se explica de manera completa las condiciones salariales asignadas al cargo.",
    "0. Tiene claras las condiciones salariales asignadas al cargo.",
]
FRECUENCIA_PAGO_OPTIONS = ["Pago Semanal.", "Pago Quincenal.", "Pago Mensual."]
FORMA_PAGO_OPTIONS = ["Abono a cuenta bancaria.", "Nequi o Daviplata.", "Efectivo.", "Cheque."]
OBS_PRESTACIONES_OPTIONS = [
    "1. Conoce, pero es la primera vez que tiene estos beneficios.",
    "2. Requiere más información.",
    "3. Desconoce.",
    "0. Conoce los beneficios y la aplicación.",
    "No aplica.",
]
OBS_CONDUCTO_REGULAR_OPTIONS = [
    "1. Conoce el conducto por experiencias anteriores.",
    "2. Requiere más información.",
    "3. Desconoce la información.",
    "0. Conoce el conducto regular.",
]
OBS_DESCARGOS_OPTIONS = [
    "Si conoce que es una diligencia de descargos.",
    "NO conoce que es una diligencia de descargos.",
]
OBS_TRAMITES_OPTIONS = [
    "Conoce cómo es el proceso para realizar trámites administrativos (certificaciones, afiliaciones, descuentos, desprendibles de nómina).",
    "NO Conoce cómo es el proceso para realizar trámites administrativos (certificaciones, afiliaciones, descuentos, desprendibles de nómina).",
]
OBS_PERMISOS_OPTIONS = [
    "Conoce cómo es el proceso de solicitud de permisos.",
    "NO Conoce cómo es el proceso de solicitud de permisos.",
]
OBS_CAUSALES_OPTIONS = [
    "1. Tiene claro las causales de cancelación del contrato por experiencias anteriores.",
    "2. Requiere aclaración de algunas causales de cancelación del contrato.",
    "3. Desconoce las causales de cancelación del contrato.",
    "0. Tiene claro las causales de cancelación del contrato.",
]
OBS_RUTAS_OPTIONS = [
    "0. Tiene claro cuales son las rutas de atención.",
    "1. Requiere aclaración de cuales son las rutas de atención.",
    "2. Conoce las rutas de atención, pero no las usa",
    "3. Desconoce las rutas de atención.",
    "4. No aplica",
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
        "start_row": 71,
        "rows": 4,
        "nombre_col": "C",
        "cargo_col": "K",
    },
}

LIST_FIELD_OPTIONS_BY_ID = {
    "modalidad": ["Presencial", "Virtual", "Mixta", "No aplica"],
    "discapacidad": list(DISCAPACIDAD_OPTIONS),
    "lgtbiq": list(LGTBIQ_OPTIONS),
    "grupo_etnico": list(GRUPO_ETNICO_OPTIONS),
    "grupo_etnico_cual": list(GRUPO_ETNICO_CUAL_OPTIONS),
    "certificado_discapacidad": list(CERTIFICADO_DISCAPACIDAD_OPTIONS),
    "tipo_contrato": list(TIPO_CONTRATO_FIRMADO_OPTIONS),
    "contrato_lee_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "contrato_lee_observacion": list(OBS_LECTURA_CONTRATO_OPTIONS_INDIVIDUAL),
    "contrato_comprendido_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "contrato_comprendido_observacion": list(OBS_COMPRENDE_CONTRATO_OPTIONS),
    "contrato_tipo_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "contrato_tipo_observacion": list(OBS_TIPO_CONTRATO_OPTIONS),
    "contrato_tipo_contrato": list(CONTRATO_TIPO_CONTRATO_OPTIONS),
    "contrato_jornada": list(JORNADA_LABORAL_OPTIONS),
    "contrato_clausulas": list(CLAUSULAS_CONTRATO_OPTIONS),
    "condiciones_salariales_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "condiciones_salariales_observacion": list(OBS_CONDICIONES_SALARIALES_OPTIONS),
    "condiciones_salariales_frecuencia_pago": list(FRECUENCIA_PAGO_OPTIONS),
    "condiciones_salariales_forma_pago": list(FORMA_PAGO_OPTIONS),
    "prestaciones_cesantias_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "prestaciones_cesantias_observacion": list(OBS_PRESTACIONES_OPTIONS),
    "prestaciones_auxilio_transporte_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "prestaciones_auxilio_transporte_observacion": list(OBS_PRESTACIONES_OPTIONS),
    "prestaciones_prima_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "prestaciones_prima_observacion": list(OBS_PRESTACIONES_OPTIONS),
    "prestaciones_seguridad_social_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "prestaciones_seguridad_social_observacion": list(OBS_PRESTACIONES_OPTIONS),
    "prestaciones_vacaciones_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "prestaciones_vacaciones_observacion": list(OBS_PRESTACIONES_OPTIONS),
    "prestaciones_auxilios_beneficios_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "prestaciones_auxilios_beneficios_observacion": list(OBS_PRESTACIONES_OPTIONS),
    "conducto_regular_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "conducto_regular_observacion": list(OBS_CONDUCTO_REGULAR_OPTIONS),
    "descargos_observacion": list(OBS_DESCARGOS_OPTIONS),
    "tramites_observacion": list(OBS_TRAMITES_OPTIONS),
    "permisos_observacion": list(OBS_PERMISOS_OPTIONS),
    "causales_fin_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "causales_fin_observacion": list(OBS_CAUSALES_OPTIONS),
    "rutas_atencion_nivel_apoyo": list(NIVEL_APOYO_OPTIONS),
    "rutas_atencion_observacion": list(OBS_RUTAS_OPTIONS),
}

EXCEL_DROPDOWN_MANUAL_CANONICAL_OPTIONS = {
    "discapacidad": list(DISCAPACIDAD_OPTIONS),
}

EXCEL_DROPDOWN_EXPLICIT_ALIASES = {
    "grupo_etnico_cual": {
        "gitano (rom)": "Rom o Gitano",
        "gitano rom": "Rom o Gitano",
        "rom": "Rom o Gitano",
        "afrocolombiano": "Afocolombiano",
        "afro colombiano": "Afocolombiano",
    },
    "tipo_contrato": {
        "prestacion de servicios": "Contrato por prestación de servicios",
        "contrato de prestacion de servicios": "Contrato por prestación de servicios",
        "contrato prestacion de servicios": "Contrato por prestación de servicios",
        "termino fijo": "Contrato de trabajo a término fijo",
        "contrato a termino fijo": "Contrato de trabajo a término fijo",
        "termino indefinido": "Contrato de trabajo a término indefinido",
        "contrato a termino indefinido": "Contrato de trabajo a término indefinido",
        "obra o labor": "Contrato por obra o labor",
        "contrato obra o labor": "Contrato por obra o labor",
        "aprendizaje": "Contrato de aprendizaje",
        "contrato de aprendizaje": "Contrato de aprendizaje",
        "temporal": "Contrato temporal",
        "contrato temporal": "Contrato temporal",
    },
}

DATE_FIELD_IDS = {
    "fecha_nacimiento",
    "fecha_firma_contrato",
    "fecha_fin",
}

VINCULADO_CELL_MAP = {
    # Row 23 — personal info line 1
    "numero": ("A", 20),
    "nombre_oferente": ("C", 20),
    "cedula": ("H", 20),
    "certificado_porcentaje": ("K", 20),
    "discapacidad": ("L", 20),
    "telefono_oferente": ("O", 20),
    # Row 24 — personal info line 2
    "genero": ("C", 21),
    "correo_oferente": ("G", 21),
    "fecha_nacimiento": ("M", 21),
    "edad": ("Q", 21),
    # Row 25 — identity
    "lgtbiq": ("E", 22),
    "grupo_etnico": ("L", 22),
    "grupo_etnico_cual": ("O", 22),
    # Row 26 — cargo / emergency
    "cargo_oferente": ("C", 23),
    "contacto_emergencia": ("I", 23),
    "parentesco": ("M", 23),
    "telefono_emergencia": ("Q", 23),
    # Row 27 — certificado / contrato firma
    "certificado_discapacidad": ("F", 24),
    "lugar_firma_contrato": ("L", 24),
    "fecha_firma_contrato": ("Q", 24),
    # Row 29 — datos adicionales
    "tipo_contrato": ("G", 26),
    "fecha_fin": ("N", 26),
    # Section 5.1 Condiciones de la vacante (rows 33-45)
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
    # Section 5.2 Prestaciones de ley (rows 48-59)
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
    # Section 5.3 Deberes y derechos (rows 62-70)
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

SECTION_1_CELL_MAP = EXCEL_MAPPING["section_1"]


def register_form():
    return {
        "id": FORM_ID,
        "name": FORM_NAME,
        "module": __name__,
        "hub_description": "Formaliza la contratación, desarrollo de la actividad y datos del vinculado.",
        "singleton_window": True,
    }


def _get_cache_dir():
    return _get_local_app_cache_dir()


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


def _normalize_section_2_payload(payload):
    if not isinstance(payload, list):
        return payload
    normalized = []
    shared_desarrollo = ""
    for entry in payload:
        current = dict(entry or {})
        normalized.append(current)
        if not shared_desarrollo:
            shared_desarrollo = (current.get("desarrollo_actividad") or "").strip()
    for entry in normalized:
        entry["desarrollo_actividad"] = shared_desarrollo
    return normalized


def load_cache_from_file():
    path = _get_cache_path()
    if not os.path.exists(path):
        return False
    with open(path, "r", encoding="utf-8") as handle:
        payload = json.load(handle) or {}
    data = payload.get("data") or {}
    FORM_CACHE.clear()
    FORM_CACHE.update(data)
    section_2 = FORM_CACHE.get("section_2")
    if isinstance(section_2, list):
        FORM_CACHE["section_2"] = _normalize_section_2_payload(section_2)
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


def _get_section_2_entries(payload=None):
    if payload is None:
        payload = FORM_CACHE.get("section_2", [])
    if not isinstance(payload, list):
        return []
    return [dict(entry or {}) for entry in payload]


def _group_export_title_for_vinculados(total_vinculados):
    total = max(0, int(total_vinculados or 0))
    if total <= 1:
        return "PROCESO DE CONTRATACION INCLUYENTE INDIVIDUAL"
    if total <= 4:
        return "PROCESO CONTRATACION INCLUYENTE GRUPAL - 2 A 4 VINCULADOS"
    if total <= 7:
        return "PROCESO CONTRATACION INCLUYENTE GRUPAL - 5 A 7 VINCULADOS"
    if total <= 10:
        return "PROCESO CONTRATACION INCLUYENTE GRUPAL - 8 A 10 VINCULADOS"
    return "PROCESO CONTRATACION INCLUYENTE GRUPAL - MAS DE 10 VINCULADOS"


def _section_2_group_block_start_row(entry_index):
    return VINCULADO_FIRST_BLOCK_START_ROW + (VINCULADO_BLOCK_HEIGHT * int(entry_index or 0))


def _section_2_group_insert_row(entry_index):
    if int(entry_index or 0) <= 0:
        raise ValueError("entry_index debe ser mayor que 0 para bloques adicionales.")
    return VINCULADO_SECOND_BLOCK_START_ROW + (VINCULADO_BLOCK_HEIGHT * (int(entry_index) - 1))


def get_section_2_field_options(field_id):
    if field_id == "contrato_lee_observacion":
        return list(OBS_LECTURA_CONTRATO_OPTIONS_GROUP)
    return list(LIST_FIELD_OPTIONS_BY_ID.get(field_id, []))


def _normalize_dropdown_text(value):
    normalized = _normalize_text(value or "")
    normalized = normalized.replace('"', "").replace("'", "")
    normalized = " ".join(normalized.split())
    return normalized.strip(" .")

@lru_cache(maxsize=None)
def _get_excel_canonical_options(field_id):
    manual = EXCEL_DROPDOWN_MANUAL_CANONICAL_OPTIONS.get(field_id)
    if manual:
        return tuple(manual)
    if field_id in EXCEL_MAPPING["section_1"]:
        expected_options = LIST_FIELD_OPTIONS_BY_ID.get(field_id, [])
    else:
        expected_options = get_section_2_field_options(field_id)
    return tuple(expected_options)


def normalize_excel_dropdown_value(field_id, raw_value):
    if raw_value in (None, ""):
        return raw_value
    current = str(raw_value).strip()
    if field_id in EXCEL_MAPPING["section_1"]:
        field_options = LIST_FIELD_OPTIONS_BY_ID.get(field_id)
    else:
        field_options = get_section_2_field_options(field_id)
    if not field_options:
        return raw_value

    canonical_options = list(_get_excel_canonical_options(field_id) or field_options)
    current_norm = _normalize_dropdown_text(current)

    explicit_alias = (
        EXCEL_DROPDOWN_EXPLICIT_ALIASES.get(field_id, {}).get(current_norm)
    )
    if explicit_alias:
        return explicit_alias

    for option in canonical_options:
        if _normalize_dropdown_text(option) == current_norm:
            return option

    for idx, option in enumerate(field_options):
        if _normalize_dropdown_text(option) == current_norm:
            if idx < len(canonical_options):
                return canonical_options[idx]
            return option

    _log_excel(
        f"WARN export_dropdown_unmatched field={field_id} "
        f"value={current!r}"
    )
    return raw_value


def _coerce_excel_date_value(value):
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value
    raw = str(value).strip()
    if not raw:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return raw


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
            "empresa_nit",
            "empresa_nombre",
        ]
    )
    params = {
        "select": select_cols,
        "cedula_usuario": f"eq.{normalized}",
        "limit": 1,
    }
    data = _supabase_get("usuarios_reca", params, env_path=env_path)
    return data[0] if data else None


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


def get_empresa_by_nit(nit, env_path=".env"):
    return evaluacion_accesibilidad.get_empresa_by_nit(nit, env_path=env_path)


def get_empresa_by_nombre(nombre, env_path=".env"):
    return evaluacion_accesibilidad.get_empresa_by_nombre(nombre, env_path=env_path)


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=50):
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
    payload = _normalize_section_2_payload(payload)
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

    empresa_nit = (SECTION_1_CACHE.get("nit_empresa") or "").strip()
    empresa_nombre = (SECTION_1_CACHE.get("nombre_empresa") or "").strip()
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
            "certificado_porcentaje": _normalize_decimal_value(
                entry.get("certificado_porcentaje"),
                decimal_separator=".",
            ),
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
            # Siempre dejamos la última empresa de contratación asociada a la cédula.
            "empresa_nit": empresa_nit,
            "empresa_nombre": empresa_nombre,
        }
        normalized_row = {
            key: (None if value == "" else value)
            for key, value in row.items()
        }
        rows.append(normalized_row)
    deduped_rows = {}
    duplicate_cedulas = []
    for row in rows:
        cedula = row.get("cedula_usuario")
        if not cedula:
            continue
        if cedula in deduped_rows:
            duplicate_cedulas.append(cedula)
            deduped_rows.pop(cedula, None)
        deduped_rows[cedula] = row
    rows = list(deduped_rows.values())

    if duplicate_cedulas:
        preview_duplicates = ", ".join(duplicate_cedulas[:10])
        extra_duplicates = "" if len(duplicate_cedulas) <= 10 else f" (+{len(duplicate_cedulas) - 10} mas)"
        _log_excel(
            f"WARN supabase_usuarios_reca_duplicate_cedulas count={len(duplicate_cedulas)} "
            f"cedulas={preview_duplicates}{extra_duplicates}"
        )

    if rows:
        sync_result = _supabase_upsert_with_queue(
            "usuarios_reca",
            rows,
            env_path=env_path,
            on_conflict="cedula_usuario",
        )
        cedulas = [row.get("cedula_usuario") for row in rows if row.get("cedula_usuario")]
        preview = ", ".join(cedulas[:10])
        extra = "" if len(cedulas) <= 10 else f" (+{len(cedulas) - 10} mas)"
        status = sync_result.get("status") or "synced"
        _log_excel(
            f"SUPABASE usuarios_reca upsert status={status} count={len(rows)} cedulas={preview}{extra}"
        )
    return len(rows)


def _build_section_2_entry_writes(entry, *, row_offset=0):
    writes = []
    for field_id, (col, row) in VINCULADO_CELL_MAP.items():
        value = entry.get(field_id, "")
        if field_id == "grupo_etnico_cual":
            grupo_etnico = _normalize_text(entry.get("grupo_etnico") or "")
            if grupo_etnico not in {"si", "sí"}:
                value = "No aplica"
        if value == "":
            continue
        target_row = row + row_offset
        if field_id == "certificado_porcentaje":
            value = _coerce_excel_decimal_value(value)
        elif field_id in DATE_FIELD_IDS:
            value = _coerce_excel_date_value(value)
            if isinstance(value, datetime):
                value = value.strftime("%d/%m/%Y")
        else:
            value = normalize_excel_dropdown_value(field_id, value)
        writes.append({"range": f"'{SHEET_NAME}'!{col}{target_row}", "value": value})
        _log_excel(f"WRITE section=section_2 cell={col}{target_row} key={field_id} value={value!r}")
    return writes


def _build_section_2_writes(vinculados):
    if not vinculados:
        return []
    writes = [
        {
            "range": f"'{SHEET_NAME}'!{GROUP_EXPORT_TITLE_CELL}",
            "value": _group_export_title_for_vinculados(len(vinculados)),
        }
    ]
    shared_desarrollo = ""
    for entry in vinculados:
        shared_desarrollo = (entry.get("desarrollo_actividad") or "").strip()
        if shared_desarrollo:
            break
    if shared_desarrollo:
        writes.append({"range": f"'{SHEET_NAME}'!{DESARROLLO_ACTIVIDAD_CELL}", "value": shared_desarrollo})
        _log_excel(f"WRITE section=section_2 cell={DESARROLLO_ACTIVIDAD_CELL} key=desarrollo_actividad value={shared_desarrollo!r}")
    _log_excel(f"SECTION section=section_2 total={len(vinculados)}")
    for idx, entry in enumerate(vinculados):
        row_offset = VINCULADO_BLOCK_HEIGHT * idx
        writes.extend(_build_section_2_entry_writes(entry, row_offset=row_offset))
    return writes


def _build_section_2_row_insertions(vinculados):
    total_vinculados = len(vinculados or [])
    if total_vinculados <= 1:
        return []
    return [
        {
            "sheet_name": SHEET_NAME,
            "insert_at_row": _section_2_group_insert_row(1),
            "template_start_row": VINCULADO_FIRST_BLOCK_START_ROW,
            "template_end_row": VINCULADO_FIRST_BLOCK_START_ROW + VINCULADO_BLOCK_HEIGHT - 1,
            "repeat_count": total_vinculados - 1,
            # copyPaste PASTE_NORMAL does not copy row heights; restore them from
            # the template block so structure rows in new blocks are not inflated
            # by the tall "nota" row they inherit from inheritFromBefore.
            "copy_row_heights": True,
        }
    ]


def _build_section_6_writes(payload, num_vinculados=1):
    if not payload:
        return []
    shift = max(0, num_vinculados - 1) * VINCULADO_BLOCK_HEIGHT
    ajustes_row = SECTION_6_BASE_AJUSTES_ROW + shift
    ajustes_value = payload.get("ajustes_recomendaciones", "")
    writes = []
    if ajustes_value:
        writes.append({"range": f"'{SHEET_NAME}'!A{ajustes_row}", "value": ajustes_value})
        _log_excel(f"WRITE section=section_6 cell=A{ajustes_row} key=ajustes_recomendaciones")
    return writes


def _build_section_7_writes(payload, num_vinculados=1):
    if not payload:
        return []
    shift = max(0, num_vinculados - 1) * VINCULADO_BLOCK_HEIGHT
    start_row = SECTION_7_BASE_START_ROW + shift
    writes = []
    for idx, entry in enumerate(payload):
        row = start_row + idx
        nombre = (entry.get("nombre") or "").strip()
        cargo = (entry.get("cargo") or "").strip()
        if nombre:
            writes.append({"range": f"'{SHEET_NAME}'!{SECTION_7_NOMBRE_COL}{row}", "value": nombre})
        if cargo:
            writes.append({"range": f"'{SHEET_NAME}'!{SECTION_7_CARGO_COL}{row}", "value": cargo})
    return writes


def _write_section_6(ws, payload, template_variant=TEMPLATE_VARIANT_INDIVIDUAL, total_vinculados=0):
    for write in _build_section_6_writes(payload, num_vinculados=max(1, int(total_vinculados or 1))):
        cell = str(write.get("range") or "").rsplit("!", 1)[-1].replace("'", "")
        ws_write(ws, cell, write.get("value", ""))


def _write_section_7(ws, payload, template_variant=TEMPLATE_VARIANT_INDIVIDUAL, total_vinculados=0):
    for write in _build_section_7_writes(payload, num_vinculados=max(1, int(total_vinculados or 1))):
        cell = str(write.get("range") or "").rsplit("!", 1)[-1].replace("'", "")
        ws_write(ws, cell, write.get("value", ""))


def _build_section_7_row_insertions(payload, num_vinculados=1):
    if not payload:
        return []
    mapping = EXCEL_MAPPING.get("section_7", {})
    base_rows = int(mapping.get("rows", 3) or 3)
    total_rows = len(payload)
    if total_rows <= base_rows:
        return []
    shift = max(0, num_vinculados - 1) * VINCULADO_BLOCK_HEIGHT
    start_row = SECTION_7_BASE_START_ROW + shift
    return [
        {
            "sheet_name": SHEET_NAME,
            "start_row": start_row,
            "base_rows": base_rows,
            "total_rows": total_rows,
        }
    ]


def _build_auto_resize_excluded_rows():
    # Only the three fixed rows of the first block are excluded from auto-resize:
    #   row 17  — structure / header row inside block 1 (no content writes)
    #   row 66  — rutas_atencion_nivel_apoyo / observacion (last data pair)
    #   row 67  — rutas_atencion_nota (last "nota" cell, intentionally fixed height)
    # New vinculado blocks receive full auto-resize on all their rows.
    return {SHEET_NAME: [17, 66, 67]}


def _build_section_1_writes(payload):
    if not payload:
        payload = SECTION_1_CACHE
    if not payload:
        try:
            if load_cache_from_file():
                payload = FORM_CACHE.get("section_1", {}) or SECTION_1_CACHE
        except Exception:
            payload = payload or {}
    writes = []
    for key, cell in SECTION_1_CELL_MAP.items():
        if key in payload:
            value = normalize_excel_dropdown_value(key, payload[key])
            writes.append({"range": f"'{SHEET_NAME}'!{cell}", "value": value})
            _log_excel(f"WRITE section=section_1 cell={cell} key={key} value={value!r}")
    return writes


def _validate_section_2_rows(issues, rows):
    row_pairs = [
        (field_id, humanize_field_id(field_id))
        for field_id in VINCULADO_CELL_MAP.keys()
        if field_id != "numero"
    ]
    row_list = rows if isinstance(rows, list) else []
    meaningful_rows = 0
    shared_desarrollo_present = False
    for row_index, row in enumerate(row_list, start=1):
        row_payload = row if isinstance(row, dict) else {}
        filled = [
            field_id
            for field_id, _label in row_pairs
            if is_meaningful(row_payload.get(field_id))
        ]
        if not filled:
            continue
        meaningful_rows += 1
        if is_meaningful(row_payload.get("desarrollo_actividad")):
            shared_desarrollo_present = True
        for field_id, label in row_pairs:
            require_value(issues, "section_2", row_payload, field_id, label, row_index=row_index)
    if meaningful_rows:
        if not shared_desarrollo_present:
            append_missing_issue(
                issues,
                "section_2",
                "desarrollo_actividad",
                "Desarrollo de la actividad",
            )
        return
    append_missing_issue(
        issues,
        "section_2",
        "",
        "Vinculados",
        message="Debes diligenciar al menos un vinculado.",
    )


def validate_before_finalize(cache=None):
    cache_data = FORM_CACHE if cache is None else (cache or {})
    issues = []

    section_1 = cache_data.get("section_1", {})
    for field_id, label in field_pairs(SECTION_1.get("fields")):
        require_value(issues, "section_1", section_1, field_id, label)

    _validate_section_2_rows(issues, cache_data.get("section_2", []))

    require_value(
        issues,
        "section_6",
        cache_data.get("section_6", {}),
        "ajustes_recomendaciones",
        "Ajustes razonables / recomendaciones",
    )

    validate_dynamic_rows(
        issues,
        "section_7",
        cache_data.get("section_7", []),
        [("nombre", "Nombre"), ("cargo", "Cargo")],
        min_rows_label="Asistentes",
    )
    return issues


def export_to_excel(clear_cache=True):
    if not FORM_CACHE.get("section_1") and cache_file_exists():
        load_cache_from_file()
    raise_validation_error(validate_before_finalize())

    from google_sheets_client import get_master_template_id
    from drive_upload import publish_sheet_from_template

    vinculados = _get_section_2_entries(FORM_CACHE.get("section_2", []))
    num_vinculados = len(vinculados)

    _log_excel(f"START export_all (Google Sheets) vinculados={num_vinculados}")

    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    base_name = _sanitize_filename(empresa_nombre)

    writes = []
    writes.extend(_build_section_1_writes(FORM_CACHE.get("section_1", {})))
    writes.extend(_build_section_2_writes(vinculados))
    writes.extend(_build_section_6_writes(FORM_CACHE.get("section_6", {}), num_vinculados=num_vinculados))
    writes.extend(_build_section_7_writes(FORM_CACHE.get("section_7", []), num_vinculados=num_vinculados))
    row_insertions = []
    row_insertions.extend(_build_section_2_row_insertions(vinculados))
    row_insertions.extend(
        _build_section_7_row_insertions(
            FORM_CACHE.get("section_7", []),
            num_vinculados=num_vinculados,
        )
    )

    # Extract checkbox cells (marked with _checkbox flag)
    checkbox_cells = [w for w in writes if w.get("_checkbox")]
    writes = [{k: v for k, v in w.items() if k != "_checkbox"} for w in writes]

    result = publish_sheet_from_template(
        template_id=get_master_template_id(),
        sheet_writes=writes,
        base_name=base_name,
        folder_name=_sanitize_filename(empresa_nombre),
        row_insertions=row_insertions or None,
        checkbox_cells=checkbox_cells or None,
        auto_resize_excluded_rows=_build_auto_resize_excluded_rows(),
    )

    _log_excel("SUCCESS export_all")

    if clear_cache:
        clear_cache_file()
        clear_form_cache()

    return {
        "output_path": result.get("webViewLink", ""),
        "drive_file_id": result.get("file_id", ""),
        "already_in_drive": True,
    }
