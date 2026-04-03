import os
import json
import time
import re

from google_sheets_client import get_template_id, read_sheet_values
from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    build_sheet_updates,
    _get_local_app_cache_dir,
    _normalize_text,
    _sanitize_filename,
)
from formularios.finalize_validation import (
    field_pairs,
    raise_validation_error,
    require_any_true,
    require_value,
    validate_dynamic_rows,
)
from logging_utils import log_excel_event
from version_info import resource_path


FORM_NAME = "Condiciones de Vacante"
SHEET_NAME = "3. REVISIÓN DE LAS CONDICIONES DE LA VACANTE"

FORM_CACHE = {}
SECTION_1_CACHE = {}
_DISABILITY_DICT = None

OFFICIAL_DICTIONARY_SHEET = "caracterizacion"
OFFICIAL_DICTIONARY_RANGE = f"'{OFFICIAL_DICTIONARY_SHEET}'!A52:B73"


def _get_official_dictionary_spreadsheet_id():
    return get_template_id(
        "google_sheets_official_dictionary_spreadsheet_id",
        "GOOGLE_SHEETS_OFFICIAL_DICTIONARY_SPREADSHEET_ID",
    )


SECTION_1 = {
    "title": "1. DATOS GENERALES",
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
            "label": "Persona que atiende la visita en la empresa",
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
        "fecha_visita": "F7",
        "modalidad": "N7",
        "nombre_empresa": "F8",
        "ciudad_empresa": "N8",
        "direccion_empresa": "F9",
        "nit_empresa": "N9",
        "correo_1": "F10",
        "telefono_empresa": "N10",
        "contacto_empresa": "F11",
        "cargo": "N11",
        "caja_compensacion": "F12",
        "sede_empresa": "N12",
        "asesor": "F13",
        "profesional_asignado": "N13",
    },
    "section_2": {
        "nombre_vacante": "I15",
        "numero_vacantes": "I16",
        "nivel_cargo": "I17",
        "genero": "I18",
        "edad": "I19",
        "modalidad_trabajo": "I20",
        "lugar_trabajo": "I21",
        "salario_asignado": "I22",
        "firma_contrato": "I23",
        "aplicacion_pruebas": "I24",
        "tipo_contrato": "I25",
        "beneficios_adicionales": "I26",
        "cargo_flexible_genero": "I27",
        "beneficios_mujeres": "I28",
        "requiere_certificado": "I29",
        "requiere_certificado_observaciones": "M29",
        "competencia_1": "I30",
        "competencia_2": "L30",
        "competencia_3": "I31",
        "competencia_4": "L31",
        "competencia_5": "I32",
        "competencia_6": "L32",
        "competencia_7": "I33",
        "competencia_8": "L33",
    },
    "section_2_1": {
        "nivel_primaria": "G36",
        "nivel_bachiller": "L36",
        "nivel_tecnico_profesional": "R36",
        "nivel_profesional": "G37",
        "nivel_especializacion": "L37",
        "nivel_tecnologo": "R37",
        "especificaciones_formacion": "I39",
        "conocimientos_basicos": "I40",
        "horarios_asignados": "I42",
        "hora_ingreso": "I43",
        "hora_salida": "I44",
        "tiempo_almuerzo": "I45",
        "break_descanso": "I46",
        "dias_laborables": "I47",
        "dias_flexibles": "I48",
        "observaciones": "I49",
        "experiencia_meses": "I50",
        "funciones_tareas": "A53",
        "herramientas_equipos": "A59",
    },
    "section_3": {
        "lectura": "L64",
        "comprension_lectora": "L65",
        "escritura": "L66",
        "comunicacion_verbal": "L67",
        "razonamiento_logico": "L68",
        "conteo_reporte": "L69",
        "clasificacion_objetos": "L70",
        "velocidad_ejecucion": "L71",
        "concentracion": "L72",
        "memoria": "L73",
        "ubicacion_espacial": "L74",
        "atencion": "L75",
        "observaciones_cognitivas": "E76",
        "agarre": "L80",
        "precision": "L81",
        "digitacion": "L82",
        "agilidad_manual": "L83",
        "coordinacion_ojo_mano": "L84",
        "observaciones_motricidad_fina": "E85",
        "esfuerzo_fisico": "L89",
        "equilibrio_corporal": "L90",
        "lanzar_objetos": "L91",
        "observaciones_motricidad_gruesa": "E92",
        "seguimiento_instrucciones": "L96",
        "resolucion_conflictos": "L97",
        "autonomia_tareas": "L98",
        "trabajo_equipo": "L99",
        "adaptabilidad": "L100",
        "flexibilidad": "L101",
        "comunicacion_asertiva": "L102",
        "manejo_tiempo": "L103",
        "liderazgo": "L104",
        "escucha_activa": "L105",
        "proactividad": "L106",
        "observaciones_transversales": "E107",
    },
    "section_4": {
        "sentado_tiempo": "H111",
        "sentado_frecuencia": "L111",
        "semisentado_tiempo": "H112",
        "semisentado_frecuencia": "L112",
        "de_pie_tiempo": "H113",
        "de_pie_frecuencia": "L113",
        "agachado_tiempo": "H114",
        "agachado_frecuencia": "L114",
        "uso_extremidades_superiores_tiempo": "H115",
        "uso_extremidades_superiores_frecuencia": "L115",
    },
    "section_5": {
        "ruido": "M120",
        "iluminacion": "M121",
        "temperaturas_externas": "M122",
        "vibraciones": "M123",
        "presion_atmosferica": "M124",
        "radiaciones": "M125",
        "polvos_organicos_inorganicos": "M126",
        "fibras": "M127",
        "liquidos": "M128",
        "gases_vapores": "M129",
        "humos_metalicos": "M130",
        "humos_no_metalicos": "M131",
        "material_particulado": "M132",
        "electrico": "M133",
        "locativo": "M134",
        "accidentes_transito": "M135",
        "publicos": "M136",
        "mecanico": "M137",
        "gestion_organizacional": "M138",
        "caracteristicas_organizacion": "M139",
        "caracteristicas_grupo_social": "M140",
        "condiciones_tarea": "M141",
        "interfase_persona_tarea": "M142",
        "jornada_trabajo": "M143",
        "postura_trabajo": "M144",
        "puesto_trabajo": "M145",
        "movimientos_repetitivos": "M146",
        "manipulacion_cargas": "M147",
        "herramientas_equipos": "M148",
        "organizacion_trabajo": "M149",
        "observaciones_peligros": "E150",
    },
    "section_6": {
        "start_row": 153,
        "discapacidad_col": "A",
        "descripcion_col": "G",
        "base_rows": 4,
    },
    "section_7": {
        "observaciones_recomendaciones": "A159",
    },
    "section_8": {
        "start_row": 161,
        "name_col": "E",
        "cargo_col": "L",
        "rows": 3,
    },
}

SECTION_7_TITLE_ROW = 158
SECTION_8_TITLE_ROW = 160


def ws_write(ws, cell, value):
    try:
        ws[cell] = value
    except Exception:
        return


def _get_cache_dir():
    return _get_local_app_cache_dir()


def _get_cache_path():
    return os.path.join(_get_cache_dir(), "condiciones_vacante.json")


def cache_file_exists():
    return os.path.exists(_get_cache_path())


def save_cache_to_file():
    payload = {
        "form_id": "condiciones_vacante",
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


def set_section_cache(section_id, payload):
    if not section_id:
        raise ValueError("section_id requerido")
    if payload is None:
        payload = {}
    FORM_CACHE[section_id] = payload


def get_form_cache():
    return dict(FORM_CACHE)


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
    set_section_cache("section_2", payload)
    FORM_CACHE["_last_section"] = "section_2"
    save_cache_to_file()
    return payload


def confirm_section_2_1(payload):
    if payload is None:
        raise ValueError("section_2_1 requerida")
    set_section_cache("section_2_1", payload)
    FORM_CACHE["_last_section"] = "section_2_1"
    save_cache_to_file()
    return payload


def confirm_section_3(payload):
    if payload is None:
        raise ValueError("section_3 requerida")
    set_section_cache("section_3", payload)
    FORM_CACHE["_last_section"] = "section_3"
    save_cache_to_file()
    return payload


def confirm_section_4(payload):
    if payload is None:
        raise ValueError("section_4 requerida")
    set_section_cache("section_4", payload)
    FORM_CACHE["_last_section"] = "section_4"
    save_cache_to_file()
    return payload


def confirm_section_5(payload):
    if payload is None:
        raise ValueError("section_5 requerida")
    set_section_cache("section_5", payload)
    FORM_CACHE["_last_section"] = "section_5"
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


def confirm_section_8(payload):
    if payload is None:
        raise ValueError("section_8 requerida")
    set_section_cache("section_8", payload)
    FORM_CACHE["_last_section"] = "section_8"
    save_cache_to_file()
    return payload


SECTION_2 = {
    "title": "2. CARACTERÍSTICAS DE LA VACANTE",
    "fields": [
        {"id": "nombre_vacante", "label": "Nombre de la vacante", "type": "texto"},
        {"id": "numero_vacantes", "label": "Número de vacantes", "type": "texto"},
        {
            "id": "nivel_cargo",
            "label": "Nivel del cargo",
            "type": "lista",
            "options": ["Administrativo.", "Operativo.", "Servicios."],
        },
        {
            "id": "genero",
            "label": "Género",
            "type": "lista",
            "options": ["Hombre", "Mujer", "Hombre - Mujer", "Otro", "Indiferente"],
        },
        {"id": "edad", "label": "Edad", "type": "texto"},
        {"id": "modalidad_trabajo", "label": "Modalidad de trabajo", "type": "texto"},
        {"id": "lugar_trabajo", "label": "Lugar de trabajo", "type": "texto"},
        {"id": "salario_asignado", "label": "Salario asignado", "type": "texto"},
        {"id": "firma_contrato", "label": "Firma de contrato", "type": "texto"},
        {"id": "aplicacion_pruebas", "label": "Aplicación de pruebas", "type": "texto"},
        {
            "id": "tipo_contrato",
            "label": "Tipo de contrato",
            "type": "lista",
            "options": [
                "Término Fijo.",
                "Término Indefinido.",
                "Obra o Labor.",
                "Prestación de Servicios.",
                "Término Indefinido con Cláusula presuntiva.",
                "Nombramiento.",
                "Contrato de Aprendizaje.",
                "Nombramiento provisional.",
            ],
        },
        {"id": "beneficios_adicionales", "label": "Beneficios adicionales", "type": "texto"},
        {"id": "cargo_flexible_genero", "label": "Cargo flexible según género", "type": "texto"},
        {
            "id": "beneficios_mujeres",
            "label": "La empresa genera beneficios adicionales a mujeres",
            "type": "texto",
        },
        {
            "id": "requiere_certificado",
            "label": "¿Requiere certificado de discapacidad?",
            "type": "lista",
            "options": ["Sí", "No", "En Trámite"],
        },
    ],
    "competencias": {
        "Administrativo.": [
            "Organización.",
            "Trabajo en equipo.",
            "Proactividad.",
            "Flexibilidad.",
            "Comunicación asertiva.",
            "Resiliencia.",
            "Resolución de problemas.",
            "Gestión del tiempo.",
        ],
        "Operativo.": [
            "Responsabilidad.",
            "Trabajo en equipo.",
            "Flexibilidad.",
            "Comunicación asertiva.",
            "Resolución de problemas.",
            "Proactividad.",
            "Liderazgo.",
            "Honestidad e integridad.",
        ],
        "Servicios.": [
            "Servicio al cliente.",
            "Paciencia.",
            "Comunicación efectiva.",
            "Empatía.",
            "Resolución de problemas.",
            "Responsabilidad.",
            "Trabajo en equipo.",
            "Proactividad.",
        ],
    },
}

SECTION_2_1 = {
    "title": "2.1 FORMACIÓN ACADÉMICA",
    "checkboxes": [
        ("nivel_primaria", "Primaria", "G36"),
        ("nivel_bachiller", "Bachiller", "L36"),
        ("nivel_tecnico_profesional", "Técnico Profesional", "R36"),
        ("nivel_profesional", "Profesional", "G37"),
        ("nivel_especializacion", "Especialización", "L37"),
        ("nivel_tecnologo", "Tecnólogo", "R37"),
    ],
    "fields": [
        {
            "id": "especificaciones_formacion",
            "label": "Especificaciones de la formación académica",
            "type": "texto_largo",
        },
        {
            "id": "conocimientos_basicos",
            "label": "Conocimientos básicos / programas",
            "type": "texto_largo",
        },
        {
            "id": "horarios_asignados",
            "label": "Horarios asignados",
            "type": "lista",
            "options": ["Horarios Fijos.", "Horarios Rotativos.", "Flexibilización de horarios"],
        },
        {
            "id": "hora_ingreso",
            "label": "Hora de ingreso",
            "type": "texto",
        },
        {
            "id": "hora_salida",
            "label": "Hora de salida",
            "type": "texto",
        },
        {
            "id": "tiempo_almuerzo",
            "label": "Tiempo de almuerzo",
            "type": "lista",
            "options": [
                "15 minutos.",
                "30 minutos.",
                "45 minutos.",
                "1 hora.",
                "No aplica.",
                "2 horas.",
            ],
        },
        {
            "id": "break_descanso",
            "label": "Break - descanso",
            "type": "lista",
            "options": ["15 minutos", "30 minutos", "45 minutos", "1 hora", "No aplica"],
        },
        {"id": "dias_laborables", "label": "Días laborables", "type": "texto"},
        {
            "id": "dias_flexibles",
            "label": 'Días laborables flexibles "familia e hijo"',
            "type": "texto",
        },
        {"id": "observaciones", "label": "Observaciones", "type": "texto_largo"},
        {
            "id": "experiencia_meses",
            "label": "Experiencia laboral - tiempo en meses",
            "type": "lista",
            "options": [
                "Seis meses.",
                "Un año.",
                "Año y medio.",
                "Dos años y medio",
                "Las prácticas son válidas como experiencia laboral.",
                "Sin experiencia laboral.",
                "Tres Meses",
                "Con o Sin Experiencia",
                "Dos Años",
                "Tres Años",
                "Cuatro Años",
                "Cinco Años",
            ],
        },
        {
            "id": "funciones_tareas",
            "label": "Principales funciones y tareas asignadas al cargo",
            "type": "texto_largo",
        },
        {
            "id": "herramientas_equipos",
            "label": "Herramientas, equipos e implementos a utilizar en el desarrollo de la labor",
            "type": "texto_largo",
        },
    ],
}

SECTION_3 = {
    "title": "3. HABILIDADES Y CAPACIDADES REQUERIDAS PARA EL CARGO",
    "options": ["Alto.", "Medio.", "Bajo.", "No aplica"],
    "categories": [
        {
            "title": "Habilidades cognitivas",
            "items": [
                ("lectura", "Lectura"),
                ("comprension_lectora", "Comprensión lectora"),
                ("escritura", "Escritura"),
                ("comunicacion_verbal", "Comunicación verbal"),
                ("razonamiento_logico", "Razonamiento lógico - matemático"),
                ("conteo_reporte", "Conteo y reporte de cantidad"),
                ("clasificacion_objetos", "Clasificación de objetos"),
                ("velocidad_ejecucion", "Velocidad de ejecución"),
                ("concentracion", "Concentración"),
                ("memoria", "Memoria"),
                ("ubicacion_espacial", "Ubicación espacial"),
                ("atencion", "Atención"),
            ],
            "observaciones_id": "observaciones_cognitivas",
            "observaciones_label": "Observaciones",
        },
        {
            "title": "Habilidades básicas (Motricidad fina)",
            "items": [
                ("agarre", "Agarre"),
                ("precision", "Precisión"),
                ("digitacion", "Digitación"),
                ("agilidad_manual", "Agilidad manual"),
                ("coordinacion_ojo_mano", "Coordinación ojo - mano"),
            ],
            "observaciones_id": "observaciones_motricidad_fina",
            "observaciones_label": "Observaciones",
        },
        {
            "title": "Habilidades básicas (Motricidad gruesa)",
            "items": [
                ("esfuerzo_fisico", "Esfuerzo físico"),
                ("equilibrio_corporal", "Equilibrio corporal"),
                ("lanzar_objetos", "Lanzar objetos"),
            ],
            "observaciones_id": "observaciones_motricidad_gruesa",
            "observaciones_label": "Observaciones",
        },
        {
            "title": "Competencias transversales",
            "items": [
                ("seguimiento_instrucciones", "Seguimiento de instrucciones"),
                ("resolucion_conflictos", "Resolución de conflictos"),
                ("autonomia_tareas", "Autonomía en desarrollo de tareas"),
                ("trabajo_equipo", "Trabajo en equipo"),
                ("adaptabilidad", "Adaptabilidad"),
                ("flexibilidad", "Flexibilidad"),
                ("comunicacion_asertiva", "Comunicación asertiva y efectiva"),
                ("manejo_tiempo", "Manejo del tiempo"),
                ("liderazgo", "Liderazgo"),
                ("escucha_activa", "Escucha activa"),
                ("proactividad", "Proactividad"),
            ],
            "observaciones_id": "observaciones_transversales",
            "observaciones_label": "Observaciones",
        },
    ],
}

SECTION_4 = {
    "title": "4. POSTURAS Y MOVIMIENTOS",
    "time_options": [
        "De 1 a 2 horas.",
        "De 2 a 4 horas.",
        "De 4 a 6 horas.",
        "De 6 a 8 horas.",
        "No aplica",
    ],
    "frequency_options": ["Diario.", "Semanal.", "Quincenal.", "Mensual.", "No aplica."],
    "fields": [
        ("sentado", "Sentado"),
        ("semisentado", "Semisentado"),
        ("de_pie", "De pie recto"),
        ("agachado", "Agachado"),
        ("uso_extremidades_superiores", "Uso extremidades superiores"),
    ],
}

SECTION_5 = {
    "title": "5. PELIGROS Y RIESGOS EN EL DESARROLLO DE LA LABOR",
    "options": ["Alto.", "Medio.", "Bajo.", "No aplica"],
    "categories": [
        {
            "title": "Físico",
            "items": [
                ("ruido", "Ruido"),
                ("iluminacion", "Iluminación"),
                ("temperaturas_externas", "Temperaturas externas"),
                ("vibraciones", "Vibraciones"),
                ("presion_atmosferica", "Presión atmosférica"),
                ("radiaciones", "Radiaciones ionizantes y no ionizantes"),
            ],
        },
        {
            "title": "Químico",
            "items": [
                ("polvos_organicos_inorganicos", "Polvos orgánicos inorgánicos"),
                ("fibras", "Fibras"),
                ("liquidos", "Líquidos"),
                ("gases_vapores", "Gases y vapores"),
                ("humos_metalicos", "Humos metálicos"),
                ("humos_no_metalicos", "Humos no metálicos"),
                ("material_particulado", "Material particulado"),
            ],
        },
        {
            "title": "Condiciones de seguridad",
            "items": [
                ("electrico", "Eléctrico"),
                ("locativo", "Locativo"),
                ("accidentes_transito", "Accidentes de tránsito"),
                ("publicos", "Públicos"),
                ("mecanico", "Mecánico"),
            ],
        },
        {
            "title": "Psicosocial",
            "items": [
                (
                    "gestion_organizacional",
                    "Gestión organizacional",
                    "Gestión organizacional. (Estilos de mando, forma de pago, contratación, participación de la persona dentro de la empresa, inducción y capacitación, bienestar social, evaluación de desempeño y manejo de cargos).",
                ),
                (
                    "caracteristicas_organizacion",
                    "Características de la organización del trabajo",
                    "Características de la organización del trabajo. (Comunicación, tecnología, organización de las cargas laborales).",
                ),
                (
                    "caracteristicas_grupo_social",
                    "Características del grupo social del trabajo",
                    "Características del grupo social del trabajo. (Relaciones laborales, clima laboral).",
                ),
                (
                    "condiciones_tarea",
                    "Condiciones de la tarea",
                    "Condiciones de la tarea. (Demandas emocionales, sistemas de control, definición de roles, monotonía, etc.).",
                ),
                (
                    "interfase_persona_tarea",
                    "Interfase persona tarea",
                    "Interfase persona tarea. (Conocimientos, habilidades con relación a la demanda de la tarea, iniciativa, autonomía y reconocimiento, identificación de la persona con la tarea y la organización).",
                ),
                (
                    "jornada_trabajo",
                    "Jornada de trabajo",
                    "Jornada de trabajo. (Pausas, trabajo nocturno, rotación, horas extras, descansos).",
                ),
            ],
        },
        {
            "title": "Ergonómico",
            "items": [
                (
                    "postura_trabajo",
                    "Postura de trabajo",
                    "Postura de trabajo. (Se mantiene posturas prolongadas (sentado o de pie) durante la jornada laboral, postura adoptada cómoda y natural para el desarrollo de la tarea, posturas forzadas del cuello, espalda o extremidades).",
                ),
                (
                    "puesto_trabajo",
                    "Puesto de trabajo",
                    "Puesto de trabajo. (La silla es ajustable en altura y cuenta con respaldo adecuado, altura de la mesa o superficie de trabajo es adecuada, puesto permite una correcta ubicación de pies y piernas).",
                ),
                (
                    "movimientos_repetitivos",
                    "Movimientos repetitivos",
                    "Movimientos repetitivos. (La tarea requiere movimientos repetitivos de manos o brazos, pausas activas durante la jornada laboral).",
                ),
                (
                    "manipulacion_cargas",
                    "Manipulación de cargas",
                    "Manipulación de cargas. (Se debe levantar, empujar o cargar peso, peso de las cargas es adecuado y manejable).",
                ),
                (
                    "herramientas_equipos",
                    "Herramientas - Equipos",
                    "Herramientas - Equipos. (Las herramientas son adecuadas al tamaño y fuerza del trabajador, las herramientas reducen el esfuerzo físico innecesario).",
                ),
                (
                    "organizacion_trabajo",
                    "Organización del trabajo",
                    "Organización del trabajo. (La jornada laboral permite pausas y descansos adecuados, la carga de trabajo es acorde con las capacidades del trabajador).",
                ),
            ],
        },
    ],
    "observaciones": {
        "id": "observaciones_peligros",
        "label": "Observaciones",
    },
}

SECTION_6 = {
    "title": "6. DISCAPACIDADES Y DESCRIPCIONES",
    "base_rows": 4,
    "options": [
        "DISCAPACIDAD VISUAL BAJA VISIÓN",
        "DISCAPACIDAD VISUAL PÉRDIDA TOTAL DE LA VISIÓN",
        "DISCAPACIDAD AUDITIVA",
        "DISCAPACIDAD AUDITIVA HIPOACUSIA",
        "DISCAPACIDAD INTELECTUAL",
        "TEA / AUTISMO",
        "DISCAPACIDAD FÍSICA USR",
        "DISCAPACIDAD FÍSICA",
        "DISCAPACIDAD PSICOSOCIAL",
        "DISCAPACIDAD MÚLTIPLE FÍSICA - VISUAL",
        "DISCAPACIDAD MÚLTIPLE FÍSICA - AUDITIVA",
        "DISCAPACIDAD MÚLTIPLE FÍSICA - PSICOSOCIAL",
        "DISCAPACIDAD MÚLTIPLE FÍSICA - INTELECTUAL",
        "DISCAPACIDAD MÚLTIPLE FÍSICA - BAJA VISIÓN",
        "DISCAPACIDAD MÚLTIPLE FÍSICA - HIPOACUSIA",
        "DISCAPACIDAD MÚLTIPLE PSICOSOCIAL - HIPOACUSIA",
        "DISCAPACIDAD MÚLTIPLE PSICOSOCIAL - AUDITIVA",
        "DISCAPACIDAD MÚLTIPLE PSICOSOCIAL - BAJA VISIÓN",
        "DISCAPACIDAD MÚLTIPLE PSICOSOCIAL - VISUAL",
        "DISCAPACIDAD MÚLTIPLE PSICOSOCIAL– INTELECTUAL",
        "DISCAPACIDAD MÚLTIPLE AUDITIVA - INTELECTUAL",
        "DISCAPACIDAD MÚLTIPLE VISUAL- INTELECTUAL",
    ],
}

SECTION_7 = {
    "title": "7. OBSERVACIONES / RECOMENDACIONES",
    "field_id": "observaciones_recomendaciones",
}

SECTION_7_TEMPLATES = {
    "proceso_vacante": """
* Ejecutar el proceso de retroalimentación a los candidatos sobre quién continúa o no en el proceso.

* Acompañamiento desde RECA durante el proceso.

* La empresa debe dar el visto bueno al perfil levantado junto al asesor de la Agencia y RECA, para que desde la Agencia se publique la vacante y se realice el envío de candidatos dentro de los 4 días hábiles.

* Remisión del perfil para el proceso correspondiente.

El presente perfil describe los tipos de discapacidad que, tras el análisis de las funciones del cargo, el entorno de trabajo, los factores de riesgo y las demandas propias del rol, se consideran compatibles para la vinculación laboral de personas con discapacidad, bajo un enfoque de inclusión social y laboral.
""".strip(),
}

SECTION_7_TEMPLATE_BUTTONS = [
    ("proceso_vacante", "Proceso vacante"),
]

SECTION_8 = {
    "title": "8. ASISTENTES",
    "rows": 3,
    "nombres": [
        "Sandra Milena Pachón Rojas",
        "Sara Zambrano",
        "Alejandra Pérez",
        "Lenny Lugo",
        "Angie Díaz",
        "Adriana Viveros",
        "Janeth Camargo",
        "Gabriela Rubiano Isaza",
        "Andrés Montes",
        "Sara Sánchez",
        "Catalina Salazar",
    ],
    "cargos": [
        "Coordinadora de inclusión laboral",
        "Coordinación de inclusión laboral",
        "Gestora de inclusión laboral",
        "Gestor de inclusión laboral",
        "Profesional de apoyo de inclusión laboral",
        "Líder empleo inclusivo",
        "Gestora de proyectos y desarrollo",
        "Profesional de inclusión laboral",
        "Directora Fundación Reca",
    ],
}


def _fix_text(text):
    if not text:
        return ""
    replacements = {
        "Ç?": "Í",
        "Ç­": "á",
        "Ç¸": "é",
        "Çð": "í",
        "Ç§": "ú",
        "Ç±": "ñ",
        "Çü": "ó",
        "Ç%": "É",
        "Çs": "Ú",
        "Æ'?": "",
        "Æ'??": "",
        "ƒÅ'": "",
    }
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    return text


def _normalize_key(text):
    text = _fix_text(text or "")
    text = text.replace("–", "-")
    text = " ".join(text.replace("\t", " ").split())
    return text.upper()


def _looks_like_disability_heading(text):
    normalized = _normalize_key(text)
    if not normalized or normalized.startswith('"') or re.match(r"^\d", normalized):
        return False
    return normalized == normalized.upper() and (
        "DISCAPACIDAD" in normalized or "TEA" in normalized or "AUTISMO" in normalized
    )


def _load_disability_descriptions_from_sheet():
    rows = read_sheet_values(_get_official_dictionary_spreadsheet_id(), OFFICIAL_DICTIONARY_RANGE)
    entries = {}
    for row in rows or []:
        if not row:
            continue
        key = str(row[0] or "").strip()
        description = str(row[1] or "").strip() if len(row) > 1 else ""
        if not key:
            continue
        entries[_normalize_key(key)] = _fix_text(description)
    return entries


def _load_disability_descriptions_from_text():
    path = resource_path("Diccionario.txt")
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            raw = handle.read()
    except UnicodeDecodeError:
        with open(path, "r", encoding="latin-1") as handle:
            raw = handle.read()
    raw = raw.replace("\r\n", "\n")
    entries = {}
    current_key = None
    current_lines = []

    def flush():
        nonlocal current_key, current_lines
        if current_key:
            text = "\n".join(current_lines).strip()
            text = text.strip('"')
            entries[_normalize_key(current_key)] = _fix_text(text)
        current_key = None
        current_lines = []

    for line in raw.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        cleaned = _fix_text(stripped)
        if _looks_like_disability_heading(cleaned):
            flush()
            current_key = cleaned
            current_lines = []
            continue
        if '"' in cleaned and not cleaned.startswith('"'):
            key_part, desc_part = cleaned.split('"', 1)
            flush()
            current_key = key_part.strip()
            current_lines = [desc_part.rstrip('"')]
            continue
        if current_key is None:
            current_key = cleaned
            current_lines = []
            continue
        current_lines.append(cleaned)
    flush()
    return entries


def get_disability_descriptions():
    global _DISABILITY_DICT
    if _DISABILITY_DICT is not None:
        return _DISABILITY_DICT
    try:
        entries = _load_disability_descriptions_from_sheet()
        if entries:
            _DISABILITY_DICT = entries
            return _DISABILITY_DICT
    except Exception:
        pass
    _DISABILITY_DICT = _load_disability_descriptions_from_text()
    return _DISABILITY_DICT


def normalize_disability_key(value):
    return _normalize_key(value)


def _get_section_6_extra_rows(cache_data=None):
    cache = FORM_CACHE if cache_data is None else (cache_data or {})
    section_6_rows = list(cache.get("section_6") or [])
    base_rows = int(EXCEL_MAPPING.get("section_6", {}).get("base_rows", 4) or 4)
    return max(0, len(section_6_rows) - base_rows)


def _section_6_extra_rows(section_6_payload=None):
    if section_6_payload is None:
        return _get_section_6_extra_rows()
    cache_snapshot = {"section_6": list(section_6_payload or [])}
    return _get_section_6_extra_rows(cache_snapshot)


def _section_7_content_row_for_payload(section_6_payload=None):
    return SECTION_7_TITLE_ROW + 1 + _section_6_extra_rows(section_6_payload)


def _section_8_start_row_for_payload(section_6_payload=None):
    return SECTION_8_TITLE_ROW + 1 + _section_6_extra_rows(section_6_payload)


def _shift_cell_reference(cell_ref, row_delta):
    text = str(cell_ref or "").strip()
    if row_delta <= 0 or not text:
        return text
    match = re.match(r"^([A-Z]+)(\d+)$", text, re.IGNORECASE)
    if not match:
        return text
    column = match.group(1).upper()
    row = int(match.group(2))
    return f"{column}{row + row_delta}"


def _build_section_writes(section_id, payload):
    """Return a list of {"range": ..., "value": ...} dicts for *section_id*."""
    if section_id == "section_6":
        mapping = EXCEL_MAPPING["section_6"]
        start_row = mapping["start_row"]
        writes = []
        for idx, entry in enumerate(payload or []):
            row = start_row + idx
            discapacidad = entry.get("discapacidad", "")
            if discapacidad:
                writes.append({
                    "range": f"'{SHEET_NAME}'!{mapping['discapacidad_col']}{row}",
                    "value": discapacidad,
                })
        return writes

    if section_id == "section_7":
        if not payload:
            return []
        row_offset = _get_section_6_extra_rows()
        cell = EXCEL_MAPPING["section_7"].get("observaciones_recomendaciones", "A159")
        cell = _shift_cell_reference(cell, row_offset)
        value = payload.get("observaciones_recomendaciones", "")
        if not value:
            return []
        return [{"range": f"'{SHEET_NAME}'!{cell}", "value": value}]

    if section_id == "section_8":
        if not payload:
            return []
        mapping = EXCEL_MAPPING["section_8"]
        start_row = int(mapping["start_row"]) + _get_section_6_extra_rows()
        name_col = mapping["name_col"]
        cargo_col = mapping["cargo_col"]
        writes = []
        for idx, entry in enumerate(payload):
            row = start_row + idx
            nombre = entry.get("nombre", "")
            cargo = entry.get("cargo", "")
            if nombre:
                writes.append({"range": f"'{SHEET_NAME}'!{name_col}{row}", "value": nombre})
            if cargo:
                writes.append({"range": f"'{SHEET_NAME}'!{cargo_col}{row}", "value": cargo})
        return writes

    mapping = EXCEL_MAPPING.get(section_id)
    if not mapping:
        return []

    if section_id == "section_2_1":
        checkbox_ids = {item[0] for item in SECTION_2_1.get("checkboxes", [])}
        writes = []
        for key, cell in mapping.items():
            if key in payload:
                value = payload.get(key)
                if key in checkbox_ids:
                    value = bool(value)
                if value is None:
                    value = ""
                writes.append({"range": f"'{SHEET_NAME}'!{cell}", "value": value, "_checkbox": key in checkbox_ids})
        return writes

    return build_sheet_updates(SHEET_NAME, mapping, payload or {})


def _build_row_insertions(cache):
    row_insertions = []

    section_6 = list((cache or {}).get("section_6") or [])
    section_6_cfg = EXCEL_MAPPING.get("section_6", {})
    section_6_base_rows = int(section_6_cfg.get("base_rows", 4) or 4)
    if section_6 and len(section_6) > section_6_base_rows:
        row_insertions.append(
            {
                "sheet_name": SHEET_NAME,
                "start_row": int(section_6_cfg["start_row"]),
                "base_rows": section_6_base_rows,
                "total_rows": len(section_6),
            }
        )

    section_8 = list((cache or {}).get("section_8") or [])
    section_8_cfg = EXCEL_MAPPING.get("section_8", {})
    section_8_base_rows = int(section_8_cfg.get("rows", 3) or 3)
    if section_8 and len(section_8) > section_8_base_rows:
        row_insertions.append(
            {
                "sheet_name": SHEET_NAME,
                "start_row": int(section_8_cfg["start_row"]),
                "base_rows": section_8_base_rows,
                "total_rows": len(section_8),
            }
        )

    return row_insertions


def _write_section_with_ws(ws, section_id, payload):
    if section_id == "section_6":
        mapping = EXCEL_MAPPING["section_6"]
        start_row = int(mapping["start_row"])
        base_rows = int(mapping.get("base_rows", 4) or 4)
        rows = list(payload or [])
        extra_rows = max(0, len(rows) - base_rows)
        insert_row = start_row + base_rows
        for _ in range(extra_rows):
            ws.Rows(insert_row).Insert()
            ws.Rows(insert_row - 1).Copy(ws.Rows(insert_row))
        for idx, entry in enumerate(rows):
            discapacidad = str((entry or {}).get("discapacidad") or "").strip()
            if not discapacidad:
                continue
            row = start_row + idx
            _log_excel(f"WRITE section=section_6 cell=A{row} key=discapacidad")
            ws_write(ws, f"A{row}", discapacidad)
        return

    if section_id == "section_7":
        value = str((payload or {}).get("observaciones_recomendaciones") or "").strip()
        if not value:
            return
        write_row = _section_7_content_row_for_payload(FORM_CACHE.get("section_6") or [])
        _log_excel(f"WRITE section=section_7 cell=A{write_row} key=observaciones_recomendaciones")
        ws_write(ws, f"A{write_row}", value)
        return

    if section_id == "section_8":
        if isinstance(payload, dict):
            rows = list(payload.get("asistentes") or [])
        else:
            rows = list(payload or [])
        start_row = _section_8_start_row_for_payload(FORM_CACHE.get("section_6") or [])
        base_rows = int(EXCEL_MAPPING.get("section_8", {}).get("rows", 3) or 3)
        extra_rows = max(0, len(rows) - base_rows)
        insert_row = start_row + base_rows
        for _ in range(extra_rows):
            ws.Rows(insert_row).Insert()
            ws.Rows(insert_row - 1).Copy(ws.Rows(insert_row))
        for idx, entry in enumerate(rows):
            row = start_row + idx
            nombre = str((entry or {}).get("nombre") or "").strip()
            cargo = str((entry or {}).get("cargo") or "").strip()
            if nombre:
                _log_excel(f"WRITE section=section_8 cell=E{row} key=nombre")
                ws_write(ws, f"E{row}", nombre)
            if cargo:
                _log_excel(f"WRITE section=section_8 cell=L{row} key=cargo")
                ws_write(ws, f"L{row}", cargo)
        return

    for write in _build_section_writes(section_id, payload):
        cell = str(write.get("range") or "").rsplit("!", 1)[-1].replace("'", "")
        _log_excel(f"WRITE section={section_id} cell={cell}")
        ws_write(ws, cell, write.get("value", ""))


def _tuple_pairs(items):
    pairs = []
    for item in items or []:
        if not isinstance(item, (tuple, list)) or len(item) < 2:
            continue
        field_id = str(item[0] or "").strip()
        if not field_id:
            continue
        pairs.append((field_id, str(item[1] or "").strip() or field_id))
    return pairs


def validate_before_finalize(cache=None):
    cache_data = FORM_CACHE if cache is None else (cache or {})
    issues = []

    section_1 = cache_data.get("section_1", {})
    for field_id, label in field_pairs(SECTION_1.get("fields")):
        require_value(issues, "section_1", section_1, field_id, label)

    section_2 = cache_data.get("section_2", {})
    for field_id, label in field_pairs(SECTION_2.get("fields")):
        require_value(issues, "section_2", section_2, field_id, label)
    require_value(
        issues,
        "section_2",
        section_2,
        "requiere_certificado_observaciones",
        "Observaciones (Requiere certificado)",
    )

    section_2_1 = cache_data.get("section_2_1", {})
    require_any_true(
        issues,
        "section_2_1",
        section_2_1,
        [field_id for field_id, _label, _cell in SECTION_2_1.get("checkboxes", [])],
        "Niveles educativos",
    )
    for field_id, label in field_pairs(SECTION_2_1.get("fields")):
        require_value(issues, "section_2_1", section_2_1, field_id, label)

    section_3 = cache_data.get("section_3", {})
    for category in SECTION_3.get("categories", []):
        for field_id, label in _tuple_pairs(category.get("items")):
            require_value(issues, "section_3", section_3, field_id, label)
        require_value(
            issues,
            "section_3",
            section_3,
            category.get("observaciones_id"),
            category.get("observaciones_label") or "Observaciones",
        )

    section_4 = cache_data.get("section_4", {})
    for field_id, label in _tuple_pairs(SECTION_4.get("fields")):
        require_value(issues, "section_4", section_4, f"{field_id}_tiempo", f"{label} - Tiempo")
        require_value(
            issues,
            "section_4",
            section_4,
            f"{field_id}_frecuencia",
            f"{label} - Frecuencia",
        )

    section_5 = cache_data.get("section_5", {})
    for category in SECTION_5.get("categories", []):
        for field_id, label in _tuple_pairs(category.get("items")):
            require_value(issues, "section_5", section_5, field_id, label)
    require_value(
        issues,
        "section_5",
        section_5,
        SECTION_5.get("observaciones", {}).get("id"),
        SECTION_5.get("observaciones", {}).get("label") or "Observaciones",
    )

    validate_dynamic_rows(
        issues,
        "section_6",
        cache_data.get("section_6", []),
        [("discapacidad", "Discapacidad"), ("descripcion", "Descripcion")],
        min_rows_label="Discapacidades y descripciones",
    )

    require_value(
        issues,
        "section_7",
        cache_data.get("section_7", {}),
        SECTION_7.get("field_id"),
        "Observaciones / Recomendaciones",
    )

    validate_dynamic_rows(
        issues,
        "section_8",
        cache_data.get("section_8", []),
        [("nombre", "Nombre"), ("cargo", "Cargo")],
        min_rows_label="Asistentes",
    )
    return issues


def export_to_excel(progress_callback=None):
    if not FORM_CACHE.get("section_1") and cache_file_exists():
        load_cache_from_file()
    raise_validation_error(validate_before_finalize())

    from google_sheets_client import get_master_template_id
    from drive_upload import publish_sheet_from_template

    _log_excel("START export_all (Google Sheets)")

    writes = []
    section_order = [
        "section_1",
        "section_2",
        "section_2_1",
        "section_3",
        "section_4",
        "section_5",
        "section_6",
        "section_7",
        "section_8",
    ]
    for section_id in section_order:
        payload = FORM_CACHE.get(section_id, {})
        _log_excel(f"SECTION export_all section={section_id}")
        if progress_callback:
            progress_callback(section_id)
        writes.extend(_build_section_writes(section_id, payload))

    checkbox_cells = [w for w in writes if w.get("_checkbox")]
    writes = [{k: v for k, v in w.items() if k != "_checkbox"} for w in writes]

    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    base_name = _sanitize_filename(empresa_nombre)
    row_insertions = _build_row_insertions(FORM_CACHE)

    result = publish_sheet_from_template(
        template_id=get_master_template_id(),
        sheet_writes=writes,
        base_name=base_name,
        folder_name=_sanitize_filename(empresa_nombre),
        row_insertions=row_insertions or None,
        checkbox_cells=checkbox_cells,
    )

    _log_excel("SUCCESS export_all (Google Sheets)")

    clear_cache_file()
    clear_form_cache()

    return {
        "output_path": result.get("webViewLink", ""),
        "drive_file_id": result.get("file_id", ""),
        "already_in_drive": True,
    }

def register_form():
    return {
        "id": "condiciones_vacante",
        "name": FORM_NAME,
        "module": __name__,
        "hub_description": "Documenta condiciones del cargo, apoyos y requisitos de la vacante.",
        "singleton_window": True,
    }
