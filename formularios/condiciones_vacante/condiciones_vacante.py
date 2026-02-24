import os
import json
import time
import re
import shutil
import unicodedata

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import _get_desktop_dir, _normalize_text, _sanitize_filename


FORM_NAME = "Condiciones de Vacante"
SHEET_NAME = "3. REVISIÓN DE LAS CONDICIONES"

FORM_CACHE = {}
SECTION_1_CACHE = {}
_DISABILITY_DICT = None


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
        "consideraciones_col": "G",
        "descripcion_col": "L",
        "base_rows": 4,
    },
    "section_7": {
        "observaciones_recomendaciones": "A158",
    },
    "section_8": {
        "start_row": 161,
        "name_col": "E",
        "cargo_col": "L",
        "rows": 3,
    },
}


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


def _find_template_path():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    templates_dir = os.path.join(base_dir, "templates")
    if not os.path.isdir(templates_dir):
        raise FileNotFoundError("No existe la carpeta templates.")
    for name in os.listdir(templates_dir):
        normalized = _normalize_text(name).replace("_", "")
        if "revision" in normalized and "condicion" in normalized and normalized.endswith(".xlsx"):
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template de revision de condiciones.")


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
    output_path = FORM_CACHE.get("_output_path")
    if output_path and os.path.exists(output_path):
        return output_path
    template_path = _find_template_path()
    desktop = _get_desktop_dir()
    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    safe_company = _sanitize_filename(empresa_nombre)
    if not safe_company:
        safe_company = "Empresa"
    output_dir = os.path.join(desktop, "Formatos Inclusion Laboral", safe_company)
    os.makedirs(output_dir, exist_ok=True)
    process_name = "Revision de las Condiciones de la Vacante"
    output_name = f"{process_name} - {safe_company}.xlsx"
    output_path = os.path.join(output_dir, output_name)
    if not os.path.exists(output_path):
        shutil.copy2(template_path, output_path)
    FORM_CACHE["_output_path"] = output_path
    return output_path


def get_output_path():
    output_path = FORM_CACHE.get("_output_path")
    if output_path and os.path.exists(output_path):
        return output_path
    return None


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
            if target.startswith("7.") or target.startswith("8."):
                if value_norm.startswith(target):
                    return row
            else:
                return row
    raise ValueError(f"No se encontró el texto '{text}' en la columna A.")


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
            "type": "hora",
        },
        {
            "id": "hora_salida",
            "label": "Hora de salida",
            "type": "hora",
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
        "Æ’?": "",
        "Æ’??": "",
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


def get_disability_descriptions():
    global _DISABILITY_DICT
    if _DISABILITY_DICT is not None:
        return _DISABILITY_DICT
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    path = os.path.join(base_dir, "Diccionario.txt")
    if not os.path.exists(path):
        _DISABILITY_DICT = {}
        return _DISABILITY_DICT
    try:
        raw = open(path, "r", encoding="utf-8").read()
    except UnicodeDecodeError:
        raw = open(path, "r", encoding="latin-1").read()
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
        if '"' in cleaned and not cleaned.startswith('"'):
            key_part, desc_part = cleaned.split('"', 1)
            flush()
            current_key = key_part.strip()
            current_lines = [desc_part.rstrip('"')]
            continue
        if cleaned == cleaned.upper() and cleaned.startswith("DISCAPACIDAD"):
            flush()
            current_key = cleaned
            current_lines = []
            continue
        if current_key is None:
            current_key = cleaned
            current_lines = []
            continue
        current_lines.append(cleaned)
    flush()
    _DISABILITY_DICT = entries
    return _DISABILITY_DICT


def normalize_disability_key(value):
    return _normalize_key(value)


def _write_section_with_ws(ws, section_id, payload):
    if section_id == "section_6":
        mapping = EXCEL_MAPPING["section_6"]
        start_row = mapping["start_row"]
        base_rows = mapping["base_rows"]
        total = len(payload or [])
        if total > base_rows:
            insert_at = start_row + base_rows
            template_row = start_row + base_rows - 1
            for _ in range(total - base_rows):
                ws.Rows(insert_at).Insert()
                ws.Rows(template_row).Copy(ws.Rows(insert_at))
                insert_at += 1
        for idx, entry in enumerate(payload or []):
            row = start_row + idx
            discapacidad = entry.get("discapacidad", "")
            consideraciones = entry.get("consideraciones", "")
            _log_excel(
                f"WRITE section=section_6 cell={mapping['discapacidad_col']}{row} key=discapacidad value={discapacidad!r}"
            )
            _log_excel(
                f"WRITE section=section_6 cell={mapping['consideraciones_col']}{row} key=consideraciones value={consideraciones!r}"
            )
            ws.Range(f"{mapping['discapacidad_col']}{row}").Value = discapacidad
            ws.Range(f"{mapping['consideraciones_col']}{row}").Value = consideraciones
        return

    if section_id == "section_7":
        if not payload:
            return
        row = _find_row_by_text(ws, "7. OBSERVACIONES / RECOMENDACIONES:")
        value = payload.get("observaciones_recomendaciones", "")
        _log_excel(
            f"WRITE section=section_7 cell=A{row + 1} key=observaciones_recomendaciones value={value!r}"
        )
        ws.Range(f"A{row + 1}").Value = value
        return

    if section_id == "section_8":
        if not payload:
            return
        row_title = _find_row_by_text(ws, "8.ASISTENTES")
        start_row = row_title + 1
        base_rows = EXCEL_MAPPING["section_8"].get("rows", 3)
        total = len(payload)
        if total > base_rows:
            insert_at = start_row + base_rows
            template_row = start_row + base_rows - 1
            for _ in range(total - base_rows):
                ws.Rows(insert_at).Insert()
                ws.Rows(template_row).Copy(ws.Rows(insert_at))
                insert_at += 1
        for idx, entry in enumerate(payload):
            row = start_row + idx
            nombre = entry.get("nombre", "")
            cargo = entry.get("cargo", "")
            _log_excel(
                f"WRITE section=section_8 cell=E{row} key=nombre value={nombre!r}"
            )
            _log_excel(
                f"WRITE section=section_8 cell=L{row} key=cargo value={cargo!r}"
            )
            ws.Range(f"E{row}").Value = nombre
            ws.Range(f"L{row}").Value = cargo
        return

    mapping = EXCEL_MAPPING.get(section_id)
    if not mapping:
        return
    if section_id == "section_2_1":
        checkbox_ids = {item[0] for item in SECTION_2_1.get("checkboxes", [])}
        for key, cell in mapping.items():
            if key in payload:
                value = payload.get(key)
                if key in checkbox_ids:
                    _log_excel(
                        f"WRITE section={section_id} cell={cell} key={key} value={bool(value)!r}"
                    )
                    ws.Range(cell).Value = True if value else False
                else:
                    _log_excel(
                        f"WRITE section={section_id} cell={cell} key={key} value={value!r}"
                    )
                    ws.Range(cell).Value = value
        return
    for key, cell in mapping.items():
        if key in payload:
            value = payload.get(key)
            _log_excel(
                f"WRITE section={section_id} cell={cell} key={key} value={value!r}"
            )
            ws.Range(cell).Value = value


def export_to_excel(progress_callback=None):
    output_path = _ensure_output_path()
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
            _write_section_with_ws(ws, section_id, payload)
        wb.Save()
        _log_excel("SUCCESS export_all")
    except Exception as exc:
        _log_excel(f"ERROR export_all error={exc!r}")
        raise
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
    clear_cache_file()
    clear_form_cache()
    return output_path

def register_form():
    return {
        "id": "condiciones_vacante",
        "name": FORM_NAME,
        "module": __name__,
    }

