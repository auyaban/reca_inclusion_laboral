import json
import os
import shutil
import time

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _get_desktop_dir,
    _normalize_cedula,
    _normalize_text,
    _sanitize_filename,
    _supabase_get,
)


FORM_ID = "induccion_operativa"
FORM_NAME = "Induccion Operativa"
SHEET_NAME = "7. INDUCCION OPERATIVA"

FORM_CACHE = {}
SECTION_1_CACHE = {}

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

SECTION_2 = {
    "title": "2. DATOS DEL VINCULADO",
    "fields": [
        {"id": "numero", "label": "No", "type": "texto"},
        {"id": "nombre_oferente", "label": "Nombre completo", "type": "texto"},
        {"id": "cedula", "label": "Cédula", "type": "texto"},
        {"id": "telefono_oferente", "label": "Teléfono", "type": "texto"},
        {"id": "cargo_oferente", "label": "Cargo", "type": "texto"},
    ],
}

SECTION_3 = {
    "title": "3. DESARROLLO DEL PROCESO DE INDUCCION OPERATIVA",
    "items": [
        {
            "id": "funciones_corresponden_perfil",
            "label": "Las funciones asignadas corresponden al perfil del cargo",
            "row": 19,
        },
        {
            "id": "explicacion_funciones",
            "label": "Se explico con detalle cada una de sus funciones",
            "row": 20,
        },
        {
            "id": "instrucciones_claras",
            "label": "Se brindaron instrucciones claras y precisas",
            "row": 21,
        },
        {
            "id": "sistema_medicion",
            "label": "Se explico el sistema de medicion de productividad, cumplimiento y calidad",
            "row": 22,
        },
        {
            "id": "induccion_maquinas",
            "label": "Se realizo induccion en uso de maquinas, instrumentos y equipos",
            "row": 23,
        },
        {
            "id": "presentacion_companeros",
            "label": "Se presento a sus companeros y equipo de trabajo cercano",
            "row": 24,
        },
        {
            "id": "presentacion_jefes",
            "label": "Se presento a su jefe y lideres",
            "row": 25,
        },
        {
            "id": "uso_epp",
            "label": "Se explico recomendaciones y uso de EPP",
            "row": 26,
        },
        {
            "id": "conducto_regular",
            "label": "Se compartio informacion sobre conducto regular y persona responsable",
            "row": 27,
        },
        {
            "id": "puesto_trabajo",
            "label": "Se asigno un puesto con requerimientos minimos",
            "row": 28,
        },
        {"id": "otros", "label": "Otros", "row": 29},
    ],
}

SECTION_3_EJECUCION_OPTIONS = ["Si", "No", "No aplica"]

SECTION_4 = {
    "title": "4. HABILIDADES SOCIOEMOCIONALES",
    "blocks": [
        {
            "id": "comprension_instrucciones",
            "title": "Comprension y ejecucion de instrucciones",
            "items": [
                {
                    "id": "reconoce_instrucciones",
                    "label": "Reconoce las instrucciones que le fueron brindadas",
                    "row": 32,
                },
                {
                    "id": "proceso_atencion",
                    "label": "Se percibe un adecuado proceso de atencion al brindar instrucciones",
                    "row": 33,
                },
            ],
            "note_row": 34,
        },
        {
            "id": "autonomia_tareas",
            "title": "Autonomia en desarrollo de tareas",
            "items": [
                {
                    "id": "identifica_funciones",
                    "label": "Identifica funciones sin supervision permanente",
                    "row": 35,
                },
                {
                    "id": "importancia_calidad",
                    "label": "Identifica la importancia de un trabajo con calidad",
                    "row": 36,
                },
            ],
            "note_row": 37,
        },
        {
            "id": "trabajo_equipo",
            "title": "Trabajo en equipo",
            "items": [
                {
                    "id": "relacion_companeros",
                    "label": "Entiende importancia de relacionarse con companeros",
                    "row": 38,
                },
                {
                    "id": "recibe_sugerencias",
                    "label": "Recibe adecuadamente sugerencias de companeros y superiores",
                    "row": 39,
                },
                {
                    "id": "objetivos_grupales",
                    "label": "Identifica importancia del cumplimiento de objetivos grupales",
                    "row": 40,
                },
            ],
            "note_row": 41,
        },
        {
            "id": "adaptacion_flexibilidad",
            "title": "Adaptacion y flexibilidad",
            "items": [
                {
                    "id": "reconoce_entorno",
                    "label": "Reconoce su entorno de trabajo",
                    "row": 42,
                },
                {
                    "id": "ajuste_cambios",
                    "label": "Comprende necesidad de ajustarse a cambios",
                    "row": 43,
                },
            ],
            "note_row": 44,
        },
        {
            "id": "solucion_problemas",
            "title": "Solucion de problemas",
            "items": [
                {
                    "id": "identifica_problema_laboral",
                    "label": "Identifica cuando tiene un problema laboral",
                    "row": 45,
                },
            ],
            "note_row": 46,
        },
        {
            "id": "comunicacion_asertiva",
            "title": "Comunicacion asertiva y efectiva",
            "items": [
                {
                    "id": "respeto_companeros",
                    "label": "Se dirige con respeto a companeros y superiores",
                    "row": 47,
                },
                {
                    "id": "lenguaje_corporal",
                    "label": "Lenguaje corporal demuestra amabilidad y respeto",
                    "row": 48,
                },
                {
                    "id": "reporte_novedades",
                    "label": "Identifica persona responsable para reporte de novedades",
                    "row": 49,
                },
            ],
            "note_row": 50,
        },
        {
            "id": "manejo_tiempo",
            "title": "Manejo del tiempo",
            "items": [
                {
                    "id": "organiza_actividades",
                    "label": "Organiza actividades segun tiempos de entrega",
                    "row": 51,
                },
                {
                    "id": "cumple_horario",
                    "label": "Cumple horario de ingreso y salida",
                    "row": 52,
                },
                {
                    "id": "identifica_horarios",
                    "label": "Identifica horarios de break, almuerzo y otros",
                    "row": 53,
                },
            ],
            "note_row": 54,
        },
        {
            "id": "iniciativa_proactividad",
            "title": "Iniciativa y proactividad",
            "items": [
                {
                    "id": "reporta_finalizacion",
                    "label": "Reporta cuando finaliza tareas asignadas",
                    "row": 55,
                },
            ],
            "note_row": 56,
        },
    ],
}

SECTION_4_NIVEL_APOYO_OPTIONS = [
    "0. No requiere apoyo.",
    "1. Nivel de apoyo bajo.",
    "2. Nivel de apoyo medio.",
    "3. Nivel de apoyo alto.",
    "No aplica.",
]

SECTION_4_OBSERVACIONES_OPTIONS = {
    32: [
        "0. Reconoce adecuadamente las instrucciones de manera autonoma.",
        "1. Requiere especificacion de instrucciones.",
        "2. Identifica instrucciones pero requiere apoyo para ejecutarlas.",
        "3. No identifica con facilidad las instrucciones.",
    ],
    33: [
        "0. Buen nivel de atencion y concentracion sin apoyos.",
        "1. Se dispersa por tiempos cortos.",
        "2. Requiere focalizar atencion antes de instruccion.",
        "3. Dificultad para mantener atencion y concentracion.",
    ],
    35: [
        "0. Identifica funciones sin supervision.",
        "1. Identifica funciones pero requiere supervision intermitente.",
        "2. Requiere entrenamiento para realizarlas sin supervision.",
        "3. Requiere supervision permanente.",
    ],
    36: [
        "0. Identifica y realiza trabajo con calidad.",
        "1. Identifica importancia de calidad.",
        "2. Requiere entrenamiento para trabajo con calidad.",
        "3. No identifica importancia y requiere entrenamiento.",
    ],
    38: [
        "0. Se relaciona con el equipo adecuadamente.",
        "1. Muestra interes pero requiere reforzar habilidades sociales.",
        "2. Poco interes, comparte parcialmente.",
        "3. No demuestra interes ni disposicion.",
    ],
    39: [
        "0. Recibe adecuadamente sugerencias y las aplica.",
        "1. Escucha sugerencias, le cuesta aplicarlas.",
        "2. Se le dificulta escucharlas, luego las aplica.",
        "3. No escucha sugerencias y se indispone.",
    ],
    40: [
        "0. Comprometido con objetivos grupales.",
        "1. Contribuye con personas afines.",
        "2. Prefiere actividades individuales.",
        "3. No se interesa por trabajo en equipo.",
    ],
    42: [
        "0. Reconoce su entorno laboral con facilidad.",
        "1. Toma tiempo, pero lo reconoce.",
        "2. Presenta inconvenientes para reconocerlo.",
        "3. No reconoce su entorno y se indispone.",
    ],
    43: [
        "0. Asimila cambios positivamente.",
        "1. Asimila cambios, tarda en adaptarse.",
        "2. Le cuesta asimilar, pero finalmente se adapta.",
        "3. No esta dispuesto a realizar cambios.",
    ],
    45: [
        "0. Identifica y resuelve problemas de manera autonoma.",
        "1. Identifica y requiere apoyo para resolver.",
        "2. No identifica de inicio, luego muestra disposicion.",
        "3. Identifica pero no muestra interes por resolver.",
    ],
    47: [
        "0. Comunica ideas con respeto y claridad.",
        "1. Comunica con respeto pero poca claridad.",
        "2. Comunica ideas con poca empatia.",
        "3. Comunica de forma poco clara y respetuosa.",
    ],
    48: [
        "0. Lenguaje corporal muestra respeto y empatia.",
        "1. A veces no demuestra amabilidad pero es respetuoso.",
        "2. Expresa ideas sin tener en cuenta al otro.",
        "3. Lenguaje corporal no muestra respeto ni empatia.",
    ],
    49: [
        "0. No se presentan novedades.",
        "1. Identifica responsable para reporte de novedades.",
        "2. Requiere apoyo para identificar responsable.",
        "3. No se interesa por identificar a quien reportar.",
    ],
    51: [
        "0. Prioriza actividades segun tiempos de entrega.",
        "1. Prioriza con apoyo parcial.",
        "2. Comprende prioridad, no sabe aplicarla.",
        "3. No muestra interes por priorizar.",
    ],
    52: [
        "0. Llega puntual, sin novedad.",
        "1. Puntual con alguna novedad.",
        "2. Llega tarde, comprende importancia de puntualidad.",
        "3. Llega tarde y no muestra preocupacion.",
    ],
    53: [
        "0. Comprende y cumple horarios establecidos.",
        "1. Comprende horarios, incumplimientos esporadicos.",
        "2. Dificultad de comprension afecta cumplimiento.",
        "3. Comprende pero no cumple horarios.",
    ],
    55: [
        "0. Reporta finalizacion de tareas asignadas.",
        "1. Sabe cuando finaliza pero no a quien reportar.",
        "2. Sabe cuando finaliza, pero no reporta.",
        "3. No se interesa por reportar finalizacion.",
    ],
}

SECTION_5 = {
    "title": "5. NIVEL DE APOYO REQUERIDO",
    "rows": [
        {
            "id": "condiciones_medicas_salud",
            "label": "Condiciones medicas y de salud",
            "row": 59,
        },
        {
            "id": "habilidades_basicas_vida_diaria",
            "label": "Habilidades basicas de la vida diaria",
            "row": 60,
        },
        {
            "id": "habilidades_socioemocionales",
            "label": "Habilidades socioemocionales",
            "row": 61,
        },
    ],
}

SECTION_5_NIVEL_OPTIONS = [
    "No requiero apoyo.",
    "Nivel de apoyo bajo.",
    "Nivel de apoyo medio.",
    "Nivel de apoyo alto.",
    "No aplica.",
]

SECTION_6 = {"title": "6. AJUSTES RAZONABLES REQUERIDOS"}
SECTION_7 = {"title": "7. PRIMER SEGUIMIENTO ESTABLECIDO PARA EL VINCULADO"}
SECTION_8 = {"title": "8. OBSERVACIONES /RECOMENDACIONES"}
SECTION_9 = {"title": "9.ASISTENTES", "rows": 4}

SECTION_1_SUPABASE_MAP = evaluacion_accesibilidad.SECTION_1_SUPABASE_MAP.copy()

EXCEL_MAPPING = {
    "section_1": {
        "fecha_visita": "E7",
        "modalidad": "M7",
        "nombre_empresa": "E8",
        "ciudad_empresa": "M8",
        "direccion_empresa": "E9",
        "nit_empresa": "M9",
        "correo_1": "E10",
        "telefono_empresa": "M10",
        "contacto_empresa": "E11",
        "cargo": "M11",
        "caja_compensacion": "E12",
        "sede_empresa": "M12",
        "asesor": "E13",
        "profesional_asignado": "M13",
    }
}
SECTION_2_TEMPLATE_ROW = 16
SECTION_2_ANCHOR = "3. DESARROLLO DEL PROCESO DE INDUCCION OPERATIVA"
SECTION_2_COL_MAP = {
    "numero": "A",
    "nombre_oferente": "B",
    "cedula": "H",
    "telefono_oferente": "M",
    "cargo_oferente": "P",
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
            "telefono_oferente",
            "cargo_oferente",
        ]
    )
    params = {
        "select": select_cols,
        "cedula_usuario": f"eq.{normalized}",
        "limit": 1,
    }
    data = _supabase_get("usuarios_reca", params, env_path=env_path)
    return data[0] if data else None


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


def confirm_section_9(payload):
    if payload is None:
        raise ValueError("section_9 requerida")
    set_section_cache("section_9", payload)
    FORM_CACHE["_last_section"] = "section_9"
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
        if (
            "induccion" in normalized
            and "operativa" in normalized
            and normalized.endswith(".xlsx")
        ):
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template de induccion operativa.")


def _ensure_output_path():
    template_path = _find_template_path()
    desktop = _get_desktop_dir()
    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    safe_company = _sanitize_filename(empresa_nombre) or "Empresa"
    output_dir = os.path.join(desktop, "Formatos Inclusion Laboral", safe_company)
    os.makedirs(output_dir, exist_ok=True)
    process_name = "Proceso de Induccion Operativa"
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


def _insert_vinculado_row(ws, insert_at):
    ws.Rows(insert_at).Insert()
    ws.Rows(SECTION_2_TEMPLATE_ROW).Copy(ws.Rows(insert_at))
    ws.Rows(insert_at).RowHeight = ws.Rows(SECTION_2_TEMPLATE_ROW).RowHeight


def _write_section_1(ws, payload):
    if not payload:
        payload = SECTION_1_CACHE
    if not payload:
        return
    mapping = EXCEL_MAPPING.get("section_1", {})
    for key, cell in mapping.items():
        if key in payload:
            ws.Range(cell).Value = payload.get(key)


def _write_section_2(ws, payload):
    if not payload:
        return
    anchor_row = _find_row_by_text(ws, SECTION_2_ANCHOR)
    total = len(payload)
    if total > 1:
        for _ in range(total - 1):
            _insert_vinculado_row(ws, anchor_row)
    for idx, row_data in enumerate(payload):
        target_row = SECTION_2_TEMPLATE_ROW + idx
        for field_id, col in SECTION_2_COL_MAP.items():
            value = row_data.get(field_id, "")
            if value in (None, ""):
                continue
            ws.Range(f"{col}{target_row}").Value = value


def _write_section_3(ws, payload):
    if not payload:
        return
    section_anchor_row = _find_row_by_text(ws, SECTION_2_ANCHOR)
    base_offset = section_anchor_row - 17
    for item in SECTION_3["items"]:
        item_id = item["id"]
        row_payload = payload.get(item_id, {}) if isinstance(payload, dict) else {}
        target_row = item["row"] + base_offset
        ejecucion = (row_payload.get("ejecucion") or "").strip()
        observaciones = (row_payload.get("observaciones") or "").strip()
        if ejecucion:
            ws.Range(f"H{target_row}").Value = ejecucion
        if observaciones:
            ws.Range(f"K{target_row}").Value = observaciones


def _write_section_4(ws, payload):
    if not payload:
        return
    section_4_anchor_row = _find_row_by_text(ws, "4. HABILIDADES SOCIOEMOCIONALES")
    base_offset = section_4_anchor_row - 30
    if not isinstance(payload, dict):
        return
    item_payload = payload.get("items", {})
    note_payload = payload.get("notes", {})
    for block in SECTION_4["blocks"]:
        for item in block["items"]:
            row = item["row"] + base_offset
            values = item_payload.get(item["id"], {})
            nivel = (values.get("nivel_apoyo") or "").strip()
            observaciones = (values.get("observaciones") or "").strip()
            if nivel:
                ws.Range(f"J{row}").Value = nivel
            if observaciones:
                ws.Range(f"N{row}").Value = observaciones
        note_row = block["note_row"] + base_offset
        note_value = (note_payload.get(block["id"]) or "").strip()
        if note_value:
            ws.Range(f"B{note_row}").Value = note_value


def _write_section_5(ws, payload):
    if not payload:
        return
    section_5_anchor_row = _find_row_by_text(ws, "5. NIVEL DE APOYO REQUERIDO")
    base_offset = section_5_anchor_row - 57
    for row_cfg in SECTION_5["rows"]:
        row = row_cfg["row"] + base_offset
        values = payload.get(row_cfg["id"], {}) if isinstance(payload, dict) else {}
        nivel = (values.get("nivel_apoyo_requerido") or "").strip()
        observaciones = (values.get("observaciones") or "").strip()
        if nivel:
            ws.Range(f"H{row}").Value = nivel
        if observaciones:
            ws.Range(f"M{row}").Value = observaciones


def _write_section_6(ws, payload):
    if not payload:
        return
    anchor = _find_row_by_text(ws, "6. AJUSTES RAZONABLES REQUERIDOS")
    value = (payload.get("ajustes_requeridos") or "").strip()
    if value:
        ws.Range(f"A{anchor + 1}").Value = value


def _write_section_7(ws, payload):
    if not payload:
        return
    anchor = _find_row_by_text(ws, "7. PRIMER SEGUIMIENTO ESTABLECIDO PARA EL VINCULADO")
    fecha = (payload.get("fecha_primer_seguimiento") or "").strip()
    if fecha:
        ws.Range(f"G{anchor + 1}").Value = fecha


def _write_section_8(ws, payload):
    if not payload:
        return
    anchor = _find_row_by_text(ws, "8. OBSERVACIONES /RECOMENDACIONES")
    texto = (payload.get("observaciones_recomendaciones") or "").strip()
    if texto:
        ws.Range(f"A{anchor + 1}").Value = texto


def _write_section_9(ws, payload):
    if not payload:
        return
    title_row = _find_row_by_text(ws, "9.ASISTENTES")
    start_row = title_row + 1
    base_rows = SECTION_9["rows"]
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
            ws.Range(f"L{row}").Value = cargo


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
        _write_section_2(ws, FORM_CACHE.get("section_2", []))
        _write_section_3(ws, FORM_CACHE.get("section_3", {}))
        _write_section_4(ws, FORM_CACHE.get("section_4", {}))
        _write_section_5(ws, FORM_CACHE.get("section_5", {}))
        _write_section_6(ws, FORM_CACHE.get("section_6", {}))
        _write_section_7(ws, FORM_CACHE.get("section_7", {}))
        _write_section_8(ws, FORM_CACHE.get("section_8", {}))
        _write_section_9(ws, FORM_CACHE.get("section_9", []))
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
    if clear_cache:
        clear_cache_file()
        clear_form_cache()
    return output_path

