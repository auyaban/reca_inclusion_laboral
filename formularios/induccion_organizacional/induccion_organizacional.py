import copy
import json
import os
import time

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _normalize_cedula,
    _sanitize_filename,
    _supabase_get,
    build_sheet_updates,
)
from logging_utils import log_excel_event


FORM_ID = "induccion_organizacional"
FORM_NAME = "Induccion Organizacional"
SHEET_NAME = "6. INDUCCIÓN ORGANIZACIONAL"

FORM_CACHE = {}
SECTION_1_CACHE = {}
SECTION_HISTORY_LIMIT = 10

SECTION_2 = {
    "title": "2. DATOS DEL VINCULADO",
    "fields": [
        {"id": "numero", "label": "No", "type": "texto"},
        {"id": "nombre_oferente", "label": "Nombre completo", "type": "texto"},
        {"id": "cedula", "label": "C\u00e9dula", "type": "texto"},
        {"id": "telefono_oferente", "label": "Tel\u00e9fono", "type": "texto"},
        {"id": "cargo_oferente", "label": "Cargo", "type": "texto"},
    ],
}

VISTO_OPTIONS = ["Si", "No", "No aplica"]
MEDIO_SOCIALIZACION_OPTIONS = [
    "Video",
    "Documentos escritos",
    "Imagenes",
    "Presentaciones",
    "Mixto",
    "Exposicion oral",
    "No aplica",
]

SECTION_3 = {
    "title": "3. DESARROLLO DEL PROCESO",
    "subsections": [
        {
            "id": "3_1",
            "title": "3.1 Generalidades de la empresa",
            "items": [
                {"id": "historia_empresa", "label": "Historia de la empresa.", "row": 21},
                {
                    "id": "mision_organizacional",
                    "label": "Explicacion y verificacion de la Mision organizacional.",
                    "row": 22,
                },
                {
                    "id": "vision_organizacional",
                    "label": "Explicacion y verificacion de la Vision organizacional",
                    "row": 23,
                },
                {
                    "id": "objetivos_valores_principios",
                    "label": "Explicacion y verificacion de Objetivos, Valores y Principios Organizacionales",
                    "row": 24,
                },
                {"id": "recorrido_empresa", "label": "Recorrido por la empresa-planta", "row": 25},
            ],
        },
        {
            "id": "3_2",
            "title": "3.2 Gestion Humana",
            "items": [
                {"id": "tramites_permisos", "label": "Explicacion tramites para permisos", "row": 27},
                {"id": "formas_pago", "label": "Explicacion de formas de pago", "row": 28},
                {
                    "id": "obligaciones_prohibiciones",
                    "label": "Explicacion de obligaciones y prohibiciones del empleado",
                    "row": 29,
                },
                {
                    "id": "normatividad_interna",
                    "label": "Explicacion de Normatividad interna de la empresa",
                    "row": 30,
                },
                {
                    "id": "practicas_inclusivas",
                    "label": "Explicacion de practicas inclusivas y/o una politica de diversidad e inclusion.",
                    "row": 31,
                },
                {"id": "horario_laboral", "label": "Horario laboral", "row": 32},
                {"id": "organigrama", "label": "Organigrama", "row": 33},
                {
                    "id": "incapacidades_permisos_calamidades",
                    "label": "Reporte y entrega de incapacidades, permisos, calamidades.",
                    "row": 34,
                },
                {"id": "equipos_tecnologicos", "label": "Entrega equipos tecnologicos", "row": 35},
                {"id": "comites", "label": "Explicacion de Comites", "row": 36},
                {
                    "id": "conductos_regulares_comunicacion",
                    "label": "Conductos regulares de comunicacion.",
                    "row": 37,
                },
            ],
        },
        {
            "id": "3_3",
            "title": "3.3 Sistema de gestion - seguridad y salud en el trabajo (SG-SST)",
            "items": [
                {
                    "id": "sgsst_general",
                    "label": "Explicacion del sistema de gestion seguridad y salud en el trabajo (SG-SST)",
                    "row": 39,
                },
                {
                    "id": "peligros_riesgos",
                    "label": "Explicacion de peligros, riesgos,accidentes y enfermedades laborales.",
                    "row": 40,
                },
                {
                    "id": "uso_epp",
                    "label": "Explicacion de uso de elementos de proteccion personal EPP.",
                    "row": 41,
                },
                {
                    "id": "politicas_medio_ambiente",
                    "label": "Explicacion de politicas de proteccion, prevencion y control del medio ambiente.",
                    "row": 42,
                },
                {
                    "id": "politicas_confidencialidad",
                    "label": "Explicacion de politicas de confidencialidad",
                    "row": 43,
                },
                {
                    "id": "plan_emergencias",
                    "label": "Explicacion de plan de emergencias, rutas de evacuacion y punto de encuentro.",
                    "row": 44,
                },
                {
                    "id": "prevencion_consumo",
                    "label": "Explicacion de politicas de prevencion del consumo de alcohol, tabaco y sustancias psicoactivas.",
                    "row": 45,
                },
                {"id": "normas_comite", "label": "Explicacion de normas de comite", "row": 46},
                {
                    "id": "normas_disciplinarias",
                    "label": "Explicacion de normas y medidas disciplinarias.",
                    "row": 47,
                },
                {
                    "id": "entrega_dotacion_epp",
                    "label": "Entrega de dotacion, elementos de proteccion personal EPP.",
                    "row": 48,
                },
                {"id": "brigada_emergencia", "label": "Explicacion brigada de emergencia", "row": 49},
                {
                    "id": "mecanismos_desempeno",
                    "label": "Mecanismos para medir o evaluar el desempeno",
                    "row": 50,
                },
                {
                    "id": "procedimiento_accidente",
                    "label": "Procedimiento que se debe seguir en caso de accidente de trabajo",
                    "row": 51,
                },
            ],
        },
        {
            "id": "3_4",
            "title": "3.4 Induccion general a puesto de trabajo",
            "items": [
                {
                    "id": "funciones_especificas",
                    "label": "Explicacion de funciones especificas.",
                    "row": 53,
                },
                {
                    "id": "horario_turnos",
                    "label": "Explicacion del horario o turnos de trabajo.",
                    "row": 54,
                },
                {"id": "dotacion_uniformes", "label": "Entrega dotacion uniformes.", "row": 55},
                {"id": "presentacion_equipo", "label": "Presentacion equipo de trabajo", "row": 56},
                {"id": "registro_ingreso", "label": "Registro ingreso empresa", "row": 57},
                {"id": "entrega_carnet", "label": "Entrega del Carnet", "row": 58},
                {"id": "recorrido_puesto", "label": "Recorrido puesto de trabajo", "row": 59},
            ],
        },
        {
            "id": "3_5",
            "title": "3.5 Proceso evaluativo de induccion",
            "items": [
                {"id": "evaluaciones", "label": "Evaluaciones", "row": 61},
                {"id": "plataformas_elearning", "label": "Plataformas e-learning", "row": 62},
            ],
        },
    ],
}

SECTION_4_OPTIONS = [
    "Video",
    "Documentos Escritos, Presentaciones, Imagenes y Evaluaciones escritas",
    "Plataformas",
    "No aplica",
]

SECTION_4_RECOMMENDATIONS = {
    "Video": (
        "1. Subtitulos precisos y sincronizados con dialogo y sonidos.\n"
        "2. Descripci\u00f3nes de audio sobre lo que sucede en video.\n"
        "3. Iluminacion adecuada y contraste alto.\n"
        "4. Audio claro, entendible y con transcripcion.\n"
        "5. Evitar parpadeos, destellos y patrones moviles.\n"
        "6. Navegabilidad e interaccion adecuadas para discapacidad cognitiva o movilidad reducida.\n"
        "7. Duracion sugerida: difusion maximo 2 minutos; formacion maximo 5 minutos.\n"
        "8. Incluir LSC para discapacidad auditiva; interprete en angulo inferior derecho.\n\n"
        "RECOMENDACION GENERAL\n"
        "- Si el video supera 10 minutos, hacer pausas cada 2-3 minutos para retroalimentacion.\n"
        "- Acompanamiento permanente durante el video para resolver preguntas."
    ),
    "Documentos Escritos, Presentaciones, Imagenes y Evaluaciones escritas": (
        "1. Usar letra legible (Arial, Calibri, Times New Roman o Tahoma).\n"
        "2. Tamano de letra no menor a 12 puntos, ajustado a necesidad.\n"
        "3. Contraste adecuado entre fondo y letra.\n"
        "4. Interlineado sugerido de 1.5 o 2.\n"
        "5. Texto en posicion vertical de izquierda a derecha.\n"
        "6. Diseno sencillo, evitando exceso de elementos decorativos.\n"
        "7. Imagenes con tamano y resolucion adecuados.\n"
        "8. Lenguaje claro y sencillo, evitando jerga tecnica.\n"
        "9. Encabezados y subtitulos para organizar informacion.\n"
        "10. Uso de listas y tablas para estructura.\n"
        "11. Incluir descripcion en imagenes, graficos y tablas.\n"
        "12. Estructura estandar con tabla de contenido y navegacion facil.\n"
        "13. Formato estandar (PDF o HTML) compatible con lectores de pantalla.\n"
        "14. Para imagenes usar formatos estandar (JPEG o PNG) compatibles."
    ),
    "Plataformas": (
        "1. Estructura de navegacion estandar con tabla de contenido.\n"
        "2. Botones y enlaces con tamano adecuado y alto contraste.\n"
        "3. Teclas de acceso rapido para navegacion.\n"
        "4. Tecnologias de reconocimiento y comandos de voz.\n"
        "5. Compatibilidad con herramientas de accesibilidad (asistente de voz, talkback, jaws, magic).\n\n"
        "RECOMENDACION GENERAL\n"
        "- Si no es posible ajustar accesibilidad en plataforma, asignar par de apoyo para lectura en voz alta y retroalimentacion constante."
    ),
}

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
            "id": "direcci\u00f3n_empresa",
            "label": "Direcci\u00f3n de la empresa",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {"id": "nit_empresa", "label": "N\u00famero de NIT", "source": "input"},
        {
            "id": "correo_1",
            "label": "Correo electr\u00f3nico",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "telefono_empresa",
            "label": "Tel\u00e9fonos",
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
            "label": "Empresa afiliada a Caja de Compensaci\u00f3n",
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
        "modalidad": "N7",
        "nombre_empresa": "D8",
        "ciudad_empresa": "N8",
        "direccion_empresa": "D9",
        "nit_empresa": "N9",
        "correo_1": "D10",
        "telefono_empresa": "N10",
        "contacto_empresa": "D11",
        "cargo": "N11",
        "caja_compensacion": "D12",
        "sede_empresa": "N12",
        "asesor": "D13",
        "profesional_asignado": "N13",
    }
}
SECTION_2_TEMPLATE_ROW = 16
SECTION_2_ANCHOR = "3. DESARROLLO DEL PROCESO"
SECTION_2_COL_MAP = {
    "numero": "A",
    "nombre_oferente": "B",
    "cedula": "H",
    "telefono_oferente": "M",
    "cargo_oferente": "P",
}

# ---------------------------------------------------------------------------
# Fixed row constants for sections that previously used _find_row_by_text().
# These correspond to the master Google Sheet template layout.
# ---------------------------------------------------------------------------
# "3. DESARROLLO DEL PROCESO" is at row 17 in the master, so base_offset = 0
# and the row numbers in SECTION_3 items are the actual rows.
SECTION_3_ANCHOR_ROW = 17

# "4. RECOMENDACIONES DE ACCESIBILIDAD..." title row
SECTION_4_START_ROW = 64
# "5. OBSERVACIONES" title row
SECTION_5_ROW = 67
# Row where the observaciones text is written
SECTION_5_TEXT_ROW = 68
# "6. ASISTENTES" title row
SECTION_6_TITLE_ROW = 70
SECTION_6_START_ROW = 71
SECTION_6_NOMBRE_COL = "C"
SECTION_6_CARGO_COL = "L"
SECTION_6_BASE_ROWS = 3


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


def _record_section_history(section_id, payload, source="manual"):
    if not section_id or str(section_id).startswith("_"):
        return
    if not _has_meaningful_values(payload):
        return
    history_root = FORM_CACHE.setdefault("_section_history", {})
    if not isinstance(history_root, dict):
        history_root = {}
        FORM_CACHE["_section_history"] = history_root
    entries = history_root.setdefault(section_id, [])
    if not isinstance(entries, list):
        entries = []
        history_root[section_id] = entries
    snapshot = copy.deepcopy(payload)
    if entries:
        last_entry = entries[-1] if isinstance(entries[-1], dict) else {}
        if last_entry.get("payload") == snapshot:
            return
    entries.append(
        {
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "source": str(source or "manual").strip() or "manual",
            "payload": snapshot,
        }
    )
    if len(entries) > SECTION_HISTORY_LIMIT:
        del entries[:-SECTION_HISTORY_LIMIT]


def set_section_cache(section_id, payload, *, source="manual"):
    if not section_id:
        raise ValueError("section_id requerido")
    normalized_payload = payload if payload is not None else {}
    FORM_CACHE[section_id] = normalized_payload
    FORM_CACHE["_last_saved_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    FORM_CACHE["_last_saved_section"] = section_id
    FORM_CACHE["_last_saved_source"] = str(source or "manual").strip() or "manual"
    _record_section_history(section_id, normalized_payload, source=source)


def get_form_cache():
    return dict(FORM_CACHE)


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


def _log_excel(message):
    try:
        log_excel_event(message)
    except Exception:
        return


# ---------------------------------------------------------------------------
# Google Sheets write builders
# ---------------------------------------------------------------------------

def _build_section_1_writes(payload):
    """Return list of update dicts for section 1 (datos generales)."""
    if not payload:
        payload = SECTION_1_CACHE
    return build_sheet_updates(SHEET_NAME, EXCEL_MAPPING.get("section_1", {}), payload or {})


def _build_section_2_writes(payload):
    """Return list of update dicts for section 2 (datos del vinculado)."""
    if not payload:
        return []
    writes = []
    for idx, row_data in enumerate(payload):
        target_row = SECTION_2_TEMPLATE_ROW + idx
        for field_id, col in SECTION_2_COL_MAP.items():
            value = row_data.get(field_id, "")
            if value in (None, ""):
                continue
            writes.append({
                "range": f"'{SHEET_NAME}'!{col}{target_row}",
                "value": value,
            })
    return writes


def _build_section_3_writes(payload):
    """Return list of update dicts for section 3 (desarrollo del proceso).

    Uses the fixed row numbers from SECTION_3 item definitions directly
    (the master anchor row is 17, giving a base_offset of 0).
    """
    if not payload:
        return []
    writes = []
    for subsection in SECTION_3["subsections"]:
        for item in subsection["items"]:
            item_id = item["id"]
            row_payload = payload.get(item_id, {}) if isinstance(payload, dict) else {}
            target_row = item["row"]
            visto = row_payload.get("visto", "")
            responsable = row_payload.get("responsable", "")
            medio = row_payload.get("medio_socializacion", "")
            descripcion = row_payload.get("descripcion", "")
            if visto not in (None, ""):
                writes.append({"range": f"'{SHEET_NAME}'!H{target_row}", "value": visto})
            if responsable not in (None, ""):
                writes.append({"range": f"'{SHEET_NAME}'!K{target_row}", "value": responsable})
            if medio not in (None, ""):
                writes.append({"range": f"'{SHEET_NAME}'!M{target_row}", "value": medio})
            if descripcion not in (None, ""):
                writes.append({"range": f"'{SHEET_NAME}'!P{target_row}", "value": descripcion})
    return writes


def _build_section_4_writes(payload):
    """Return list of update dicts for section 4 (recomendaciones de accesibilidad).

    The three recommendation rows sit at SECTION_5_ROW - 3, -2, -1.
    Only writes the dropdown value (column A); column G has a formula
    that auto-populates the recommendation text based on the selection.
    """
    if not payload:
        return []
    writes = []
    rows = [SECTION_5_ROW - 3, SECTION_5_ROW - 2, SECTION_5_ROW - 1]
    for idx, row in enumerate(rows):
        entry = payload[idx] if idx < len(payload) else {}
        medio = (entry.get("medio") or "").strip()
        if medio:
            writes.append({"range": f"'{SHEET_NAME}'!A{row}", "value": medio})
    return writes


def _build_section_5_writes(payload):
    """Return list of update dicts for section 5 (observaciones)."""
    if not payload:
        return []
    observaciones = (payload.get("observaciones") or "").strip()
    if not observaciones:
        return []
    return [{"range": f"'{SHEET_NAME}'!A{SECTION_5_TEXT_ROW}", "value": observaciones}]


def _build_section_6_writes(payload):
    """Return list of update dicts for section 6 (asistentes)."""
    if not payload:
        return []
    writes = []
    for idx, entry in enumerate(payload):
        row = SECTION_6_START_ROW + idx
        nombre = (entry.get("nombre") or "").strip()
        cargo = (entry.get("cargo") or "").strip()
        if nombre:
            writes.append({
                "range": f"'{SHEET_NAME}'!{SECTION_6_NOMBRE_COL}{row}",
                "value": nombre,
            })
        if cargo:
            writes.append({
                "range": f"'{SHEET_NAME}'!{SECTION_6_CARGO_COL}{row}",
                "value": cargo,
            })
    return writes


def _build_section_6_row_insertions(payload):
    if not payload:
        return []
    total_rows = len(payload)
    if total_rows <= SECTION_6_BASE_ROWS:
        return []
    return [
        {
            "sheet_name": SHEET_NAME,
            "start_row": SECTION_6_START_ROW,
            "base_rows": SECTION_6_BASE_ROWS,
            "total_rows": total_rows,
        }
    ]


def _has_meaningful_values(value):
    if isinstance(value, dict):
        for key, item in value.items():
            if str(key or "").startswith("_"):
                continue
            if _has_meaningful_values(item):
                return True
        return False
    if isinstance(value, list):
        return any(_has_meaningful_values(item) for item in value)
    return str(value or "").strip() != ""


def _validate_cache_before_export():
    section_3 = FORM_CACHE.get("section_3", {})
    if _has_meaningful_values(section_3):
        return
    later_sections_have_data = any(
        _has_meaningful_values(FORM_CACHE.get(section_id))
        for section_id in ("section_4", "section_5", "section_6")
    )
    if later_sections_have_data:
        raise RuntimeError(
            "La seccion 3 quedo vacia en el cache. Se cancelo la exportacion para evitar "
            "generar un Excel incompleto. Revisa la seccion 3 antes de finalizar."
        )
    raise RuntimeError("La seccion 3 no tiene informacion diligenciada. Revisa esa seccion antes de finalizar.")


def export_to_excel(clear_cache=True):
    if not FORM_CACHE.get("section_1") and cache_file_exists():
        load_cache_from_file()
    _validate_cache_before_export()

    from google_sheets_client import get_master_template_id
    from drive_upload import publish_sheet_from_template

    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    base_name = _sanitize_filename(empresa_nombre)

    writes = []
    writes.extend(_build_section_1_writes(FORM_CACHE.get("section_1", {})))
    writes.extend(_build_section_2_writes(FORM_CACHE.get("section_2", [])))
    writes.extend(_build_section_3_writes(FORM_CACHE.get("section_3", {})))
    writes.extend(_build_section_4_writes(FORM_CACHE.get("section_4", [])))
    writes.extend(_build_section_5_writes(FORM_CACHE.get("section_5", {})))
    writes.extend(_build_section_6_writes(FORM_CACHE.get("section_6", [])))
    row_insertions = _build_section_6_row_insertions(FORM_CACHE.get("section_6", []))

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
