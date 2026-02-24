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


FORM_ID = "induccion_organizacional"
FORM_NAME = "Induccion Organizacional"
SHEET_NAME = "6. INDUCCION ORGANIZACIONAL"

FORM_CACHE = {}
SECTION_1_CACHE = {}

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
        "2. Descripciónes de audio sobre lo que sucede en video.\n"
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
            and "organizacional" in normalized
            and normalized.endswith(".xlsx")
        ):
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template de induccion organizacional.")


def _ensure_output_path():
    template_path = _find_template_path()
    desktop = _get_desktop_dir()
    empresa_nombre = SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    safe_company = _sanitize_filename(empresa_nombre) or "Empresa"
    output_dir = os.path.join(desktop, "Formatos Inclusion Laboral", safe_company)
    os.makedirs(output_dir, exist_ok=True)
    process_name = "Proceso de Induccion Organizacional"
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
    try:
        anchor_row = _find_row_by_text(ws, SECTION_2_ANCHOR)
    except Exception:
        anchor_row = SECTION_2_TEMPLATE_ROW + 1
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
    section_anchor_row = _find_row_by_text(ws, "3. DESARROLLO DEL PROCESO")
    base_offset = section_anchor_row - 17
    for subsection in SECTION_3["subsections"]:
        for item in subsection["items"]:
            item_id = item["id"]
            row_payload = payload.get(item_id, {}) if isinstance(payload, dict) else {}
            target_row = item["row"] + base_offset
            visto = row_payload.get("visto", "")
            responsable = row_payload.get("responsable", "")
            medio = row_payload.get("medio_socializacion", "")
            descripcion = row_payload.get("descripcion", "")
            if visto not in (None, ""):
                ws.Range(f"H{target_row}").Value = visto
            if responsable not in (None, ""):
                ws.Range(f"K{target_row}").Value = responsable
            if medio not in (None, ""):
                ws.Range(f"M{target_row}").Value = medio
            if descripcion not in (None, ""):
                ws.Range(f"P{target_row}").Value = descripcion


def _write_section_4(ws, payload):
    if not payload:
        return
    section_5_row = _find_row_by_text(ws, "5. OBSERVACIONES")
    rows = [section_5_row - 3, section_5_row - 2, section_5_row - 1]
    for idx, row in enumerate(rows):
        entry = payload[idx] if idx < len(payload) else {}
        medio = (entry.get("medio") or "").strip()
        if medio:
            ws.Range(f"A{row}").Value = medio
        texto = (entry.get("recomendacion") or "").strip()
        if not texto and medio in SECTION_4_RECOMMENDATIONS:
            texto = SECTION_4_RECOMMENDATIONS.get(medio, "")
        if texto:
            ws.Range(f"G{row}").Value = texto


def _write_section_5(ws, payload):
    if not payload:
        return
    observaciones = (payload.get("observaciones") or "").strip()
    if observaciones:
        section_5_row = _find_row_by_text(ws, "5. OBSERVACIONES")
        ws.Range(f"A{section_5_row + 1}").Value = observaciones


def _write_section_6(ws, payload):
    if not payload:
        return
    title_row = _find_row_by_text(ws, "6. ASISTENTES")
    start_row = title_row + 1
    base_rows = 4
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
        _write_section_4(ws, FORM_CACHE.get("section_4", []))
        _write_section_5(ws, FORM_CACHE.get("section_5", {}))
        _write_section_6(ws, FORM_CACHE.get("section_6", []))
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
    if clear_cache:
        clear_cache_file()
        clear_form_cache()
    return output_path

