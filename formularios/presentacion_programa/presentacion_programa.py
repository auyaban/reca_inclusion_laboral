import os
import json
import re
import shutil
import time
from formularios.common import (
    format_checkbox_symbol,
    write_checkbox_symbol,
    _build_process_output_path,
    _get_desktop_dir,
    _next_available_file_path,
    _normalize_text,
    sanitize_logo_error_cells,
    autofit_rows,
    clear_written_rows,
    ws_write,
    _sanitize_filename,
    _supabase_get,
)
from logging_utils import log_excel_event
from version_info import resource_path

FORM_NAME = "Presentacion/Reactivacion del programa de inclusion laboral"

SECTIONS = [
    "1. DATOS GENERALES",
    "2. TEMARIO",
    "3. DESCRIPCION DE LOS TEMAS",
    "4. ACUERDOS Y OBSERVACIONES DE LA REUNION",
    "5. ASISTENTES",
]

SECTION_1 = {
    "title": "1. DATOS GENERALES",
    "nit_lookup_field": "nit_empresa",
    "fields": [
        {
            "id": "tipo_visita",
            "label": "Tipo de visita",
            "source": "input",
            "options": ["Presentación", "Reactivación"],
        },
        {"id": "fecha_visita", "label": "Fecha de la visita", "source": "input"},
        {
            "id": "nombre_empresa",
            "label": "Nombre de la empresa",
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
        {
            "id": "correo_1",
            "label": "Correo electrónico",
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
            "id": "caja_compensacion",
            "label": "Empresa afiliada a Caja de Compensación",
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
        {
            "id": "asesor",
            "label": "Asesor fidelización",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "modalidad",
            "label": "Modalidad",
            "source": "input",
            "options": ["Virtual", "Presencial", "Mixto", "No aplica"],
        },
        {
            "id": "ciudad_empresa",
            "label": "Ciudad/Municipio",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {"id": "nit_empresa", "label": "Número de NIT", "source": "input"},
        {
            "id": "telefono_empresa",
            "label": "Teléfonos responsable empresa",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "cargo",
            "label": "Cargo responsable empresa",
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
            "id": "correo_profesional",
            "label": "Correo profesional RECA",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
        {
            "id": "correo_asesor",
            "label": "Correo asesor",
            "source": "supabase",
            "table": "empresas",
            "readonly": True,
        },
    ],
}

SECTION_2 = {
    "title": "2. TEMARIO",
    "items": [
        "Presentación del Programa de Inclusión Laboral",
        "Presentación del Operador RECA",
        "Servicios de Inclusión Laboral (IL) que ofrece el programa",
        "La Inclusión Laboral como variable de Responsabilidad Social Empresarial (RSE)",
        "Beneficios de contratar personas con discapacidad",
        "Rol del operador del programa y la Agencia de Empleo",
        "Alcance del programa",
        "Motivación empresarial para hacer procesos de Inclusión Laboral.",
        "Términos y condiciones en la ruta de Inclusión Laboral (consentimiento informado de uso y protección de datos).",
        "Observaciones entre las partes interesadas.",
    ],
}

SECTION_3 = {
    "title": "3. DESCRIPCION DE LOS TEMAS",
    "items": [
        {
            "id": 1,
            "title": "Presentación del Programa de Inclusión Laboral",
            "content": """El profesional de RECA presenta los servicios del Programa de Inclusión Laboral (IL) como operador de la Agencia de Empleo de Compensar, explicando la ruta del programa, el cual busca brindar estrategias y ajustes razonables para que la empresa realice procesos incluyentes.

Este programa está enmarcado dentro del Servicio Público de Empleo, por lo tanto, es gratuito tanto para las empresas como para las personas. Con base en esto, RECA apoyará las siguientes etapas del proceso de Inclusión Laboral:

- Evaluación de Accesibilidad en las instalaciones de la Empresa y reporte de ajustes razonables.
- Revisión de Condiciones de la Vacante.
- Proceso de Selección Incluyente.
- Proceso de Contratación Incluyente.
- Inducción Organizacional.
- Inducción Operativa.
- Sensibilización una (1) sesión.
- Seguimiento y Acompañamiento al Proceso de Inclusión Laboral.

Nota: Para acceder a los servicios descritos, es requisito que la empresa haya recibido la remisión de los oferentes por parte de la Agencia de Empleo y del Emprendimiento Compensar.""",
        },
        {
            "id": 2,
            "title": "Presentación del operador",
            "content": """RECA- Red Empleo con Apoyo, es una Organización sin ánimo de lucro especializada en procesos de inclusión laboral, que brinda acompañamiento y apoyo profesional a las empresas interesadas en incluir personas en condición de discapacidad, conformado por un equipo interdisciplinar con amplia experiencia en la realización de dichos procesos.

Somos el operador autorizado del programa de inclusión laboral de la Agencia de Empleo de Compensar para las empresas inscritas.

Durante el proceso se brinda apoyo en los siguientes aspectos:
- Asesoría a las organizaciones en "Buenas Prácticas de Inclusión".
- Brindar información sobre normatividad vigente respecto a la contratación del personal con discapacidad.
- Explicación de los beneficios y cómo acceder a ellos.""",
        },
        {
            "id": 3,
            "title": "Servicios de la ruta de Inclusión Laboral",
            "content": """Evaluación de Accesibilidad: se informa a la empresa el propósito de la visita de Evaluación de Accesibilidad:

1. Identificar barreras arquitectónicas, urbanísticas y de movilidad enfocadas en los tres pilares de la accesibilidad del Diseño Universal para el proceso de IL.
2. Identificar ajustes razonables y/o acciones que faciliten los procesos de inclusión laboral y la identificación de ayudas técnicas requeridas, que se confirmarán en el momento de la vinculación de PcD.
3. La Empresa y la Agencia de Empleo de Compensar reciben copia del reporte de la Evaluación de Accesibilidad y del compromiso por parte de la empresa en la implementación de los ajustes razonables para seguir con el proceso de vinculación incluyente.

Nota Aclaratoria: El proceso de Evaluación de accesibilidad es exclusivo para empresas afiliadas a caja de compensación familiar Compensar.

Revisión Condiciones de la Vacante: Según la o las vacantes que haya identificado la empresa, se revisarán las condiciones específicas para determinar ajustes y verificar el perfil requerido, así como el tipo de accesibilidad que requiere la vacante teniendo en cuenta los tipos de discapacidad. Se concretarán las competencias para el cargo así como necesidades de apoyo. Todo lo anterior, con el fin de lograr en la empresa un buen proceso de selección, contando con ajustes y apoyos que logren minimizar las brechas a nivel laboral que hay con las Personas con Discapacidad.

Procesos de Selección Incluyente: Apoyar al equipo de selección de la empresa brindando las estrategias necesarias para identificar las competencias laborales requeridas para cubrir la vacante, garantizando los ajustes razonables que faciliten la selección del personal idóneo. Apoyo en la adaptación de la prueba técnica que se realice.

Nota Aclaratoria: Para la aplicación de pruebas psicotécnicas se requiere de verificación previa, ya que no siempre aplican para personas en condición de discapacidad.

Contratación incluyente: Brindar el apoyo para realizar un proceso de contratación incluyente, asegurando la comprensión de las responsabilidades de las partes y la plena comprensión de la información por parte de las personas en condición de discapacidad, dejando claro aspectos tales como: normas, salario, forma de pago, conducto regular y reglamento interno de trabajo, entre otros.

Inducción Organizacional: Empoderar a la empresa a través de la transferencia de conocimiento para que realicen de manera autónoma el proceso de Inducción Organizacional en futuras contrataciones de Personas con Discapacidad.

Nota Aclaratoria: El proceso de Inducción Organizacional es exclusivo para empresas afiliadas a caja de compensación Compensar.

Inducción Operativa: Brindar estrategias a la empresa y persona vinculada para la comprensión clara de las funciones del cargo, entender procesos internos, despejar dudas y de esta manera obtener un buen desempeño, teniendo en cuenta el ajuste a la curva de aprendizaje del vinculado con relación a las labores a realizar.

Sensibilización: Capacitar a los colaboradores de la empresa explicando las pautas de interacción y comunicación, características, necesidades de apoyo y ajustes razonables para lograr un proceso de inclusión dentro de la empresa. Además, se despejan dudas para generar un ambiente favorable y de esta manera la Persona con Discapacidad que ingrese sea recibida y apoyada por el equipo de trabajo.

Nota Aclaratoria: El proceso de Capacitación y Sensibilización es exclusivo para empresas afiliadas a caja de compensación Compensar.

Seguimiento Inclusión laboral: Brindar apoyo a la empresa durante un tiempo limitado para guiar, orientar y asesorar sobre el proceso de inclusión laboral con la población con discapacidad. Se buscan estrategias conjuntas para poder lograr la adaptación de todo el equipo humano, promoviendo el empoderamiento sobre el proceso de inclusión laboral, adicionalmente se apoya al vinculado brindando estrategias para alcanzar un buen desempeño. Sin embargo, se deben tener en cuenta los siguientes lineamientos:

- Empresas que son afiliadas a caja de compensación Compensar podrán llegar a tener hasta 6 seguimientos dependiendo el nivel de apoyo que requiera la PcD, en un periodo de tiempo máximo de hasta 8 meses o menos.
- Empresas que no son afiliadas a caja de compensación Compensar podrán llegar a tener hasta 3 seguimientos en un periodo máximo de hasta 6 meses o menos, dependiendo el nivel de apoyo que requiera la PcD.

Servicios Fallidos: Durante el proceso de inclusión la empresa se compromete a disponer de los tiempos concertados para atender cada proceso y cumplir las citas programadas, ya que cuando no se cumplen se consideran servicios fallidos, que no tendrán posibilidad de reposición, lo que afecta el éxito del proceso. Se acuerda establecer canales de comunicación efectivos para coordinar visitas al proceso de inclusión laboral para evitar así generar inconvenientes y/o retrasos.""",
        },
        {
            "id": 4,
            "title": "Presentación de la Inclusión Laboral como Variable de Responsabilidad Social Empresarial (RSE)",
            "content": """Se enfatiza en las acciones de inclusión que pueden fortalecer el programa de RSE, favoreciendo las condiciones laborales para la vinculación de población con discapacidad.""",
        },
        {
            "id": 5,
            "title": "Beneficios de Contratación de PcD",
            "content": """Se presenta la tabla de beneficios con los que cuenta la empresa al contratar personas en condición de discapacidad y el impacto positivo que tienen los procesos de inclusión en la organización.""",
        },
        {
            "id": 6,
            "title": "Rol del Operador del Programa y la Agencia de Empleo",
            "content": """Agencia empleo Compensar: Facilita el acceso a oportunidades laborales a nivel nacional, a través de orientación y direccionamiento en la ruta de empleabilidad.

Red Empleo con Apoyo: RECA es uno de los Operadores de la Agencia de Empleo y Fomento Empresarial Compensar y aplicará los procesos de inclusión laboral en las empresas inscritas, teniendo en cuenta que el alcance de RECA no implica responsabilidades sobre la contratación de las personas y efectos legales que puedan generar las mismas.""",
        },
        {
            "id": 7,
            "title": "Alcance del Programa de Inclusión",
            "content": """- Todos los servicios para personas con discapacidad auditiva cuentan con el acompañamiento de intérprete de Lengua de Señas Colombiana (LSC) en modalidad virtual y presencial con un límite de tiempo con base al proceso que se está acompañando.

- Para servicios adicionales que requieran de acompañamiento de intérprete LSC, se deberá informar con tres (3) días de anterioridad con la finalidad de garantizar el intérprete.

- En caso de requerir apoyos profesionales adicionales a los descritos arriba por parte de RECA, se solicitará por escrito la autorización del servicio al asesor de la Agencia de empleo Compensar asignado a la empresa inscrita.

- Se especifica que el alcance del acompañamiento de RECA, no implica responsabilidades sobre la contratación de las personas y efectos legales de las mismas.

Nota Aclaratoria: RECA tiene definido un horario laboral de 8:00 am - 5:00 pm de LV y sábados hasta mediodía, a su vez entregará la documentación establecida para cada uno de los procesos de la ruta de Inclusión Laboral a las partes interesadas.""",
        },
        {
            "id": 8,
            "title": "Motivación de la organización para la contratación de PcD",
            "type": "checkboxes",
            "content": {
                "Responsabilidad Social Empresarial": False,
                "Objetivos y metas para la diversidad, equidad e inclusión.": False,
                "Avances a nivel global de impacto en Colombia": False,
                "Beneficios Tributarios": True,
                "Beneficios en la contratación de población en riesgo de exclusión": False,
                "Ventaja en licitaciones públicas": True,
                "Cumplimiento de la normativa establecida por el Estado Colombiano.": False,
                "Experiencia en la vinculación de personas en condición de discapacidad.": False,
            },
        },
        {
            "id": 9,
            "title": "Términos y condiciones",
            "content": """Se socializan los términos y condiciones establecidos para la implementación de la ruta de Inclusión Laboral por parte de RECA aprobados por la Agencia de Empleo y Fomento Empresarial Compensar. También se resalta la importancia por parte de la empresa de realizar la firma del documento de Autorización para el Tratamiento de Datos, esto con el fin de poder tomar evidencias en los avances que se tengan en el programa IL y plasmarlos en diferentes plataformas (redes sociales, página web y otras).""",
        },
        {
            "id": 10,
            "title": "Observaciones",
            "content": """Al finalizar la reunión con el asesor de Compensar y el empresario, se dará el espacio para aclarar dudas, preguntas de interés entre las partes, dejando por escrito en el presente documento todas las recomendaciones y observaciones que resulten del proceso.""",
        },
    ],
}

SECTION_4 = {
    "title": "4. ACUERDOS Y OBSERVACIONES DE LA REUNION",
    "field": {
        "id": "acuerdos_observaciones",
        "label": "Acuerdos y observaciones de la reunión",
        "source": "input",
        "type": "textarea",
    },
}

RUTA_INCLUSION_TEMPLATES = {
    "objetivo_y_participantes": """
Se llevó a cabo una reunión virtual con el objetivo de dar a conocer Ruta de inclusión ante las iniciativas de inclusión laboral en la empresa bajo el cumplimiento Normativo.

En el encuentro participó representante clave del área de Gestión Humana, Asesora desde Agencia de Empleo y Fomento Empresarial Compensar, y desde RECA Coordinación de Empleo Inclusivo.

El espacio inicia con presentación de representante de la empresa, el cual dialoga sobre el interés de conocer la ruta y acompañamiento ante iniciativas de vinculación de personas con discapacidad.
""".strip(),
    "roles_proceso_seleccion": """
Seguidamente la Asesora de la Agencia clarificó que la Agencia es la entidad encargada de la selección y envío de candidatos que se ajusten a los perfiles requeridos y el proceso de envío de hojas de vida y la evaluación por competencias será ejecutado por la analista de la Agencia, y se informa el rol de RECA como operador del programa de inclusión siendo este brindar acompañamiento técnico y especializado durante todo el proceso de inclusión, sin costo alguno para la empresa.

Se informa tiempo de respuesta en el envío de candidatos siendo de 4 días hábiles a partir de la publicación de la vacante y la importancia de la flexibilización del perfil, y la no creación de un cargo en específico para población con discapacidad.
""".strip(),
    "certificado_discapacidad": """
Se reitera que la Agencia no realiza el envío del Certificado de Discapacidad, no obstante, desde Compensar, la psicóloga encargada verifica durante el contacto con el candidato si este cuenta con dicho documento siendo este proceso de preselección, posteriormente, en el proceso de firma de contrato, corresponde a la empresa validar el certificado emitido por la Secretaría de Salud.
""".strip(),
    "acompanamiento_reca": """
Seguidamente desde RECA se comunica los procesos a ejecutar, los cuales no tienen costo al estar con la caja Compensar:
* Evaluación accesibilidad
* Revisión de la vacante
* Acompañamiento en procesos de entrevistas
* Acompañamiento a firma de contrato
* Inducción organizacional
* Inducción operativa
* Sensibilización
* Seguimiento a cada vinculado se realizarán seis (6) de manera individual tanto con el nuevo colaborador como con su jefe directo para asegurar una adaptación exitosa, se reitera el acompañamiento al proceso a candidatos exclusivamente remitidos por la agencia, y en el caso de la empresa tener candidatos estos deberán ser remitido al asesor de la Agencia vía correo electrónico para asi ingresar a la ruta de la Agencia y ejecutar el acompañamiento desde RECA.
""".strip(),
    "seguimiento_y_normativa": """
Se reitera la importancia de contar con la retroalimentación vía correo electrónico a la Agencia con copia a RECA, ante procesos de entrevista en el caso de no pasar candidatos filtros de selección y solicitar nuevos candidatos; y firma de contrato.

Se dialoga de la nueva ley 2466 del 2025, en donde se orienta ante totalidad de colaboradores la vinculación de 2 personas con discapacidad, se informa beneficios tangibles y no tangibles bajo la ley 361 art. 31 deducción en la renta por vinculación de personas con discapacidad y el apoyo que está entregando la secretaria de desarrollo.

El Decreto 0223 de 2026 es explícito al indicar en el numeral 1 de su artículo 2.2.6.3.3.33. que:

“Los aprendices no integran la base de trabajadores de carácter permanente de la empresa, para efectos del cálculo de la cuota de empleo para personas en situación de discapacidad, prevista en el numeral 17 del artículo 57 del Código Sustantivo del Trabajo.”

En consecuencia y a la luz de esta nueva norma, contratar aprendices con discapacidad no sirve para aumentar el número de personas con discapacidad computables dentro de la cuota de empleo exigida, por lo cual la cuota se calculará ahora sobre la base de trabajadores permanentes, y el decreto 0223 excluye a los aprendices de esa base.

Sin embargo, el Decreto genera un incentivo distinto en el numeral 2 del mismo artículo, donde establece que:

“La cuota de aprendices se reducirá en un 50% si las personas contratadas tienen una discapacidad comprobada no inferior al 25%”, en cumplimiento del parágrafo del artículo 31 de la Ley 361 de 1997.

Es decir, que sí es posible contratar aprendices con discapacidad, pero el efecto jurídico directo es sobre la cuota de aprendices (Ley 789 de 2002), no sobre la cuota de empleo para personas con discapacidad (Ley 2466 de 2025 art. 57 num. 17 CST).
""".strip(),
    "casos_alcance_cierre": """
Durante la reunión, se socializaron casos exitosos de inclusión y el apoyo de interprete lengua de señas en el caso de vincular personas con discapacidad auditiva.

Se reitera que el alcance operativo de la ruta de inclusión abarca Bogotá y Cundinamarca.

Se agradece espacio, se informa envío de presentación y se estará a la espera de contacto para dar continuidad a la ruta.

Se finaliza reunión sin novedad
""".strip(),
}

RUTA_INCLUSION_TEMPLATE_BUTTONS = [
    ("objetivo_y_participantes", "Objetivo y participantes"),
    ("roles_proceso_seleccion", "Roles y selección"),
    ("certificado_discapacidad", "Certificado"),
    ("acompanamiento_reca", "Acompañamiento RECA"),
    ("seguimiento_y_normativa", "Seguimiento y normativa"),
    ("casos_alcance_cierre", "Casos, alcance y cierre"),
]

SECTION_5 = {
    "title": "5. ASISTENTES",
    "max_items": 10,
    "items": [
        {"id": "asistente_1", "nombre": "", "cargo": ""},
        {"id": "asistente_2", "nombre": "", "cargo": ""},
        {"id": "asistente_3", "nombre": "", "cargo": ""},
    ],
}

SECTION_1_SUPABASE_MAP = {
    "nombre_empresa": "nombre_empresa",
    "direccion_empresa": "direccion_empresa",
    "correo_1": "correo_1",
    "contacto_empresa": "contacto_empresa",
    "caja_compensacion": "caja_compensacion",
    "profesional_asignado": "profesional_asignado",
    "asesor": "asesor",
    "ciudad_empresa": "ciudad_empresa",
    "telefono_empresa": "telefono_empresa",
    "cargo": "cargo",
    "sede_empresa": "zona_empresa",
    "correo_profesional": "correo_profesional",
    "correo_asesor": "correo_asesor",
}

FORM_CACHE = {}
SECTION_1_CACHE = {}


def _map_company_row(row):
    if not isinstance(row, dict):
        return row
    mapped = dict(row)
    for field_id, source_key in SECTION_1_SUPABASE_MAP.items():
        if source_key in row:
            mapped[field_id] = row.get(source_key)
    return mapped


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
    return os.path.join(_get_cache_dir(), "presentacion_programa.json")


def cache_file_exists():
    return os.path.exists(_get_cache_path())


def save_cache_to_file():
    payload = {
        "form_id": "presentacion_programa",
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

EXCEL_MAPPING = {
    "section_1": {
        "fecha_visita": "D7",
        "modalidad": "Q7",
        "nit_empresa": "Q9",
        "nombre_empresa": "D8",
        "direccion_empresa": "D9",
        "correo_1": "D10",
        "contacto_empresa": "D11",
        "caja_compensacion": "D12",
        "profesional_asignado": "D13",
        "asesor": "D14",
        "ciudad_empresa": "Q8",
        "telefono_empresa": "Q10",
        "cargo": "Q11",
        "sede_empresa": "Q12",
        "correo_profesional": "Q13",
        "correo_asesor": "Q14",
    },
    "section_3_item_8": {
        "Responsabilidad Social Empresarial": "U38",
        "Objetivos y metas para la diversidad, equidad e inclusion.": "U39",
        "Avances a nivel global de impacto en Colombia": "U40",
        "Beneficios Tributarios": "U41",
        "Beneficios en la contratacion de poblacion en riesgo de exclusion": "U42",
        "Ventaja en licitaciones publicas": "U43",
        "Cumplimiento de la normativa establecida por el Estado Colombiano.": "U44",
        "Experiencia en la vinculacion de personas en condicion de discapacidad.": "U45",
    },
    "section_4": {
        "acuerdos_observaciones": "A49",
    },
    "section_5": {
        "start_row": 53,
        "name_col": "C",
        "cargo_col": "N",
    },
}


def get_empresa_by_nit(nit, env_path=".env"):
    """
    Busca empresa por NIT en Supabase (solo lectura).
    Retorna un dict con las columnas mapeadas o None si no existe.
    """
    if not nit:
        return None
    nit = "".join(str(nit).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {
        "select": select_cols,
        "nit_empresa": f"eq.{nit}",
        "limit": 1,
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    return _map_company_row(data[0]) if data else None


def get_empresa_by_nombre(nombre, env_path=".env"):
    """
    Busca empresa por nombre exacto (case-insensitive) en Supabase (solo lectura).
    Retorna un dict con las columnas mapeadas o None si no existe.
    """
    if not nombre:
        return None
    nombre = " ".join(str(nombre).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {
        "select": select_cols,
        "nombre_empresa": f"ilike.{nombre}",
        "limit": 2,
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    if not data:
        return None
    if len(data) > 1:
        raise ValueError("Hay más de una empresa con ese nombre. Usa el NIT.")
    return _map_company_row(data[0])


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=50):
    if not prefix:
        return []
    prefix = " ".join(str(prefix).split())
    if not prefix:
        return []
    params = {
        "select": "nombre_empresa",
        "nombre_empresa": f"ilike.{prefix}%",
        "limit": max(1, int(limit)),
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    if not data:
        return []
    names = []
    seen = set()
    for row in data:
        name = row.get("nombre_empresa")
        if not name:
            continue
        if name in seen:
            continue
        seen.add(name)
        names.append(name)
    return names


def confirm_section_1(company_data, user_inputs):
    """
    Confirma y guarda en cache la seccion 1 para uso posterior (Excel).
    """
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


def confirm_section_3_item8(checkbox_values):
    """
    Guarda en cache la motivacion (item 8) con valores de checkboxes.
    """
    if checkbox_values is None:
        raise ValueError("checkbox_values requerido")
    defaults = SECTION_3["items"][7]["content"]
    unexpected_keys = set(checkbox_values.keys()) - set(defaults.keys())
    if unexpected_keys:
        raise ValueError(f"Claves no permitidas: {sorted(unexpected_keys)}")
    payload = {}
    for key, default_value in defaults.items():
        payload[key] = bool(checkbox_values.get(key, default_value))
    set_section_cache("section_3_item_8", payload)
    FORM_CACHE["_last_section"] = "section_3_item_8"
    save_cache_to_file()
    return payload


def confirm_section_4(notes_text):
    """
    Guarda en cache las notas de la seccion 4.
    """
    if notes_text is None:
        raise ValueError("notes_text requerido")
    payload = {"acuerdos_observaciones": str(notes_text)}
    set_section_cache("section_4", payload)
    FORM_CACHE["_last_section"] = "section_4"
    save_cache_to_file()
    return payload


def confirm_section_5(asistentes):
    """
    Guarda en cache los asistentes (hasta 3).
    """
    if asistentes is None:
        raise ValueError("asistentes requerido")
    max_items = SECTION_5.get("max_items", 10)
    if len(asistentes) > max_items:
        raise ValueError(f"Máximo {max_items} asistentes")

    def _normalize_text(value):
        parts = str(value).strip().split()
        return " ".join(part[:1].upper() + part[1:].lower() for part in parts)

    payload = []
    has_nombre = False
    for item in asistentes:
        if not isinstance(item, dict):
            raise ValueError("Cada asistente debe ser un dict")
        nombre = _normalize_text(item.get("nombre", ""))
        cargo = _normalize_text(item.get("cargo", ""))
        if nombre:
            has_nombre = True
        payload.append({"nombre": nombre, "cargo": cargo})
    if not has_nombre:
        raise ValueError("Debe ingresar al menos un nombre de asistente")
    set_section_cache("section_5", payload)
    FORM_CACHE["_last_section"] = "section_5"
    save_cache_to_file()
    return payload


def set_section_cache(section_id, payload):
    if not section_id:
        raise ValueError("section_id requerido")
    if payload is None:
        payload = {}
    FORM_CACHE[section_id] = payload


def get_form_cache():
    return dict(FORM_CACHE)


def _get_log_dir(output_path=None):
    if output_path:
        base_dir = os.path.dirname(output_path)
    else:
        base_dir = os.getcwd()
    log_dir = os.path.join(base_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    return log_dir


def _log_excel(message, output_path=None):
    try:
        _ = output_path
        log_excel_event(message)
    except Exception:
        return


def _find_template_path(tipo_visita=None):
    templates_dir = resource_path("templates")
    if not templates_dir.is_dir():
        raise FileNotFoundError("No existe la carpeta templates.")
    visit_type = (tipo_visita or "").strip().lower()
    keyword = "reactivacion" if visit_type == "reactivacion" else "presentacion"
    for name in os.listdir(templates_dir):
        normalized = _normalize_text(name)
        if keyword in normalized and normalized.endswith(".xlsx"):
            return os.fspath(templates_dir / name)
    raise FileNotFoundError("No se encontró el template correspondiente.")


def export_to_excel(cache=None):
    clear_written_rows()
    if cache is None:
        cache = FORM_CACHE
    section_1 = cache.get("section_1") or {}
    section_3_item_8 = cache.get("section_3_item_8") or {}
    section_4 = cache.get("section_4") or {}
    section_5 = cache.get("section_5") or []

    tipo_visita_raw = section_1.get("tipo_visita") or "Presentacion"
    tipo_visita = _normalize_text(tipo_visita_raw)
    template_path = _find_template_path(tipo_visita=tipo_visita)

    empresa_nombre = section_1.get("nombre_empresa") or "Empresa"

    prefix = "Reactivacion" if tipo_visita == "reactivacion" else "Presentacion"
    process_name = f"{prefix} del Programa de Inclusion Laboral"
    output_path = _build_process_output_path(empresa_nombre, process_name)
    shutil.copy2(template_path, output_path)
    _log_excel(f"START export_all output={output_path}", output_path)

    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("pywin32 no está instalado. Instala con pip install pywin32.") from exc

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(output_path)
        ws = wb.Worksheets(1)

        for key, cell in EXCEL_MAPPING["section_1"].items():
            if key in section_1:
                value = section_1.get(key)
                _log_excel(
                    f"WRITE section=section_1 cell={cell} key={key} value={value!r}",
                    output_path,
                )
                ws_write(ws, cell, value)

        normalized_section_3_item_8 = {
            _normalize_text(key): bool(value)
            for key, value in section_3_item_8.items()
        }
        for key, cell in EXCEL_MAPPING["section_3_item_8"].items():
            value = normalized_section_3_item_8.get(
                _normalize_text(key),
                bool(section_3_item_8.get(key, False)),
            )
            symbol = format_checkbox_symbol(value)
            _log_excel(
                f"WRITE section=section_3_item_8 cell={cell} key={key} "
                f"checkbox_value={value!r} checkbox_symbol={symbol!r}",
                output_path,
            )
            write_checkbox_symbol(ws, cell, value)

        for key, cell in EXCEL_MAPPING["section_4"].items():
            if key in section_4:
                value = section_4.get(key)
                _log_excel(
                    f"WRITE section=section_4 cell={cell} key={key} value={value!r}",
                    output_path,
                )
                ws_write(ws, cell, value)

        section_5_cfg = EXCEL_MAPPING["section_5"]
        start_row = section_5_cfg["start_row"]
        name_col = section_5_cfg["name_col"]
        cargo_col = section_5_cfg["cargo_col"]
        total = len(section_5)
        template_row = start_row + 2
        if total > 3:
            insert_at = start_row + 3
            extra_rows = total - 3
            for _ in range(extra_rows):
                ws.Rows(insert_at).Insert()
                ws.Rows(template_row).Copy()
                ws.Rows(insert_at).PasteSpecial(-4122)
                insert_at += 1

        for idx in range(total):
            row = start_row + idx
            entry = section_5[idx]
            nombre = entry.get("nombre", "")
            cargo = entry.get("cargo", "")
            _log_excel(
                f"WRITE section=section_5 cell={name_col}{row} key=nombre value={nombre!r}",
                output_path,
            )
            _log_excel(
                f"WRITE section=section_5 cell={cargo_col}{row} key=cargo value={cargo!r}",
                output_path,
            )
            ws_write(ws, f"{name_col}{row}", nombre)
            ws_write(ws, f"{cargo_col}{row}", cargo)
            if idx >= 3:
                ws_write(ws, f"A{row}", "Nombre completo:")
                ws_write(ws, f"L{row}", "Cargo:")

        sanitize_logo_error_cells(wb)
        autofit_rows(ws, log_fn=lambda msg: _log_excel(msg, output_path))
        wb.Save()
        _log_excel("SUCCESS export_all", output_path)
    except Exception as exc:
        _log_excel(f"ERROR export_all error={exc!r}", output_path)
        raise
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
    clear_cache_file()
    clear_form_cache()
    return output_path


def register_form():
    """Return metadata for hub registration."""
    return {
        "id": "presentacion_programa",
        "name": FORM_NAME,
        "module": __name__,
    }


