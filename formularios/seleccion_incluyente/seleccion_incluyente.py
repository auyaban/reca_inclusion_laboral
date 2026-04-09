"""
formularios/seleccion_incluyente/seleccion_incluyente.py
Formulario: "4. SELECCIÓN INCLUYENTE"

Responsabilidades:
  - Mapeo de campos a celdas del master spreadsheet (hoja 4)
  - Secciones 1, 2, 5, 6: datos empresa, proceso de selección, candidatos
    evaluados (dinámica), compromisos
  - Manejo de candidatos múltiples: sección 2 puede tener N candidatos;
    cada uno genera filas adicionales en la hoja
  - Vincula con evaluacion_accesibilidad para reutilizar datos de la empresa

Entry points para app.py:
  confirm_section_1(company_data, user_inputs)
  confirm_section_2 / confirm_section_5 / confirm_section_6(payload)
  validate_before_finalize()   → retorna lista de ValidationIssue
  export_to_excel()            → escribe en Google Sheets y sube a Drive
  register_form()              → metadata para HubWindow

Depende de: google_sheets_client, formularios/common, formularios/finalize_validation,
            formularios/evaluacion_programa (para datos de empresa)
"""
import copy
import json
import os
import time
from functools import lru_cache

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _get_cached_payload,
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
    is_meaningful,
    raise_validation_error,
    require_value,
    validate_dynamic_rows,
)
from logging_utils import log_excel_event

FORM_ID = "seleccion_incluyente"
FORM_NAME = "Proceso de Seleccion Incluyente"
_USUARIOS_RECA_CEDULAS_CACHE_TTL_SECONDS = 86400
SECTION_HISTORY_LIMIT = 10

SECTION_1 = {
    "title": "1. DATOS DE LA EMPRESA",
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
        {
            "id": "nit_empresa",
            "label": "Número de NIT",
            "source": "input",
        },
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

FORM_CACHE = {}
SECTION_1_CACHE = {}


SHEET_NAME = "4. SELECCIÓN INCLUYENTE"

TEMPLATE_VARIANT_INDIVIDUAL = "individual"
TEMPLATE_VARIANT_GROUP_2_PLUS = "group_2_plus"

# Oferente block geometry (unified format — one sheet for individual & group)
OFERENTE_BLOCK_HEIGHT = 61          # rows per oferente block (rows 16-76 for first)
OFERENTE_FIRST_BLOCK_START_ROW = 16
OFERENTE_SECOND_BLOCK_START_ROW = OFERENTE_FIRST_BLOCK_START_ROW + OFERENTE_BLOCK_HEIGHT
DESARROLLO_ACTIVIDAD_CELL = "A14"   # shared across all oferentes
GROUP_EXPORT_TITLE_CELL = "G1"
OFERENTE_TITLE_COL = "A"
SECTION_2_LAST_COLUMN = "U"

# Base row positions for 1 oferente (shift by (N-1)*BLOCK_HEIGHT for N oferentes)
SECTION_5_BASE_AJUSTES_ROW = 78     # ajustes text row
SECTION_5_BASE_NOTA_ROW = 79        # nota row
SECTION_6_BASE_START_ROW = 84       # first asistente data row
SECTION_5_TITLE_ROW_BY_TEMPLATE = {
    TEMPLATE_VARIANT_INDIVIDUAL: 77,
    TEMPLATE_VARIANT_GROUP_2_PLUS: 77,
}
SECTION_6_TITLE_ROW_BY_TEMPLATE = {
    TEMPLATE_VARIANT_INDIVIDUAL: 83,
    TEMPLATE_VARIANT_GROUP_2_PLUS: 83,
}


def ws_write(ws, cell, value):
    try:
        ws[cell] = value
    except Exception:
        return

AJUSTES_ENTREVISTA_TEMPLATES = {
    "preparacion_proceso": """
Ajustes razonables para entrevista.

Contactar a la Agencia de Empleo para el apoyo de adaptación y organización de las pruebas psicotécnicas de la empresa.

Realizar una organización de la entrevista de forma que se pueda anticipar a los oferentes el paso a paso de lo que se realizará en el proceso, permitiendo que logren organizar su tiempo y la constancia que requiere el mismo.

Promover la aplicación de formatos para la hoja de vida con diseños sencillos, fáciles de diligenciar y a través de medios accesibles; estos pueden ser virtuales o físicos.

Para mantener un contacto directo con los oferentes que han sido preseleccionados, se recomienda evaluar alternativas hasta identificar el mejor canal de comunicación: contacto telefónico, mensajes de texto, mensajes por chat, correo electrónico y, en algunos casos, a través de familiares que hayan sido referenciados en la hoja de vida. Los mensajes deben ser precisos y concretos, incluyendo solamente la información básica y utilizando frases simples.
""".strip(),
    "trato_respetuoso": """
Evitar conductas, palabras, frases, sentimientos, preconcepciones y estigmas que impidan u obstaculicen el acceso en igualdad de condiciones de las personas con y/o en situación de discapacidad a los espacios, objetos, servicios y, en general, a las posibilidades que ofrece la sociedad.

Evitar, durante el proceso de entrevista, preguntar sobre cómo fue la adquisición de la discapacidad.

Informar a los candidatos que no fueron seleccionados sobre la finalización del proceso, basándose en sus habilidades residuales individuales. De esta manera no tendrán que esperar innecesariamente y se respeta su tiempo.

La evaluación de desempeño debe ajustarse de acuerdo con el perfil del cargo y no enfocarse en la discapacidad de la persona.

Mantener contacto visual con la persona, mirando a sus ojos y evitando enfocar su discapacidad.
""".strip(),
    "accesibilidad_entrevista": """
Asegurarse de que el lugar de la entrevista sea accesible para personas con discapacidad física. Esto incluye la disponibilidad de rampas, ascensores o espacios adecuados para sillas de ruedas, si es necesario.

Si el candidato tiene dificultades de comunicación, ofrecer alternativas como permitir el uso de comunicación por texto, proporcionar un intérprete de lengua de señas si es necesario o permitir que el candidato responda por escrito.

Considerar la posibilidad de ofrecer tiempo adicional para completar la entrevista si la discapacidad del candidato afecta su velocidad de procesamiento o comunicación.

Ser flexible en cuanto al formato de la entrevista. Algunas personas pueden necesitar entrevistas en un formato diferente, como entrevistas virtuales o en un entorno menos estimulante para quienes presentan sensibilidad sensorial.

Formular preguntas claras y directas y ser paciente al esperar la respuesta del candidato. Evitar jergas o frases complicadas que puedan ser difíciles de entender.

Al proporcionar retroalimentación al candidato, ser claro, conciso y constructivo. Destacar sus fortalezas y ofrecer sugerencias de mejora, si es necesario, de manera útil y respetuosa.

Realizar preguntas orientadoras que permitan identificar que la información está siendo recibida correctamente por el oferente durante el proceso de selección.
""".strip(),
    "accesibilidad_documentos": """
Aumentar el tamaño de letra, utilizar colores de fondo en el texto y emplear fuentes de fácil lectura como Arial o Verdana. Adicionalmente, proporcionar esos documentos de manera virtual para facilitar el uso de herramientas tiflotecnológicas.
""".strip(),
    "pruebas_seleccion": """
La aplicación de pruebas de tipo visual y gráfico, como por ejemplo Test de Percepción Temática, Wartegg, Técnica de dibujo proyectivo HTP por sus siglas en inglés (Casa, Árbol, Persona) y Test de la Figura Humana, no es recomendable para personas con discapacidad visual.

Para los procesos de selección en los que participen candidatos usuarios de lengua de señas, es clave contar con un servicio de interpretación profesional y no apoyarse en amigos o familiares de los candidatos que tengan conocimiento.

No se recomienda hacer interpretación en lengua de señas de las pruebas psicotécnicas, dado que se puede sesgar la información que se espera recoger y, por ende, los resultados. Es preferible reemplazar este tipo de pruebas por entrevistas por competencias.

En el caso de personas con discapacidad que cuentan con familias sobreprotectoras, se deben establecer límites claros con estas, restringiendo su participación durante el proceso de selección.

Se recomienda minimizar el uso de pruebas psicotécnicas y reemplazarlas por otras estrategias de selección que permitan alcanzar los mismos objetivos. Solo en caso de que lo anterior no sea posible, se sugiere priorizar la aplicación de pruebas gráficas proyectivas que buscan identificar rasgos de personalidad.
""".strip(),
}

AJUSTES_ENTREVISTA_TEMPLATE_BUTTONS = [
    ("preparacion_proceso", "Preparación del proceso"),
    ("trato_respetuoso", "Trato respetuoso"),
    ("accesibilidad_entrevista", "Accesibilidad entrevista"),
    ("accesibilidad_documentos", "Accesibilidad documentos"),
    ("pruebas_seleccion", "Pruebas de selección"),
]

SECTION_2 = {
    "title": "2. DATOS DEL OFERENTE",
    "fields": [
        {"id": "numero", "label": "No", "type": "texto"},
        {"id": "nombre_oferente", "label": "Nombre oferente", "type": "texto"},
        {"id": "cedula", "label": "Cédula", "type": "texto"},
        {"id": "certificado_porcentaje", "label": "Certificado %", "type": "texto"},
        {
            "id": "discapacidad",
            "label": "Discapacidad",
            "type": "lista",
            "options": [
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
            ],
        },
        {"id": "telefono_oferente", "label": "Teléfono oferente", "type": "texto"},
        {
            "id": "resultado_certificado",
            "label": "Resultado",
            "type": "lista",
            "options": ["Aprobado", "No aprobado", "Pendiente"],
        },
        {"id": "cargo_oferente", "label": "Cargo oferente", "type": "texto"},
        {
            "id": "nombre_contacto_emergencia",
            "label": "Nombre contacto emergencia",
            "type": "texto",
        },
        {"id": "parentesco", "label": "Parentesco", "type": "texto"},
        {"id": "telefono_emergencia", "label": "Teléfono", "type": "texto"},
        {"id": "fecha_nacimiento", "label": "Fecha de nacimiento", "type": "texto"},
        {"id": "edad", "label": "Edad", "type": "texto"},
        {
            "id": "pendiente_otros_oferentes",
            "label": "Pendiente otros oferentes",
            "type": "lista",
            "options": ["Si", "No", "Por Confirmar"],
        },
        {
            "id": "lugar_firma_contrato",
            "label": "Lugar firma de contrato",
            "type": "texto",
        },
        {
            "id": "fecha_firma_contrato",
            "label": "Fecha firma de contrato",
            "type": "texto",
        },
        {
            "id": "cuenta_pension",
            "label": "Cuenta con pension",
            "type": "lista",
            "options": ["Si", "No", "Por Confirmar"],
        },
        {
            "id": "tipo_pension",
            "label": "Tipo de pension",
            "type": "lista",
            "options": [
                "Pension Invalidez",
                "Subsidiada",
                "Especial de vejez",
                "Victimas conflicto",
                "Familiar",
                "Regimen especial",
                "No aplica",
            ],
        },
        {
            "id": "desarrollo_actividad",
            "label": "Desarrollo de la actividad",
            "type": "texto_largo",
        },
        {
            "id": "medicamentos_nivel_apoyo",
            "label": "Toma medicamentos - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "medicamentos_conocimiento",
            "label": "Toma medicamentos - Conocimiento de medicamentos",
            "type": "lista",
            "options": [
                "1. Conoce los medicamentos que consume.",
                "2. Un tercero es quien conoce los medicamentos que consume.",
                "3. No conoce los medicamentos que consume.",
                "No aplica.",
                "0. No requiere apoyo.",
            ],
        },
        {
            "id": "medicamentos_horarios",
            "label": "Toma medicamentos - Conocimiento de horarios",
            "type": "lista",
            "options": [
                "1. Conoce los horarios de toma de medicamentos que consume.",
                "2. Es un tercero quien conoce los horarios de la toma de medicamentos.",
                "3. No conoce los horarios de toma de medicamentos que consume.",
                "0. No requiere apoyo.",
                "No aplica.",
            ],
        },
        {
            "id": "medicamentos_nota",
            "label": "Toma medicamentos - Nota",
            "type": "texto",
        },
        {
            "id": "alergias_nivel_apoyo",
            "label": "Presenta alergia - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "alergias_tipo",
            "label": "Presenta alergia - Tipo de alergia",
            "type": "lista",
            "options": [
                "0. No presenta alergias.",
                "1. Presenta alergias y sabe darle manejo.",
                "2. No conoce si presenta alguna alergia.",
                "3. Presenta alergias a: medicamentos, sustancias y productos quimicos, alimentos, animales, entre otros.",
                "No aplica.",
            ],
        },
        {
            "id": "alergias_nota",
            "label": "Presenta alergia - Nota",
            "type": "texto",
        },
        {
            "id": "restriccion_nivel_apoyo",
            "label": "Tiene restriccion medica - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "restriccion_conocimiento",
            "label": "Tiene restriccion medica - Conocimiento",
            "type": "lista",
            "options": [
                "0. No tiene restricciones medicas.",
                "1. Tiene restricciones medicas y conoce su manejo.",
                "2. No conoce si tiene restricciones medicas.",
                "3. Si tiene restricciones medicas y desconoce su manejo.",
                "No aplica.",
            ],
        },
        {
            "id": "restriccion_nota",
            "label": "Tiene restriccion medica - Nota",
            "type": "texto",
        },
        {
            "id": "controles_nivel_apoyo",
            "label": "Asiste a controles medicos - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "controles_asistencia",
            "label": "Asiste a controles medicos - Asistencia a controles",
            "type": "lista",
            "options": [
                "No aplica.",
                "2. Si asiste a controles medicos con especialista.",
                "3. No sabe si tiene controles medicos con especialista.",
                "1. Asiste a controles medicos con especialista y conoce el manejo.",
                "0. No requiere apoyo.",
            ],
        },
        {
            "id": "controles_frecuencia",
            "label": "Asiste a controles medicos - Frecuencia",
            "type": "lista",
            "options": [
                "Mensual",
                "Trimestral",
                "Semestral",
                "Otra frecuencia",
                "No aplica",
            ],
        },
        {
            "id": "controles_nota",
            "label": "Asiste a controles medicos - Nota",
            "type": "texto",
        },
        {
            "id": "desplazamiento_nivel_apoyo",
            "label": "Desplazamiento independiente - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "desplazamiento_modo",
            "label": "Desplazamiento independiente - Modo de desplazamiento",
            "type": "lista",
            "options": [
                "0. Se desplaza de manera independiente sin necesidad de apoyos (ortesis, baston, silla de ruedas entre otros).",
                "1. Se desplaza de forma independiente con un apoyo temporal (ortesis, baston, silla de ruedas entre otros).",
                "2. Se desplaza de manera independiente con un apoyo permanente (ortesis, baston, silla de ruedas entre otros).",
                "3. No se desplaza de manera independiente. Requiere el acompanamiento de un tercero y un apoyo (ortesis, baston, silla de ruedas entre otros).",
                "No aplica.",
            ],
        },
        {
            "id": "desplazamiento_transporte",
            "label": "Desplazamiento independiente - Medio de transporte",
            "type": "lista",
            "options": [
                "Caminando.",
                "Bicicleta.",
                "Transmilenio, Sitp.",
                "Vehiculo propio.",
                "Vehiculo especial.",
                "No aplica.",
            ],
        },
        {
            "id": "desplazamiento_nota",
            "label": "Desplazamiento independiente - Nota",
            "type": "texto",
        },
        {
            "id": "ubicacion_nivel_apoyo",
            "label": "Ubicacion en la ciudad - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "ubicacion_ciudad",
            "label": "Ubicacion en la ciudad",
            "type": "lista",
            "options": [
                "0. Sabe ubicarse en la ciudad de manera autonoma.",
                "1. Sabe ubicarse en la ciudad pero haciendo uso de aplicaciones (Maps, Waze, entre otros).",
                "2. Requiere de acompanamiento para ubicarse.",
                "3. No sabe ubicarse en la ciudad.",
            ],
        },
        {
            "id": "ubicacion_aplicaciones",
            "label": "Manejo de aplicaciones",
            "type": "lista",
            "options": [
                "Se ubica por puntos de referencia y direcciones.",
                "No se ubica por puntos de referencia.",
                "Se ubica por puntos cardinales.",
                "No aplica",
            ],
        },
        {
            "id": "ubicacion_nota",
            "label": "Ubicacion en la ciudad - Nota",
            "type": "texto",
        },
        {
            "id": "dinero_nivel_apoyo",
            "label": "Reconoce y maneja el dinero - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "dinero_reconocimiento",
            "label": "Reconocimiento del dinero",
            "type": "lista",
            "options": ["Autonomo.", "Con apoyo familiar."],
        },
        {
            "id": "dinero_manejo",
            "label": "Manejo del dinero",
            "type": "lista",
            "options": [
                "0. Reconoce y maneja el dinero de manera autonoma.",
                "1. Reconoce y maneja el dinero pero en ocasiones requiere apoyo.",
                "2. Solo reconoce el dinero pero no lo sabe manejar.",
                "3. No reconoce el dinero y no lo sabe manejar.",
                "No aplica.",
            ],
        },
        {
            "id": "dinero_medios",
            "label": "Uso de medios electronicos",
            "type": "lista",
            "options": [
                "Dinero fisico, plastico y digital.",
                "Dinero fisico y plastico.",
                "Dinero fisico.",
                "Dinero plastico y digital.",
                "Dinero plastico.",
                "Dinero digital.",
                "Dinero digital y fisico.",
            ],
        },
        {
            "id": "dinero_nota",
            "label": "Reconoce y maneja el dinero - Nota",
            "type": "texto",
        },
        {
            "id": "presentacion_nivel_apoyo",
            "label": "Presentacion personal - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "presentacion_personal",
            "label": "Presentacion personal",
            "type": "lista",
            "options": [
                "0. Su codigo de vestuario es acorde al contexto.",
                "1. Su codigo de vestuario es acorde al contexto, pero presenta oportunidades de mejora.",
                "2. Su codigo de vestuario es medianamente acorde al contexto.",
                "3. Su codigo de vestuario no es acorde al contexto.",
                "No aplica.",
            ],
        },
        {
            "id": "presentacion_nota",
            "label": "Presentacion personal - Nota",
            "type": "texto",
        },
        {
            "id": "comunicacion_escrita_nivel_apoyo",
            "label": "Apoyo comunicacion escrita - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "comunicacion_escrita_apoyo",
            "label": "Apoyo comunicacion escrita",
            "type": "lista",
            "options": [
                "0. Si conoce y maneja los apoyos (Jaws, Magic, el lector de pantalla de Windows/IOS).",
                "1. Maneja algunos apoyos de comunicacion escrita, pero no todos en general.",
                "2. Conoce pero no maneja apoyos.",
                "3. No conoce, ni maneja los apoyos.",
                "No aplica.",
            ],
        },
        {
            "id": "comunicacion_escrita_nota",
            "label": "Apoyo comunicacion escrita - Nota",
            "type": "texto",
        },
        {
            "id": "comunicacion_verbal_nivel_apoyo",
            "label": "Apoyo comunicacion verbal - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "comunicacion_verbal_apoyo",
            "label": "Apoyo comunicacion verbal",
            "type": "lista",
            "options": [
                "0. Si conoce y maneja los apoyos (Centro de relevo, entre otros).",
                "1. Maneja algunos apoyos, pero no los conoce todos en general (Centro de relevo, entre otros).",
                "2. Conoce pero no maneja apoyos.",
                "3. No conoce, ni maneja los apoyos.",
                "No aplica.",
            ],
        },
        {
            "id": "comunicacion_verbal_nota",
            "label": "Apoyo comunicacion verbal - Nota",
            "type": "texto",
        },
        {
            "id": "decisiones_nivel_apoyo",
            "label": "Toma de decisiones - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "toma_decisiones",
            "label": "Toma de decisiones",
            "type": "lista",
            "options": [
                "0. Toma las decisiones de manera autonoma.",
                "1. Toma decisiones pero en ocasiones requiere el apoyo de un tercero.",
                "2. Debe consultar con un tercero para la toma de decisiones.",
                "3. Requiere el apoyo de un tercero para tomar decisiones.",
                "No aplica.",
            ],
        },
        {
            "id": "toma_decisiones_nota",
            "label": "Toma de decisiones - Nota",
            "type": "texto",
        },
        {
            "id": "aseo_nivel_apoyo",
            "label": "Apoyo en aseo personal - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "alimentacion",
            "label": "Alimentacion",
            "type": "lista",
            "options": [
                "0. No requiere apoyo en sus actividades de la vida diaria.",
                "1. Requiere apoyo en algunas actividades de la vida diaria.",
                "2. Requiere apoyo en la mayoria de actividades de la vida diaria.",
                "3. Requiere apoyo en todas las actividades de la vida diaria.",
                "No aplica.",
            ],
        },
        {
            "id": "aseo_criar_apoyo",
            "label": "Criar y cuidado de ninos - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "aseo_comunicacion_apoyo",
            "label": "Uso de sistemas de comunicacion - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "aseo_ayudas_apoyo",
            "label": "Cuidado de ayudas tecnicas - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "aseo_alimentacion",
            "label": "Alimentacion",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "aseo_movilidad_funcional",
            "label": "Movilidad funcional",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "aseo_higiene_aseo",
            "label": "Higiene personal y aseo (Control de esfinter)",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "aseo_nota",
            "label": "Apoyo en aseo personal - Nota",
            "type": "texto",
        },
        {
            "id": "instrumentales_nivel_apoyo",
            "label": "Apoyo en actividades instrumentales - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "instrumentales_actividades",
            "label": "Actividades instrumentales",
            "type": "lista",
            "options": [
                "0. No requiere apoyo en actividades instrumentales de la vida diaria.",
                "1. Requiere apoyo en algunas actividades instrumentales de la vida diaria.",
                "2. Requiere apoyo en la mayoria de actividades instrumentales de la vida diaria.",
                "3. Requiere apoyo en todas las actividades instrumentales de la vida diaria.",
                "No aplica.",
            ],
        },
        {
            "id": "instrumentales_criar_apoyo",
            "label": "Criar y cuidado de ninos - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "instrumentales_comunicacion_apoyo",
            "label": "Uso de sistemas de comunicacion - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "instrumentales_movilidad_apoyo",
            "label": "Movilidad en la comunidad - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "instrumentales_finanzas",
            "label": "Manejo de tematicas financieras",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "instrumentales_cocina_limpieza",
            "label": "Cocina y limpieza",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "instrumentales_crear_hogar",
            "label": "Crear y mantener un hogar",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "instrumentales_salud_cuenta_apoyo",
            "label": "Cuidado de salud y manutencion - Cuenta con apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "instrumentales_nota",
            "label": "Apoyo en actividades instrumentales - Nota",
            "type": "texto",
        },
        {
            "id": "actividades_nivel_apoyo",
            "label": "Apoyo durante actividades - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "actividades_apoyo",
            "label": "Apoyo durante actividades",
            "type": "lista",
            "options": [
                "0. No requiere apoyo en sus actividades laborales.",
                "1. Requiere apoyo en algunas actividades laborales.",
                "2. Requiere apoyo en la mayoria de actividades laborales.",
                "3. Requiere apoyo en todas las actividades laborales.",
                "No aplica",
            ],
        },
        {
            "id": "actividades_esparcimiento_apoyo",
            "label": "Actividades de esparcimiento con familia - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "actividades_esparcimiento_cuenta_apoyo",
            "label": "Actividades de esparcimiento con familia - Cuenta con apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "actividades_complementarios_apoyo",
            "label": "Complementarios medicos - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "actividades_complementarios_cuenta_apoyo",
            "label": "Complementarios medicos - Cuenta con apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "actividades_subsidios_cuenta_apoyo",
            "label": "Subsidios economicos para estudio de hijos - Cuenta con apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "actividades_nota",
            "label": "Apoyo durante actividades - Nota",
            "type": "texto",
        },
        {
            "id": "discriminacion_nivel_apoyo",
            "label": "Discriminacion - Nivel de apoyo",
            "type": "lista",
            "options": [
                "0. No requiere apoyo.",
                "1. Nivel de apoyo Bajo.",
                "2. Nivel de apoyo medio.",
                "3. Nivel de apoyo alto.",
                "No aplica.",
            ],
        },
        {
            "id": "discriminacion",
            "label": "Discriminacion",
            "type": "lista",
            "options": [
                "0. No ha sufrido de discriminacion.",
                "1. Ha sufrido de discriminacion en algunos contextos.",
                "2. Ha sufrido de discriminacion en repetidas ocasiones.",
                "3. Ha sufrido de discriminacion a los largo del ciclo vital.",
                "No aplica.",
            ],
        },
        {
            "id": "discriminacion_violencia_apoyo",
            "label": "Violencia fisica - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "discriminacion_violencia_cuenta_apoyo",
            "label": "Violencia fisica - Cuenta con apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "discriminacion_vulneracion_apoyo",
            "label": "Vulneracion de derechos - Requiere apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "discriminacion_vulneracion_cuenta_apoyo",
            "label": "Vulneracion de derechos - Cuenta con apoyo",
            "type": "lista",
            "options": ["Si", "No", "No aplica"],
        },
        {
            "id": "discriminacion_nota",
            "label": "Discriminacion - Nota",
            "type": "texto",
        },
    ],
}

SECTION_5 = {
    "title": "5. AJUSTES RAZONABLES / RECOMENDACIONES AL PROCESO DE SELECCION",
    "fields": [
        {"id": "ajustes_recomendaciones", "label": "Ajustes razonables", "type": "texto_largo"},
        {"id": "nota", "label": "Nota", "type": "texto"},
    ],
}

SECTION_6 = {
    "title": "6. ASISTENTES",
    "rows": 2,
}

SECTION_1_FIELD_MAP = {field["id"]: field for field in SECTION_1["fields"]}
SECTION_2_FIELD_MAP = {field["id"]: field for field in SECTION_2["fields"]}
LIST_FIELD_OPTIONS_BY_ID = {
    field["id"]: list(field.get("options", []))
    for field in SECTION_1["fields"] + SECTION_2["fields"]
    if field.get("type") == "lista"
}

SECTION_1_CELL_MAP = {
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
    "asesor": "F12",
    "sede_empresa": "N12",
}

# Cell map for one oferente block.  (col, row) tuples — row is 1-indexed.
# For oferente N, add (N-1)*OFERENTE_BLOCK_HEIGHT to each row.
OFERENTE_CELL_MAP = {
    # Row 19 — personal info line 1
    "numero": ("A", 19),
    "nombre_oferente": ("C", 19),
    "cedula": ("H", 19),
    "certificado_porcentaje": ("K", 19),
    "discapacidad": ("L", 19),
    "telefono_oferente": ("O", 19),
    "resultado_certificado": ("R", 19),
    # Row 21 — personal info line 2 (row 20 = labels)
    "cargo_oferente": ("A", 21),
    "nombre_contacto_emergencia": ("F", 21),
    "parentesco": ("I", 21),
    "telefono_emergencia": ("K", 21),
    "fecha_nacimiento": ("N", 21),
    "edad": ("S", 21),
    # Row 22 — pendiente / contrato
    "pendiente_otros_oferentes": ("G", 22),
    "lugar_firma_contrato": ("L", 22),
    "fecha_firma_contrato": ("R", 22),
    # Row 23 — pension
    "cuenta_pension": ("I", 23),
    "tipo_pension": ("Q", 23),
    # Section 4 — Caracterización (4.1 Condiciones médicas, rows 27+)
    "medicamentos_nivel_apoyo": ("I", 27),
    "medicamentos_conocimiento": ("N", 27),
    "medicamentos_horarios": ("N", 28),
    "medicamentos_nota": ("O", 29),
    "alergias_nivel_apoyo": ("I", 30),
    "alergias_tipo": ("N", 30),
    "alergias_nota": ("O", 31),
    "restriccion_nivel_apoyo": ("I", 32),
    "restriccion_conocimiento": ("N", 32),
    "restriccion_nota": ("O", 33),
    "controles_nivel_apoyo": ("I", 34),
    "controles_asistencia": ("N", 34),
    "controles_frecuencia": ("N", 35),
    "controles_nota": ("O", 36),
    # 4.2 Habilidades básicas (rows 39+)
    "desplazamiento_nivel_apoyo": ("I", 39),
    "desplazamiento_modo": ("N", 39),
    "desplazamiento_transporte": ("N", 40),
    "desplazamiento_nota": ("O", 41),
    "ubicacion_nivel_apoyo": ("I", 42),
    "ubicacion_ciudad": ("N", 42),
    "ubicacion_aplicaciones": ("N", 43),
    "ubicacion_nota": ("O", 44),
    "dinero_nivel_apoyo": ("I", 45),
    "dinero_reconocimiento": ("N", 45),
    "dinero_manejo": ("N", 46),
    "dinero_medios": ("N", 47),
    "dinero_nota": ("O", 48),
    "presentacion_nivel_apoyo": ("I", 49),
    "presentacion_personal": ("N", 49),
    "presentacion_nota": ("O", 50),
    "comunicacion_escrita_nivel_apoyo": ("I", 51),
    "comunicacion_escrita_apoyo": ("N", 51),
    "comunicacion_escrita_nota": ("O", 52),
    "comunicacion_verbal_nivel_apoyo": ("I", 53),
    "comunicacion_verbal_apoyo": ("N", 53),
    "comunicacion_verbal_nota": ("O", 54),
    "decisiones_nivel_apoyo": ("I", 55),
    "toma_decisiones": ("N", 55),
    "toma_decisiones_nota": ("O", 56),
    "aseo_nivel_apoyo": ("I", 57),
    "alimentacion": ("N", 57),
    "aseo_criar_apoyo": ("Q", 58),
    "aseo_comunicacion_apoyo": ("Q", 59),
    "aseo_ayudas_apoyo": ("Q", 60),
    "aseo_alimentacion": ("U", 58),
    "aseo_movilidad_funcional": ("U", 59),
    "aseo_higiene_aseo": ("U", 60),
    "aseo_nota": ("O", 61),
    "instrumentales_nivel_apoyo": ("I", 62),
    "instrumentales_actividades": ("N", 62),
    "instrumentales_criar_apoyo": ("Q", 63),
    "instrumentales_finanzas": ("U", 63),
    "instrumentales_comunicacion_apoyo": ("Q", 64),
    "instrumentales_cocina_limpieza": ("U", 64),
    "instrumentales_movilidad_apoyo": ("Q", 65),
    "instrumentales_crear_hogar": ("U", 65),
    "instrumentales_salud_cuenta_apoyo": ("U", 66),
    "instrumentales_nota": ("O", 67),
    "actividades_nivel_apoyo": ("I", 68),
    "actividades_apoyo": ("N", 68),
    "actividades_esparcimiento_apoyo": ("Q", 69),
    "actividades_esparcimiento_cuenta_apoyo": ("U", 69),
    "actividades_complementarios_apoyo": ("Q", 70),
    "actividades_complementarios_cuenta_apoyo": ("U", 70),
    "actividades_subsidios_cuenta_apoyo": ("U", 71),
    "actividades_nota": ("O", 72),
    "discriminacion_nivel_apoyo": ("I", 73),
    "discriminacion": ("N", 73),
    "discriminacion_violencia_apoyo": ("Q", 74),
    "discriminacion_violencia_cuenta_apoyo": ("U", 74),
    "discriminacion_vulneracion_apoyo": ("Q", 75),
    "discriminacion_vulneracion_cuenta_apoyo": ("U", 75),
    "discriminacion_nota": ("O", 76),
}

SECTION_6_NOMBRE_COL = "E"
SECTION_6_CARGO_COL = "M"


EXCEL_DROPDOWN_MANUAL_CANONICAL_OPTIONS = {
    "tipo_pension": [
        "Pensión Invalidez",
        "Pensión Subsidiada",
        "Pensión especial de vejez anticipada",
        "Pensión para víctimas de conflicto armado",
        "Pensión familiar",
        "Pensión régimen especial (fuerzas militares)",
        "No aplica",
    ],
    # El dropdown del template tiene una comilla residual en una opción;
    # preservamos el texto limpio de la UI para no exportar ese artefacto.
    "discapacidad": list(LIST_FIELD_OPTIONS_BY_ID.get("discapacidad", [])),
}


def register_form():
    return {
        "id": FORM_ID,
        "name": FORM_NAME,
        "hub_description": "Registra el proceso de selección, oferentes y ajustes razonables.",
        "singleton_window": True,
    }


def _get_cache_dir():
    return _get_local_app_cache_dir()




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




def _get_cache_path():
    return os.path.join(_get_cache_dir(), "seleccion_incluyente.json")


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
    if section_id == "section_1":
        SECTION_1_CACHE.clear()
        if isinstance(normalized_payload, dict):
            SECTION_1_CACHE.update(normalized_payload)
    elif section_id == "section_2" and isinstance(normalized_payload, list):
        normalized_payload = _normalize_section_2_payload(normalized_payload)
    FORM_CACHE[section_id] = normalized_payload
    FORM_CACHE["_last_saved_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
    FORM_CACHE["_last_saved_section"] = section_id
    FORM_CACHE["_last_saved_source"] = str(source or "manual").strip() or "manual"
    _record_section_history(section_id, normalized_payload, source=source)


def get_form_cache():
    return dict(FORM_CACHE)


def _get_section_2_entries(payload=None):
    if payload is None:
        payload = FORM_CACHE.get("section_2", [])
    if not isinstance(payload, list):
        return []
    return [dict(entry or {}) for entry in payload]


def _group_export_title_for_oferentes(total_oferentes):
    total = max(0, int(total_oferentes or 0))
    if total <= 1:
        return "PROCESO DE SELECCION INCLUYENTE INDIVIDUAL"
    if total <= 4:
        return "PROCESO DE SELECCION INCLUYENTE GRUPAL - 2 A 4 OFERENTES"
    if total <= 7:
        return "PROCESO DE SELECCION INCLUYENTE GRUPAL - 5 A 7 OFERENTES"
    if total <= 10:
        return "PROCESO DE SELECCION INCLUYENTE GRUPAL - 8 A 10 OFERENTES"
    return "PROCESO DE SELECCION INCLUYENTE GRUPAL - MAS DE 10 OFERENTES"


def _section_2_group_block_start_row(entry_index):
    return OFERENTE_FIRST_BLOCK_START_ROW + (OFERENTE_BLOCK_HEIGHT * int(entry_index or 0))


def _section_2_group_insert_row(entry_index):
    if int(entry_index or 0) <= 0:
        raise ValueError("entry_index debe ser mayor que 0 para bloques adicionales.")
    return OFERENTE_SECOND_BLOCK_START_ROW + (OFERENTE_BLOCK_HEIGHT * (int(entry_index) - 1))



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
    expected_options = LIST_FIELD_OPTIONS_BY_ID.get(field_id, [])
    return tuple(expected_options)


def normalize_excel_dropdown_value(field_id, raw_value):
    if raw_value in (None, ""):
        return raw_value
    current = str(raw_value).strip()
    field_options = LIST_FIELD_OPTIONS_BY_ID.get(field_id)
    if not field_options:
        return raw_value

    canonical_options = list(_get_excel_canonical_options(field_id) or field_options)
    current_norm = _normalize_dropdown_text(current)

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




def get_usuarios_reca_cedulas(env_path=".env"):
    def _loader():
        params = {
            "select": "cedula_usuario",
            "cedula_usuario": "not.is.null",
            "order": "cedula_usuario.asc",
        }
        data = _supabase_get("usuarios_reca", params, env_path=env_path)
        return [row.get("cedula_usuario") for row in data if row.get("cedula_usuario")]

    return list(
        _get_cached_payload(
            "usuarios_reca_cedulas_v1",
            _loader,
            ttl_seconds=_USUARIOS_RECA_CEDULAS_CACHE_TTL_SECONDS,
            allow_stale_on_error=True,
        )
        or []
    )


def get_usuario_reca_by_cedula(cedula, env_path=".env"):
    normalized = _normalize_cedula(cedula)
    if not normalized:
        return None
    select_cols = ",".join(
        [
            "cedula_usuario",
            "nombre_usuario",
            "discapacidad_usuario",
            "discapacidad_detalle",
            "certificado_porcentaje",
            "telefono_oferente",
            "fecha_nacimiento",
            "cargo_oferente",
            "contacto_emergencia",
            "parentesco",
            "telefono_emergencia",
            "resultado_certificado",
            "pendiente_otros_oferentes",
            "cuenta_pension",
            "tipo_pension",
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
    return payload


def confirm_section_2(payload):
    if payload is None:
        raise ValueError("section_2 requerida")
    payload = _normalize_section_2_payload(payload)
    set_section_cache("section_2", payload)
    FORM_CACHE["_last_section"] = "section_2"
    return payload


def confirm_section_5(payload):
    if payload is None:
        raise ValueError("section_5 requerida")
    set_section_cache("section_5", payload)
    FORM_CACHE["_last_section"] = "section_5"
    return payload


def confirm_section_6(payload):
    if payload is None:
        raise ValueError("section_6 requerida")
    set_section_cache("section_6", payload)
    FORM_CACHE["_last_section"] = "section_6"
    return payload


def sync_usuarios_reca(cache=None, env_path=".env"):
    cache_data = FORM_CACHE if cache is None else (cache or {})
    data = cache_data.get("section_2")
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
            "discapacidad_usuario": discapacidad_usuario,
            "discapacidad_detalle": discapacidad_detalle or None,
            "certificado_porcentaje": _normalize_decimal_value(
                entry.get("certificado_porcentaje"),
                decimal_separator=".",
            ),
            "telefono_oferente": (entry.get("telefono_oferente") or "").strip(),
            "fecha_nacimiento": _parse_date_value(entry.get("fecha_nacimiento")),
            "cargo_oferente": (entry.get("cargo_oferente") or "").strip(),
            "contacto_emergencia": (entry.get("nombre_contacto_emergencia") or "").strip(),
            "parentesco": (entry.get("parentesco") or "").strip(),
            "telefono_emergencia": (entry.get("telefono_emergencia") or "").strip(),
            "resultado_certificado": (entry.get("resultado_certificado") or "").strip(),
            "pendiente_otros_oferentes": (entry.get("pendiente_otros_oferentes") or "").strip(),
            "cuenta_pension": (entry.get("cuenta_pension") or "").strip(),
            "tipo_pension": (entry.get("tipo_pension") or "").strip(),
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


def _build_section_1_writes(payload):
    if not payload:
        payload = SECTION_1_CACHE
    writes = []
    for key, cell in SECTION_1_CELL_MAP.items():
        if key in payload:
            value = normalize_excel_dropdown_value(key, payload[key])
            writes.append({"range": f"'{SHEET_NAME}'!{cell}", "value": value})
            _log_excel(f"WRITE section=section_1 cell={cell} key={key} value={value!r}")
    return writes


def _build_section_2_entry_writes(entry, *, row_offset=0):
    writes = []
    for field_id, (col, row) in OFERENTE_CELL_MAP.items():
        value = entry.get(field_id, "")
        if value == "":
            continue
        if field_id == "certificado_porcentaje":
            value = _coerce_excel_decimal_value(value)
        else:
            value = normalize_excel_dropdown_value(field_id, value)
        target_row = row + row_offset
        writes.append({"range": f"'{SHEET_NAME}'!{col}{target_row}", "value": value})
        _log_excel(f"WRITE section=section_2 cell={col}{target_row} key={field_id} value={value!r}")
    return writes


def _build_section_2_writes(oferentes):
    if not oferentes:
        return []
    writes = [
        {
            "range": f"'{SHEET_NAME}'!{GROUP_EXPORT_TITLE_CELL}",
            "value": _group_export_title_for_oferentes(len(oferentes)),
        }
    ]
    # Write shared desarrollo_actividad
    shared_desarrollo = ""
    for entry in oferentes:
        shared_desarrollo = (entry.get("desarrollo_actividad") or "").strip()
        if shared_desarrollo:
            break
    if shared_desarrollo:
        writes.append({"range": f"'{SHEET_NAME}'!{DESARROLLO_ACTIVIDAD_CELL}", "value": shared_desarrollo})
        _log_excel(f"WRITE section=section_2 cell={DESARROLLO_ACTIVIDAD_CELL} key=desarrollo_actividad value={shared_desarrollo!r}")
    _log_excel(f"SECTION section=section_2 total={len(oferentes)}")
    for idx, entry in enumerate(oferentes):
        title_row = _section_2_group_block_start_row(idx)
        writes.append(
            {
                "range": f"'{SHEET_NAME}'!{OFERENTE_TITLE_COL}{title_row}",
                "value": f"OFERENTE {idx + 1}",
            }
        )
        row_offset = OFERENTE_BLOCK_HEIGHT * idx
        writes.extend(_build_section_2_entry_writes(entry, row_offset=row_offset))
    return writes


def _build_section_2_row_insertions(oferentes):
    total_oferentes = len(oferentes or [])
    if total_oferentes <= 1:
        return []
    return [
        {
            "sheet_name": SHEET_NAME,
            "insert_at_row": _section_2_group_insert_row(1),
            "template_start_row": OFERENTE_FIRST_BLOCK_START_ROW,
            "template_end_row": OFERENTE_FIRST_BLOCK_START_ROW + OFERENTE_BLOCK_HEIGHT - 1,
            "repeat_count": total_oferentes - 1,
        }
    ]


def _build_section_5_writes(payload, num_oferentes=1):
    if not payload:
        return []
    shift = max(0, num_oferentes - 1) * OFERENTE_BLOCK_HEIGHT
    ajustes_row = SECTION_5_BASE_AJUSTES_ROW + shift
    nota_row = SECTION_5_BASE_NOTA_ROW + shift
    ajustes_value = payload.get("ajustes_recomendaciones", "")
    nota_value = payload.get("nota", "")
    nota_value = f"Nota: {nota_value}" if nota_value else "Nota:"
    writes = []
    if ajustes_value:
        writes.append({"range": f"'{SHEET_NAME}'!A{ajustes_row}", "value": ajustes_value})
    writes.append({"range": f"'{SHEET_NAME}'!A{nota_row}", "value": nota_value})
    _log_excel(f"WRITE section=section_5 ajustes_row={ajustes_row} nota_row={nota_row}")
    return writes


def _build_section_6_writes(payload, num_oferentes=1):
    if not payload:
        return []
    shift = max(0, num_oferentes - 1) * OFERENTE_BLOCK_HEIGHT
    start_row = SECTION_6_BASE_START_ROW + shift
    writes = []
    for idx, entry in enumerate(payload):
        row = start_row + idx
        nombre = (entry.get("nombre") or "").strip()
        cargo = (entry.get("cargo") or "").strip()
        if nombre:
            writes.append({"range": f"'{SHEET_NAME}'!{SECTION_6_NOMBRE_COL}{row}", "value": nombre})
        if cargo:
            writes.append({"range": f"'{SHEET_NAME}'!{SECTION_6_CARGO_COL}{row}", "value": cargo})
    return writes


def _write_section_5(ws, payload, template_variant=TEMPLATE_VARIANT_INDIVIDUAL, total_oferentes=0):
    for write in _build_section_5_writes(payload, num_oferentes=max(1, int(total_oferentes or 1))):
        cell = str(write.get("range") or "").rsplit("!", 1)[-1].replace("'", "")
        ws_write(ws, cell, write.get("value", ""))


def _write_section_6(ws, payload, template_variant=TEMPLATE_VARIANT_INDIVIDUAL, total_oferentes=0):
    for write in _build_section_6_writes(payload, num_oferentes=max(1, int(total_oferentes or 1))):
        cell = str(write.get("range") or "").rsplit("!", 1)[-1].replace("'", "")
        ws_write(ws, cell, write.get("value", ""))


def _build_section_6_row_insertions(payload, num_oferentes=1):
    if not payload:
        return []
    base_rows = int(SECTION_6.get("rows", 4) or 4)
    total_rows = len(payload)
    if total_rows <= base_rows:
        return []
    shift = max(0, num_oferentes - 1) * OFERENTE_BLOCK_HEIGHT
    start_row = SECTION_6_BASE_START_ROW + shift
    return [
        {
            "sheet_name": SHEET_NAME,
            "start_row": start_row,
            "base_rows": base_rows,
            "total_rows": total_rows,
        }
    ]


def _validate_section_2_rows(issues, rows):
    row_pairs = [
        (field_id, label)
        for field_id, label in field_pairs(SECTION_2.get("fields"))
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
        "Oferentes",
        message="Debes diligenciar al menos un oferente.",
    )


def validate_before_finalize(cache=None):
    cache_data = FORM_CACHE if cache is None else (cache or {})
    issues = []

    section_1 = cache_data.get("section_1", {})
    for field_id, label in field_pairs(SECTION_1.get("fields")):
        require_value(issues, "section_1", section_1, field_id, label)

    _validate_section_2_rows(issues, cache_data.get("section_2", []))

    section_5 = cache_data.get("section_5", {})
    for field_id, label in field_pairs(SECTION_5.get("fields")):
        require_value(issues, "section_5", section_5, field_id, label)

    validate_dynamic_rows(
        issues,
        "section_6",
        cache_data.get("section_6", []),
        [("nombre", "Nombre"), ("cargo", "Cargo")],
        min_rows_label="Asistentes",
    )
    return issues


def export_to_excel(clear_cache=True, cache=None):
    cache_data = FORM_CACHE if cache is None else (cache or {})
    raise_validation_error(validate_before_finalize(cache_data))

    from google_sheets_client import get_master_template_id
    from drive_upload import publish_sheet_from_template

    oferentes = _get_section_2_entries(cache_data.get("section_2", []))
    num_oferentes = len(oferentes)

    _log_excel(f"START export_all (Google Sheets) oferentes={num_oferentes}")

    section_1 = cache_data.get("section_1") or {}
    empresa_nombre = section_1.get("nombre_empresa") or SECTION_1_CACHE.get("nombre_empresa") or "Empresa"
    base_name = _sanitize_filename(empresa_nombre)

    writes = []
    writes.extend(_build_section_1_writes(section_1))
    writes.extend(_build_section_2_writes(oferentes))
    writes.extend(_build_section_5_writes(cache_data.get("section_5", {}), num_oferentes=num_oferentes))
    writes.extend(_build_section_6_writes(cache_data.get("section_6", []), num_oferentes=num_oferentes))
    row_insertions = []
    row_insertions.extend(_build_section_2_row_insertions(oferentes))
    row_insertions.extend(
        _build_section_6_row_insertions(
            cache_data.get("section_6", []),
            num_oferentes=num_oferentes,
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
    )

    _log_excel("SUCCESS export_all")

    # Determinar tipo_acta y extra_name para el PDF
    fecha_visita_raw = str(section_1.get("fecha_visita") or "").strip()

    if num_oferentes == 1:
        tipo_acta = "seleccion_individual"
        # CRITERIOS ROTULACIÓN: primer nombre + primer apellido
        nombre_completo = str((oferentes[0] if oferentes else {}).get("nombre_oferente") or "").strip()
        extra_name = _primer_nombre_apellido(nombre_completo)
    else:
        tipo_acta = "seleccion_grupal"
        extra_name = str(num_oferentes)

    # Construir participantes para RECA ODS
    participantes = [
        {
            "nombre_usuario": str(o.get("nombre_oferente") or "").strip(),
            "cedula_usuario": str(o.get("cedula") or "").strip(),
            "discapacidad_usuario": str(o.get("discapacidad") or "").strip(),
            "genero_usuario": "",
        }
        for o in oferentes
        if isinstance(o, dict)
    ]

    acta_metadata = {
        "tipo_acta": tipo_acta,
        "nit_empresa": str(section_1.get("nit_empresa") or "").strip(),
        "nombre_empresa": str(empresa_nombre or "").strip(),
        "fecha_servicio": fecha_visita_raw,
        "nombre_profesional": str(section_1.get("profesional_asignado") or section_1.get("asesor") or "").strip(),
        "modalidad_servicio": str(section_1.get("modalidad") or "").strip(),
        "participantes": participantes,
        "asistentes": [],
    }

    if clear_cache and cache is None:
        clear_cache_file()
        clear_form_cache()

    return {
        "output_path": result.get("webViewLink", ""),
        "drive_file_id": result.get("file_id", ""),
        "already_in_drive": True,
        "tipo_acta": tipo_acta,
        "fecha_servicio": fecha_visita_raw,
        "acta_metadata": acta_metadata,
        "extra_name": extra_name,
    }


def _primer_nombre_apellido(nombre_completo: str) -> str:
    """Extrae primer nombre y primer apellido de un nombre completo.

    Sigue los CRITERIOS DE ROTULACIÓN de RECA: "Debe ir el primer nombre
    y primer apellido."

    Heurística para nombres colombianos:
    - 1-2 palabras  → retorna tal cual ("Juan Pérez")
    - 3+ palabras   → words[0] + words[-2]  ("Juan Carlos Pérez García" → "Juan Pérez")
    """
    words = nombre_completo.strip().split()
    if len(words) <= 2:
        return nombre_completo.strip()
    return f"{words[0]} {words[-2]}"
