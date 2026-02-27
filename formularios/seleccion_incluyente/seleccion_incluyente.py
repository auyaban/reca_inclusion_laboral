import json
import os
import shutil
import time

from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.common import (
    _get_desktop_dir,
    _normalize_cedula,
    _normalize_text,
    _parse_date_value,
    _sanitize_filename,
    _supabase_get,
    _supabase_upsert_with_queue,
)

FORM_ID = "seleccion_incluyente"
FORM_NAME = "Proceso de Seleccion Incluyente"

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

SHEET_NAME = "4. PROCESO DE SELECCION INCLUYE"
SECTION_2_ANCHOR = "2. DATOS DEL OFERENTE"
SECTION_5_ANCHOR = "5. AJUSTES RAZONABLES / RECOMENDACIONES AL PROCESO DE SELECCION"
SECTION_2_TEMPLATE_ANCHOR_ROW = 14
SECTION_2_LAST_COLUMN = "U"

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
            "label": "Resultado certificado",
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
    "rows": 4,
}

SECTION_2_CELL_MAP = {
    "numero": ("A", 17),
    "nombre_oferente": ("C", 17),
    "cedula": ("H", 17),
    "certificado_porcentaje": ("K", 17),
    "discapacidad": ("L", 17),
    "telefono_oferente": ("O", 17),
    "resultado_certificado": ("R", 17),
    "cargo_oferente": ("A", 19),
    "nombre_contacto_emergencia": ("F", 19),
    "parentesco": ("I", 19),
    "telefono_emergencia": ("K", 19),
    "fecha_nacimiento": ("N", 19),
    "edad": ("S", 19),
    "pendiente_otros_oferentes": ("G", 20),
    "lugar_firma_contrato": ("L", 20),
    "fecha_firma_contrato": ("R", 20),
    "cuenta_pension": ("I", 21),
    "tipo_pension": ("Q", 21),
    "desarrollo_actividad": ("A", 23),
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
    "desplazamiento_nivel_apoyo": ("I", 40),
    "desplazamiento_modo": ("N", 40),
    "desplazamiento_transporte": ("N", 41),
    "desplazamiento_nota": ("O", 42),
    "ubicacion_nivel_apoyo": ("I", 43),
    "ubicacion_ciudad": ("N", 43),
    "ubicacion_aplicaciones": ("N", 44),
    "ubicacion_nota": ("O", 45),
    "dinero_nivel_apoyo": ("I", 46),
    "dinero_reconocimiento": ("N", 46),
    "dinero_manejo": ("N", 47),
    "dinero_medios": ("N", 48),
    "dinero_nota": ("O", 49),
    "presentacion_nivel_apoyo": ("I", 50),
    "presentacion_personal": ("N", 50),
    "presentacion_nota": ("O", 51),
    "comunicacion_escrita_nivel_apoyo": ("I", 52),
    "comunicacion_escrita_apoyo": ("N", 52),
    "comunicacion_escrita_nota": ("N", 53),
    "comunicacion_verbal_nivel_apoyo": ("I", 54),
    "comunicacion_verbal_apoyo": ("N", 54),
    "comunicacion_verbal_nota": ("O", 55),
    "decisiones_nivel_apoyo": ("I", 56),
    "toma_decisiones": ("N", 56),
    "toma_decisiones_nota": ("O", 57),
    "aseo_nivel_apoyo": ("I", 58),
    "alimentacion": ("N", 58),
    "aseo_criar_apoyo": ("Q", 59),
    "aseo_comunicacion_apoyo": ("Q", 60),
    "aseo_ayudas_apoyo": ("Q", 61),
    "aseo_alimentacion": ("U", 59),
    "aseo_movilidad_funcional": ("U", 60),
    "aseo_higiene_aseo": ("U", 61),
    "aseo_nota": ("O", 62),
    "instrumentales_nivel_apoyo": ("I", 63),
    "instrumentales_actividades": ("N", 63),
    "instrumentales_criar_apoyo": ("Q", 64),
    "instrumentales_finanzas": ("U", 64),
    "instrumentales_comunicacion_apoyo": ("Q", 65),
    "instrumentales_cocina_limpieza": ("U", 65),
    "instrumentales_movilidad_apoyo": ("Q", 66),
    "instrumentales_crear_hogar": ("U", 66),
    "instrumentales_salud_cuenta_apoyo": ("U", 67),
    "instrumentales_nota": ("O", 68),
    "actividades_nivel_apoyo": ("I", 69),
    "actividades_apoyo": ("N", 69),
    "actividades_esparcimiento_apoyo": ("Q", 70),
    "actividades_esparcimiento_cuenta_apoyo": ("U", 70),
    "actividades_complementarios_apoyo": ("Q", 71),
    "actividades_complementarios_cuenta_apoyo": ("U", 71),
    "actividades_subsidios_cuenta_apoyo": ("U", 72),
    "actividades_nota": ("O", 73),
    "discriminacion_nivel_apoyo": ("I", 74),
    "discriminacion": ("N", 74),
    "discriminacion_violencia_apoyo": ("Q", 75),
    "discriminacion_violencia_cuenta_apoyo": ("U", 75),
    "discriminacion_vulneracion_apoyo": ("Q", 76),
    "discriminacion_vulneracion_cuenta_apoyo": ("U", 76),
    "discriminacion_nota": ("O", 77),
}

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
    "section_6": {
        "start_row": 85,
        "rows": 4,
        "nombre_col": "E",
        "cargo_col": "M",
    },
}


def register_form():
    return {
        "id": FORM_ID,
        "name": FORM_NAME,
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
    FORM_CACHE[section_id] = payload


def get_form_cache():
    return dict(FORM_CACHE)




def _find_template_path():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    templates_dir = os.path.join(base_dir, "templates")
    if not os.path.isdir(templates_dir):
        raise FileNotFoundError("No existe la carpeta templates.")
    for name in os.listdir(templates_dir):
        if name.startswith("~$"):
            continue
        normalized = _normalize_text(name).replace("_", "")
        if "seleccion" in normalized and "incluyente" in normalized and normalized.endswith(".xlsx"):
            return os.path.join(templates_dir, name)
    raise FileNotFoundError("No se encontró el template de seleccion incluyente.")


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
    process_name = "Proceso de Seleccion Incluyente"
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
            if target.startswith("2.") or target.startswith("5."):
                if value_norm.startswith(target):
                    return row
            else:
                return row
    raise ValueError(f"No se encontró el texto '{text}' en la columna A.")


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
            "discapacidad_usuario": discapacidad_usuario,
            "discapacidad_detalle": discapacidad_detalle or None,
            "certificado_porcentaje": (entry.get("certificado_porcentaje") or "").strip(),
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
        cleaned = {k: v for k, v in row.items() if v not in ("", None)}
        rows.append(cleaned)
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


def _write_section_1(ws, payload):
    if not payload:
        payload = SECTION_1_CACHE
    if not payload:
        try:
            if load_cache_from_file():
                payload = FORM_CACHE.get("section_1", {}) or SECTION_1_CACHE
        except Exception:
            payload = payload or {}
    for key, cell in EXCEL_MAPPING.get("section_1", {}).items():
        if key in payload:
            ws.Range(cell).Value = payload.get(key)
            _log_excel(
                f"WRITE section=section_1 cell={cell} key={key} value={payload.get(key)!r}"
            )


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
    start_row = _find_row_by_text(ws, SECTION_2_ANCHOR)
    next_row = _find_row_by_text(ws, SECTION_5_ANCHOR)
    block_height = next_row - start_row
    if block_height <= 0:
        raise ValueError("Anclas de seccion 2 invalidas.")
    _log_excel(
        f"SECTION section=section_2 start_row={start_row} next_row={next_row} block_height={block_height} total={len(oferentes)}"
    )
    total_oferentes = len(oferentes)
    if 2 <= total_oferentes <= 4:
        ws.Range("G1").Value = "PROCESO DE SELECCION INCLUYENTE GRUPAL - 2 A 4 OFERENTES"
    elif 5 <= total_oferentes <= 7:
        ws.Range("G1").Value = "PROCESO DE SELECCION INCLUYENTE GRUPAL - 5 A 7 OFERENTES"
    elif 8 <= total_oferentes <= 10:
        ws.Range("G1").Value = "PROCESO DE SELECCION INCLUYENTE GRUPAL - 8 A 10 OFERENTES"
    for idx in range(1, len(oferentes)):
        insert_at = start_row + (block_height * idx)
        _insert_person_block(ws, start_row, block_height, insert_at)
        _log_excel(
            f"INSERT section=section_2 rows={block_height} at={insert_at}"
        )

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


def _write_section_5(ws, payload):
    if not payload:
        return
    anchor_row = _find_row_by_text(ws, SECTION_5_ANCHOR)
    ajustes_row = anchor_row + 1
    nota_row = anchor_row + 2
    ajustes_value = payload.get("ajustes_recomendaciones", "")
    nota_value = payload.get("nota", "")
    nota_value = f"Nota: {nota_value}" if nota_value else "Nota:"
    _log_excel(
        f"WRITE section=section_5 cell=A{ajustes_row} key=ajustes_recomendaciones value={ajustes_value!r}"
    )
    _log_excel(
        f"WRITE section=section_5 cell=A{nota_row} key=nota value={nota_value!r}"
    )
    ws.Range(f"A{ajustes_row}").Value = ajustes_value
    ws.Range(f"A{nota_row}").Value = nota_value


def _write_section_6(ws, payload):
    if not payload:
        return
    mapping = EXCEL_MAPPING.get("section_6", {})
    title_row = _find_row_by_text(ws, "6. ASISTENTES")
    start_row = title_row + 1
    base_rows = mapping.get("rows", 4)
    nombre_col = mapping.get("nombre_col", "E")
    cargo_col = mapping.get("cargo_col", "M")
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
            f"WRITE section=section_6 cell={nombre_col}{row} key=nombre value={nombre!r}"
        )
        _log_excel(
            f"WRITE section=section_6 cell={cargo_col}{row} key=cargo value={cargo!r}"
        )
        ws.Range(f"{nombre_col}{row}").Value = nombre
        ws.Range(f"{cargo_col}{row}").Value = cargo


def export_to_excel(clear_cache=True):
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
        _write_section_1(ws, FORM_CACHE.get("section_1", {}))
        _write_section_2(ws, FORM_CACHE.get("section_2", []))
        _write_section_5(ws, FORM_CACHE.get("section_5", {}))
        _write_section_6(ws, FORM_CACHE.get("section_6", []))
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

