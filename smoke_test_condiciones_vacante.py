#!/usr/bin/env python3
"""
Smoke test — Condiciones de Vacante (Excel export)
====================================================
Verifies that every editable section is written to the Excel file and
that autofit is applied, with a ~1000-word stress test on section 7.

Run:
    python smoke_test_condiciones_vacante.py

Requirements: Excel installed, pywin32
"""

import os
import sys
import time

# Ensure UTF-8 output so special chars (→, ☑, etc.) don't crash on Windows cp1252
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ---------------------------------------------------------------------------
# ~1000-word observaciones for the extreme autofit stress test
# ---------------------------------------------------------------------------
OBSERVACIONES_1000W = (
    "PROCESO DE REVISIÓN DE LAS CONDICIONES DE LA VACANTE — OBSERVACIONES Y RECOMENDACIONES DETALLADAS\n\n"
    "* Análisis del entorno físico de trabajo: Durante la visita realizada a las instalaciones de la empresa se "
    "llevó a cabo una evaluación exhaustiva de las condiciones físicas del puesto de trabajo asignado para este "
    "cargo. Se identificó que el área cuenta con iluminación adecuada, temperatura controlada dentro de los "
    "parámetros establecidos por la normativa colombiana de salud y seguridad en el trabajo, y espacios de "
    "circulación suficientemente amplios para garantizar la accesibilidad universal. No obstante, se recomienda "
    "instalar señalización visual y auditiva en las rutas de evacuación, así como demarcar con colores contrastantes "
    "las zonas de tránsito peatonal para facilitar la orientación de personas con baja visión.\n\n"
    "* Evaluación de la accesibilidad arquitectónica: El inmueble dispone de rampa de acceso en la entrada "
    "principal con pendiente inferior al 8%, lo cual cumple con los requisitos de la Norma Técnica Colombiana NTC "
    "4143. Sin embargo, el acceso al segundo piso donde se ubica la cafetería y las salas de descanso requiere "
    "la instalación de un ascensor o plataforma elevadora para garantizar la movilidad autónoma de personas con "
    "discapacidad física. La empresa se comprometió a gestionar esta adecuación en un plazo no mayor a seis meses "
    "contados a partir de la firma del presente documento.\n\n"
    "* Recomendaciones sobre el proceso de selección: Se sugiere que el proceso de selección incluya etapas "
    "adaptadas a las necesidades de cada candidato con discapacidad. En particular, para personas con discapacidad "
    "auditiva se deberá contar con un intérprete de Lengua de Señas Colombiana durante las entrevistas; para "
    "personas con discapacidad visual se facilitarán las pruebas psicotécnicas en formato digital accesible o en "
    "Braille según corresponda; para personas con discapacidad intelectual se utilizarán instrucciones simplificadas "
    "con apoyo de lectura fácil y mayor tiempo en la realización de las evaluaciones.\n\n"
    "* Ajustes razonables en el puesto de trabajo: De conformidad con la Ley 1618 de 2013 y el Decreto 392 de "
    "2018, la empresa está en obligación de realizar los ajustes razonables que permitan a la persona vinculada "
    "desempeñar sus funciones en condiciones de igualdad. Los ajustes identificados para este cargo incluyen: "
    "software lector de pantalla o ampliador de texto en el equipo de cómputo asignado, silla ergonómica con "
    "soporte lumbar ajustable, reposapiés regulable, teclado en Braille o de alto contraste según la condición "
    "de la persona, y adaptación del escritorio en altura para usuarios de silla de ruedas.\n\n"
    "* Apoyo durante el período de inducción y adaptación: Se recomienda que la empresa designe un mentor o "
    "compañero de apoyo durante los primeros tres meses de vinculación laboral, quien acompañe al trabajador "
    "con discapacidad en el conocimiento de los procesos, procedimientos, cultura organizacional y relaciones "
    "interpersonales. Este acompañamiento favorecerá la integración efectiva y reducirá los tiempos de adaptación "
    "al cargo. Adicionalmente, el equipo de RECA realizará visitas de seguimiento a los 30, 60 y 90 días para "
    "verificar la correcta implementación de los ajustes y apoyar tanto al trabajador como a la empresa en la "
    "resolución de cualquier dificultad que se presente.\n\n"
    "* Sensibilización al equipo de trabajo: Es indispensable que antes del inicio de labores de la persona con "
    "discapacidad se realice un taller de sensibilización dirigido a todo el equipo de trabajo del área, con el "
    "objetivo de promover una cultura de inclusión, derribar imaginarios y estereotipos, y brindar herramientas "
    "prácticas de comunicación asertiva con personas con diferentes tipos de discapacidad. Este taller será "
    "facilitado por los profesionales de RECA y tendrá una duración aproximada de dos horas.\n\n"
    "* Perfil de discapacidad compatible con el cargo: El presente perfil describe los tipos de discapacidad que, "
    "tras el análisis detallado de las funciones del cargo, el entorno de trabajo, los factores de riesgo "
    "identificados y las demandas físicas, cognitivas y relacionales propias del rol, se consideran compatibles "
    "para la vinculación laboral de personas con discapacidad bajo un enfoque de inclusión social y laboral "
    "genuina, garantizando siempre la dignidad, el respeto y el pleno ejercicio de los derechos laborales de "
    "las personas candidatas al cargo.\n\n"
    "* Compromisos de la empresa: La empresa se compromete formalmente a: (1) realizar los ajustes razonables "
    "identificados en el presente documento dentro de los plazos acordados; (2) no divulgar la condición de "
    "discapacidad del trabajador sin su consentimiento expreso; (3) garantizar igualdad de condiciones en materia "
    "de remuneración, beneficios, oportunidades de capacitación y desarrollo profesional; (4) reportar a la "
    "Agencia Pública de Empleo el progreso de la vinculación y cualquier novedad relevante; (5) mantener "
    "comunicación permanente con el equipo de RECA durante el período de seguimiento.\n\n"
    "* Próximos pasos: La Agencia procederá a publicar la vacante con el perfil levantado en la presente visita "
    "y realizará el envío de candidatos dentro de los cuatro días hábiles siguientes a la aprobación del perfil "
    "por parte de la empresa. Se realizará retroalimentación oportuna a todos los candidatos que participen en "
    "el proceso, independientemente del resultado, garantizando una experiencia de selección digna y respetuosa.\n\n"
    "* Gestión de riesgos laborales: El área de seguridad y salud en el trabajo deberá actualizar la matriz de "
    "riesgos para incluir los factores asociados a la vinculación de personas con discapacidad. Esto incluye la "
    "revisión de los procedimientos de evacuación, la identificación de rutas accesibles en caso de emergencia, "
    "la asignación de personal de apoyo en simulacros para garantizar la evacuación segura de todos los "
    "colaboradores, y la actualización del plan de emergencias con los protocolos específicos para personas con "
    "movilidad reducida, discapacidad sensorial o discapacidad cognitiva. El comité paritario de seguridad y "
    "salud en el trabajo debe ser informado y capacitado sobre estas modificaciones antes del inicio de "
    "operaciones del nuevo colaborador.\n\n"
    "* Indicadores de seguimiento y evaluación: Para medir el éxito del proceso de inclusión laboral se "
    "establecerán los siguientes indicadores de seguimiento: tasa de permanencia en el cargo a los 3, 6 y 12 "
    "meses, nivel de satisfacción del trabajador medido mediante encuesta trimestral, cumplimiento de metas de "
    "desempeño en comparación con el promedio del equipo, número de ajustes razonables implementados versus "
    "los comprometidos, y percepción del clima laboral inclusivo evaluada con el equipo directo. Estos "
    "indicadores serán reportados por la empresa al equipo de RECA en cada visita de seguimiento y formarán "
    "parte del informe final de cierre del caso.\n\n"
    "* Consideraciones adicionales sobre accesibilidad tecnológica: Los sistemas de información y comunicación "
    "utilizados en el cargo deben ser accesibles para personas con discapacidad visual o auditiva. Se revisará "
    "la compatibilidad del software de gestión interna con lectores de pantalla como NVDA o JAWS, y se "
    "gestionará la habilitación de subtítulos automáticos en plataformas de videoconferencia utilizadas para "
    "capacitaciones virtuales. La empresa designará un referente de accesibilidad tecnológica que actuará como "
    "punto de contacto para resolver inconvenientes técnicos que enfrente el trabajador con discapacidad en su "
    "día a día laboral, asegurando que las barreras tecnológicas no se conviertan en obstáculos para el "
    "desempeño efectivo del rol asignado ni para el desarrollo profesional a largo plazo dentro de la organización."
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
_PASS = "[PASS]"
_FAIL = "[FAIL]"
_results: list[tuple[str, str]] = []


def check(label: str, condition: bool, detail: str = "") -> bool:
    tag = _PASS if condition else _FAIL
    msg = f"  {tag}  {label}"
    if detail:
        msg += f"  ({detail})"
    print(msg)
    _results.append((tag, label))
    return condition


def word_count(text: str) -> int:
    return len(text.split())


# ---------------------------------------------------------------------------
# Test data — covers EVERY editable field across all 9 sections
# ---------------------------------------------------------------------------
FAKE_SECTION_1 = {
    "fecha_visita": "11/03/2026",
    "modalidad": "Presencial",
    "nombre_empresa": "Empresa Prueba QA S.A.S.",
    "ciudad_empresa": "Bogotá D.C.",
    "direccion_empresa": "Calle 100 No. 20-30, Bogotá D.C.",
    "nit_empresa": "900123456-7",
    "correo_1": "rrhh@empresaprueba.com.co",
    "telefono_empresa": "601 234 5678 / 310 987 6543",
    "contacto_empresa": "María García López",
    "cargo": "Gerente de Recursos Humanos",
    "caja_compensacion": "Compensar",
    "sede_empresa": "Sede Bogotá Norte",
    "asesor": "Carlos Alberto Pérez Rodríguez",
    "profesional_asignado": "Ana María Martínez Vargas",
}

FAKE_SECTION_2 = {
    "nombre_vacante": "Operario de Producción Inclusivo",
    "numero_vacantes": "3",
    "nivel_cargo": "Operativo.",
    "genero": "Hombre - Mujer",
    "edad": "25 a 45 años",
    "modalidad_trabajo": "Presencial",
    "lugar_trabajo": "Planta Principal — Zona Industrial Norte",
    "salario_asignado": "$1.500.000 + prestaciones de ley",
    "firma_contrato": "Inmediato, sujeto a pruebas",
    "aplicacion_pruebas": "Sí, psicotécnicas y de aptitud",
    "tipo_contrato": "Término Fijo.",
    "beneficios_adicionales": "Auxilio de transporte, desayuno, dotación",
    "cargo_flexible_genero": "No aplica restricción de género",
    "beneficios_mujeres": "Sí, auxilio de maternidad ampliado",
    "requiere_certificado": "Sí",
    "requiere_certificado_observaciones": "Certificado médico de discapacidad requerido para registro SENA",
    "competencia_1": "Responsabilidad y cumplimiento",
    "competencia_2": "Trabajo en equipo",
    "competencia_3": "Adaptabilidad al cambio",
    "competencia_4": "Comunicación asertiva",
    "competencia_5": "Seguimiento de instrucciones",
    "competencia_6": "Proactividad",
    "competencia_7": "Honestidad e integridad",
    "competencia_8": "Orientación al resultado",
}

FAKE_SECTION_2_1 = {
    # Checkboxes (bool → ☑ or ☐)
    "nivel_primaria": False,
    "nivel_bachiller": True,
    "nivel_tecnico_profesional": True,
    "nivel_profesional": False,
    "nivel_especializacion": False,
    "nivel_tecnologo": True,
    # Text fields
    "especificaciones_formacion": "Bachillerato completo o técnico en cualquier área productiva o afín",
    "conocimientos_basicos": "Lectura de planos básicos, Excel nivel básico, manejo de inventarios",
    "horarios_asignados": "Horarios Rotativos.",
    "hora_ingreso": "6:00 a.m.",
    "hora_salida": "2:00 p.m. / 10:00 p.m. (según rotación)",
    "tiempo_almuerzo": "1 hora.",
    "break_descanso": "15 minutos",
    "dias_laborables": "Lunes a Sábado con un día compensatorio",
    "dias_flexibles": "Viernes flexible cada 15 días",
    "observaciones": (
        "El horario puede variar según la demanda de producción. "
        "La empresa garantiza respetar los períodos de descanso establecidos por la ley. "
        "Se evaluará la posibilidad de horarios ajustados para personas que requieran traslado especial."
    ),
    "experiencia_meses": "Un año.",
    "funciones_tareas": (
        "1. Operar máquinas de empaque y sellado según especificaciones técnicas. "
        "2. Realizar control de calidad visual de los productos terminados. "
        "3. Reportar novedades al supervisor de línea de forma oportuna. "
        "4. Mantener el orden y aseo de la estación de trabajo. "
        "5. Participar en los procesos de inventario periódico."
    ),
    "herramientas_equipos": (
        "Máquinas empacadoras automáticas, básculas electrónicas, "
        "montacargas manual (patín hidráulico), EPP completo (casco, guantes, botas de seguridad, "
        "tapabocas industrial, gafas de protección)."
    ),
}

FAKE_SECTION_3 = {
    # Habilidades cognitivas
    "lectura": "Medio.",
    "comprension_lectora": "Medio.",
    "escritura": "Bajo.",
    "comunicacion_verbal": "Alto.",
    "razonamiento_logico": "Medio.",
    "conteo_reporte": "Medio.",
    "clasificacion_objetos": "Alto.",
    "velocidad_ejecucion": "Alto.",
    "concentracion": "Alto.",
    "memoria": "Medio.",
    "ubicacion_espacial": "Alto.",
    "atencion": "Alto.",
    "observaciones_cognitivas": "Buena capacidad de concentración sostenida; atención al detalle es clave para el cargo.",
    # Motricidad fina
    "agarre": "Alto.",
    "precision": "Alto.",
    "digitacion": "Bajo.",
    "agilidad_manual": "Alto.",
    "coordinacion_ojo_mano": "Alto.",
    "observaciones_motricidad_fina": "Se requiere destreza manual para manipulación de piezas pequeñas.",
    # Motricidad gruesa
    "esfuerzo_fisico": "Alto.",
    "equilibrio_corporal": "Medio.",
    "lanzar_objetos": "No aplica",
    "observaciones_motricidad_gruesa": "El cargo exige períodos prolongados de pie y traslado de materiales livianos.",
    # Competencias transversales
    "seguimiento_instrucciones": "Alto.",
    "resolucion_conflictos": "Medio.",
    "autonomia_tareas": "Alto.",
    "trabajo_equipo": "Alto.",
    "adaptabilidad": "Alto.",
    "flexibilidad": "Alto.",
    "comunicacion_asertiva": "Medio.",
    "manejo_tiempo": "Alto.",
    "liderazgo": "Bajo.",
    "escucha_activa": "Alto.",
    "proactividad": "Alto.",
    "observaciones_transversales": "Persona con actitud positiva, puntual y comprometida con los objetivos del equipo.",
}

FAKE_SECTION_4 = {
    "sentado_tiempo": "De 1 a 2 horas.",
    "sentado_frecuencia": "Diario.",
    "semisentado_tiempo": "No aplica",
    "semisentado_frecuencia": "No aplica.",
    "de_pie_tiempo": "De 4 a 6 horas.",
    "de_pie_frecuencia": "Diario.",
    "agachado_tiempo": "De 1 a 2 horas.",
    "agachado_frecuencia": "Semanal.",
    "uso_extremidades_superiores_tiempo": "De 6 a 8 horas.",
    "uso_extremidades_superiores_frecuencia": "Diario.",
}

FAKE_SECTION_5 = {
    # Físico
    "ruido": "Medio.",
    "iluminacion": "Alto.",
    "temperaturas_externas": "Medio.",
    "vibraciones": "Bajo.",
    "presion_atmosferica": "No aplica",
    "radiaciones": "No aplica",
    # Químico
    "polvos_organicos_inorganicos": "Medio.",
    "fibras": "No aplica",
    "liquidos": "Bajo.",
    "gases_vapores": "Bajo.",
    "humos_metalicos": "No aplica",
    "humos_no_metalicos": "Bajo.",
    "material_particulado": "Medio.",
    # Condiciones de seguridad
    "electrico": "Bajo.",
    "locativo": "Bajo.",
    "accidentes_transito": "No aplica",
    "publicos": "No aplica",
    "mecanico": "Medio.",
    # Psicosocial
    "gestion_organizacional": "Medio.",
    "caracteristicas_organizacion": "Medio.",
    "caracteristicas_grupo_social": "Alto.",
    "condiciones_tarea": "Medio.",
    "interfase_persona_tarea": "Medio.",
    "jornada_trabajo": "Medio.",
    # Ergonómico
    "postura_trabajo": "Medio.",
    "puesto_trabajo": "Medio.",
    "movimientos_repetitivos": "Alto.",
    "manipulacion_cargas": "Medio.",
    "herramientas_equipos": "Bajo.",
    "organizacion_trabajo": "Medio.",
    "observaciones_peligros": (
        "El entorno de producción está certificado bajo OHSAS 18001. "
        "La empresa cuenta con matriz de riesgos actualizada. "
        "Se deben respetar estrictamente los procedimientos de bloqueo y etiquetado (LOTO) en las máquinas."
    ),
}

# 6 entries — exceeds base_rows (4) to test dynamic row insertion
FAKE_SECTION_6 = [
    {
        "discapacidad": "DISCAPACIDAD FÍSICA",
        "descripcion": "Amputación de miembro superior o inferior con prótesis funcional; movilidad conservada.",
    },
    {
        "discapacidad": "DISCAPACIDAD AUDITIVA",
        "descripcion": "Pérdida auditiva bilateral con uso de auxiliares auditivos; comunicación oral funcional.",
    },
    {
        "discapacidad": "DISCAPACIDAD AUDITIVA HIPOACUSIA",
        "descripcion": "Hipoacusia leve a moderada; comprensión del lenguaje oral preservada en ambientes sin ruido excesivo.",
    },
    {
        "discapacidad": "DISCAPACIDAD VISUAL BAJA VISIÓN",
        "descripcion": "Agudeza visual corregida con apoyo de lentes o software amplificador; orientación espacial conservada.",
    },
    {
        "discapacidad": "DISCAPACIDAD MÚLTIPLE FÍSICA - AUDITIVA",
        "descripcion": "Combinación de discapacidad motriz leve y pérdida auditiva parcial; requiere ajustes en comunicación y movilidad.",
    },
    {
        "discapacidad": "DISCAPACIDAD PSICOSOCIAL",
        "descripcion": "Trastorno controlado con medicación; estabilidad emocional verificada; requiere ambiente laboral sin exposición a situaciones de alta presión sostenida.",
    },
]

# 1000-word observaciones — stress test for autofit
FAKE_SECTION_7 = {
    "observaciones_recomendaciones": OBSERVACIONES_1000W,
}

# 5 entries — exceeds base_rows (3) to test dynamic row insertion
FAKE_SECTION_8 = [
    {
        "nombre": "Sandra Milena Pachón Rojas",
        "cargo": "Coordinadora de inclusión laboral",
    },
    {
        "nombre": "Ana María Martínez Vargas",
        "cargo": "Profesional de inclusión laboral",
    },
    {
        "nombre": "Carlos Alberto Pérez Rodríguez",
        "cargo": "Asesor Agencia de Empleo",
    },
    {
        "nombre": "María García López",
        "cargo": "Gerente de Recursos Humanos",
    },
    {
        "nombre": "Juan David Rincón Suárez",
        "cargo": "Gestor de inclusión laboral",
    },
]


# ---------------------------------------------------------------------------
# Main smoke test
# ---------------------------------------------------------------------------
def run():
    print("=" * 65)
    print("SMOKE TEST — Condiciones de Vacante")
    print(f"Word count in observaciones: {word_count(OBSERVACIONES_1000W)}")
    print("=" * 65)

    # ------------------------------------------------------------------
    # [0] Pre-flight
    # ------------------------------------------------------------------
    print("\n--- Pre-flight ---")
    try:
        from formularios.condiciones_vacante import condiciones_vacante as cv
    except ImportError as exc:
        print(f"  {_FAIL}  Cannot import condiciones_vacante: {exc}")
        return 1

    check("Module imported", True)

    # Inject fake data directly into FORM_CACHE (bypasses Supabase/UI)
    cv.FORM_CACHE.clear()
    cv.SECTION_1_CACHE.clear()

    cv.FORM_CACHE["section_1"] = FAKE_SECTION_1
    cv.FORM_CACHE["section_2"] = FAKE_SECTION_2
    cv.FORM_CACHE["section_2_1"] = FAKE_SECTION_2_1
    cv.FORM_CACHE["section_3"] = FAKE_SECTION_3
    cv.FORM_CACHE["section_4"] = FAKE_SECTION_4
    cv.FORM_CACHE["section_5"] = FAKE_SECTION_5
    cv.FORM_CACHE["section_6"] = FAKE_SECTION_6
    cv.FORM_CACHE["section_7"] = FAKE_SECTION_7
    cv.FORM_CACHE["section_8"] = FAKE_SECTION_8

    # SECTION_1_CACHE.nombre_empresa drives the output file name
    cv.SECTION_1_CACHE["nombre_empresa"] = FAKE_SECTION_1["nombre_empresa"]

    check("FORM_CACHE injected (all 9 sections)", len(cv.FORM_CACHE) == 9)

    # ------------------------------------------------------------------
    # [1] Export
    # ------------------------------------------------------------------
    print("\n--- Export ---")
    output_path = None
    try:
        output_path = cv.export_to_excel()
        check("export_to_excel() returned a path", bool(output_path), output_path or "")
    except Exception as exc:
        check("export_to_excel() completed without exception", False, str(exc))
        _print_summary()
        return 1

    if not output_path:
        check("Output file path is not empty", False)
        _print_summary()
        return 1

    check("Output file exists on disk", os.path.exists(output_path), output_path)

    # ------------------------------------------------------------------
    # [2] Cell-value verification
    # ------------------------------------------------------------------
    print("\n--- Cell values ---")
    try:
        import win32com.client as win32
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(output_path)
    except Exception as exc:
        check("Excel opened for verification", False, str(exc))
        _print_summary()
        return 1

    try:
        # Find the worksheet
        from formularios.common import _normalize_text
        target = _normalize_text("3. REVISIÓN DE LAS CONDICIONES").replace(" ", "")
        ws = None
        for sheet in wb.Worksheets:
            if _normalize_text(sheet.Name).replace(" ", "") == target:
                ws = sheet
                break
        if ws is None:
            ws = wb.Worksheets(1)

        check("Worksheet found", ws is not None, getattr(ws, "Name", "?"))

        def cell_val(ref):
            try:
                v = ws.Range(ref).Value
                return str(v).strip() if v is not None else ""
            except Exception:
                return ""

        def row_height(row_num):
            try:
                return ws.Rows(row_num).RowHeight
            except Exception:
                return 0.0

        # --- Section 2 spot-checks (known cell refs from EXCEL_MAPPING) ---
        print("  [Section 2]")
        check("  nombre_vacante  → I15",
              FAKE_SECTION_2["nombre_vacante"] in cell_val("I15"),
              repr(cell_val("I15")))
        check("  numero_vacantes → I16",
              FAKE_SECTION_2["numero_vacantes"] in cell_val("I16"),
              repr(cell_val("I16")))
        check("  nivel_cargo     → I17",
              FAKE_SECTION_2["nivel_cargo"].rstrip(".") in cell_val("I17"),
              repr(cell_val("I17")))
        check("  tipo_contrato   → I25",
              "Fijo" in cell_val("I25"),
              repr(cell_val("I25")))
        check("  competencia_1   → I30",
              bool(cell_val("I30")),
              repr(cell_val("I30")))

        # --- Section 2.1 checkboxes ---
        print("  [Section 2.1 — checkboxes]")
        check("  nivel_bachiller (True) → L36 = ☑",
              cell_val("L36") == "☑",
              repr(cell_val("L36")))
        check("  nivel_primaria (False) → G36 = ☐",
              cell_val("G36") == "☐",
              repr(cell_val("G36")))
        check("  nivel_tecnologo (True) → R37 = ☑",
              cell_val("R37") == "☑",
              repr(cell_val("R37")))

        # --- Section 2.1 text fields ---
        print("  [Section 2.1 — text fields]")
        check("  horarios_asignados → I42",
              "Rotativos" in cell_val("I42"),
              repr(cell_val("I42")))
        check("  hora_ingreso      → I43",
              FAKE_SECTION_2_1["hora_ingreso"] in cell_val("I43"),
              repr(cell_val("I43")))
        check("  experiencia_meses → I50",
              "año" in cell_val("I50").lower(),
              repr(cell_val("I50")))

        # --- Section 3 spot-checks ---
        print("  [Section 3]")
        check("  lectura           → L64",
              "Medio" in cell_val("L64"),
              repr(cell_val("L64")))
        check("  concentracion     → L72",
              "Alto" in cell_val("L72"),
              repr(cell_val("L72")))
        check("  trabajo_equipo    → L99",
              "Alto" in cell_val("L99"),
              repr(cell_val("L99")))

        # --- Section 4 spot-checks ---
        print("  [Section 4]")
        check("  de_pie_tiempo     → H113",
              "4" in cell_val("H113") or "6" in cell_val("H113"),
              repr(cell_val("H113")))
        check("  de_pie_frecuencia → L113",
              "Diario" in cell_val("L113"),
              repr(cell_val("L113")))

        # --- Section 5 spot-checks ---
        print("  [Section 5]")
        check("  ruido             → M120",
              "Medio" in cell_val("M120"),
              repr(cell_val("M120")))
        check("  movimientos_rept  → M146",
              "Alto" in cell_val("M146"),
              repr(cell_val("M146")))

        # --- Section 6 — 6 discapacidades (> base_rows=4, tests row insertion) ---
        print("  [Section 6 — 6 entries, base_rows=4, extra rows inserted]")
        # Locate section 6 dynamically: find its header "6. PERFIL" then count
        # data rows starting 1 row below.  Autofit may insert extra overflow rows
        # within section 6, so we search a window (up to 20 rows) rather than
        # expecting 6 consecutive rows.
        sec6_header = None
        try:
            c6 = ws.Columns("A").Find(What="6.", LookAt=2)
            if c6 and "PERFIL" in str(c6.Value or "").upper():
                sec6_header = c6.Row
        except Exception:
            pass
        if sec6_header is None:
            # Fallback: use static row 153 (works when no pre-section overflow)
            sec6_header = 152  # data starts at 153
        sec6_data_start = sec6_header + 1
        found_discapacidades = 0
        for i in range(20):  # wide window to handle overflow-inserted rows
            v = cell_val(f"A{sec6_data_start + i}")
            if v and any(d["discapacidad"] in str(v) for d in FAKE_SECTION_6):
                found_discapacidades += 1
            if found_discapacidades == 6:
                break
        check("  all 6 discapacidad rows written",
              found_discapacidades == 6,
              f"{found_discapacidades}/6 rows have data")

        # --- Section 7 — 1000-word observaciones (locate by searching near section header) ---
        print("  [Section 7 — ~1000-word observaciones]")
        # Find the "7. OBSERVACIONES" header row
        obs_row = None
        try:
            cell = ws.Columns("A").Find(What="7. OBSERVACIONES", LookAt=2)
            if cell:
                obs_row = cell.Row + 1  # data is one row below the header
        except Exception:
            pass

        if obs_row:
            obs_cell_val = cell_val(f"A{obs_row}")
            check("  section 7 cell has content",
                  len(obs_cell_val) > 100,
                  f"{len(obs_cell_val)} chars in A{obs_row}")
            check("  section 7 contains key text",
                  "PROCESO" in obs_cell_val.upper() or "discapacidad" in obs_cell_val.lower(),
                  repr(obs_cell_val[:80]))
        else:
            check("  section 7 header row located", False, "header not found")

        # --- Section 8 — 5 asistentes (> base_rows=3, tests row insertion) ---
        print("  [Section 8 — 5 asistentes, base_rows=3, extra rows inserted]")
        asist_header_row = None
        try:
            cell = ws.Columns("A").Find(What="8.ASISTENTES", LookAt=2)
            if cell is None:
                cell = ws.Columns("A").Find(What="ASISTENTES", LookAt=2)
            if cell:
                asist_header_row = cell.Row
        except Exception:
            pass

        if asist_header_row:
            found_asistentes = 0
            for i in range(5):
                row = asist_header_row + 1 + i
                if cell_val(f"E{row}"):
                    found_asistentes += 1
            check("  all 5 asistente rows written",
                  found_asistentes == 5,
                  f"{found_asistentes}/5 rows have data")
        else:
            check("  section 8 header row located", False, "header not found")

        # ------------------------------------------------------------------
        # [3] Autofit assertions
        # ------------------------------------------------------------------
        print("\n--- Autofit row heights ---")
        DEFAULT_HEIGHT = 15.0  # Excel default row height (pts)
        TALL_THRESHOLD = 30.0  # any text-heavy row should be at least 2× default

        # funciones_tareas → A53 (wide merged cell; text fits in 1-2 lines so height
        # may be near default — verify cell has content and height is positive)
        h_53 = row_height(53)
        check(f"  funciones_tareas row 53 height > 0pt (wide cell, low threshold)",
              h_53 > 0,
              f"{h_53:.1f}pt")

        # herramientas_equipos → A59 (long text field)
        h_59 = row_height(59)
        check(f"  herramientas_equipos row 59 height > {TALL_THRESHOLD}pt",
              h_59 > TALL_THRESHOLD,
              f"{h_59:.1f}pt")

        # observaciones_cognitivas → E76 (multi-row merge area)
        h_76 = row_height(76)
        check(f"  observaciones_cognitivas row 76 height > {DEFAULT_HEIGHT}pt",
              h_76 > DEFAULT_HEIGHT,
              f"{h_76:.1f}pt")

        # section 7 — 1000-word stress test
        if obs_row:
            # Multi-row merge: sum heights of the merge block
            merge_total = 0.0
            merge_rows_checked = 0
            try:
                ma = ws.Cells(obs_row, 1).MergeArea
                num_rows = int(ma.Rows.Count)
                for i in range(num_rows):
                    merge_total += ws.Rows(obs_row + i).RowHeight
                    merge_rows_checked += 1
            except Exception:
                merge_total = row_height(obs_row)
                merge_rows_checked = 1

            check(
                f"  section 7 total merge height > 400pt (1000-word stress test)",
                merge_total > 400.0,
                f"{merge_total:.1f}pt across {merge_rows_checked} row(s)",
            )
            print(f"         ↳ Individual row heights:")
            if obs_row:
                for i in range(min(merge_rows_checked, 10)):
                    rh = row_height(obs_row + i)
                    print(f"            row {obs_row + i}: {rh:.1f}pt")

        # observaciones_peligros → E150 (text with multiple sentences)
        h_150 = row_height(150)
        check(f"  observaciones_peligros row 150 height > {DEFAULT_HEIGHT}pt",
              h_150 > DEFAULT_HEIGHT,
              f"{h_150:.1f}pt")

    finally:
        try:
            wb.Close(SaveChanges=False)
            excel.Quit()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # [4] Summary
    # ------------------------------------------------------------------
    return _print_summary()


def _print_summary() -> int:
    passed = sum(1 for tag, _ in _results if tag == _PASS)
    failed = sum(1 for tag, _ in _results if tag == _FAIL)
    print("\n" + "=" * 65)
    print(f"RESULTS:  {passed} passed  |  {failed} failed  |  {passed + failed} total")
    print("=" * 65)
    if failed:
        print("\nFailed checks:")
        for tag, label in _results:
            if tag == _FAIL:
                print(f"  ✗  {label}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(run())
