import re
from typing import Dict, List

from formularios.common import _normalize_text
from formularios.condiciones_vacante import condiciones_vacante


FORM_ID = "condiciones_vacante"
VOICE_FUNCTION_NAME = "vacancy-section2-extract"
SCHEMA_VERSION = 1

SECTION_2_FIELD_META = {
    field["id"]: dict(field)
    for field in condiciones_vacante.SECTION_2.get("fields", [])
}
SECTION_2_FIELD_META["requiere_certificado_observaciones"] = {
    "id": "requiere_certificado_observaciones",
    "label": "Observaciones (Requiere certificado)",
    "type": "texto_largo",
}
SECTION_2_FIELD_ORDER = [
    "nombre_vacante",
    "numero_vacantes",
    "nivel_cargo",
    "genero",
    "edad",
    "modalidad_trabajo",
    "lugar_trabajo",
    "salario_asignado",
    "firma_contrato",
    "aplicacion_pruebas",
    "tipo_contrato",
    "beneficios_adicionales",
    "cargo_flexible_genero",
    "beneficios_mujeres",
    "requiere_certificado",
    "requiere_certificado_observaciones",
]

SECTION_2_1_FIELD_META = {
    field["id"]: dict(field)
    for field in condiciones_vacante.SECTION_2_1.get("fields", [])
}
SECTION_2_1_FIELD_ORDER = [
    "horarios_asignados",
    "hora_ingreso",
    "hora_salida",
    "tiempo_almuerzo",
    "break_descanso",
    "dias_laborables",
    "dias_flexibles",
    "observaciones",
    "experiencia_meses",
    "funciones_tareas",
    "herramientas_equipos",
]

FIELD_META_BY_SUBSECTION = {
    "section_2_vacancy": SECTION_2_FIELD_META,
    "section_2_1_schedule_experience": SECTION_2_1_FIELD_META,
}
FIELD_ORDER_BY_SUBSECTION = {
    "section_2_vacancy": SECTION_2_FIELD_ORDER,
    "section_2_1_schedule_experience": SECTION_2_1_FIELD_ORDER,
}
FIELD_LABELS_BY_SUBSECTION = {
    subsection_key: {
        field_id: str(meta.get("label") or field_id).strip()
        for field_id, meta in field_meta.items()
    }
    for subsection_key, field_meta in FIELD_META_BY_SUBSECTION.items()
}
LIST_FIELDS_BY_SUBSECTION = {
    subsection_key: {
        field_id
        for field_id, meta in field_meta.items()
        if str(meta.get("type") or "").strip() == "lista"
    }
    for subsection_key, field_meta in FIELD_META_BY_SUBSECTION.items()
}
SECTION_ID_BY_SUBSECTION = {
    "section_2_vacancy": "section_2",
    "section_2_1_schedule_experience": "section_2_1",
}

SUBSECTION_SPECS = {
    "section_2_vacancy": {
        "title": "2. Caracteristicas de la vacante",
        "script": (
            "Responde en un solo audio. Di solo lo que este confirmado. "
            "En los campos abiertos puedes hablar con lenguaje natural; no necesitas usar fechas ni numeros en un formato especial."
        ),
        "questions": [
            "Como se llama la vacante y cuantas vacantes hay. Si puedes, di el numero de vacantes en numero claro.",
            "Que nivel de cargo es, si hay preferencia de genero y que edad o rango de edad buscan.",
            "Cual es la modalidad, el lugar de trabajo y el salario.",
            "Cuando o como seria la firma del contrato, si aplican pruebas y que tipo de contrato es.",
            "Que beneficios adicionales ofrece la empresa.",
            "Si el cargo es flexible segun genero, si hay beneficios para mujeres y si requiere certificado de discapacidad con su observacion si aplica.",
        ],
        "examples": [
            (
                "Vacante auxiliar logistico, son 2 vacantes, cargo operativo, genero indiferente, "
                "entre 20 y 45 anos, trabajo presencial en Fontibon, salario minimo con prestaciones, "
                "contratacion inmediata, si aplican pruebas psicotecnicas, contrato a termino fijo, "
                "beneficios de ruta y alimentacion, el cargo es flexible segun genero, hay beneficios para mujeres, "
                "y el certificado de discapacidad esta en tramite."
            ),
        ],
    },
    "section_2_1_schedule_experience": {
        "title": "2.1 Horarios, experiencia, funciones y herramientas",
        "script": (
            "Responde en un solo audio. Esta subseccion no cubre niveles educativos ni formacion academica. "
            "Enfocate solo en horarios, jornada, experiencia, funciones y herramientas."
        ),
        "questions": [
            "Que horario asignado tiene la vacante: fijo, rotativo o con flexibilizacion.",
            "Cual es la hora de ingreso y la hora de salida.",
            "Cuanto dura el almuerzo y el break o descanso.",
            "Cuales son los dias laborables, si hay dias flexibles y si hay alguna observacion de la jornada.",
            "Cuanta experiencia pide la vacante.",
            "Cuales son las funciones principales y que herramientas o equipos se usan.",
        ],
        "examples": [
            (
                "Horarios rotativos, ingreso a las 6 de la manana y salida a las 2 de la tarde o 10 de la noche "
                "segun rotacion, almuerzo de 1 hora, break de 15 minutos, se trabaja de lunes a sabado con un dia compensatorio, "
                "viernes flexible cada 15 dias, experiencia de 1 ano, funciones operar maquinas, hacer control de calidad y reportar novedades, "
                "herramientas empacadoras, basculas y elementos de proteccion personal."
            ),
        ],
    },
}

SEMANTIC_OPTION_MAPS = {
    "section_2_vacancy": {
        "nivel_cargo": {
            "administrative": "Administrativo.",
            "operational": "Operativo.",
            "services": "Servicios.",
        },
        "genero": {
            "male": "Hombre",
            "female": "Mujer",
            "male_female": "Hombre - Mujer",
            "other": "Otro",
            "indifferent": "Indiferente",
        },
        "tipo_contrato": {
            "fixed_term": "Término Fijo.",
            "indefinite": "Término Indefinido.",
            "work_or_labor": "Obra o Labor.",
            "service_contract": "Prestación de Servicios.",
            "indefinite_presumptive": "Término Indefinido con Cláusula presuntiva.",
            "appointment": "Nombramiento.",
            "apprenticeship": "Contrato de Aprendizaje.",
            "provisional_appointment": "Nombramiento provisional.",
        },
        "requiere_certificado": {
            "yes": "Sí",
            "no": "No",
            "in_process": "En Trámite",
        },
    },
    "section_2_1_schedule_experience": {
        "horarios_asignados": {
            "fixed": "Horarios Fijos.",
            "rotating": "Horarios Rotativos.",
            "flexible": "Flexibilización de horarios",
        },
        "tiempo_almuerzo": {
            "15m": "15 minutos.",
            "30m": "30 minutos.",
            "45m": "45 minutos.",
            "1h": "1 hora.",
            "2h": "2 horas.",
            "not_applicable": "No aplica.",
        },
        "break_descanso": {
            "15m": "15 minutos",
            "30m": "30 minutos",
            "45m": "45 minutos",
            "1h": "1 hora",
            "not_applicable": "No aplica",
        },
        "experiencia_meses": {
            "three_months": "Tres Meses",
            "six_months": "Seis meses.",
            "one_year": "Un año.",
            "one_and_half_years": "Año y medio.",
            "two_years": "Dos Años",
            "two_and_half_years": "Dos años y medio",
            "three_years": "Tres Años",
            "four_years": "Cuatro Años",
            "five_years": "Cinco Años",
            "internships_valid": "Las prácticas son válidas como experiencia laboral.",
            "no_experience": "Sin experiencia laboral.",
            "with_or_without_experience": "Con o Sin Experiencia",
        },
    },
}

OPTION_ALIASES = {
    "section_2_vacancy": {
        "nivel_cargo": {
            "administrativo": "Administrativo.",
            "administrativa": "Administrativo.",
            "cargo administrativo": "Administrativo.",
            "operativo": "Operativo.",
            "operativa": "Operativo.",
            "cargo operativo": "Operativo.",
            "servicios": "Servicios.",
            "servicio": "Servicios.",
            "cargo de servicios": "Servicios.",
        },
        "genero": {
            "hombre": "Hombre",
            "masculino": "Hombre",
            "mujer": "Mujer",
            "femenino": "Mujer",
            "hombre mujer": "Hombre - Mujer",
            "hombres y mujeres": "Hombre - Mujer",
            "hombre y mujer": "Hombre - Mujer",
            "ambos": "Hombre - Mujer",
            "ambos sexos": "Hombre - Mujer",
            "otro": "Otro",
            "indiferente": "Indiferente",
            "sin preferencia": "Indiferente",
            "no hay preferencia": "Indiferente",
        },
        "tipo_contrato": {
            "termino fijo": "Término Fijo.",
            "a termino fijo": "Término Fijo.",
            "fijo": "Término Fijo.",
            "termino indefinido": "Término Indefinido.",
            "a termino indefinido": "Término Indefinido.",
            "indefinido": "Término Indefinido.",
            "obra o labor": "Obra o Labor.",
            "obra labor": "Obra o Labor.",
            "prestacion de servicios": "Prestación de Servicios.",
            "servicios": "Prestación de Servicios.",
            "termino indefinido con clausula presuntiva": "Término Indefinido con Cláusula presuntiva.",
            "clausula presuntiva": "Término Indefinido con Cláusula presuntiva.",
            "nombramiento": "Nombramiento.",
            "contrato de aprendizaje": "Contrato de Aprendizaje.",
            "aprendizaje": "Contrato de Aprendizaje.",
            "nombramiento provisional": "Nombramiento provisional.",
            "provisional": "Nombramiento provisional.",
        },
        "requiere_certificado": {
            "si": "Sí",
            "sí": "Sí",
            "requiere": "Sí",
            "no": "No",
            "no requiere": "No",
            "en tramite": "En Trámite",
            "en trámite": "En Trámite",
            "tramite": "En Trámite",
            "trámite": "En Trámite",
        },
    },
    "section_2_1_schedule_experience": {
        "horarios_asignados": {
            "fijo": "Horarios Fijos.",
            "horario fijo": "Horarios Fijos.",
            "horarios fijos": "Horarios Fijos.",
            "rotativo": "Horarios Rotativos.",
            "rotativos": "Horarios Rotativos.",
            "horario rotativo": "Horarios Rotativos.",
            "horarios rotativos": "Horarios Rotativos.",
            "flexible": "Flexibilización de horarios",
            "flexibilizacion": "Flexibilización de horarios",
            "flexibilizacion de horarios": "Flexibilización de horarios",
            "flexibilización de horarios": "Flexibilización de horarios",
        },
        "tiempo_almuerzo": {
            "15 minutos": "15 minutos.",
            "30 minutos": "30 minutos.",
            "45 minutos": "45 minutos.",
            "1 hora": "1 hora.",
            "una hora": "1 hora.",
            "2 horas": "2 horas.",
            "dos horas": "2 horas.",
            "no aplica": "No aplica.",
        },
        "break_descanso": {
            "15 minutos": "15 minutos",
            "30 minutos": "30 minutos",
            "45 minutos": "45 minutos",
            "1 hora": "1 hora",
            "una hora": "1 hora",
            "no aplica": "No aplica",
        },
        "experiencia_meses": {
            "tres meses": "Tres Meses",
            "3 meses": "Tres Meses",
            "seis meses": "Seis meses.",
            "6 meses": "Seis meses.",
            "un ano": "Un año.",
            "1 ano": "Un año.",
            "un año": "Un año.",
            "1 año": "Un año.",
            "ano y medio": "Año y medio.",
            "año y medio": "Año y medio.",
            "dos anos": "Dos Años",
            "dos años": "Dos Años",
            "2 anos": "Dos Años",
            "2 años": "Dos Años",
            "dos anos y medio": "Dos años y medio",
            "dos años y medio": "Dos años y medio",
            "tres anos": "Tres Años",
            "tres años": "Tres Años",
            "3 anos": "Tres Años",
            "3 años": "Tres Años",
            "cuatro anos": "Cuatro Años",
            "cuatro años": "Cuatro Años",
            "cinco anos": "Cinco Años",
            "cinco años": "Cinco Años",
            "practicas validas": "Las prácticas son válidas como experiencia laboral.",
            "las practicas son validas como experiencia laboral": "Las prácticas son válidas como experiencia laboral.",
            "sin experiencia": "Sin experiencia laboral.",
            "con o sin experiencia": "Con o Sin Experiencia",
        },
    },
}

SEMANTIC_SOURCE_FIELDS = {
    "section_2_vacancy": {
        "nivel_cargo": "position_level",
        "genero": "gender_preference",
        "tipo_contrato": "contract_type",
        "requiere_certificado": "certificate_requirement",
    },
    "section_2_1_schedule_experience": {
        "horarios_asignados": "schedule_type",
        "tiempo_almuerzo": "lunch_duration",
        "break_descanso": "break_duration",
        "experiencia_meses": "experience_requirement",
    },
}

TEXT_SOURCE_FIELDS = {
    "section_2_vacancy": {
        "nombre_vacante": "vacancy_name",
        "numero_vacantes": "openings_count",
        "edad": "age_requirement_text",
        "modalidad_trabajo": "work_modality_text",
        "lugar_trabajo": "work_location_text",
        "salario_asignado": "salary_text",
        "firma_contrato": "contract_signing_text",
        "aplicacion_pruebas": "tests_text",
        "beneficios_adicionales": "additional_benefits_text",
        "cargo_flexible_genero": "gender_flexibility_text",
        "beneficios_mujeres": "women_benefits_text",
        "requiere_certificado_observaciones": "certificate_notes",
    },
    "section_2_1_schedule_experience": {
        "hora_ingreso": "entry_time_text",
        "hora_salida": "exit_time_text",
        "dias_laborables": "workdays_text",
        "dias_flexibles": "flexible_days_text",
        "observaciones": "schedule_notes_text",
        "funciones_tareas": "main_functions_text",
        "herramientas_equipos": "tools_and_equipment_text",
    },
}

SEMANTIC_CONTAINER_KEYS = {
    "section_2_vacancy": ("section_2_vacancy", "vacancy"),
    "section_2_1_schedule_experience": ("section_2_1_schedule_experience", "schedule_experience"),
}

NUMBER_WORDS = {
    "cero": 0,
    "un": 1,
    "uno": 1,
    "una": 1,
    "dos": 2,
    "tres": 3,
    "cuatro": 4,
    "cinco": 5,
    "seis": 6,
    "siete": 7,
    "ocho": 8,
    "nueve": 9,
    "diez": 10,
    "once": 11,
    "doce": 12,
    "trece": 13,
    "catorce": 14,
    "quince": 15,
    "dieciseis": 16,
    "dieciséis": 16,
    "diecisiete": 17,
    "dieciocho": 18,
    "diecinueve": 19,
    "veinte": 20,
}


def _normalize_key(value) -> str:
    text = _normalize_text(str(value or ""))
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9áéíóúñ]+", " ", text)
    return " ".join(text.split())


def _clean_text(value) -> str:
    text = str(value or "").strip()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _resolve_option(subsection_key: str, field_id: str, value) -> str | None:
    field_meta = FIELD_META_BY_SUBSECTION.get(subsection_key) or {}
    text = _clean_text(value)
    if not text:
        return None
    options = list(field_meta.get(field_id, {}).get("options") or [])
    if text in options:
        return text
    normalized = _normalize_key(text)
    for option in options:
        if _normalize_key(option) == normalized:
            return option
    alias = OPTION_ALIASES.get(subsection_key, {}).get(field_id, {}).get(normalized)
    if alias:
        return alias
    return None


def _get_semantic_data(payload: dict, subsection_key: str) -> dict:
    semantic = payload.get("semantic") if isinstance(payload, dict) else None
    if not isinstance(semantic, dict):
        return {}
    container_key, inner_key = SEMANTIC_CONTAINER_KEYS.get(subsection_key, ("", ""))
    section = semantic.get(container_key)
    if not isinstance(section, dict):
        return {}
    inner = section.get(inner_key)
    if not isinstance(inner, dict):
        return {}
    return inner


def _normalize_openings_count(value) -> str | None:
    text = _clean_text(value)
    if not text:
        return None
    digits = re.findall(r"\d+", text)
    if digits:
        return digits[0]
    normalized = _normalize_key(text)
    tokens = normalized.split()
    for token in tokens:
        if token in NUMBER_WORDS:
            return str(NUMBER_WORDS[token])
    if normalized in NUMBER_WORDS:
        return str(NUMBER_WORDS[normalized])
    return text


def get_subsection_spec(subsection_key: str) -> dict:
    key = str(subsection_key or "").strip()
    spec = SUBSECTION_SPECS.get(key)
    if not spec:
        raise KeyError(f"Subseccion no soportada: {key}")
    return dict(spec)


def postprocess_extraction_payload(payload: dict, *, subsection_key: str, candidate_index: int | None = None) -> dict:
    _ = candidate_index
    key = str(subsection_key or "").strip()
    if key not in FIELD_ORDER_BY_SUBSECTION:
        raise ValueError(f"Subseccion no soportada: {key}")
    if not isinstance(payload, dict):
        raise ValueError("La respuesta estructurada no es un objeto JSON.")

    candidate_raw = payload.get("candidate") or {}
    if not isinstance(candidate_raw, dict):
        raise ValueError("candidate debe ser un objeto.")

    updates: Dict[str, str] = {}
    warnings: List[str] = []
    field_order = FIELD_ORDER_BY_SUBSECTION[key]
    field_labels = FIELD_LABELS_BY_SUBSECTION[key]

    for field_id in field_order:
        value = _clean_text(candidate_raw.get(field_id))
        if value:
            updates[field_id] = value

    semantic_data = _get_semantic_data(payload, key)
    for field_id, semantic_key in TEXT_SOURCE_FIELDS.get(key, {}).items():
        semantic_value = _clean_text(semantic_data.get(semantic_key))
        if semantic_value:
            updates[field_id] = semantic_value
    for field_id, semantic_key in SEMANTIC_SOURCE_FIELDS.get(key, {}).items():
        semantic_value = _clean_text(semantic_data.get(semantic_key))
        mapped = SEMANTIC_OPTION_MAPS.get(key, {}).get(field_id, {}).get(semantic_value)
        if mapped:
            updates[field_id] = mapped

    if key == "section_2_vacancy":
        normalized_openings = _normalize_openings_count(
            updates.get("numero_vacantes") or semantic_data.get("openings_count")
        )
        if normalized_openings:
            updates["numero_vacantes"] = normalized_openings

    for field_id in LIST_FIELDS_BY_SUBSECTION.get(key, set()):
        current = updates.get(field_id)
        if not current:
            continue
        resolved = _resolve_option(key, field_id, current)
        if resolved:
            updates[field_id] = resolved
            continue
        warnings.append(
            f"Se descarto {field_labels.get(field_id, field_id)} porque no coincidió con una opción válida."
        )
        updates.pop(field_id, None)

    return {
        "candidate": updates,
        "warnings": warnings,
    }


def summarize_candidate_updates(updates: dict, *, subsection_key: str) -> List[str]:
    lines: List[str] = []
    key = str(subsection_key or "").strip()
    field_order = FIELD_ORDER_BY_SUBSECTION.get(key) or []
    field_labels = FIELD_LABELS_BY_SUBSECTION.get(key) or {}
    data = updates or {}
    for field_id in field_order:
        value = _clean_text(data.get(field_id))
        if not value:
            continue
        label = field_labels.get(field_id, field_id)
        lines.append(f"{label}: {value}")
    return lines
