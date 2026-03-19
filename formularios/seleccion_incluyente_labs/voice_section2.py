from __future__ import annotations

import json
import re
import unicodedata
from datetime import date, datetime
from functools import lru_cache
from pathlib import Path
from typing import Dict, Iterable, List, Optional

from formularios.seleccion_incluyente_labs import seleccion_incluyente


FORM_ID = "seleccion_incluyente_labs"
SECTION_ID = "section_2"
AUDIO_UNIT = "single_candidate"
VOICE_FUNCTION_NAME = "selection-section2-extract"
SCHEMA_VERSION = 1

SECTION2_FIELD_META = {
    field["id"]: dict(field) for field in seleccion_incluyente.SECTION_2.get("fields", [])
}
SECTION2_FIELD_IDS = list(SECTION2_FIELD_META.keys())
SECTION2_FIELD_OPTIONS = {
    field_id: list(meta.get("options") or [])
    for field_id, meta in SECTION2_FIELD_META.items()
}

OPTION_ALIASES = {
    "discapacidad": {
        "discapacidad visual perdida total de la vision": "Discapacidad visual pérdida total de la visión",
        "visual perdida total": "Discapacidad visual pérdida total de la visión",
        "ciego": "Discapacidad visual pérdida total de la visión",
        "ciega": "Discapacidad visual pérdida total de la visión",
        "discapacidad visual baja vision": "Discapacidad visual baja visión",
        "baja vision": "Discapacidad visual baja visión",
        "discapacidad auditiva": "Discapacidad auditiva",
        "auditiva": "Discapacidad auditiva",
        "discapacidad auditiva hipoacusia": "Discapacidad auditiva hipoacusia",
        "hipoacusia": "Discapacidad auditiva hipoacusia",
        "trastorno de espectro autista": "Trastorno de espectro autista",
        "tea": "Trastorno de espectro autista",
        "autismo": "Trastorno de espectro autista",
        "discapacidad intelectual": "Discapacidad intelectual",
        "intelectual": "Discapacidad intelectual",
        "discapacidad fisica": "Discapacidad física",
        "fisica": "Discapacidad física",
        "discapacidad fisica usuario en silla de ruedas": "Discapacidad física usuario en silla de ruedas",
        "fisica usuario en silla de ruedas": "Discapacidad física usuario en silla de ruedas",
        "usuario en silla de ruedas": "Discapacidad física usuario en silla de ruedas",
        "silla de ruedas": "Discapacidad física usuario en silla de ruedas",
        "discapacidad psicosocial": "Discapacidad psicosocial",
        "psicosocial": "Discapacidad psicosocial",
        "discapacidad multiple": "Discapacidad múltiple",
        "multiple": "Discapacidad múltiple",
        "no aplica": "No aplica",
    },
    "resultado_certificado": {
        "aprobado": "Aprobado",
        "no aprobado": "No aprobado",
        "pendiente": "Pendiente",
    },
    "pendiente_otros_oferentes": {
        "si": "Si",
        "sí": "Si",
        "no": "No",
        "por confirmar": "Por Confirmar",
    },
    "cuenta_pension": {
        "si": "Si",
        "sí": "Si",
        "no": "No",
        "por confirmar": "Por Confirmar",
    },
    "tipo_pension": {
        "pension invalidez": "Pension Invalidez",
        "subsidiada": "Subsidiada",
        "especial de vejez": "Especial de vejez",
        "victimas conflicto": "Victimas conflicto",
        "victimas del conflicto": "Victimas conflicto",
        "familiar": "Familiar",
        "regimen especial": "Regimen especial",
        "régimen especial": "Regimen especial",
        "no aplica": "No aplica",
    },
}

GENERIC_LEVEL_OPTIONS = {
    "0": "0. No requiere apoyo.",
    "1": "1. Nivel de apoyo Bajo.",
    "2": "2. Nivel de apoyo medio.",
    "3": "3. Nivel de apoyo alto.",
}

SECTION_4_2A_LEVEL_SOURCES = {
    "desplazamiento_nivel_apoyo": "desplazamiento_modo",
    "ubicacion_nivel_apoyo": "ubicacion_ciudad",
    "dinero_nivel_apoyo": "dinero_manejo",
    "presentacion_nivel_apoyo": "presentacion_personal",
    "comunicacion_escrita_nivel_apoyo": "comunicacion_escrita_apoyo",
    "comunicacion_verbal_nivel_apoyo": "comunicacion_verbal_apoyo",
    "decisiones_nivel_apoyo": "toma_decisiones",
}

SECTION_4_2B_LEVEL_SOURCES = {
    "aseo_nivel_apoyo": "alimentacion",
    "instrumentales_nivel_apoyo": "instrumentales_actividades",
    "actividades_nivel_apoyo": "actividades_apoyo",
    "discriminacion_nivel_apoyo": "discriminacion",
}

SECTION_4_2B_BINARY_GROUPS = {
    "alimentacion": [
        "aseo_criar_apoyo",
        "aseo_comunicacion_apoyo",
        "aseo_ayudas_apoyo",
        "aseo_alimentacion",
        "aseo_movilidad_funcional",
        "aseo_higiene_aseo",
    ],
    "instrumentales_actividades": [
        "instrumentales_criar_apoyo",
        "instrumentales_comunicacion_apoyo",
        "instrumentales_movilidad_apoyo",
        "instrumentales_finanzas",
        "instrumentales_cocina_limpieza",
        "instrumentales_crear_hogar",
        "instrumentales_salud_cuenta_apoyo",
    ],
    "actividades_apoyo": [
        "actividades_esparcimiento_apoyo",
        "actividades_esparcimiento_cuenta_apoyo",
        "actividades_complementarios_apoyo",
        "actividades_complementarios_cuenta_apoyo",
        "actividades_subsidios_cuenta_apoyo",
    ],
    "discriminacion": [
        "discriminacion_violencia_apoyo",
        "discriminacion_violencia_cuenta_apoyo",
        "discriminacion_vulneracion_apoyo",
        "discriminacion_vulneracion_cuenta_apoyo",
    ],
}

SUBSECTION_FIELD_IDS = {
    "section_2_fields": [
        "numero",
        "nombre_oferente",
        "cedula",
        "certificado_porcentaje",
        "discapacidad",
        "telefono_oferente",
        "resultado_certificado",
        "cargo_oferente",
        "nombre_contacto_emergencia",
        "parentesco",
        "telefono_emergencia",
        "fecha_nacimiento",
        "edad",
        "pendiente_otros_oferentes",
        "lugar_firma_contrato",
        "fecha_firma_contrato",
        "cuenta_pension",
        "tipo_pension",
    ],
    "section_3_desarrollo": [
        "desarrollo_actividad",
    ],
    "section_4_1_salud": [
        "medicamentos_nivel_apoyo",
        "medicamentos_conocimiento",
        "medicamentos_horarios",
        "medicamentos_nota",
        "alergias_nivel_apoyo",
        "alergias_tipo",
        "alergias_nota",
        "restriccion_nivel_apoyo",
        "restriccion_conocimiento",
        "restriccion_nota",
        "controles_nivel_apoyo",
        "controles_asistencia",
        "controles_frecuencia",
        "controles_nota",
    ],
    "section_4_2_a_habilidades": [
        "desplazamiento_nivel_apoyo",
        "desplazamiento_modo",
        "desplazamiento_transporte",
        "desplazamiento_nota",
        "ubicacion_nivel_apoyo",
        "ubicacion_ciudad",
        "ubicacion_aplicaciones",
        "ubicacion_nota",
        "dinero_nivel_apoyo",
        "dinero_reconocimiento",
        "dinero_manejo",
        "dinero_medios",
        "dinero_nota",
        "presentacion_nivel_apoyo",
        "presentacion_personal",
        "presentacion_nota",
        "comunicacion_escrita_nivel_apoyo",
        "comunicacion_escrita_apoyo",
        "comunicacion_escrita_nota",
        "comunicacion_verbal_nivel_apoyo",
        "comunicacion_verbal_apoyo",
        "comunicacion_verbal_nota",
        "decisiones_nivel_apoyo",
        "toma_decisiones",
        "toma_decisiones_nota",
    ],
    "section_4_2_b_actividades": [
        "aseo_nivel_apoyo",
        "alimentacion",
        "aseo_criar_apoyo",
        "aseo_comunicacion_apoyo",
        "aseo_ayudas_apoyo",
        "aseo_alimentacion",
        "aseo_movilidad_funcional",
        "aseo_higiene_aseo",
        "aseo_nota",
        "instrumentales_nivel_apoyo",
        "instrumentales_actividades",
        "instrumentales_criar_apoyo",
        "instrumentales_comunicacion_apoyo",
        "instrumentales_movilidad_apoyo",
        "instrumentales_finanzas",
        "instrumentales_cocina_limpieza",
        "instrumentales_crear_hogar",
        "instrumentales_salud_cuenta_apoyo",
        "instrumentales_nota",
        "actividades_nivel_apoyo",
        "actividades_apoyo",
        "actividades_esparcimiento_apoyo",
        "actividades_esparcimiento_cuenta_apoyo",
        "actividades_complementarios_apoyo",
        "actividades_complementarios_cuenta_apoyo",
        "actividades_subsidios_cuenta_apoyo",
        "actividades_nota",
        "discriminacion_nivel_apoyo",
        "discriminacion",
        "discriminacion_violencia_apoyo",
        "discriminacion_violencia_cuenta_apoyo",
        "discriminacion_vulneracion_apoyo",
        "discriminacion_vulneracion_cuenta_apoyo",
        "discriminacion_nota",
    ],
}

SUBSECTION_TITLES = {
    "section_2_fields": "2. Datos del oferente",
    "section_3_desarrollo": "3. Desarrollo de la actividad",
    "section_4_1_salud": "4.1 Condiciones medicas y de salud",
    "section_4_2_a_habilidades": "4.2A Habilidades basicas de la vida diaria",
    "section_4_2_b_actividades": "4.2B Actividades, apoyos y discriminacion",
}

SEMANTIC_SECTION_4_1_BLOCKS = (
    "medications",
    "allergies",
    "restrictions",
    "specialist_controls",
)

SEMANTIC_SECTION_4_2_A_BLOCKS = {
    "mobility": ("support_level", "mode", "transport", "details"),
    "orientation": ("support_level", "city_status", "references_status", "details"),
    "money": ("support_level", "recognition", "management", "mediums", "details"),
    "presentation": ("support_level", "dress_code", "details"),
    "written_communication": ("support_level", "support_status", "details"),
    "verbal_communication": ("support_level", "support_status", "details"),
    "decision_making": ("support_level", "decision_status", "details"),
}

SEMANTIC_SECTION_4_2_B_BLOCKS = {
    "daily_living": (
        "support_level",
        "scope",
        "child_care",
        "communication_systems",
        "assistive_devices",
        "feeding",
        "functional_mobility",
        "hygiene",
        "details",
    ),
    "instrumental": (
        "support_level",
        "scope",
        "child_care",
        "communication_systems",
        "community_mobility",
        "finances",
        "cooking_cleaning",
        "household",
        "health_support",
        "details",
    ),
    "work_activities": (
        "support_level",
        "scope",
        "family_recreation_requires_support",
        "family_recreation_has_support",
        "medical_followup_requires_support",
        "medical_followup_has_support",
        "children_subsidies_has_support",
        "details",
    ),
    "discrimination": (
        "support_level",
        "scope",
        "physical_violence_requires_support",
        "physical_violence_has_support",
        "rights_violation_requires_support",
        "rights_violation_has_support",
        "details",
    ),
}

SEMANTIC_TOP_LEVEL_KEYS = {
    "section_2_identity": {
        "identity": ("document_number", "applicant_phone", "emergency_phone", "birthdate_iso"),
    },
    "section_4_1_health": {
        block_name: ("support_level", *tuple(field for field in fields if field != "support_level"))
        for block_name, fields in {
            "medications": ("support_level", "status", "schedule_status", "details"),
            "allergies": ("support_level", "status", "details"),
            "restrictions": ("support_level", "status", "details"),
            "specialist_controls": ("support_level", "attendance", "frequency", "details"),
        }.items()
    },
    "section_4_2_a_skills": SEMANTIC_SECTION_4_2_A_BLOCKS,
    "section_4_2_b_support": SEMANTIC_SECTION_4_2_B_BLOCKS,
}


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _assets_dir() -> Path:
    return _repo_root() / "supabase" / "functions" / VOICE_FUNCTION_NAME


@lru_cache(maxsize=1)
def load_subsection_specs() -> dict:
    path = _assets_dir() / "subsection_specs.json"
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle) or {}
    return payload


def get_subsection_spec(subsection_key: str) -> dict:
    specs = load_subsection_specs().get("subsections") or {}
    if subsection_key not in specs:
        raise KeyError(f"Subseccion Labs desconocida: {subsection_key}")
    return dict(specs[subsection_key])


def build_empty_candidate() -> dict:
    return {field_id: None for field_id in SECTION2_FIELD_IDS}


def build_empty_extraction_payload(subsection_key: str) -> dict:
    if subsection_key not in SUBSECTION_FIELD_IDS:
        raise KeyError(f"Subseccion Labs desconocida: {subsection_key}")
    return {
        "schema_version": SCHEMA_VERSION,
        "form_id": FORM_ID,
        "section_id": SECTION_ID,
        "subsection_key": subsection_key,
        "audio_unit": AUDIO_UNIT,
        "transcription_summary": "",
        "warnings": [],
        "semantic": {
            "section_2_identity": None,
            "section_4_1_health": None,
            "section_4_2_a_skills": None,
            "section_4_2_b_support": None,
        },
        "candidate": build_empty_candidate(),
    }


def _clean_string(value) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _normalize_option_key(value: str) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().strip()
    text = re.sub(r"[^\w\s]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _option_with_prefix(field_id: str, prefix: str) -> Optional[str]:
    options = SECTION2_FIELD_OPTIONS.get(field_id) or []
    target = str(prefix or "").strip()
    for option in options:
        if _normalize_option_key(option).startswith(f"{target} "):
            return option
    return None


def _option_no_aplica(field_id: str) -> Optional[str]:
    options = SECTION2_FIELD_OPTIONS.get(field_id) or []
    for option in options:
        if "no aplica" in _normalize_option_key(option):
            return option
    return None


def _options_are_binary(field_id: str) -> bool:
    options = SECTION2_FIELD_OPTIONS.get(field_id) or []
    return {str(item).strip() for item in options} == {"Si", "No", "No aplica"}


def _option_from_binary_phrase(field_id: str, normalized: str) -> Optional[str]:
    if not _options_are_binary(field_id) or not normalized:
        return None
    if "no aplica" in normalized:
        return "No aplica"
    if "no sabe" in normalized or "sin dato" in normalized:
        return None
    if "no cuenta con apoyo" in normalized:
        return "No"
    if "cuenta con apoyo" in normalized:
        return "Si"
    if "no requiere apoyo" in normalized or "sin apoyo" in normalized:
        return "No"
    if "requiere apoyo" in normalized or "necesita apoyo" in normalized:
        return "Si"
    if normalized in {"si", "sii", "si aplica"}:
        return "Si"
    if normalized in {"no", "ninguno", "ninguna"}:
        return "No"
    if re.search(r"\bsi\b", normalized):
        return "Si"
    if re.search(r"\bno\b", normalized):
        return "No"
    return None


def _extract_numeric_option_prefix(value) -> Optional[str]:
    cleaned = _clean_string(value)
    if not cleaned:
        return None
    match = re.match(r"\s*([0-3])\.", cleaned)
    if match:
        return match.group(1)
    return None


def _derive_generic_level_from_main_value(level_field_id: str, main_value) -> Optional[str]:
    prefix = _extract_numeric_option_prefix(main_value)
    if prefix in GENERIC_LEVEL_OPTIONS:
        return _option_with_prefix(level_field_id, prefix)
    normalized = _normalize_option_key(main_value or "")
    if "no aplica" in normalized:
        return _option_no_aplica(level_field_id)
    return None


def _set_if_missing(updates: dict, field_id: str, value) -> None:
    if value is None or field_id in updates:
        return
    updates[field_id] = value


def _fill_binary_group_defaults(updates: dict, group_field_id: str, negative_value: str) -> None:
    subitems = SECTION_4_2B_BINARY_GROUPS.get(group_field_id) or []
    for field_id in subitems:
        _set_if_missing(updates, field_id, negative_value)


def _semantic_support_level_to_option(field_id: str, value) -> Optional[str]:
    normalized = _normalize_semantic_value(value)
    mapping = {
        "none": "0",
        "low": "1",
        "medium": "2",
        "high": "3",
    }
    if normalized in mapping:
        return _option_with_prefix(field_id, mapping[normalized])
    if normalized == "not applicable" or normalized == "not_applicable":
        return _option_no_aplica(field_id)
    return None


def _normalize_semantic_value(value) -> str:
    return _normalize_option_key(value or "").replace("_", " ")


def _get_semantic_section(payload: dict, section_key: str) -> Optional[dict]:
    semantic = payload.get("semantic")
    if not isinstance(semantic, dict):
        return None
    section = semantic.get(section_key)
    if not isinstance(section, dict):
        return None
    return section


def _get_section_4_1_semantic(payload: dict) -> Optional[dict]:
    return _get_semantic_section(payload, "section_4_1_health")


def _get_section_2_semantic(payload: dict) -> Optional[dict]:
    return _get_semantic_section(payload, "section_2_identity")


def _get_section_4_2_a_semantic(payload: dict) -> Optional[dict]:
    return _get_semantic_section(payload, "section_4_2_a_skills")


def _get_section_4_2_b_semantic(payload: dict) -> Optional[dict]:
    return _get_semantic_section(payload, "section_4_2_b_support")


def _semantic_binary_to_option(value) -> Optional[str]:
    normalized = _normalize_semantic_value(value)
    mapping = {
        "yes": "Si",
        "no": "No",
        "not applicable": "No aplica",
    }
    return mapping.get(normalized)


def _semantic_scope_to_option(field_id: str, value, mapping: dict) -> Optional[str]:
    normalized = _normalize_semantic_value(value)
    prefix = mapping.get(normalized)
    if prefix is None:
        return None
    if prefix == "na":
        return _option_no_aplica(field_id)
    return _option_with_prefix(field_id, prefix)


def _semantic_money_mediums_to_option(values) -> Optional[str]:
    if values is None:
        return None
    if isinstance(values, str):
        normalized_values = {_normalize_semantic_value(values)}
    elif isinstance(values, list):
        normalized_values = {
            _normalize_semantic_value(item)
            for item in values
            if isinstance(item, str) and _clean_string(item)
        }
    else:
        return None

    if not normalized_values:
        return None

    has_fisico = any(item in {"cash", "physical", "cash physical"} for item in normalized_values)
    has_plastico = any(item in {"card", "plastic", "debit card", "credit card"} for item in normalized_values)
    has_digital = any(item in {"digital", "wallet", "transfer"} for item in normalized_values)

    if has_fisico and has_plastico and has_digital:
        return "Dinero fisico, plastico y digital."
    if has_fisico and has_plastico:
        return "Dinero fisico y plastico."
    if has_plastico and has_digital:
        return "Dinero plastico y digital."
    if has_fisico and has_digital:
        return "Dinero digital y fisico."
    if has_fisico:
        return "Dinero fisico."
    if has_plastico:
        return "Dinero plastico."
    if has_digital:
        return "Dinero digital."
    return None


def _apply_section_2_semantic(payload: dict, updates: dict) -> None:
    identity_section = _get_section_2_semantic(payload)
    if not identity_section:
        return
    identity = identity_section.get("identity")
    if not isinstance(identity, dict):
        return

    document_number = _clean_string(identity.get("document_number"))
    if document_number:
        updates["cedula"] = document_number

    applicant_phone = _clean_string(identity.get("applicant_phone"))
    if applicant_phone:
        updates["telefono_oferente"] = applicant_phone

    emergency_phone = _clean_string(identity.get("emergency_phone"))
    if emergency_phone:
        updates["telefono_emergencia"] = emergency_phone

    birthdate_iso = _clean_string(identity.get("birthdate_iso"))
    if birthdate_iso:
        updates["fecha_nacimiento"] = birthdate_iso


def _apply_section_4_1_semantic(payload: dict, updates: dict) -> None:
    health = _get_section_4_1_semantic(payload)
    if not health:
        return

    medications = health.get("medications")
    if isinstance(medications, dict):
        level = _semantic_support_level_to_option("medicamentos_nivel_apoyo", medications.get("support_level"))
        if level is not None:
            updates["medicamentos_nivel_apoyo"] = level
        status = _normalize_option_key(medications.get("status") or "")
        if status == "not taking":
            updates["medicamentos_nivel_apoyo"] = _option_no_aplica("medicamentos_nivel_apoyo")
            updates["medicamentos_conocimiento"] = _option_no_aplica("medicamentos_conocimiento")
            updates["medicamentos_horarios"] = _option_no_aplica("medicamentos_horarios")
        elif status == "self managed":
            updates["medicamentos_conocimiento"] = _option_with_prefix("medicamentos_conocimiento", "1")
        elif status == "third party managed":
            updates["medicamentos_conocimiento"] = _option_with_prefix("medicamentos_conocimiento", "2")
        elif status == "unknown":
            updates["medicamentos_conocimiento"] = _option_with_prefix("medicamentos_conocimiento", "3")
        elif status == "not applicable":
            updates["medicamentos_conocimiento"] = _option_no_aplica("medicamentos_conocimiento")
        schedule_status = _normalize_option_key(medications.get("schedule_status") or "")
        if schedule_status == "self managed":
            updates["medicamentos_horarios"] = _option_with_prefix("medicamentos_horarios", "1")
        elif schedule_status == "third party managed":
            updates["medicamentos_horarios"] = _option_with_prefix("medicamentos_horarios", "2")
        elif schedule_status == "unknown":
            updates["medicamentos_horarios"] = _option_with_prefix("medicamentos_horarios", "3")
        elif schedule_status == "not applicable":
            updates["medicamentos_horarios"] = _option_no_aplica("medicamentos_horarios")
        details = _clean_string(medications.get("details"))
        if details:
            updates["medicamentos_nota"] = details

    allergies = health.get("allergies")
    if isinstance(allergies, dict):
        level = _semantic_support_level_to_option("alergias_nivel_apoyo", allergies.get("support_level"))
        if level is not None:
            updates["alergias_nivel_apoyo"] = level
        status = _normalize_option_key(allergies.get("status") or "")
        if status == "none reported":
            updates["alergias_tipo"] = _option_with_prefix("alergias_tipo", "0")
        elif status == "self managed":
            updates["alergias_tipo"] = _option_with_prefix("alergias_tipo", "1")
        elif status == "unknown":
            updates["alergias_tipo"] = _option_with_prefix("alergias_tipo", "2")
        elif status == "described":
            updates["alergias_tipo"] = _option_with_prefix("alergias_tipo", "3")
        elif status == "not applicable":
            updates["alergias_tipo"] = _option_no_aplica("alergias_tipo")
        details = _clean_string(allergies.get("details"))
        if details:
            updates["alergias_nota"] = details

    restrictions = health.get("restrictions")
    if isinstance(restrictions, dict):
        level = _semantic_support_level_to_option("restriccion_nivel_apoyo", restrictions.get("support_level"))
        if level is not None:
            updates["restriccion_nivel_apoyo"] = level
        status = _normalize_option_key(restrictions.get("status") or "")
        if status == "none reported":
            updates["restriccion_conocimiento"] = _option_with_prefix("restriccion_conocimiento", "0")
        elif status == "self managed":
            updates["restriccion_conocimiento"] = _option_with_prefix("restriccion_conocimiento", "1")
        elif status == "unknown":
            updates["restriccion_conocimiento"] = _option_with_prefix("restriccion_conocimiento", "2")
        elif status == "does not know management":
            updates["restriccion_conocimiento"] = _option_with_prefix("restriccion_conocimiento", "3")
        elif status == "not applicable":
            updates["restriccion_conocimiento"] = _option_no_aplica("restriccion_conocimiento")
        details = _clean_string(restrictions.get("details"))
        if details:
            updates["restriccion_nota"] = details

    controls = health.get("specialist_controls")
    if isinstance(controls, dict):
        level = _semantic_support_level_to_option("controles_nivel_apoyo", controls.get("support_level"))
        if level is not None:
            updates["controles_nivel_apoyo"] = level
        attendance = _normalize_option_key(controls.get("attendance") or "")
        if attendance == "attends and self manages":
            updates["controles_asistencia"] = _option_with_prefix("controles_asistencia", "1")
        elif attendance == "attends":
            updates["controles_asistencia"] = _option_with_prefix("controles_asistencia", "2")
        elif attendance == "unknown":
            updates["controles_asistencia"] = _option_with_prefix("controles_asistencia", "3")
        elif attendance == "not applicable":
            updates["controles_asistencia"] = _option_no_aplica("controles_asistencia")
        frequency = _normalize_option_key(controls.get("frequency") or "")
        if frequency == "monthly":
            updates["controles_frecuencia"] = "Mensual"
        elif frequency == "quarterly":
            updates["controles_frecuencia"] = "Trimestral"
        elif frequency == "semiannual":
            updates["controles_frecuencia"] = "Semestral"
        elif frequency == "other":
            updates["controles_frecuencia"] = "Otra frecuencia"
        elif frequency == "not applicable":
            updates["controles_frecuencia"] = _option_no_aplica("controles_frecuencia")
        details = _clean_string(controls.get("details"))
        if details:
            updates["controles_nota"] = details


def _apply_section_4_2_a_semantic(payload: dict, updates: dict) -> None:
    skills = _get_section_4_2_a_semantic(payload)
    if not skills:
        return

    mobility = skills.get("mobility")
    if isinstance(mobility, dict):
        level = _semantic_support_level_to_option("desplazamiento_nivel_apoyo", mobility.get("support_level"))
        if level is not None:
            updates["desplazamiento_nivel_apoyo"] = level
        mode = _normalize_semantic_value(mobility.get("mode"))
        mode_map = {
            "autonomous": "0",
            "temporary support": "1",
            "permanent support": "2",
            "third party support": "3",
            "not applicable": "na",
        }
        option = _semantic_scope_to_option("desplazamiento_modo", mode, mode_map)
        if option is not None:
            updates["desplazamiento_modo"] = option
        transport_map = {
            "walking": "Caminando.",
            "bicycle": "Bicicleta.",
            "mass transit": "Transmilenio, Sitp.",
            "own vehicle": "Vehiculo propio.",
            "special vehicle": "Vehiculo especial.",
            "not applicable": "No aplica.",
        }
        transport = transport_map.get(_normalize_semantic_value(mobility.get("transport")))
        if transport:
            updates["desplazamiento_transporte"] = transport
        details = _clean_string(mobility.get("details"))
        if details:
            updates["desplazamiento_nota"] = details

    orientation = skills.get("orientation")
    if isinstance(orientation, dict):
        level = _semantic_support_level_to_option("ubicacion_nivel_apoyo", orientation.get("support_level"))
        if level is not None:
            updates["ubicacion_nivel_apoyo"] = level
        city_status = _normalize_semantic_value(orientation.get("city_status"))
        city_map = {
            "autonomous": "0",
            "apps": "1",
            "accompanied": "2",
            "does not orient": "3",
        }
        option = _semantic_scope_to_option("ubicacion_ciudad", city_status, city_map)
        if option is not None:
            updates["ubicacion_ciudad"] = option
        references_map = {
            "references": "Se ubica por puntos de referencia y direcciones.",
            "no references": "No se ubica por puntos de referencia.",
            "cardinal points": "Se ubica por puntos cardinales.",
            "not applicable": "No aplica",
        }
        references = references_map.get(_normalize_semantic_value(orientation.get("references_status")))
        if references:
            updates["ubicacion_aplicaciones"] = references
        details = _clean_string(orientation.get("details"))
        if details:
            updates["ubicacion_nota"] = details

    money = skills.get("money")
    if isinstance(money, dict):
        level = _semantic_support_level_to_option("dinero_nivel_apoyo", money.get("support_level"))
        if level is not None:
            updates["dinero_nivel_apoyo"] = level
        recognition_map = {
            "autonomous": "Autonomo.",
            "family support": "Con apoyo familiar.",
        }
        recognition = recognition_map.get(_normalize_semantic_value(money.get("recognition")))
        if recognition:
            updates["dinero_reconocimiento"] = recognition
        management_map = {
            "autonomous": "0",
            "occasional support": "1",
            "recognizes only": "2",
            "does not recognize": "3",
            "not applicable": "na",
        }
        option = _semantic_scope_to_option("dinero_manejo", money.get("management"), management_map)
        if option is not None:
            updates["dinero_manejo"] = option
        medios = _semantic_money_mediums_to_option(money.get("mediums"))
        if medios is not None:
            updates["dinero_medios"] = medios
        details = _clean_string(money.get("details"))
        if details:
            updates["dinero_nota"] = details

    presentation = skills.get("presentation")
    if isinstance(presentation, dict):
        level = _semantic_support_level_to_option("presentacion_nivel_apoyo", presentation.get("support_level"))
        if level is not None:
            updates["presentacion_nivel_apoyo"] = level
        dress_code_map = {
            "appropriate": "0",
            "appropriate with improvement": "1",
            "partially appropriate": "2",
            "inappropriate": "3",
            "not applicable": "na",
        }
        option = _semantic_scope_to_option("presentacion_personal", presentation.get("dress_code"), dress_code_map)
        if option is not None:
            updates["presentacion_personal"] = option
        details = _clean_string(presentation.get("details"))
        if details:
            updates["presentacion_nota"] = details

    written = skills.get("written_communication")
    if isinstance(written, dict):
        level = _semantic_support_level_to_option(
            "comunicacion_escrita_nivel_apoyo",
            written.get("support_level"),
        )
        if level is not None:
            updates["comunicacion_escrita_nivel_apoyo"] = level
        support_map = {
            "knows and uses": "0",
            "uses some": "1",
            "knows not uses": "2",
            "neither": "3",
            "not applicable": "na",
        }
        option = _semantic_scope_to_option(
            "comunicacion_escrita_apoyo",
            written.get("support_status"),
            support_map,
        )
        if option is not None:
            updates["comunicacion_escrita_apoyo"] = option
        details = _clean_string(written.get("details"))
        if details:
            updates["comunicacion_escrita_nota"] = details

    verbal = skills.get("verbal_communication")
    if isinstance(verbal, dict):
        level = _semantic_support_level_to_option(
            "comunicacion_verbal_nivel_apoyo",
            verbal.get("support_level"),
        )
        if level is not None:
            updates["comunicacion_verbal_nivel_apoyo"] = level
        support_map = {
            "knows and uses": "0",
            "uses some": "1",
            "knows not uses": "2",
            "neither": "3",
            "not applicable": "na",
        }
        option = _semantic_scope_to_option(
            "comunicacion_verbal_apoyo",
            verbal.get("support_status"),
            support_map,
        )
        if option is not None:
            updates["comunicacion_verbal_apoyo"] = option
        details = _clean_string(verbal.get("details"))
        if details:
            updates["comunicacion_verbal_nota"] = details

    decisions = skills.get("decision_making")
    if isinstance(decisions, dict):
        level = _semantic_support_level_to_option("decisiones_nivel_apoyo", decisions.get("support_level"))
        if level is not None:
            updates["decisiones_nivel_apoyo"] = level
        decision_map = {
            "autonomous": "0",
            "occasional support": "1",
            "consults third party": "2",
            "requires third party": "3",
            "not applicable": "na",
        }
        option = _semantic_scope_to_option("toma_decisiones", decisions.get("decision_status"), decision_map)
        if option is not None:
            updates["toma_decisiones"] = option
        details = _clean_string(decisions.get("details"))
        if details:
            updates["toma_decisiones_nota"] = details


def _apply_section_4_2_b_semantic(payload: dict, updates: dict) -> None:
    support = _get_section_4_2_b_semantic(payload)
    if not support:
        return

    binary_field_map = {
        "daily_living": {
            "child_care": "aseo_criar_apoyo",
            "communication_systems": "aseo_comunicacion_apoyo",
            "assistive_devices": "aseo_ayudas_apoyo",
            "feeding": "aseo_alimentacion",
            "functional_mobility": "aseo_movilidad_funcional",
            "hygiene": "aseo_higiene_aseo",
        },
        "instrumental": {
            "child_care": "instrumentales_criar_apoyo",
            "communication_systems": "instrumentales_comunicacion_apoyo",
            "community_mobility": "instrumentales_movilidad_apoyo",
            "finances": "instrumentales_finanzas",
            "cooking_cleaning": "instrumentales_cocina_limpieza",
            "household": "instrumentales_crear_hogar",
            "health_support": "instrumentales_salud_cuenta_apoyo",
        },
        "work_activities": {
            "family_recreation_requires_support": "actividades_esparcimiento_apoyo",
            "family_recreation_has_support": "actividades_esparcimiento_cuenta_apoyo",
            "medical_followup_requires_support": "actividades_complementarios_apoyo",
            "medical_followup_has_support": "actividades_complementarios_cuenta_apoyo",
            "children_subsidies_has_support": "actividades_subsidios_cuenta_apoyo",
        },
        "discrimination": {
            "physical_violence_requires_support": "discriminacion_violencia_apoyo",
            "physical_violence_has_support": "discriminacion_violencia_cuenta_apoyo",
            "rights_violation_requires_support": "discriminacion_vulneracion_apoyo",
            "rights_violation_has_support": "discriminacion_vulneracion_cuenta_apoyo",
        },
    }

    daily_living = support.get("daily_living")
    if isinstance(daily_living, dict):
        level = _semantic_support_level_to_option("aseo_nivel_apoyo", daily_living.get("support_level"))
        if level is not None:
            updates["aseo_nivel_apoyo"] = level
        scope_map = {"none": "0", "some": "1", "most": "2", "all": "3", "not applicable": "na"}
        option = _semantic_scope_to_option("alimentacion", daily_living.get("scope"), scope_map)
        if option is not None:
            updates["alimentacion"] = option
        for key, field_id in binary_field_map["daily_living"].items():
            value = _semantic_binary_to_option(daily_living.get(key))
            if value is not None:
                updates[field_id] = value
        details = _clean_string(daily_living.get("details"))
        if details:
            updates["aseo_nota"] = details

    instrumental = support.get("instrumental")
    if isinstance(instrumental, dict):
        level = _semantic_support_level_to_option("instrumentales_nivel_apoyo", instrumental.get("support_level"))
        if level is not None:
            updates["instrumentales_nivel_apoyo"] = level
        scope_map = {"none": "0", "some": "1", "most": "2", "all": "3", "not applicable": "na"}
        option = _semantic_scope_to_option(
            "instrumentales_actividades",
            instrumental.get("scope"),
            scope_map,
        )
        if option is not None:
            updates["instrumentales_actividades"] = option
        for key, field_id in binary_field_map["instrumental"].items():
            value = _semantic_binary_to_option(instrumental.get(key))
            if value is not None:
                updates[field_id] = value
        details = _clean_string(instrumental.get("details"))
        if details:
            updates["instrumentales_nota"] = details

    work_activities = support.get("work_activities")
    if isinstance(work_activities, dict):
        level = _semantic_support_level_to_option("actividades_nivel_apoyo", work_activities.get("support_level"))
        if level is not None:
            updates["actividades_nivel_apoyo"] = level
        scope_map = {"none": "0", "some": "1", "most": "2", "all": "3", "not applicable": "na"}
        option = _semantic_scope_to_option("actividades_apoyo", work_activities.get("scope"), scope_map)
        if option is not None:
            updates["actividades_apoyo"] = option
        for key, field_id in binary_field_map["work_activities"].items():
            value = _semantic_binary_to_option(work_activities.get(key))
            if value is not None:
                updates[field_id] = value
        details = _clean_string(work_activities.get("details"))
        if details:
            updates["actividades_nota"] = details

    discrimination = support.get("discrimination")
    if isinstance(discrimination, dict):
        level = _semantic_support_level_to_option("discriminacion_nivel_apoyo", discrimination.get("support_level"))
        if level is not None:
            updates["discriminacion_nivel_apoyo"] = level
        scope_map = {
            "none": "0",
            "some contexts": "1",
            "repeated": "2",
            "lifelong": "3",
            "not applicable": "na",
        }
        option = _semantic_scope_to_option("discriminacion", discrimination.get("scope"), scope_map)
        if option is not None:
            updates["discriminacion"] = option
        for key, field_id in binary_field_map["discrimination"].items():
            value = _semantic_binary_to_option(discrimination.get(key))
            if value is not None:
                updates[field_id] = value
        details = _clean_string(discrimination.get("details"))
        if details:
            updates["discriminacion_nota"] = details


def _apply_section_4_1_defaults(updates: dict) -> None:
    medications_level = _clean_string(updates.get("medicamentos_nivel_apoyo"))
    if medications_level == _option_with_prefix("medicamentos_nivel_apoyo", "0"):
        _set_if_missing(updates, "medicamentos_conocimiento", _option_with_prefix("medicamentos_conocimiento", "0"))
        _set_if_missing(updates, "medicamentos_horarios", _option_with_prefix("medicamentos_horarios", "0"))
    elif medications_level == _option_no_aplica("medicamentos_nivel_apoyo"):
        _set_if_missing(updates, "medicamentos_conocimiento", _option_no_aplica("medicamentos_conocimiento"))
        _set_if_missing(updates, "medicamentos_horarios", _option_no_aplica("medicamentos_horarios"))

    if _clean_string(updates.get("alergias_nivel_apoyo")) == _option_no_aplica("alergias_nivel_apoyo"):
        _set_if_missing(updates, "alergias_tipo", _option_no_aplica("alergias_tipo"))

    if _clean_string(updates.get("restriccion_nivel_apoyo")) == _option_no_aplica("restriccion_nivel_apoyo"):
        _set_if_missing(updates, "restriccion_conocimiento", _option_no_aplica("restriccion_conocimiento"))

    controls_level = _clean_string(updates.get("controles_nivel_apoyo"))
    if controls_level == _option_with_prefix("controles_nivel_apoyo", "0"):
        _set_if_missing(updates, "controles_asistencia", _option_with_prefix("controles_asistencia", "0"))
    elif controls_level == _option_no_aplica("controles_nivel_apoyo"):
        _set_if_missing(updates, "controles_asistencia", _option_no_aplica("controles_asistencia"))
        _set_if_missing(updates, "controles_frecuencia", _option_no_aplica("controles_frecuencia"))

    if _clean_string(updates.get("controles_asistencia")) == _option_no_aplica("controles_asistencia"):
        _set_if_missing(updates, "controles_frecuencia", _option_no_aplica("controles_frecuencia"))


def _apply_section_4_2_a_defaults(updates: dict) -> None:
    for level_field_id, main_field_id in SECTION_4_2A_LEVEL_SOURCES.items():
        derived_level = _derive_generic_level_from_main_value(level_field_id, updates.get(main_field_id))
        _set_if_missing(updates, level_field_id, derived_level)


def _apply_section_4_2_b_defaults(updates: dict) -> None:
    for level_field_id, main_field_id in SECTION_4_2B_LEVEL_SOURCES.items():
        main_value = _clean_string(updates.get(main_field_id))
        derived_level = _derive_generic_level_from_main_value(level_field_id, main_value)
        _set_if_missing(updates, level_field_id, derived_level)

        prefix = _extract_numeric_option_prefix(main_value)
        if prefix == "0":
            _fill_binary_group_defaults(updates, main_field_id, "No")
        elif "no aplica" in _normalize_option_key(main_value or ""):
            _fill_binary_group_defaults(updates, main_field_id, "No aplica")


def _resolve_dropdown_value(field_id: str, value) -> Optional[str]:
    cleaned = _clean_string(value)
    if cleaned is None:
        return None
    options = SECTION2_FIELD_OPTIONS.get(field_id) or []
    if not options:
        return cleaned
    if cleaned in options:
        return cleaned
    normalized = _normalize_option_key(cleaned)
    for option in options:
        if _normalize_option_key(option) == normalized:
            return option
    aliases = OPTION_ALIASES.get(field_id) or {}
    aliased = aliases.get(normalized)
    if aliased in options:
        return aliased
    keyword_match = _resolve_dropdown_value_by_keywords(field_id, normalized)
    if keyword_match in options:
        return keyword_match
    return cleaned


def _resolve_dropdown_value_by_keywords(field_id: str, normalized: str) -> Optional[str]:
    options = SECTION2_FIELD_OPTIONS.get(field_id) or []
    if not normalized or not options:
        return None

    def _pick(*needles: str):
        normalized_needles = tuple(_normalize_option_key(item) for item in needles)
        for option in options:
            option_key = _normalize_option_key(option)
            if all(needle in option_key for needle in normalized_needles):
                return option
        return None

    binary_value = _option_from_binary_phrase(field_id, normalized)
    if binary_value in options:
        return binary_value

    if field_id in {"medicamentos_nivel_apoyo", "alergias_nivel_apoyo", "restriccion_nivel_apoyo", "controles_nivel_apoyo"}:
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no requiere apoyo")
        if "nivel bajo" in normalized or "apoyo bajo" in normalized:
            return _pick("nivel de apoyo bajo")
        if "nivel medio" in normalized or "apoyo medio" in normalized:
            return _pick("nivel de apoyo medio")
        if "nivel alto" in normalized or "apoyo alto" in normalized:
            return _pick("nivel de apoyo alto")

    if field_id == "medicamentos_conocimiento":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no toma" in normalized and "medicament" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no requiere apoyo")
        if "tercero" in normalized and "conoce" in normalized:
            return _pick("tercero", "conoce")
        if "no conoce" in normalized:
            return _pick("no conoce")
        if "conoce" in normalized and ("medicamento" in normalized or "toma" in normalized):
            return _pick("conoce los medicamentos")

    if field_id == "medicamentos_horarios":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no toma" in normalized and "medicament" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no requiere apoyo")
        if "tercero" in normalized and "horario" in normalized:
            return _pick("tercero", "horarios")
        if "no conoce" in normalized and "horario" in normalized:
            return _pick("no conoce", "horarios")
        if "conoce" in normalized and ("horario" in normalized or "hora" in normalized):
            return _pick("conoce", "horarios")

    if field_id == "alergias_tipo":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if (
            ("no presenta" in normalized or "no tiene" in normalized or "no refiere" in normalized or "tampoco tiene" in normalized)
            and "alerg" in normalized
        ):
            return _pick("no presenta alergias")
        if "no conoce" in normalized and "alerg" in normalized:
            return _pick("no conoce")
        if "sabe" in normalized and ("manejo" in normalized or "manejar" in normalized):
            return _pick("sabe darle manejo")
        if "alerg" in normalized:
            return _pick("presenta alergias a")

    if field_id == "restriccion_conocimiento":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if ("no tiene" in normalized or "no presenta" in normalized) and "restriccion" in normalized:
            return _pick("no tiene restricciones medicas")
        if "no conoce" in normalized and "restriccion" in normalized:
            return _pick("no conoce")
        if (
            ("si tiene" in normalized or "tiene" in normalized)
            and "restriccion" in normalized
            and ("conoce" in normalized and "manejo" in normalized)
        ):
            return _pick("tiene restricciones medicas y conoce su manejo")
        if (
            ("si tiene" in normalized or "tiene" in normalized)
            and "restriccion" in normalized
            and ("desconoce" in normalized or ("no conoce" in normalized and "manejo" in normalized))
        ):
            return _pick("si tiene restricciones medicas y desconoce su manejo")
        if "conoce" in normalized and "manejar" in normalized:
            return _pick("tiene restricciones medicas y conoce su manejo")
        if "conoce" in normalized and "manejo" in normalized:
            return _pick("tiene restricciones medicas y conoce su manejo")

    if field_id == "controles_asistencia":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no requiere apoyo")
        if "no sabe" in normalized and "control" in normalized:
            return _pick("no sabe")
        if "tiene cita" in normalized and ("medic" in normalized or "especialista" in normalized):
            if "conoce" in normalized and "manejo" in normalized:
                return _pick("asiste", "conoce el manejo")
            return _pick("si asiste")
        if "asiste" in normalized and ("conoce el manejo" in normalized or ("conoce" in normalized and "manejo" in normalized)):
            return _pick("asiste", "conoce el manejo")
        if "tiene controles" in normalized and ("especialista" in normalized or "control" in normalized):
            if "conoce" in normalized and "manejo" in normalized:
                return _pick("asiste", "conoce el manejo")
            return _pick("si asiste")
        if "asiste" in normalized and "control" in normalized:
            return _pick("si asiste")

    if field_id == "controles_frecuencia":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "mensual" in normalized or "cada mes" in normalized:
            return _pick("mensual")
        if "trimestral" in normalized or "cada tres meses" in normalized:
            return _pick("trimestral")
        if "semestral" in normalized or "cada seis meses" in normalized:
            return _pick("semestral")
        if "otra frecuencia" in normalized or "otra periodicidad" in normalized:
            return _pick("otra frecuencia")

    if field_id == "desplazamiento_modo":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no se desplaza" in normalized or ("acompanamiento" in normalized and "tercero" in normalized):
            return _pick("no se desplaza")
        if "apoyo permanente" in normalized or "permanente" in normalized:
            return _pick("apoyo permanente")
        if "apoyo temporal" in normalized or "temporal" in normalized:
            return _pick("apoyo temporal")
        if "independiente" in normalized or "autonom" in normalized or "no requiere apoyo" in normalized:
            return _pick("sin necesidad de apoyos")

    if field_id == "desplazamiento_transporte":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "transmilenio" in normalized or "sitp" in normalized:
            return _pick("transmilenio")
        if "bicicleta" in normalized or "bici" in normalized:
            return _pick("bicicleta")
        if "camin" in normalized:
            return _pick("caminando")
        if "vehiculo especial" in normalized or "transporte especial" in normalized:
            return _pick("vehiculo especial")
        if "vehiculo propio" in normalized or "carro propio" in normalized or "moto" in normalized:
            return _pick("vehiculo propio")

    if field_id == "ubicacion_ciudad":
        if "no sabe ubicarse" in normalized:
            return _pick("no sabe ubicarse")
        if "acompanamiento" in normalized or "requiere ayuda para ubicarse" in normalized:
            return _pick("acompanamiento")
        if "maps" in normalized or "waze" in normalized or "aplicacion" in normalized:
            return _pick("uso de aplicaciones")
        if "autonom" in normalized or "independiente" in normalized or "no requiere apoyo" in normalized:
            return _pick("manera autonoma")

    if field_id == "ubicacion_aplicaciones":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "puntos cardinales" in normalized:
            return _pick("puntos cardinales")
        if "no se ubica por puntos" in normalized:
            return _pick("no se ubica por puntos")
        if "puntos de referencia" in normalized or "direcciones" in normalized:
            return _pick("puntos de referencia")

    if field_id == "dinero_reconocimiento":
        if "apoyo familiar" in normalized or ("famil" in normalized and "apoyo" in normalized):
            return _pick("apoyo familiar")
        if "autonom" in normalized or "independiente" in normalized or "no requiere apoyo" in normalized:
            return _pick("autonomo")

    if field_id == "dinero_manejo":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no reconoce" in normalized:
            return _pick("no reconoce")
        if "solo reconoce" in normalized:
            return _pick("solo reconoce")
        if "apoyo ocasional" in normalized or ("requiere apoyo" in normalized and "ocasiones" in normalized):
            return _pick("ocasiones requiere apoyo")
        if "autonom" in normalized or "independiente" in normalized or "no requiere apoyo" in normalized:
            return _pick("manera autonoma")

    if field_id == "dinero_medios":
        has_fisico = any(token in normalized for token in ("fisico", "efectivo", "billete"))
        has_plastico = any(token in normalized for token in ("plastico", "tarjeta", "debito", "credito"))
        has_digital = any(token in normalized for token in ("digital", "transferencia", "nequi", "daviplata", "billetera"))
        if has_fisico and has_plastico and has_digital:
            return "Dinero fisico, plastico y digital."
        if has_fisico and has_plastico:
            return "Dinero fisico y plastico."
        if has_plastico and has_digital:
            return "Dinero plastico y digital."
        if has_fisico and has_digital:
            return "Dinero digital y fisico."
        if has_fisico:
            return "Dinero fisico."
        if has_plastico:
            return "Dinero plastico."
        if has_digital:
            return "Dinero digital."

    if field_id == "presentacion_personal":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no es acorde" in normalized or "no acorde" in normalized:
            return _pick("no es acorde")
        if "medianamente" in normalized:
            return _pick("medianamente acorde")
        if "oportunidades de mejora" in normalized or "mejora" in normalized:
            return _pick("oportunidades de mejora")
        if "acorde al contexto" in normalized or "no requiere apoyo" in normalized:
            return _pick("acorde al contexto")

    if field_id in {"comunicacion_escrita_apoyo", "comunicacion_verbal_apoyo"}:
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no aplica")
        if "no conoce" in normalized and "maneja" in normalized:
            return _pick("no conoce", "maneja")
        if "conoce" in normalized and "no maneja" in normalized:
            return _pick("conoce pero no maneja")
        if "maneja algunos" in normalized or ("algunos apoyos" in normalized and "no todos" in normalized):
            return _pick("maneja algunos")
        if "conoce" in normalized and "maneja" in normalized:
            return _pick("si conoce y maneja")

    if field_id == "toma_decisiones":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "requiere el apoyo de un tercero" in normalized:
            return _pick("requiere el apoyo de un tercero")
        if "debe consultar" in normalized or "consulta con un tercero" in normalized:
            return _pick("debe consultar")
        if "en ocasiones requiere" in normalized or "a veces consulta" in normalized or "ocasiones requiere" in normalized:
            return _pick("en ocasiones requiere")
        if "autonom" in normalized or "independiente" in normalized or "no requiere apoyo" in normalized:
            return _pick("manera autonoma")

    if field_id in {"alimentacion", "instrumentales_actividades", "actividades_apoyo"}:
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no requiere apoyo")
        if "mayoria" in normalized:
            return _pick("mayoria")
        if "todas" in normalized:
            return _pick("todas")
        if "algunas" in normalized or "solo" in normalized or "unicamente" in normalized:
            return _pick("algunas")

    if field_id == "discriminacion":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no ha sufrido" in normalized or "sin discriminacion" in normalized:
            return _pick("no ha sufrido")
        if "algunos contextos" in normalized:
            return _pick("algunos contextos")
        if "repetidas ocasiones" in normalized:
            return _pick("repetidas ocasiones")
        if "ciclo vital" in normalized:
            return _pick("ciclo vital")

    return None


def normalize_birthdate_text(value) -> Optional[str]:
    text = _clean_string(value)
    if not text:
        return None
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) == 8:
        day = int(digits[:2])
        month = int(digits[2:4])
        year = int(digits[4:])
        try:
            return date(year, month, day).strftime("%d/%m/%Y")
        except ValueError:
            pass
    formats = (
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d.%m.%Y",
        "%Y/%m/%d",
    )
    for fmt in formats:
        try:
            parsed = datetime.strptime(text, fmt).date()
            return parsed.strftime("%d/%m/%Y")
        except ValueError:
            continue
    normalized = _normalize_option_key(text)
    month_map = {
        "enero": 1,
        "febrero": 2,
        "marzo": 3,
        "abril": 4,
        "mayo": 5,
        "junio": 6,
        "julio": 7,
        "agosto": 8,
        "septiembre": 9,
        "setiembre": 9,
        "octubre": 10,
        "noviembre": 11,
        "diciembre": 12,
    }
    match = re.search(
        r"\b(\d{1,2})\s*(?:de\s+)?([a-z]+)\s*(?:de\s+)?(\d{4})\b",
        normalized,
    )
    if match:
        day = int(match.group(1))
        month = month_map.get(match.group(2))
        year = int(match.group(3))
        if month is not None:
            try:
                return date(year, month, day).strftime("%d/%m/%Y")
            except ValueError:
                return None
    return text


def normalize_birthdate_strict(value) -> Optional[str]:
    normalized = normalize_birthdate_text(value)
    if not normalized:
        return None
    try:
        parsed = datetime.strptime(normalized, "%d/%m/%Y").date()
    except ValueError:
        return None
    if parsed.year < 1900 or parsed > date.today():
        return None
    return parsed.strftime("%d/%m/%Y")


def _digits_only_text(value, *, min_len: int = 1, max_len: Optional[int] = None) -> Optional[str]:
    cleaned = _clean_string(value)
    if cleaned is None:
        return None
    digits = "".join(ch for ch in cleaned if ch.isdigit())
    if not digits:
        return None
    if len(digits) < int(min_len):
        return None
    if max_len is not None and len(digits) > int(max_len):
        return None
    return digits


def normalize_phone_text(value) -> Optional[str]:
    return _digits_only_text(value, min_len=10, max_len=10)


def normalize_document_text(value) -> Optional[str]:
    return _digits_only_text(value, min_len=5, max_len=10)


def normalize_percentage_text(value) -> Optional[str]:
    cleaned = _clean_string(value)
    if cleaned is None:
        return None
    lowered = cleaned.lower().replace("por ciento", "").replace("%", "").strip()
    normalized = lowered.replace(",", ".")
    try:
        number = float(normalized)
    except ValueError:
        return cleaned
    if number < 0 or number > 100:
        return None
    if number.is_integer():
        return str(int(number))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def normalize_date_or_pending_text(value) -> Optional[str]:
    cleaned = _clean_string(value)
    if cleaned is None:
        return None
    normalized = _normalize_option_key(cleaned)
    pending_markers = (
        "por confirmar",
        "pendiente",
        "esperando por confirmar",
        "pendiente por confirmar",
        "sin fecha",
        "aun no",
        "aun no definida",
        "todavia no",
        "todavia pendiente",
    )
    if any(marker in normalized for marker in pending_markers):
        return "Por Confirmar"
    strict_date = normalize_birthdate_strict(cleaned)
    if strict_date is not None:
        return strict_date
    return None


def derive_age_from_birthdate(value) -> Optional[str]:
    normalized = normalize_birthdate_strict(value)
    if not normalized:
        return None
    try:
        parsed = datetime.strptime(normalized, "%d/%m/%Y").date()
    except ValueError:
        return None
    today = date.today()
    if parsed > today:
        return None
    age = today.year - parsed.year
    if (today.month, today.day) < (parsed.month, parsed.day):
        age -= 1
    if age < 0:
        return None
    return str(age)


def validate_extraction_payload(payload) -> List[str]:
    errors: List[str] = []
    if not isinstance(payload, dict):
        return ["El payload debe ser un objeto JSON."]

    required_keys = {
        "schema_version",
        "form_id",
        "section_id",
        "subsection_key",
        "audio_unit",
        "transcription_summary",
        "warnings",
        "semantic",
        "candidate",
    }
    allowed_keys = set(required_keys)
    payload_keys = set(payload.keys())
    missing = sorted(required_keys - payload_keys)
    extra = sorted(payload_keys - allowed_keys)
    if missing:
        errors.append(f"Faltan claves top-level: {', '.join(missing)}.")
    if extra:
        errors.append(f"Hay claves top-level no permitidas: {', '.join(extra)}.")

    if payload.get("schema_version") != SCHEMA_VERSION:
        errors.append(f"schema_version debe ser {SCHEMA_VERSION}.")
    if payload.get("form_id") != FORM_ID:
        errors.append(f"form_id debe ser '{FORM_ID}'.")
    if payload.get("section_id") != SECTION_ID:
        errors.append(f"section_id debe ser '{SECTION_ID}'.")

    subsection_key = payload.get("subsection_key")
    if subsection_key not in SUBSECTION_FIELD_IDS:
        errors.append("subsection_key no es valido.")

    if payload.get("audio_unit") != AUDIO_UNIT:
        errors.append(f"audio_unit debe ser '{AUDIO_UNIT}'.")

    if not isinstance(payload.get("transcription_summary"), str):
        errors.append("transcription_summary debe ser string.")

    warnings = payload.get("warnings")
    if not isinstance(warnings, list) or any(not isinstance(item, str) for item in warnings):
        errors.append("warnings debe ser una lista de strings.")

    semantic = payload.get("semantic")
    if not isinstance(semantic, dict):
        errors.append("semantic debe ser un objeto.")
    else:
        expected_semantic_keys = set(SEMANTIC_TOP_LEVEL_KEYS)
        semantic_keys = set(semantic.keys())
        missing_semantic = sorted(expected_semantic_keys - semantic_keys)
        extra_semantic = sorted(semantic_keys - expected_semantic_keys)
        if missing_semantic:
            errors.append(f"Faltan claves semantic: {', '.join(missing_semantic)}.")
        if extra_semantic:
            errors.append(f"Hay claves semantic no permitidas: {', '.join(extra_semantic)}.")
        for section_key, block_map in SEMANTIC_TOP_LEVEL_KEYS.items():
            section = semantic.get(section_key)
            if section is not None and not isinstance(section, dict):
                errors.append(f"semantic.{section_key} debe ser un objeto o null.")
                continue
            if not isinstance(section, dict):
                continue

            section_fields = set(block_map)
            section_keys = set(section.keys())
            missing_section = sorted(section_fields - section_keys)
            extra_section = sorted(section_keys - section_fields)
            if missing_section:
                errors.append(f"Faltan claves semantic.{section_key}: {', '.join(missing_section)}.")
            if extra_section:
                errors.append(
                    f"Hay claves semantic.{section_key} no permitidas: {', '.join(extra_section)}."
                )

            for block_name, expected_fields in block_map.items():
                block = section.get(block_name)
                if block is not None and not isinstance(block, dict):
                    errors.append(f"semantic.{section_key}.{block_name} debe ser un objeto o null.")
                    continue
                if not isinstance(block, dict):
                    continue

                block_keys = set(block.keys())
                expected_block_keys = set(expected_fields)
                missing_block = sorted(expected_block_keys - block_keys)
                extra_block = sorted(block_keys - expected_block_keys)
                if missing_block:
                    errors.append(
                        f"Faltan claves semantic.{section_key}.{block_name}: {', '.join(missing_block)}."
                    )
                if extra_block:
                    errors.append(
                        f"Hay claves semantic.{section_key}.{block_name} no permitidas: {', '.join(extra_block)}."
                    )
                for key, value in block.items():
                    if value is None:
                        continue
                    if (
                        section_key == "section_4_2_a_skills"
                        and block_name == "money"
                        and key == "mediums"
                    ):
                        if not isinstance(value, list) or any(not isinstance(item, str) for item in value):
                            errors.append(
                                "semantic.section_4_2_a_skills.money.mediums debe ser lista de strings o null."
                            )
                        continue
                    if not isinstance(value, str):
                        errors.append(
                            f"semantic.{section_key}.{block_name}.{key} debe ser string o null."
                        )

    candidate = payload.get("candidate")
    if not isinstance(candidate, dict):
        errors.append("candidate debe ser un objeto.")
        return errors

    candidate_keys = set(candidate.keys())
    expected_candidate_keys = set(SECTION2_FIELD_IDS)
    missing_candidate = sorted(expected_candidate_keys - candidate_keys)
    extra_candidate = sorted(candidate_keys - expected_candidate_keys)
    if missing_candidate:
        errors.append(f"Faltan claves candidate: {', '.join(missing_candidate)}.")
    if extra_candidate:
        errors.append(f"Hay claves candidate no permitidas: {', '.join(extra_candidate)}.")

    for field_id in SECTION2_FIELD_IDS:
        value = candidate.get(field_id)
        if value is not None and not isinstance(value, str):
            errors.append(f"candidate.{field_id} debe ser string o null.")

    return errors


def filter_candidate_fields(candidate: dict, subsection_key: str) -> dict:
    allowed_fields = set(SUBSECTION_FIELD_IDS.get(subsection_key) or [])
    filtered = {}
    for field_id in allowed_fields:
        filtered[field_id] = _resolve_dropdown_value(field_id, (candidate or {}).get(field_id))
    return filtered


def _drop_unresolved_dropdown_values(updates: dict, warnings: list) -> None:
    for field_id in list(updates.keys()):
        options = SECTION2_FIELD_OPTIONS.get(field_id) or []
        if not options:
            continue
        value = _clean_string(updates.get(field_id))
        if value is None:
            updates.pop(field_id, None)
            continue
        if value not in options:
            updates.pop(field_id, None)
            label = SECTION2_FIELD_META.get(field_id, {}).get("label") or field_id
            warnings.append(f"Se descarto {label} porque no coincidio con una opcion valida.")


def postprocess_extraction_payload(payload: dict, *, subsection_key: str, candidate_index: int) -> dict:
    errors = validate_extraction_payload(payload)
    if errors:
        raise ValueError(" | ".join(errors))

    filtered = filter_candidate_fields(payload.get("candidate") or {}, subsection_key)
    updates = {field_id: value for field_id, value in filtered.items() if value is not None}
    updates["numero"] = str(int(candidate_index))

    warnings = []
    for item in payload.get("warnings") or []:
        text = _clean_string(item)
        if text:
            warnings.append(text)

    if subsection_key == "section_2_fields":
        _apply_section_2_semantic(payload, updates)
        normalized_document = normalize_document_text(updates.get("cedula"))
        if normalized_document is not None:
            updates["cedula"] = normalized_document
        elif "cedula" in updates:
            updates.pop("cedula", None)
            warnings.append("La cedula se descarto por formato invalido.")

        normalized_percentage = normalize_percentage_text(updates.get("certificado_porcentaje"))
        if normalized_percentage is not None:
            if "certificado_porcentaje" in updates:
                updates["certificado_porcentaje"] = normalized_percentage
        elif "certificado_porcentaje" in updates:
            updates.pop("certificado_porcentaje", None)
            warnings.append("Se descarto el porcentaje del certificado por formato invalido.")

        for field_id, label in (
            ("telefono_oferente", "telefono del oferente"),
            ("telefono_emergencia", "telefono de emergencia"),
        ):
            normalized_phone = normalize_phone_text(updates.get(field_id))
            if normalized_phone is not None:
                updates[field_id] = normalized_phone
            elif field_id in updates:
                updates.pop(field_id, None)
                warnings.append(f"Se descarto el {label} por formato invalido.")

        fecha_nacimiento = updates.get("fecha_nacimiento")
        normalized_birthdate = normalize_birthdate_strict(fecha_nacimiento)
        if normalized_birthdate:
            updates["fecha_nacimiento"] = normalized_birthdate
            derived_age = derive_age_from_birthdate(normalized_birthdate)
            if derived_age is not None:
                updates["edad"] = derived_age
        else:
            if "fecha_nacimiento" in updates:
                updates.pop("fecha_nacimiento", None)
                warnings.append("Se descarto la fecha de nacimiento por formato invalido.")
            updates.pop("edad", None)

        normalized_contract_date = normalize_date_or_pending_text(updates.get("fecha_firma_contrato"))
        if normalized_contract_date is not None:
            if "fecha_firma_contrato" in updates:
                updates["fecha_firma_contrato"] = normalized_contract_date
        elif "fecha_firma_contrato" in updates:
            updates.pop("fecha_firma_contrato", None)
            warnings.append("Se descarto la fecha de firma de contrato por formato invalido.")

        cuenta_pension = updates.get("cuenta_pension")
        if cuenta_pension == "No":
            updates["tipo_pension"] = "No aplica"
        elif cuenta_pension == "Por Confirmar" and _clean_string(
            (payload.get("candidate") or {}).get("tipo_pension")
        ) is None:
            updates.pop("tipo_pension", None)
    elif subsection_key == "section_4_1_salud":
        _apply_section_4_1_semantic(payload, updates)
        _apply_section_4_1_defaults(updates)
    elif subsection_key == "section_4_2_a_habilidades":
        _apply_section_4_2_a_semantic(payload, updates)
        _apply_section_4_2_a_defaults(updates)
    elif subsection_key == "section_4_2_b_actividades":
        _apply_section_4_2_b_semantic(payload, updates)
        _apply_section_4_2_b_defaults(updates)
    else:
        updates.pop("edad", None)
        if subsection_key != "section_2_fields":
            updates.pop("tipo_pension", None)
            updates.pop("cuenta_pension", None)

    if subsection_key != "section_2_fields":
        updates.pop("edad", None)
        updates.pop("tipo_pension", None)
        updates.pop("cuenta_pension", None)

    _drop_unresolved_dropdown_values(updates, warnings)

    transcription_summary = _clean_string(payload.get("transcription_summary")) or ""

    return {
        "schema_version": SCHEMA_VERSION,
        "form_id": FORM_ID,
        "section_id": SECTION_ID,
        "subsection_key": subsection_key,
        "audio_unit": AUDIO_UNIT,
        "transcription_summary": transcription_summary,
        "warnings": list(dict.fromkeys(warnings)),
        "semantic": payload.get("semantic")
        if isinstance(payload.get("semantic"), dict)
        else {
            "section_2_identity": None,
            "section_4_1_health": None,
            "section_4_2_a_skills": None,
            "section_4_2_b_support": None,
        },
        "candidate": updates,
    }


def merge_non_null_fields(existing: dict, updates: dict) -> dict:
    merged = dict(existing or {})
    for field_id, value in (updates or {}).items():
        cleaned = _clean_string(value)
        if cleaned is None:
            continue
        merged[field_id] = cleaned
    return merged


def summarize_candidate_updates(updates: dict) -> List[str]:
    summary = []
    for field_id in SECTION2_FIELD_IDS:
        value = _clean_string((updates or {}).get(field_id))
        if value is None:
            continue
        label = SECTION2_FIELD_META.get(field_id, {}).get("label") or field_id
        summary.append(f"{label}: {value}")
    return summary


def get_allowed_fields_for_subsection(subsection_key: str) -> List[str]:
    return list(SUBSECTION_FIELD_IDS.get(subsection_key) or [])


def iter_subsection_keys() -> Iterable[str]:
    return SUBSECTION_FIELD_IDS.keys()
