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

    if field_id == "medicamentos_conocimiento":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no requiere apoyo")
        if "tercero" in normalized and "conoce" in normalized:
            return _pick("tercero", "conoce")
        if "no conoce" in normalized:
            return _pick("no conoce")
        if "conoce" in normalized:
            return _pick("conoce los medicamentos")

    if field_id == "medicamentos_horarios":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no requiere apoyo")
        if "tercero" in normalized and "horario" in normalized:
            return _pick("tercero", "horarios")
        if "no conoce" in normalized and "horario" in normalized:
            return _pick("no conoce", "horarios")
        if "conoce" in normalized and "horario" in normalized:
            return _pick("conoce", "horarios")

    if field_id == "alergias_tipo":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no presenta" in normalized and "alerg" in normalized:
            return _pick("no presenta alergias")
        if "no conoce" in normalized and "alerg" in normalized:
            return _pick("no conoce")
        if "sabe" in normalized and ("manejo" in normalized or "manejar" in normalized):
            return _pick("sabe darle manejo")
        if "alerg" in normalized:
            return _pick("presenta alergias a")

    if field_id == "controles_asistencia":
        if "no aplica" in normalized:
            return _pick("no aplica")
        if "no requiere apoyo" in normalized:
            return _pick("no requiere apoyo")
        if "no sabe" in normalized and "control" in normalized:
            return _pick("no sabe")
        if "asiste" in normalized and ("conoce el manejo" in normalized or ("conoce" in normalized and "manejo" in normalized)):
            return _pick("asiste", "conoce el manejo")
        if "asiste" in normalized and "control" in normalized:
            return _pick("si asiste")

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
    return _digits_only_text(value, min_len=7, max_len=10)


def normalize_document_text(value) -> Optional[str]:
    return _digits_only_text(value, min_len=5, max_len=20)


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

    expected_keys = {
        "schema_version",
        "form_id",
        "section_id",
        "subsection_key",
        "audio_unit",
        "transcription_summary",
        "warnings",
        "candidate",
    }
    payload_keys = set(payload.keys())
    missing = sorted(expected_keys - payload_keys)
    extra = sorted(payload_keys - expected_keys)
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
        value = _resolve_dropdown_value(field_id, candidate.get(field_id))
        if value is not None and not isinstance(value, str):
            errors.append(f"candidate.{field_id} debe ser string o null.")
            continue
        if value is None:
            continue
        options = SECTION2_FIELD_OPTIONS.get(field_id) or []
        if options and value not in options:
            errors.append(f"candidate.{field_id} tiene un valor fuera del dropdown permitido.")

    return errors


def filter_candidate_fields(candidate: dict, subsection_key: str) -> dict:
    allowed_fields = set(SUBSECTION_FIELD_IDS.get(subsection_key) or [])
    filtered = {}
    for field_id in allowed_fields:
        filtered[field_id] = _resolve_dropdown_value(field_id, (candidate or {}).get(field_id))
    return filtered


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
    else:
        updates.pop("edad", None)
        if subsection_key != "section_2_fields":
            updates.pop("tipo_pension", None)
            updates.pop("cuenta_pension", None)

    transcription_summary = _clean_string(payload.get("transcription_summary")) or ""

    return {
        "schema_version": SCHEMA_VERSION,
        "form_id": FORM_ID,
        "section_id": SECTION_ID,
        "subsection_key": subsection_key,
        "audio_unit": AUDIO_UNIT,
        "transcription_summary": transcription_summary,
        "warnings": list(dict.fromkeys(warnings)),
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
