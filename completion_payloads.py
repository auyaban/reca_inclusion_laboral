"""
completion_payloads.py — Construcción y serialización de payloads de finalización.

Responsabilidades:
  - Construir el dict final que se escribe en Supabase al completar un formulario
  - Normalizar fechas, textos, nombres de archivo para el payload
  - PAYLOAD_SCHEMA_VERSION: versión del esquema — cambiar rompe compatibilidad
    con registros ya guardados en completed_forms.json

Depende de: nada interno (solo stdlib)
Usado por: app.py (antes de llamar a google_sheets_client y drive_upload)
"""
from __future__ import annotations

import copy
import hashlib
import os
import re
from urllib.parse import urlparse
import unicodedata
from datetime import UTC, date, datetime
from pathlib import Path
from typing import Any, Callable


PAYLOAD_SCHEMA_VERSION = 1
_PAYLOAD_REASON = "Generado directamente por la app de Inclusion Laboral."
_INTERNAL_CACHE_KEYS = {
    "_section_history",
    "_last_saved_at",
    "_last_saved_section",
    "_last_saved_source",
}


def _clean_text(value: Any) -> str:
    return " ".join(str(value or "").split()).strip()


def _clean_upper_text(value: Any) -> str:
    return _clean_text(value).upper()


def _normalize_filename(value: str) -> str:
    clean = _clean_text(value)
    return clean or "Acta"


def _normalize_date(value: Any) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = _clean_text(value)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return text


def _normalize_number_text(value: Any) -> str:
    text = _clean_text(value)
    if not text:
        return ""
    match = re.search(r"\d+", text.replace(",", ""))
    return match.group(0) if match else text


def _section_dict(cache_snapshot: dict[str, Any], section_id: str) -> dict[str, Any]:
    value = cache_snapshot.get(section_id)
    return dict(value) if isinstance(value, dict) else {}


def _section_list(cache_snapshot: dict[str, Any], section_id: str) -> list[dict[str, Any]]:
    value = cache_snapshot.get(section_id)
    if not isinstance(value, list):
        return []
    rows: list[dict[str, Any]] = []
    for item in value:
        if isinstance(item, dict):
            rows.append(dict(item))
    return rows


def _strip_internal_cache_metadata(cache_snapshot: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(cache_snapshot, dict):
        return {}
    return {
        key: value
        for key, value in cache_snapshot.items()
        if str(key or "") not in _INTERNAL_CACHE_KEYS
    }


def _normalize_asistentes(value: Any) -> list[dict[str, str]]:
    if not isinstance(value, list):
        return []
    asistentes: list[dict[str, str]] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        nombre = _clean_text(item.get("nombre"))
        cargo = _clean_text(item.get("cargo"))
        if not nombre and not cargo:
            continue
        asistentes.append({"nombre": nombre, "cargo": cargo})
    return asistentes


def _assistant_names(asistentes: list[dict[str, str]]) -> list[str]:
    names: list[str] = []
    for item in asistentes:
        nombre = _clean_text(item.get("nombre"))
        if nombre and nombre not in names:
            names.append(nombre)
    return names


def _normalize_participant(entry: dict[str, Any]) -> dict[str, str] | None:
    cedula = _normalize_number_text(entry.get("cedula") or entry.get("cedula_usuario"))
    nombre = _clean_upper_text(entry.get("nombre_oferente") or entry.get("nombre_usuario"))
    discapacidad = _clean_text(entry.get("discapacidad") or entry.get("discapacidad_usuario"))
    genero = _clean_text(entry.get("genero") or entry.get("genero_usuario"))
    if not cedula and not nombre:
        return None
    return {
        "nombre_usuario": nombre,
        "cedula_usuario": cedula,
        "discapacidad_usuario": discapacidad,
        "genero_usuario": genero,
    }


def _normalize_participants(value: Any) -> list[dict[str, str]]:
    if not isinstance(value, list):
        return []
    participants: list[dict[str, str]] = []
    for item in value:
        if not isinstance(item, dict):
            continue
        participant = _normalize_participant(item)
        if participant is not None:
            participants.append(participant)
    return participants


def _unique_cargo(entries: list[dict[str, Any]]) -> str:
    cargos = {
        _clean_text(item.get("cargo_oferente") or item.get("cargo_servicio"))
        for item in entries
        if _clean_text(item.get("cargo_oferente") or item.get("cargo_servicio"))
    }
    return cargos.pop() if len(cargos) == 1 else ""


def _build_attachment(
    *,
    output_path: str,
    form_name: str,
    document_kind: str,
    document_label: str,
) -> dict[str, Any]:
    filename_source = form_name
    clean_output_path = _clean_text(output_path)
    if clean_output_path:
        parsed = urlparse(clean_output_path)
        if parsed.scheme and parsed.netloc:
            if "docs.google.com" not in parsed.netloc or "/spreadsheets/d/" not in parsed.path:
                filename_source = Path(parsed.path).name or form_name
        else:
            filename_source = Path(clean_output_path).name or form_name
    filename = _normalize_filename(filename_source)
    return {
        "filename": filename,
        "document_kind": document_kind,
        "document_label": document_label,
        "is_ods_candidate": True,
        "classification_score": 1.0,
        "classification_reason": _PAYLOAD_REASON,
        "process_hint": "",
        "process_score": 1.0,
    }


def _build_base_parsed(
    section_1: dict[str, Any],
    *,
    asistentes: list[dict[str, str]] | None = None,
    participantes: list[dict[str, str]] | None = None,
    cargo_objetivo: str = "",
    total_vacantes: str = "",
    numero_seguimiento: str = "",
    extra_fields: dict[str, Any] | None = None,
) -> dict[str, Any]:
    asistentes_clean = asistentes or []
    asistentes_names = _assistant_names(asistentes_clean)
    nombre_profesional = _clean_text(section_1.get("profesional_asignado"))
    candidatos = []
    for value in [nombre_profesional, *asistentes_names]:
        clean = _clean_text(value)
        if clean and clean not in candidatos:
            candidatos.append(clean)

    parsed = {
        "nit_empresa": _clean_text(section_1.get("nit_empresa")),
        "nombre_empresa": _clean_text(section_1.get("nombre_empresa")),
        "fecha_servicio": _normalize_date(section_1.get("fecha_visita")),
        "nombre_profesional": nombre_profesional,
        "candidatos_profesional": candidatos,
        "modalidad_servicio": _clean_text(section_1.get("modalidad")),
        "cargo_objetivo": _clean_text(cargo_objetivo),
        "total_vacantes": _normalize_number_text(total_vacantes),
        "numero_seguimiento": _normalize_number_text(numero_seguimiento),
        "participantes": list(participantes or []),
        "warnings": [],
        "asistentes": asistentes_names,
        "ciudad_empresa": _clean_text(section_1.get("ciudad_empresa")),
        "sede_empresa": _clean_text(section_1.get("sede_empresa")),
        "caja_compensacion": _clean_text(section_1.get("caja_compensacion")),
        "asesor_empresa": _clean_text(section_1.get("asesor")),
    }
    if extra_fields:
        parsed.update(extra_fields)
    return parsed


def _presentation_document_kind(section_1: dict[str, Any]) -> tuple[str, str]:
    visita = _normalize_lookup(section_1.get("tipo_visita"))
    if "reactiv" in visita:
        return "program_reactivation", "Reactivacion del programa"
    return "program_presentation", "Presentacion del programa"


def _normalize_lookup(value: Any) -> str:
    text = unicodedata.normalize("NFD", _clean_text(value))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text.casefold()


def _normalize_presentacion(
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    context: dict[str, Any],
) -> dict[str, Any]:
    del context
    section_1 = _section_dict(cache_snapshot, "section_1")
    asistentes = _normalize_asistentes(cache_snapshot.get("section_5"))
    document_kind, document_label = _presentation_document_kind(section_1)
    return {
        "attachment": _build_attachment(
            output_path=output_path,
            form_name=form_name,
            document_kind=document_kind,
            document_label=document_label,
        ),
        "parsed_raw": _build_base_parsed(section_1, asistentes=asistentes),
    }


def _normalize_evaluacion(
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    context: dict[str, Any],
) -> dict[str, Any]:
    del context
    section_1 = _section_dict(cache_snapshot, "section_1")
    asistentes = _normalize_asistentes(cache_snapshot.get("section_8"))
    parsed = _build_base_parsed(
        section_1,
        asistentes=asistentes,
        extra_fields={
            "cargos_compatibles": _clean_text(_section_dict(cache_snapshot, "section_7").get("cargos_compatibles")),
            "observaciones_generales": _clean_text(_section_dict(cache_snapshot, "section_6").get("observaciones_generales")),
        },
    )
    return {
        "attachment": _build_attachment(
            output_path=output_path,
            form_name=form_name,
            document_kind="accessibility_assessment",
            document_label="Evaluacion de accesibilidad",
        ),
        "parsed_raw": parsed,
    }


def _normalize_condiciones(
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    context: dict[str, Any],
) -> dict[str, Any]:
    del context
    section_1 = _section_dict(cache_snapshot, "section_1")
    section_2 = _section_dict(cache_snapshot, "section_2")
    asistentes = _normalize_asistentes(cache_snapshot.get("section_8"))
    return {
        "attachment": _build_attachment(
            output_path=output_path,
            form_name=form_name,
            document_kind="vacancy_review",
            document_label="Revision de condicion o vacante",
        ),
        "parsed_raw": _build_base_parsed(
            section_1,
            asistentes=asistentes,
            cargo_objetivo=_clean_text(section_2.get("nombre_vacante")),
            total_vacantes=_clean_text(section_2.get("numero_vacantes")),
        ),
    }


def _normalize_seleccion(
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    context: dict[str, Any],
) -> dict[str, Any]:
    del context
    section_1 = _section_dict(cache_snapshot, "section_1")
    section_2 = _section_list(cache_snapshot, "section_2")
    asistentes = _normalize_asistentes(cache_snapshot.get("section_6"))
    participants = _normalize_participants(section_2)
    return {
        "attachment": _build_attachment(
            output_path=output_path,
            form_name=form_name,
            document_kind="inclusive_selection",
            document_label="Seleccion incluyente",
        ),
        "parsed_raw": _build_base_parsed(
            section_1,
            asistentes=asistentes,
            participantes=participants,
            cargo_objetivo=_unique_cargo(section_2),
        ),
    }


def _normalize_contratacion(
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    context: dict[str, Any],
) -> dict[str, Any]:
    del context
    section_1 = _section_dict(cache_snapshot, "section_1")
    section_2 = _section_list(cache_snapshot, "section_2")
    asistentes = _normalize_asistentes(cache_snapshot.get("section_7"))
    participants = _normalize_participants(section_2)
    return {
        "attachment": _build_attachment(
            output_path=output_path,
            form_name=form_name,
            document_kind="inclusive_hiring",
            document_label="Contratacion incluyente",
        ),
        "parsed_raw": _build_base_parsed(
            section_1,
            asistentes=asistentes,
            participantes=participants,
            cargo_objetivo=_unique_cargo(section_2),
        ),
    }


def _normalize_induccion_organizacional(
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    context: dict[str, Any],
) -> dict[str, Any]:
    del context
    section_1 = _section_dict(cache_snapshot, "section_1")
    section_2 = _section_list(cache_snapshot, "section_2")
    asistentes = _normalize_asistentes(cache_snapshot.get("section_6"))
    participants = _normalize_participants(section_2)
    return {
        "attachment": _build_attachment(
            output_path=output_path,
            form_name=form_name,
            document_kind="organizational_induction",
            document_label="Induccion organizacional",
        ),
        "parsed_raw": _build_base_parsed(
            section_1,
            asistentes=asistentes,
            participantes=participants,
            cargo_objetivo=_unique_cargo(section_2),
        ),
    }


def _normalize_induccion_operativa(
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    context: dict[str, Any],
) -> dict[str, Any]:
    del context
    section_1 = _section_dict(cache_snapshot, "section_1")
    section_2 = _section_list(cache_snapshot, "section_2")
    asistentes = _normalize_asistentes(cache_snapshot.get("section_9"))
    participants = _normalize_participants(section_2)
    return {
        "attachment": _build_attachment(
            output_path=output_path,
            form_name=form_name,
            document_kind="operational_induction",
            document_label="Induccion operativa",
        ),
        "parsed_raw": _build_base_parsed(
            section_1,
            asistentes=asistentes,
            participantes=participants,
            cargo_objetivo=_unique_cargo(section_2),
        ),
    }


def _normalize_sensibilizacion(
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    context: dict[str, Any],
) -> dict[str, Any]:
    del context
    section_1 = _section_dict(cache_snapshot, "section_1")
    asistentes = _normalize_asistentes(cache_snapshot.get("section_5"))
    return {
        "attachment": _build_attachment(
            output_path=output_path,
            form_name=form_name,
            document_kind="sensibilizacion",
            document_label="Sensibilizacion",
        ),
        "parsed_raw": _build_base_parsed(section_1, asistentes=asistentes),
    }


_STANDARD_NORMALIZERS: dict[str, Callable[..., dict[str, Any]]] = {
    "presentacion_programa": _normalize_presentacion,
    "evaluacion_accesibilidad": _normalize_evaluacion,
    "condiciones_vacante": _normalize_condiciones,
    "seleccion_incluyente": _normalize_seleccion,
    "seleccion_incluyente_labs": _normalize_seleccion,
    "contratacion_incluyente": _normalize_contratacion,
    "induccion_organizacional": _normalize_induccion_organizacional,
    "induccion_operativa": _normalize_induccion_operativa,
    "sensibilizacion": _normalize_sensibilizacion,
}

_NORMALIZER_FORM_ALIASES = {
    "condiciones_vacante_labs": "condiciones_vacante",
}


def build_completion_payload(
    form_id: str,
    form_name: str,
    cache_snapshot: dict[str, Any],
    *,
    output_path: str,
    session_id: str,
    app_version: str,
    extra_context: dict[str, Any] | None = None,
) -> dict[str, Any]:
    clean_form_id = _clean_text(form_id)
    normalized_form_id = _NORMALIZER_FORM_ALIASES.get(clean_form_id, clean_form_id)
    normalizer = _STANDARD_NORMALIZERS.get(normalized_form_id)
    if normalizer is None:
        raise ValueError(f"No existe normalizador de payload para '{form_id}'.")

    generated_at = datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    payload_source = _clean_text((extra_context or {}).get("payload_source")) or "form_cache"
    raw_cache = _strip_internal_cache_metadata(copy.deepcopy(dict(cache_snapshot or {})))
    normalized = normalizer(
        _clean_text(form_name),
        raw_cache,
        output_path=str(output_path or ""),
        context=dict(extra_context or {}),
    )
    normalized.update(
        {
            "schema_version": PAYLOAD_SCHEMA_VERSION,
            "form_id": clean_form_id,
            "form_name": _clean_text(form_name),
            "metadata": {
                "session_id": _clean_text(session_id),
                "generated_at": generated_at,
                "app_version": _clean_text(app_version),
                "payload_source": payload_source,
            },
        }
    )
    payload_raw = {
        "schema_version": PAYLOAD_SCHEMA_VERSION,
        "form_id": clean_form_id,
        "form_name": _clean_text(form_name),
        "output_path": str(output_path or ""),
        "cache_snapshot": raw_cache,
        "metadata": {
            "session_id": _clean_text(session_id),
            "generated_at": generated_at,
            "app_version": _clean_text(app_version),
            "payload_source": payload_source,
        },
    }
    return {
        "payload_raw": payload_raw,
        "payload_normalized": normalized,
        "source_item_key": f"{clean_form_id}:{_clean_text(session_id)}",
    }


def _case_identifier(case_ref: Any, case_meta: dict[str, Any], extra_context: dict[str, Any]) -> str:
    case_record = extra_context.get("case_record")
    raw_id = ""
    if isinstance(case_record, dict):
        raw_id = (
            _clean_text(case_record.get("id"))
            or _clean_text(case_record.get("drive_file_id"))
            or _clean_text(case_record.get("webViewLink"))
        )
    if not raw_id:
        if isinstance(case_ref, dict):
            raw_id = (
                _clean_text(case_ref.get("id"))
                or _clean_text(case_ref.get("webViewLink"))
                or _clean_text(case_ref.get("local_path"))
            )
        else:
            raw_id = os.path.abspath(str(case_ref or ""))
    if not raw_id:
        raw_id = _clean_text(case_meta.get("cedula"))
    digest = hashlib.sha1(raw_id.encode("utf-8")).hexdigest()
    return digest[:20]


def build_followup_completion_payload(
    *,
    case_ref: Any,
    followup_index: int,
    base_payload: dict[str, Any],
    followup_payload: dict[str, Any],
    form_name: str,
    session_id: str,
    app_version: str,
    extra_context: dict[str, Any] | None = None,
) -> dict[str, Any]:
    context = dict(extra_context or {})
    payload_source = _clean_text(context.get("payload_source")) or "seguimientos_sheet"
    generated_at = datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    asistentes = _normalize_asistentes(followup_payload.get("asistentes"))
    participants = _normalize_participants(
        [
            {
                "nombre_usuario": base_payload.get("nombre_vinculado"),
                "cedula_usuario": base_payload.get("cedula"),
                "discapacidad_usuario": base_payload.get("discapacidad"),
                "genero_usuario": "",
                "cargo_servicio": base_payload.get("cargo_vinculado"),
            }
        ]
    )
    parsed = _build_base_parsed(
        dict(base_payload or {}),
        asistentes=asistentes,
        participantes=participants,
        numero_seguimiento=str(followup_index),
        extra_fields={
            "cargo_objetivo": _clean_text(base_payload.get("cargo_vinculado")),
            "seguimiento_numero": _normalize_number_text(followup_payload.get("seguimiento_numero") or followup_index),
            "situacion_encontrada": _clean_text(followup_payload.get("situacion_encontrada")),
            "estrategias_ajustes": _clean_text(followup_payload.get("estrategias_ajustes")),
            "tipo_apoyo": _clean_text(followup_payload.get("tipo_apoyo")),
        },
    )
    output_path = _clean_text(context.get("local_path"))
    attachment = _build_attachment(
        output_path=output_path or _clean_text(context.get("remote_url")),
        form_name=form_name,
        document_kind="follow_up",
        document_label="Seguimiento y acompanamiento",
    )
    case_meta = dict(context.get("case_meta") or {})
    case_identifier = _case_identifier(case_ref, case_meta, context)
    payload_normalized = {
        "schema_version": PAYLOAD_SCHEMA_VERSION,
        "form_id": "seguimientos",
        "form_name": _clean_text(form_name),
        "attachment": attachment,
        "parsed_raw": parsed,
        "metadata": {
            "session_id": _clean_text(session_id),
            "generated_at": generated_at,
            "app_version": _clean_text(app_version),
            "payload_source": payload_source,
            "case_identifier": case_identifier,
            "followup_index": int(followup_index),
        },
    }
    payload_raw = {
        "schema_version": PAYLOAD_SCHEMA_VERSION,
        "form_id": "seguimientos",
        "form_name": _clean_text(form_name),
        "metadata": {
            "session_id": _clean_text(session_id),
            "generated_at": generated_at,
            "app_version": _clean_text(app_version),
            "payload_source": payload_source,
            "case_identifier": case_identifier,
            "followup_index": int(followup_index),
        },
        "base_payload": copy.deepcopy(dict(base_payload or {})),
        "followup_payload": copy.deepcopy(dict(followup_payload or {})),
        "case_meta": copy.deepcopy(case_meta),
        "local_path": output_path,
        "remote_url": _clean_text(context.get("remote_url")),
        "drive_file_id": _clean_text(context.get("drive_file_id")),
    }
    return {
        "payload_raw": payload_raw,
        "payload_normalized": payload_normalized,
        "source_item_key": f"seguimientos:{case_identifier}:followup:{int(followup_index)}",
    }
