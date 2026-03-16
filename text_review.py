from __future__ import annotations

import copy
import json
import os
import time
import urllib.error
import urllib.request
from dataclasses import dataclass

from formularios.common import _load_env_file, _load_supabase_credentials, _supabase_get_access_token
from formularios.presentacion_programa import presentacion_programa
from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.condiciones_vacante import condiciones_vacante
from formularios.induccion_organizacional import induccion_organizacional
from formularios.induccion_operativa import induccion_operativa
from logging_utils import log_app_event


OPENAI_RESPONSES_URL = "https://api.openai.com/v1/responses"
DEFAULT_MODEL = "gpt-4.1-mini"
DEFAULT_TIMEOUT_SECONDS = 20
MAX_TEXT_CHARS = 6000
DEFAULT_EDGE_FUNCTION_NAME = "text-review-orthography"
REVIEW_PROMPT = (
    "Corrige solo ortografia, tildes, signos de puntuacion y uso basico de mayusculas/minusculas. "
    "No resumas, no reformules, no cambies el tono, no inventes informacion y no alteres el sentido del texto. "
    "No cambies nombres propios, numeros, correos, URLs, siglas, articulos legales, referencias normativas, codigos, "
    "ni el formato general de listas o parrafos. Devuelve unicamente el texto final corregido en texto plano."
)


@dataclass
class ReviewResult:
    status: str
    cache: dict
    reason: str = ""
    reviewed_count: int = 0
    elapsed_ms: int = 0


def _log_review(message, level="INFO"):
    try:
        log_app_event(f"[OPENAI_REVIEW] {message}", level=level)
    except Exception:
        pass


def _bool_env(value, default=False):
    raw = str(value or "").strip().lower()
    if not raw:
        return bool(default)
    return raw not in {"0", "false", "no", "off"}


def _read_settings(env_path=".env"):
    env = _load_env_file(env_path)
    enabled = _bool_env(env.get("OPENAI_TEXT_REVIEW_ENABLED") or os.getenv("OPENAI_TEXT_REVIEW_ENABLED"), True)
    api_key = str(os.getenv("OPENAI_API_KEY") or env.get("OPENAI_API_KEY") or "").strip()
    model = str(os.getenv("OPENAI_TEXT_REVIEW_MODEL") or env.get("OPENAI_TEXT_REVIEW_MODEL") or DEFAULT_MODEL).strip() or DEFAULT_MODEL
    function_name = str(
        os.getenv("OPENAI_TEXT_REVIEW_FUNCTION_NAME")
        or env.get("OPENAI_TEXT_REVIEW_FUNCTION_NAME")
        or DEFAULT_EDGE_FUNCTION_NAME
    ).strip() or DEFAULT_EDGE_FUNCTION_NAME
    timeout_raw = str(
        os.getenv("OPENAI_TEXT_REVIEW_TIMEOUT")
        or env.get("OPENAI_TEXT_REVIEW_TIMEOUT")
        or DEFAULT_TIMEOUT_SECONDS
    ).strip()
    try:
        timeout = max(5, int(float(timeout_raw)))
    except Exception:
        timeout = DEFAULT_TIMEOUT_SECONDS
    return {
        "enabled": enabled,
        "api_key": api_key,
        "model": model,
        "function_name": function_name,
        "timeout": timeout,
    }


def _is_meaningful_text(value):
    text = str(value or "").strip()
    if not text:
        return False
    return any(ch.isalpha() for ch in text)


def _extract_error_message(exc):
    if isinstance(exc, urllib.error.HTTPError):
        try:
            body = exc.read().decode("utf-8", errors="replace")
            payload = json.loads(body) if body else {}
            if isinstance(payload, dict):
                err = payload.get("error")
                if isinstance(err, dict):
                    message = str(err.get("message") or "").strip()
                    if message:
                        return message
                for key in ("message", "error_description", "error", "detail"):
                    value = payload.get(key)
                    if value:
                        return str(value)
            if body:
                return body
        except Exception:
            pass
    return str(exc)


def _extract_output_text(payload):
    direct = payload.get("output_text")
    if isinstance(direct, str) and direct.strip():
        return direct.strip()
    output = payload.get("output")
    if not isinstance(output, list):
        return ""
    chunks = []
    for item in output:
        if not isinstance(item, dict):
            continue
        content = item.get("content")
        if not isinstance(content, list):
            continue
        for part in content:
            if not isinstance(part, dict):
                continue
            if part.get("type") in {"output_text", "text"}:
                text = str(part.get("text") or "").strip()
                if text:
                    chunks.append(text)
    return "\n".join(chunks).strip()


def _review_text(text, settings):
    if settings.get("api_key"):
        return _review_text_direct(text, settings)
    return _review_text_via_edge(text, settings)


def _review_text_direct(text, settings):
    body = json.dumps(
        {
            "model": settings["model"],
            "instructions": REVIEW_PROMPT,
            "input": str(text or ""),
        },
        ensure_ascii=False,
    ).encode("utf-8")
    request = urllib.request.Request(
        OPENAI_RESPONSES_URL,
        data=body,
        headers={
            "Authorization": f"Bearer {settings['api_key']}",
            "Content-Type": "application/json",
            "User-Agent": "RECA-Inclusion-Laboral/1.0",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=settings["timeout"]) as response:
        raw = response.read().decode("utf-8", errors="replace")
    payload = json.loads(raw) if raw else {}
    reviewed = _extract_output_text(payload)
    if not reviewed:
        raise RuntimeError("OpenAI no devolvio texto corregido.")
    return reviewed


def _review_text_via_edge(text, settings):
    supabase_url, supabase_key = _load_supabase_credentials(".env")
    jwt_token = str(_supabase_get_access_token(".env") or "").strip()
    if not jwt_token:
        raise RuntimeError("No hay sesión válida para revisar ortografía.")
    function_name = str(settings.get("function_name") or DEFAULT_EDGE_FUNCTION_NAME).strip() or DEFAULT_EDGE_FUNCTION_NAME
    url = f"{supabase_url.rstrip('/')}/functions/v1/{function_name}"
    body = json.dumps(
        {
            "text": str(text or ""),
            "model": settings["model"],
        },
        ensure_ascii=False,
    ).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        headers={
            "apikey": supabase_key,
            "Authorization": f"Bearer {jwt_token}",
            "Content-Type": "application/json",
            "User-Agent": "RECA-Inclusion-Laboral/1.0",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=settings["timeout"]) as response:
        raw = response.read().decode("utf-8", errors="replace")
    payload = json.loads(raw) if raw else {}
    ok = bool(payload.get("ok"))
    reviewed = str(payload.get("text") or "").strip()
    if not ok or not reviewed:
        err = payload.get("error")
        if isinstance(err, dict):
            message = str(err.get("message") or "").strip()
            if message:
                raise RuntimeError(message)
        raise RuntimeError(str(payload.get("message") or "La función de revisión no devolvió texto."))
    return reviewed


def _conditions_section2_1_text_paths():
    paths = []
    for field in condiciones_vacante.SECTION_2_1.get("fields", []):
        if field.get("type") == "texto_largo":
            paths.append(("section_2_1", field["id"]))
    return paths


def _conditions_section3_observation_paths():
    paths = []
    for category in condiciones_vacante.SECTION_3.get("categories", []):
        obs_id = category.get("observaciones_id")
        if obs_id:
            paths.append(("section_3", obs_id))
    return paths


def _evaluacion_section6_paths():
    return [("section_6", field["id"]) for field in evaluacion_accesibilidad.SECTION_6.get("fields", [])]


def _evaluacion_section7_paths():
    return [("section_7", field["id"]) for field in evaluacion_accesibilidad.SECTION_7.get("fields", [])]


def _induccion_operativa_section4_free_text_paths():
    paths = []
    for block in induccion_operativa.SECTION_4.get("blocks", []):
        for item in block.get("items", []):
            if induccion_operativa.SECTION_4_OBSERVACIONES_OPTIONS.get(item.get("row"), []):
                continue
            paths.append(("section_4", "items", item["id"], "observaciones"))
        paths.append(("section_4", "notes", block["id"]))
    return paths


TEXT_REVIEW_FIELDS_BY_FORM = {
    "presentacion_programa": [
        ("section_4", "acuerdos_observaciones"),
    ],
    "evaluacion_accesibilidad": [
        ("section_2_1", "especificaciones_formacion"),
        ("section_2_1", "conocimientos_basicos"),
        ("section_2_1", "observaciones"),
        ("section_2_1", "funciones_tareas"),
        ("section_2_1", "herramientas_equipos"),
        {"section": "section_2_1", "suffixes": ("_observaciones", "_detalle")},
        {"section": "section_2_2", "suffixes": ("_observaciones", "_detalle")},
        {"section": "section_2_3", "suffixes": ("_observaciones", "_detalle")},
        {"section": "section_2_4", "suffixes": ("_observaciones", "_detalle")},
        {"section": "section_2_5", "suffixes": ("_observaciones", "_detalle")},
        {"section": "section_2_6", "suffixes": ("_observaciones", "_detalle")},
        {"section": "section_3", "suffixes": ("_observaciones", "_detalle")},
        {"section": "section_5", "suffixes": ("_nota",)},
        *_evaluacion_section6_paths(),
        *_evaluacion_section7_paths(),
    ],
    "condiciones_vacante": [
        ("section_2", "requiere_certificado_observaciones"),
        *_conditions_section2_1_text_paths(),
        *_conditions_section3_observation_paths(),
        ("section_5", condiciones_vacante.SECTION_5["observaciones"]["id"]),
        ("section_7", condiciones_vacante.SECTION_7["field_id"]),
    ],
    "seleccion_incluyente": [
        ("section_2", "*", "desarrollo_actividad"),
        ("section_5", "ajustes_recomendaciones"),
        ("section_5", "nota"),
    ],
    "contratacion_incluyente": [
        ("section_2", "*", "desarrollo_actividad"),
        ("section_6", "ajustes_recomendaciones"),
    ],
    "induccion_organizacional": [
        ("section_3", "*", "descripcion"),
        ("section_4", "*", "recomendacion"),
        ("section_5", "observaciones"),
    ],
    "induccion_operativa": [
        ("section_3", "*", "observaciones"),
        *_induccion_operativa_section4_free_text_paths(),
        ("section_5", "*", "observaciones"),
        ("section_6", "ajustes_requeridos"),
        ("section_8", "observaciones_recomendaciones"),
    ],
    "sensibilizacion": [
        ("section_3", "observaciones"),
    ],
}


def _iter_path_targets(node, path_spec, current_path=()):
    if not path_spec:
        if isinstance(node, str):
            yield current_path, node
        return
    part = path_spec[0]
    rest = path_spec[1:]
    if part == "*":
        if isinstance(node, list):
            for idx, value in enumerate(node):
                yield from _iter_path_targets(value, rest, current_path + (idx,))
        elif isinstance(node, dict):
            for key, value in node.items():
                yield from _iter_path_targets(value, rest, current_path + (key,))
        return
    if isinstance(node, dict) and part in node:
        yield from _iter_path_targets(node[part], rest, current_path + (part,))


def extract_review_targets(form_id, cache_snapshot):
    specs = TEXT_REVIEW_FIELDS_BY_FORM.get(str(form_id or "").strip(), [])
    if not isinstance(cache_snapshot, dict):
        return []
    targets = []
    seen = set()
    for spec in specs:
        if isinstance(spec, tuple):
            for path, value in _iter_path_targets(cache_snapshot, spec):
                text = str(value or "").strip()
                if not _is_meaningful_text(text) or len(text) > MAX_TEXT_CHARS:
                    continue
                if path in seen:
                    continue
                seen.add(path)
                targets.append({"path": path, "text": text})
            continue
        if not isinstance(spec, dict):
            continue
        section = str(spec.get("section") or "").strip()
        if not section:
            continue
        section_data = cache_snapshot.get(section)
        if not isinstance(section_data, dict):
            continue
        suffixes = tuple(spec.get("suffixes") or ())
        for key, value in section_data.items():
            if not isinstance(key, str):
                continue
            if not suffixes or not key.endswith(suffixes):
                continue
            text = str(value or "").strip()
            if not _is_meaningful_text(text) or len(text) > MAX_TEXT_CHARS:
                continue
            path = (section, key)
            if path in seen:
                continue
            seen.add(path)
            targets.append({"path": path, "text": text})
    return targets


def apply_reviewed_targets(cache_copy, reviewed_targets):
    for item in reviewed_targets:
        path = tuple(item.get("path") or ())
        value = str(item.get("text") or "")
        if not path:
            continue
        target = cache_copy
        for part in path[:-1]:
            if isinstance(part, int):
                if not isinstance(target, list) or part < 0 or part >= len(target):
                    target = None
                    break
                target = target[part]
            else:
                if not isinstance(target, dict) or part not in target:
                    target = None
                    break
                target = target[part]
        if target is None:
            continue
        last = path[-1]
        if isinstance(last, int):
            if isinstance(target, list) and 0 <= last < len(target):
                target[last] = value
        elif isinstance(target, dict):
            target[last] = value
    return cache_copy


def review_export_cache(form_id, cache_snapshot, env_path=".env"):
    started_at = time.perf_counter()
    original_cache = copy.deepcopy(cache_snapshot or {})
    settings = _read_settings(env_path=env_path)
    transport = "direct" if settings.get("api_key") else "edge"
    _log_review(
        f"start form={form_id} transport={transport} model={settings['model']} timeout={settings['timeout']}"
    )
    if not settings["enabled"]:
        _log_review(f"skip form={form_id} reason=disabled")
        return ReviewResult(
            status="skipped",
            cache=original_cache,
            reason="disabled",
            elapsed_ms=int((time.perf_counter() - started_at) * 1000),
        )
    if not settings["api_key"]:
        try:
            _load_supabase_credentials(".env")
            jwt_token = str(_supabase_get_access_token(".env") or "").strip()
        except Exception:
            jwt_token = ""
        if not jwt_token:
            _log_review(f"skip form={form_id} reason=missing_api_key_and_session", level="WARN")
            return ReviewResult(
                status="skipped",
                cache=original_cache,
                reason="missing_api_key_and_session",
                elapsed_ms=int((time.perf_counter() - started_at) * 1000),
            )
    targets = extract_review_targets(form_id, original_cache)
    if not targets:
        _log_review(f"skip form={form_id} reason=no_reviewable_text")
        return ReviewResult(
            status="skipped",
            cache=original_cache,
            reason="no_reviewable_text",
            elapsed_ms=int((time.perf_counter() - started_at) * 1000),
        )
    _log_review(f"targets form={form_id} count={len(targets)} transport={transport}")

    reviewed_texts: dict[str, str] = {}
    reviewed_targets = []
    changed_count = 0
    try:
        for index, target in enumerate(targets, start=1):
            original_text = target["text"]
            path = tuple(target.get("path") or ())
            if original_text in reviewed_texts:
                reviewed_text = reviewed_texts[original_text]
                _log_review(
                    f"reuse_cached_result form={form_id} index={index}/{len(targets)} path={path!r}"
                )
            else:
                _log_review(
                    f"request form={form_id} index={index}/{len(targets)} path={path!r} chars={len(original_text)} transport={transport}"
                )
                reviewed_text = _review_text(original_text, settings)
                reviewed_text = str(reviewed_text or "").strip()
                if not reviewed_text:
                    reviewed_text = original_text
                reviewed_texts[original_text] = reviewed_text
            reviewed_targets.append({"path": path, "text": reviewed_text})
            if reviewed_text != original_text:
                changed_count += 1
    except Exception as exc:
        _log_review(
            f"failed form={form_id} transport={transport} reviewed_count={changed_count} "
            f"path={path!r} error={_extract_error_message(exc)!r}",
            level="ERROR",
        )
        return ReviewResult(
            status="failed",
            cache=original_cache,
            reason=_extract_error_message(exc),
            reviewed_count=changed_count,
            elapsed_ms=int((time.perf_counter() - started_at) * 1000),
        )

    reviewed_cache = apply_reviewed_targets(copy.deepcopy(original_cache), reviewed_targets)
    _log_review(
        f"done form={form_id} transport={transport} targets={len(targets)} "
        f"reviewed_count={changed_count} elapsed_ms={int((time.perf_counter() - started_at) * 1000)}"
    )
    return ReviewResult(
        status="reviewed",
        cache=reviewed_cache,
        reason="ok",
        reviewed_count=changed_count,
        elapsed_ms=int((time.perf_counter() - started_at) * 1000),
    )
