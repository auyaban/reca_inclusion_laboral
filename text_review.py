from __future__ import annotations

import copy
import json
import os
import re
import socket
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
DEFAULT_MODEL = "gpt-4.1-nano"
DEFAULT_TIMEOUT_SECONDS = 45
TIMEOUT_RETRY_SECONDS = 75
MAX_TEXT_CHARS = 6000
DEFAULT_BATCH_MAX_ITEMS = 8
DEFAULT_BATCH_MAX_CHARS = 12000
DEFAULT_EDGE_FUNCTION_NAME = "text-review-orthography"
LIST_FORMATTING_PROMPT_SUFFIX = (
    "Si el texto ya representa una enumeracion evidente, por ejemplo marcadores numerados, "
    "items en lineas separadas, una frase introductoria terminada en dos puntos seguida de varias lineas cortas, "
    "o elementos claramente separados por punto y coma, puedes "
    "devolverlo como lista simple en texto plano usando prefijos '- '. "
    "Hazlo solo cuando la estructura enumerativa sea inequívoca y no haya riesgo de convertir "
    "un parrafo normal en lista. Si no es inequívoco, conserva el formato original."
)
REVIEW_PROMPT = (
    "Corrige solo ortografia, tildes, signos de puntuacion y uso basico de mayusculas/minusculas. "
    "No resumas, no reformules, no cambies el tono, no inventes informacion y no alteres el sentido del texto. "
    "No cambies nombres propios, numeros, correos, URLs, siglas, articulos legales, referencias normativas, codigos, "
    "ni el formato general de listas o parrafos. "
    f"{LIST_FORMATTING_PROMPT_SUFFIX} "
    "Devuelve unicamente el texto final corregido en texto plano."
)
BATCH_REVIEW_PROMPT = (
    "Corrige solo ortografia, tildes, signos de puntuacion y uso basico de mayusculas/minusculas. "
    "Recibiras un JSON con este formato exacto: {\"items\":[{\"id\":\"...\",\"text\":\"...\"}]}. "
    "Corrige cada campo text por separado, sin mezclar items entre si. "
    "No resumas, no reformules, no cambies el tono, no inventes informacion y no alteres el sentido del texto. "
    "No cambies nombres propios, numeros, correos, URLs, siglas, articulos legales, referencias normativas, codigos, "
    "ni el formato general de listas o parrafos. "
    f"{LIST_FORMATTING_PROMPT_SUFFIX} "
    "Devuelve exclusivamente JSON valido con este mismo formato: {\"items\":[{\"id\":\"...\",\"text\":\"texto corregido\"}]}. "
    "Manten exactamente los mismos ids y la misma cantidad de items. No agregues markdown ni explicaciones."
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
    batch_max_items_raw = str(
        os.getenv("OPENAI_TEXT_REVIEW_BATCH_MAX_ITEMS")
        or env.get("OPENAI_TEXT_REVIEW_BATCH_MAX_ITEMS")
        or DEFAULT_BATCH_MAX_ITEMS
    ).strip()
    batch_max_chars_raw = str(
        os.getenv("OPENAI_TEXT_REVIEW_BATCH_MAX_CHARS")
        or env.get("OPENAI_TEXT_REVIEW_BATCH_MAX_CHARS")
        or DEFAULT_BATCH_MAX_CHARS
    ).strip()
    try:
        batch_max_items = max(1, int(float(batch_max_items_raw)))
    except Exception:
        batch_max_items = DEFAULT_BATCH_MAX_ITEMS
    try:
        batch_max_chars = max(1, int(float(batch_max_chars_raw)))
    except Exception:
        batch_max_chars = DEFAULT_BATCH_MAX_CHARS
    return {
        "enabled": enabled,
        "api_key": api_key,
        "model": model,
        "function_name": function_name,
        "timeout": timeout,
        "batch_max_items": batch_max_items,
        "batch_max_chars": batch_max_chars,
    }


def _is_meaningful_text(value):
    text = str(value or "").strip()
    if not text:
        return False
    return any(ch.isalpha() for ch in text)


_LIST_LINE_RE = re.compile(r"^\s*(?:[-*•]|\d+[.)])\s+(?P<body>.+?)\s*$")
_LIST_INLINE_NUMBER_RE = re.compile(r"(?:(?<=^)|(?<=\s))\d+[.)]\s+")
_LIST_DEFAULT_INDENT = "  "
_LIST_TRAILING_PUNCTUATION = (".", "!", "?", ";", ":")


def _looks_like_list(text):
    lines = [line.strip() for line in str(text or "").splitlines() if line.strip()]
    if len(lines) < 2:
        return False
    matched = sum(1 for line in lines if _LIST_LINE_RE.match(line))
    return matched >= 2 and matched == len(lines)


def _extract_list_items_from_lines(text):
    lines = [line.strip() for line in str(text or "").splitlines() if line.strip()]
    if len(lines) < 2:
        return None
    items = []
    for line in lines:
        match = _LIST_LINE_RE.match(line)
        if not match:
            return None
        item = str(match.group("body") or "").strip()
        if not item:
            return None
        items.append(item)
    return items if len(items) >= 2 else None


def _extract_list_items_from_inline_numbers(text):
    value = " ".join(str(text or "").split())
    matches = list(_LIST_INLINE_NUMBER_RE.finditer(value))
    if len(matches) < 2:
        return None
    items = []
    for idx, match in enumerate(matches):
        start = match.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(value)
        item = value[start:end].strip(" ;")
        if not item:
            return None
        items.append(item)
    return items if len(items) >= 2 else None


def _extract_list_items_from_semicolons(text):
    value = " ".join(str(text or "").split())
    if ";" not in value:
        return None
    parts = [part.strip(" ;") for part in value.split(";") if part.strip(" ;")]
    if len(parts) < 2:
        return None
    if any(len(part.split()) > 18 for part in parts):
        return None
    return parts


def _format_list_items_as_bullets(items):
    values = [str(item or "").strip() for item in items if str(item or "").strip()]
    if len(values) < 2:
        return ""
    return "\n".join(f"{_LIST_DEFAULT_INDENT}- {item}" for item in values)


def _is_plain_list_candidate_line(line_text):
    stripped = str(line_text or "").strip()
    if not stripped:
        return False
    if _LIST_LINE_RE.match(stripped):
        return False
    if len(stripped) > 90:
        return False
    if stripped.endswith(_LIST_TRAILING_PUNCTUATION):
        return False
    words = stripped.split()
    if len(words) > 12:
        return False
    return any(any(ch.isalpha() for ch in word) for word in words)


def _normalize_text_list_blocks(text):
    lines = str(text or "").splitlines()
    if not lines:
        return str(text or "")
    normalized = list(lines)
    index = 0
    while index < len(normalized):
        line = normalized[index]
        match = _LIST_LINE_RE.match(line)
        if match:
            body = str(match.group("body") or "").strip()
            if line.lstrip().startswith(tuple(["*", "•"])):
                normalized[index] = f"{_LIST_DEFAULT_INDENT}- {body}"
            elif not line.startswith((" ", "\t")):
                marker_match = re.match(r"^\s*(?P<marker>[-*•]|\d+[.)])\s+", line)
                marker = marker_match.group("marker") if marker_match else "-"
                if marker in {"*", "•"}:
                    marker = "-"
                normalized[index] = f"{_LIST_DEFAULT_INDENT}{marker} {body}"
            index += 1
            continue

        if not _is_plain_list_candidate_line(line):
            index += 1
            continue

        prev_nonempty = ""
        prev_idx = index - 1
        while prev_idx >= 0:
            prev_line = str(normalized[prev_idx] or "").strip()
            if prev_line:
                prev_nonempty = prev_line
                break
            prev_idx -= 1
        if not prev_nonempty.endswith(":"):
            index += 1
            continue

        end_index = index
        while end_index < len(normalized) and _is_plain_list_candidate_line(normalized[end_index]):
            end_index += 1
        if end_index - index < 2:
            index += 1
            continue
        for line_idx in range(index, end_index):
            normalized[line_idx] = f"{_LIST_DEFAULT_INDENT}- {str(normalized[line_idx]).strip()}"
        index = end_index
    return "\n".join(normalized)


def _maybe_format_reviewed_list(original_text, reviewed_text):
    reviewed = str(reviewed_text or "").strip()
    if not reviewed:
        return reviewed
    if _looks_like_list(reviewed):
        return _normalize_text_list_blocks(reviewed)

    items = (
        _extract_list_items_from_lines(reviewed)
        or _extract_list_items_from_inline_numbers(reviewed)
        or _extract_list_items_from_semicolons(reviewed)
    )
    if items:
        formatted = _format_list_items_as_bullets(items)
        return formatted or reviewed

    normalized_blocks = _normalize_text_list_blocks(reviewed)
    if normalized_blocks != reviewed:
        return normalized_blocks

    if not items:
        original = str(original_text or "").strip()
        if _looks_like_list(original):
            return _normalize_text_list_blocks(reviewed)
        items = (
            _extract_list_items_from_lines(original)
            or _extract_list_items_from_inline_numbers(original)
            or _extract_list_items_from_semicolons(original)
        )
        if not items:
            return reviewed
    formatted = _format_list_items_as_bullets(items)
    return formatted or reviewed


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
        finally:
            try:
                exc.close()
            except Exception:
                pass
    return str(exc)


def _is_timeout_error(exc):
    if isinstance(exc, TimeoutError):
        return True
    if isinstance(exc, socket.timeout):
        return True
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, "reason", None)
        if isinstance(reason, TimeoutError):
            return True
        if isinstance(reason, socket.timeout):
            return True
    message = _extract_error_message(exc).lower()
    return "timed out" in message or "timeout" in message


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


def _strip_code_fences(text):
    value = str(text or "").strip()
    if not value.startswith("```"):
        return value
    lines = value.splitlines()
    if lines and lines[0].strip().startswith("```"):
        lines = lines[1:]
    if lines and lines[-1].strip().startswith("```"):
        lines = lines[:-1]
    return "\n".join(lines).strip()


def _extract_json_candidate(text):
    value = _strip_code_fences(text)
    if value.startswith("{") and value.endswith("}"):
        return value
    start = value.find("{")
    end = value.rfind("}")
    if start != -1 and end > start:
        return value[start : end + 1]
    return value


def _parse_batch_review_output(text, expected_ids):
    candidate = _extract_json_candidate(text)
    payload = json.loads(candidate)
    items = payload.get("items")
    if not isinstance(items, list):
        raise RuntimeError("La respuesta por lotes no contiene items validos.")
    expected_order = [str(item_id) for item_id in expected_ids]
    expected_set = set(expected_order)
    reviewed_map = {}
    for item in items:
        if not isinstance(item, dict):
            raise RuntimeError("La respuesta por lotes contiene items invalidos.")
        item_id = str(item.get("id") or "").strip()
        if not item_id or item_id not in expected_set:
            raise RuntimeError("La respuesta por lotes devolvio ids inesperados.")
        if item_id in reviewed_map:
            raise RuntimeError("La respuesta por lotes devolvio ids duplicados.")
        reviewed_map[item_id] = str(item.get("text") or "").strip()
    missing_ids = [item_id for item_id in expected_order if item_id not in reviewed_map]
    if missing_ids:
        raise RuntimeError("La respuesta por lotes no devolvio todos los items.")
    return [reviewed_map[item_id] for item_id in expected_order]


def _build_review_batches(texts, settings):
    max_items = max(1, int(settings.get("batch_max_items") or DEFAULT_BATCH_MAX_ITEMS))
    max_chars = max(1, int(settings.get("batch_max_chars") or DEFAULT_BATCH_MAX_CHARS))
    batches = []
    current_batch = []
    current_chars = 0
    for text in texts:
        safe_text = str(text or "")
        text_chars = len(safe_text)
        should_flush = (
            current_batch
            and (
                len(current_batch) >= max_items
                or current_chars + text_chars > max_chars
            )
        )
        if should_flush:
            batches.append(current_batch)
            current_batch = []
            current_chars = 0
        current_batch.append(safe_text)
        current_chars += text_chars
    if current_batch:
        batches.append(current_batch)
    return batches


def _review_text_with_retry(review_fn, settings, transport_label):
    try:
        return review_fn(settings)
    except Exception as exc:
        if not _is_timeout_error(exc):
            raise
        retry_settings = dict(settings)
        retry_settings["timeout"] = max(int(settings.get("timeout") or 0), TIMEOUT_RETRY_SECONDS)
        _log_review(
            f"timeout_retry transport={transport_label} timeout={settings.get('timeout')} "
            f"retry_timeout={retry_settings['timeout']}",
            level="WARN",
        )
        return review_fn(retry_settings)


def _call_openai_responses_api(input_text, settings, *, instructions):
    body = json.dumps(
        {
            "model": settings["model"],
            "instructions": instructions,
            "input": input_text,
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
    return json.loads(raw) if raw else {}


def _call_edge_review(payload, settings):
    supabase_url, supabase_key = _load_supabase_credentials(".env")
    jwt_token = str(_supabase_get_access_token(".env") or "").strip()
    if not jwt_token:
        raise RuntimeError("No hay sesiÃ³n vÃ¡lida para revisar ortografÃ­a.")
    function_name = str(settings.get("function_name") or DEFAULT_EDGE_FUNCTION_NAME).strip() or DEFAULT_EDGE_FUNCTION_NAME
    url = f"{supabase_url.rstrip('/')}/functions/v1/{function_name}"
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
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
    response_payload = json.loads(raw) if raw else {}
    ok = bool(response_payload.get("ok"))
    if not ok:
        err = response_payload.get("error")
        if isinstance(err, dict):
            message = str(err.get("message") or "").strip()
            if message:
                raise RuntimeError(message)
        raise RuntimeError(str(response_payload.get("message") or "La funciÃ³n de revisiÃ³n no devolviÃ³ texto."))
    return response_payload


def _review_text(text, settings):
    transport = "direct" if settings.get("api_key") else "edge"
    if settings.get("api_key"):
        return _review_text_with_retry(lambda current: _review_text_direct(text, current), settings, transport)
    return _review_text_with_retry(lambda current: _review_text_via_edge(text, current), settings, transport)


def _review_text_direct(text, settings):
    payload = _call_openai_responses_api(str(text or ""), settings, instructions=REVIEW_PROMPT)
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


def _review_text_batch_direct(batch_items, settings):
    input_payload = json.dumps({"items": batch_items}, ensure_ascii=False)
    payload = _call_openai_responses_api(input_payload, settings, instructions=BATCH_REVIEW_PROMPT)
    reviewed = _extract_output_text(payload)
    if not reviewed:
        raise RuntimeError("OpenAI no devolvio lote corregido.")
    expected_ids = [item["id"] for item in batch_items]
    return _parse_batch_review_output(reviewed, expected_ids)


def _review_text_batch_via_edge(batch_items, settings):
    response_payload = _call_edge_review(
        {
            "items": batch_items,
            "model": settings["model"],
        },
        settings,
    )
    reviewed_items = response_payload.get("items")
    if not isinstance(reviewed_items, list):
        raise RuntimeError("La funcion de revision no devolvio items por lote.")
    expected_ids = [item["id"] for item in batch_items]
    return _parse_batch_review_output(
        json.dumps({"items": reviewed_items}, ensure_ascii=False),
        expected_ids,
    )


def _review_text_batch(batch_texts, settings):
    batch_items = [
        {"id": f"item_{index}", "text": str(text or "")}
        for index, text in enumerate(batch_texts, start=1)
    ]
    transport = "direct" if settings.get("api_key") else "edge"
    review_fn = _review_text_batch_direct if settings.get("api_key") else _review_text_batch_via_edge
    try:
        return _review_text_with_retry(
            lambda current: review_fn(batch_items, current),
            settings,
            f"{transport}_batch",
        )
    except Exception as exc:
        _log_review(
            f"batch_fallback transport={transport} items={len(batch_items)} "
            f"chars={sum(len(item['text']) for item in batch_items)} error={_extract_error_message(exc)!r}",
            level="WARN",
        )
        return [_review_text(item["text"], settings) for item in batch_items]


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
    "seleccion_incluyente_labs": [
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

TEXT_REVIEW_FORM_ALIASES = {
    "condiciones_vacante_labs": "condiciones_vacante",
    "seleccion_incluyente_labs": "seleccion_incluyente",
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
    clean_form_id = str(form_id or "").strip()
    review_form_id = TEXT_REVIEW_FORM_ALIASES.get(clean_form_id, clean_form_id)
    specs = TEXT_REVIEW_FIELDS_BY_FORM.get(review_form_id, [])
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
        unique_texts = []
        seen_texts = set()
        for target in targets:
            original_text = str(target.get("text") or "")
            if original_text in seen_texts:
                continue
            seen_texts.add(original_text)
            unique_texts.append(original_text)

        batches = _build_review_batches(unique_texts, settings)
        for batch_index, batch_texts in enumerate(batches, start=1):
            if len(batch_texts) == 1:
                original_text = batch_texts[0]
                _log_review(
                    f"request form={form_id} batch={batch_index}/{len(batches)} items=1 chars={len(original_text)} transport={transport}"
                )
                reviewed_batch = [_review_text(original_text, settings)]
            else:
                _log_review(
                    f"batch_request form={form_id} batch={batch_index}/{len(batches)} "
                    f"items={len(batch_texts)} chars={sum(len(text) for text in batch_texts)} transport={transport}"
                )
                reviewed_batch = _review_text_batch(batch_texts, settings)
            for original_text, reviewed_text in zip(batch_texts, reviewed_batch):
                normalized_review = str(reviewed_text or "").strip() or original_text
                normalized_review = _maybe_format_reviewed_list(original_text, normalized_review)
                reviewed_texts[original_text] = normalized_review

        for index, target in enumerate(targets, start=1):
            original_text = target["text"]
            path = tuple(target.get("path") or ())
            reviewed_text = reviewed_texts.get(original_text, original_text)
            if original_text in reviewed_texts and original_text != reviewed_text:
                _log_review(
                    f"reuse_reviewed_result form={form_id} index={index}/{len(targets)} path={path!r}"
                )
            reviewed_targets.append({"path": path, "text": reviewed_text})
            if reviewed_text != original_text:
                changed_count += 1
    except Exception as exc:
        _log_review(
            f"failed form={form_id} transport={transport} reviewed_count={changed_count} "
            f"error={_extract_error_message(exc)!r}",
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
