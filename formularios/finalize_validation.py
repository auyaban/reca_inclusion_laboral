"""
formularios/finalize_validation.py — Validación pre-envío de formularios.

Responsabilidades:
  - ValidationIssue: dataclass que representa un campo con error (field_id, message, section_id)
  - format_issues_for_message: formatea la lista de issues para mostrar al usuario
  - Funciones auxiliares de normalización de texto para comparación

Depende de: nada interno
Usado por: app.py (_guard_form_finalization), módulos de formularios
"""
from __future__ import annotations

from dataclasses import dataclass
import re
import unicodedata


@dataclass(frozen=True)
class ValidationIssue:
    section_id: str
    field_id: str
    label: str
    message: str
    row_index: int | None = None


_EMPTY_NOTE_RE = re.compile(r"^\s*nota:?\s*$", re.IGNORECASE)


def _normalize_text(value):
    text = str(value or "").strip()
    if not text:
        return ""
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch)).strip().lower()


def humanize_section_id(section_id):
    raw = str(section_id or "").strip()
    if not raw:
        return "Seccion"
    match = re.fullmatch(r"section_(\d+)(?:_(\d+))?", raw)
    if match:
        main = match.group(1)
        sub = match.group(2)
        return f"Seccion {main}.{sub}" if sub else f"Seccion {main}"
    return raw.replace("_", " ").strip().title()


def humanize_field_id(field_id):
    raw = str(field_id or "").strip()
    if not raw:
        return "Campo"
    text = raw.replace("_", " ").strip()
    if not text:
        return "Campo"
    return text[:1].upper() + text[1:]


def field_pairs(fields):
    pairs = []
    for field in fields or []:
        if not isinstance(field, dict):
            continue
        field_id = str(field.get("id") or "").strip()
        if not field_id:
            continue
        label = str(field.get("label") or "").strip() or humanize_field_id(field_id)
        pairs.append((field_id, label))
    return pairs


def is_meaningful(value):
    if value is None:
        return False
    if isinstance(value, bool):
        return True
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, dict):
        for key, item in value.items():
            if str(key or "").startswith("_"):
                continue
            if is_meaningful(item):
                return True
        return False
    if isinstance(value, list):
        return any(is_meaningful(item) for item in value)
    text = str(value).strip()
    if not text:
        return False
    if _EMPTY_NOTE_RE.fullmatch(text):
        return False
    return True


def append_missing_issue(
    issues,
    section_id,
    field_id,
    label="",
    *,
    row_index=None,
    message="Campo obligatorio sin diligenciar.",
):
    issues.append(
        ValidationIssue(
            section_id=str(section_id or "").strip(),
            field_id=str(field_id or "").strip(),
            label=str(label or "").strip() or humanize_field_id(field_id),
            message=str(message or "").strip() or "Campo obligatorio sin diligenciar.",
            row_index=row_index,
        )
    )


def require_value(issues, section_id, payload, field_id, label="", *, row_index=None):
    value = payload.get(field_id) if isinstance(payload, dict) else None
    if is_meaningful(value):
        return
    append_missing_issue(
        issues,
        section_id,
        field_id,
        label,
        row_index=row_index,
    )


def require_any_true(issues, section_id, payload, field_ids, label, *, message="Selecciona al menos una opcion."):
    values = payload if isinstance(payload, dict) else {}
    for field_id in field_ids or []:
        if bool(values.get(field_id)):
            return
    issues.append(
        ValidationIssue(
            section_id=str(section_id or "").strip(),
            field_id=str((field_ids or [""])[0] or "").strip(),
            label=str(label or "").strip() or "Seleccion",
            message=str(message or "").strip() or "Selecciona al menos una opcion.",
            row_index=None,
        )
    )


def validate_dynamic_rows(
    issues,
    section_id,
    rows,
    row_fields,
    *,
    min_rows=1,
    min_rows_label="",
):
    row_list = rows if isinstance(rows, list) else []
    meaningful_rows = 0
    field_defs = [(str(field_id or "").strip(), str(label or "").strip() or humanize_field_id(field_id)) for field_id, label in (row_fields or []) if str(field_id or "").strip()]
    for row_index, row in enumerate(row_list, start=1):
        row_payload = row if isinstance(row, dict) else {}
        filled = [
            field_id
            for field_id, _label in field_defs
            if is_meaningful(row_payload.get(field_id))
        ]
        if not filled:
            continue
        meaningful_rows += 1
        for field_id, label in field_defs:
            if field_id in filled:
                continue
            append_missing_issue(
                issues,
                section_id,
                field_id,
                label,
                row_index=row_index,
            )
    if meaningful_rows >= int(min_rows or 0):
        return
    issues.append(
        ValidationIssue(
            section_id=str(section_id or "").strip(),
            field_id="",
            label=str(min_rows_label or "").strip() or humanize_section_id(section_id),
            message="Debes diligenciar al menos una fila.",
            row_index=None,
        )
    )


def format_issues_for_message(issues, *, title="No se puede finalizar el formato.", limit=8):
    issue_list = list(issues or [])
    if not issue_list:
        return title
    lines = [title, "", "Faltan campos obligatorios:"]
    capped = issue_list[: max(1, int(limit or 1))]
    for issue in capped:
        field_label = str(issue.label or "").strip() or humanize_field_id(issue.field_id)
        if issue.row_index is not None:
            field_label = f"{field_label} (fila {issue.row_index})"
        section_label = humanize_section_id(issue.section_id)
        lines.append(f"- {section_label}: {field_label}")
    remaining = len(issue_list) - len(capped)
    if remaining > 0:
        lines.append(f"- ...y {remaining} campos mas.")
    return "\n".join(lines)


def raise_validation_error(issues, *, title="No se puede finalizar el formato."):
    if not issues:
        return
    raise RuntimeError(format_issues_for_message(issues, title=title))
