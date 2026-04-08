"""
formularios/seguimientos/seguimientos.py
Formulario: "9. SEGUIMIENTO AL PROCESO DE INCLUSIÓN LABORAL"

Responsabilidades:
  - Flujo diferente al resto: NO usa el master spreadsheet directamente.
    Cada caso (persona vinculada) tiene su propio Google Sheet copiado desde
    una plantilla (SEGUIMIENTOS_TEMPLATE_ID) en una carpeta compartida de Drive
  - build_case_folder_name(): construye el nombre de carpeta por usuario/cédula
  - Gestión de casos abiertos/cerrados: un caso = una carpeta en Drive con su Sheet
  - Sincronización: lee y escribe casos existentes directamente en Google Sheets/Drive
  - SeguimientoEditorWindow en app.py edita un caso individual

Entry points para app.py:
  register_form()              → metadata para HubWindow
  get_empresa_by_nit/nombre/prefix() → búsqueda de empresa
  build_case_folder_name()     → nombre de carpeta del caso en Drive

Variables clave de entorno:
  GOOGLE_SHEETS_SEGUIMIENTOS_TEMPLATE_ID → ID de la plantilla de seguimiento

Depende de: google_sheets_client, drive_upload, google_api_requests,
            formularios/common
"""
import os
import re
import uuid
from datetime import datetime

from formularios.common import (
    _get_cached_payload,
    _load_env_file,
    _merge_company_row_with_cache,
    _normalize_cedula,
    _sanitize_filename,
    _supabase_get,
)
from google_api_requests import (
    execute_google_create_with_confirmation,
    execute_google_request_with_retry,
)
from google_sheets_client import (
    batch_write_sheet_updates,
    clear_protected_ranges,
    extract_spreadsheet_id,
    get_master_template_id,
    get_google_sheets_service,
    get_spreadsheet,
    read_sheet_values,
)
import drive_upload


FORM_ID = "seguimientos"
FORM_NAME = "Seguimientos"
SEGUIMIENTOS_FOLDER_NAME = "SEGUIMIENTOS"
DEFAULT_SEGUIMIENTOS_TEMPLATE_ID = ""
GOOGLE_SHEETS_MIME = "application/vnd.google-apps.spreadsheet"
_USUARIOS_RECA_CEDULAS_CACHE_TTL_SECONDS = 86400

SHEET_BASE = "9. SEGUIMIENTO AL PROCESO DE INCLUSIÓN LABORAL"
LEGACY_SHEET_BASE = "9.  SEGUIMIENTO AL PROCESO DE INCLUSION LABORAL"
LEGACY_SHEET_BASE_SHORT = "9.  SEGUIMIENTO AL PROCESO DE I"
BASE_SHEET_CANDIDATES = (
    SHEET_BASE,
    LEGACY_SHEET_BASE,
    LEGACY_SHEET_BASE_SHORT,
)
SHEET_PREFIX = "SEGUIMIENTO PROCESO IL "
SHEET_FINAL = "PONDERADO FINAL"
FOLLOWUP_DATE_LABEL = "Fecha seguimiento:"
FOLLOWUP_DATE_LABEL_CELL = "U8"
FOLLOWUP_DATE_VALUE_CELL = "X8"
SHEET_COVERAGE_THRESHOLD = 90

BASE_PROGRESS_SCALAR_FIELDS = (
    "fecha_visita",
    "modalidad",
    "contacto_emergencia",
    "parentesco",
    "telefono_emergencia",
    "certificado_discapacidad",
    "certificado_porcentaje",
    "tipo_contrato",
    "fecha_firma_contrato",
    "fecha_inicio_contrato",
    "fecha_fin_contrato",
    "apoyos_ajustes",
)
FOLLOWUP_PROGRESS_SCALAR_FIELDS = (
    "modalidad",
    "fecha_seguimiento",
    "tipo_apoyo",
    "situacion_encontrada",
    "estrategias_ajustes",
)


def _load_runtime_env():
    try:
        return _load_env_file(".env") or {}
    except Exception:
        return {}


def _get_shared_root():
    runtime_env = _load_runtime_env()
    return str(
        os.getenv("SEGUIMIENTOS_SHARED_ROOT")
        or runtime_env.get("SEGUIMIENTOS_SHARED_ROOT")
        or drive_upload._load_config().get("seguimientos_shared_root")
        or ""
    ).strip()

MODALIDAD_OPTIONS = ["Presencial", "Virtual", "Mixta", "No aplica"]
SI_NO_NA_OPTIONS = ["Si", "No", "No aplica"]
EVAL_OPTIONS = ["Excelente", "Bien", "Necesita mejorar", "Mal", "No aplica"]
TIPO_APOYO_OPTIONS = [
    "Requiere apoyo bajo.",
    "Requiere apoyo medio.",
    "Requiere apoyo Alto.",
    "No requiere apoyo.",
]

SECTION_1_SUPABASE_MAP = {
    "nombre_empresa": "nombre_empresa",
    "ciudad_empresa": "ciudad_empresa",
    "direccion_empresa": "direccion_empresa",
    "nit_empresa": "nit_empresa",
    "correo_1": "correo_1",
    "telefono_empresa": "telefono_empresa",
    "contacto_empresa": "contacto_empresa",
    "cargo": "cargo",
    "asesor": "asesor",
    "sede_empresa": "zona_empresa",
    "caja_compensacion": "caja_compensacion",
    "profesional_asignado": "profesional_asignado",
}

PONDERADO_COMPANY_MAP = {
    "fecha_visita": "D6",
    "modalidad": "Q6",
    "nombre_empresa": "D7",
    "ciudad_empresa": "Q7",
    "direccion_empresa": "D8",
    "nit_empresa": "Q8",
    "correo_1": "D9",
    "telefono_empresa": "Q9",
    "contacto_empresa": "D10",
    "cargo": "Q10",
    "caja_compensacion": "D11",
    "sede_empresa": "Q11",
    "asesor": "D12",
    "profesional_asignado": "Q12",
}

PONDERADO_USER_MAP = {
    "nombre_vinculado": "K15",
    "cedula": "Q15",
    "telefono_vinculado": "S15",
    "correo_vinculado": "U15",
    "cargo_vinculado": "K17",
    "certificado_discapacidad": "Q17",
    "certificado_porcentaje": "U17",
    "fecha_firma_contrato": "N18",
    "discapacidad": "U18",
}

BASE_SHEET_FIELD_MAP = {
    "fecha_visita": "D8",
    "modalidad": "R8",
    "nombre_empresa": "D9",
    "ciudad_empresa": "R9",
    "direccion_empresa": "D10",
    "nit_empresa": "R10",
    "correo_1": "D11",
    "telefono_empresa": "R11",
    "contacto_empresa": "D12",
    "cargo": "R12",
    "asesor": "D13",
    "sede_empresa": "R13",
    "nombre_vinculado": "A16",
    "cedula": "E16",
    "telefono_vinculado": "I16",
    "correo_vinculado": "K16",
    "contacto_emergencia": "P16",
    "parentesco": "S16",
    "telefono_emergencia": "U16",
    "cargo_vinculado": "A18",
    "certificado_discapacidad": "E18",
    "certificado_porcentaje": "I18",
    "discapacidad": "N18",
    "tipo_contrato": "C20",
    "fecha_inicio_contrato": "M20",
    "fecha_fin_contrato": "T20",
    "apoyos_ajustes": "E21",
}

BASE_SHEET_FIELD_LABELS = {
    "fecha_visita": "Fecha de visita",
    "modalidad": "Modalidad",
    "nombre_empresa": "Nombre empresa",
    "ciudad_empresa": "Ciudad/Municipio",
    "direccion_empresa": "Direccion",
    "nit_empresa": "NIT",
    "correo_1": "Correo",
    "telefono_empresa": "Telefonos",
    "contacto_empresa": "Contacto empresa",
    "cargo": "Cargo empresa",
    "asesor": "Asesor",
    "sede_empresa": "Sede Compensar",
    "nombre_vinculado": "Nombre",
    "cedula": "Cedula",
    "telefono_vinculado": "Telefono",
    "correo_vinculado": "Correo vinculado",
    "contacto_emergencia": "Contacto emergencia",
    "parentesco": "Parentesco",
    "telefono_emergencia": "Telefono emergencia",
    "cargo_vinculado": "Cargo",
    "certificado_discapacidad": "Certificado discapacidad",
    "certificado_porcentaje": "Porcentaje certificado",
    "discapacidad": "Discapacidad",
    "tipo_contrato": "Tipo contrato",
    "fecha_inicio_contrato": "Fecha inicio contrato",
    "fecha_fin_contrato": "Fecha fin contrato",
    "apoyos_ajustes": "Apoyos y ajustes razonables",
}

FOLLOWUP_SHEET_FIELD_MAP = {
    "modalidad": "E8",
    "seguimiento_numero": "P8",
    "fecha_seguimiento": FOLLOWUP_DATE_VALUE_CELL,
    "tipo_apoyo": "J31",
    "situacion_encontrada": "A43",
    "estrategias_ajustes": "A45",
}

FOLLOWUP_SHEET_FIELD_LABELS = {
    "modalidad": "Modalidad",
    "seguimiento_numero": "Seguimiento #",
    "fecha_seguimiento": "Fecha seguimiento",
    "tipo_apoyo": "Tipo de apoyo",
    "situacion_encontrada": "Situacion encontrada",
    "estrategias_ajustes": "Estrategias y ajustes",
}

def register_form():
    return {
        "id": FORM_ID,
        "name": FORM_NAME,
        "module": __name__,
        "supports_drafts": False,
        "hub_description": "Abre y actualiza casos con ficha inicial, seguimientos periódicos y resultado final.",
        "singleton_window": True,
    }


def _map_company_row(row):
    if not isinstance(row, dict):
        return row
    mapped = dict(row)
    for field_id, source_key in SECTION_1_SUPABASE_MAP.items():
        if source_key in row:
            mapped[field_id] = row.get(source_key)
    return mapped


def get_linked_company_for_user(user_row, env_path=".env"):
    if not isinstance(user_row, dict):
        return {}
    nit = _get_str(user_row.get("empresa_nit"))
    nombre = _get_str(user_row.get("empresa_nombre"))
    company = None
    if nit:
        try:
            company = get_empresa_by_nit(nit, env_path=env_path)
        except Exception:
            company = None
    if not company and nombre:
        try:
            company = get_empresa_by_nombre(nombre, env_path=env_path)
        except Exception:
            company = None
    payload = _map_company_row(company or {})
    if nit and not payload.get("nit_empresa"):
        payload["nit_empresa"] = nit
    if nombre and not payload.get("nombre_empresa"):
        payload["nombre_empresa"] = nombre
    return payload


def _get_base_sheet_name_from_workbook(wb):
    for candidate in BASE_SHEET_CANDIDATES:
        if candidate in wb.sheetnames:
            return candidate
    raise ValueError("No existe la hoja base de seguimientos en el archivo.")


def _is_native_case_ref(case_ref):
    return (
        isinstance(case_ref, dict)
        and str(case_ref.get("source") or "").strip() == "drive"
        and str(case_ref.get("mime_type") or "").strip() == GOOGLE_SHEETS_MIME
        and str(case_ref.get("file_id") or "").strip()
    )


def _get_spreadsheet_id_from_case_ref(case_ref):
    if _is_native_case_ref(case_ref):
        return str(case_ref.get("file_id") or "").strip()
    text = str(case_ref or "").strip()
    lowered = text.lower()
    if not text:
        raise RuntimeError("No se recibió una referencia válida de Google Sheets.")
    if lowered.endswith((".xlsx", ".xlsm", ".xls")):
        raise RuntimeError(
            "Seguimientos ya no admite archivos locales de Excel. "
            "Usa una referencia nativa de Google Sheets."
        )
    if re.match(r"^[a-zA-Z]:[\\/]", text) or ("\\" in text and "docs.google.com" not in lowered):
        raise RuntimeError(
            "Seguimientos ya no admite rutas locales de archivos. "
            "Usa una referencia nativa de Google Sheets."
        )
    spreadsheet_id = extract_spreadsheet_id(text)
    if spreadsheet_id == text and not re.fullmatch(r"[A-Za-z0-9_-]{20,}", spreadsheet_id):
        raise RuntimeError(
            "Seguimientos requiere un spreadsheet_id o URL válida de Google Sheets."
        )
    return spreadsheet_id


def _batch_read_sheet_values(spreadsheet_id_or_url, ranges):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    cleaned = [str(item or "").strip() for item in (ranges or []) if str(item or "").strip()]
    if not cleaned:
        return {}
    service = get_google_sheets_service()
    response = execute_google_request_with_retry(
        service.spreadsheets()
        .values()
        .batchGet(spreadsheetId=spreadsheet_id, ranges=cleaned),
        operation_name="seguimientos.batch_read_values",
    )
    return {item.get("range"): item.get("values", []) for item in response.get("valueRanges", [])}


def _first_batch_value(values_map, range_name):
    rows = values_map.get(range_name) or []
    if not rows or not rows[0]:
        return ""
    return str(rows[0][0]).strip()


def _column_batch_values(values_map, range_name, expected_count):
    rows = values_map.get(range_name) or []
    values = []
    for row in rows:
        if row:
            values.append(str(row[0]).strip())
        else:
            values.append("")
    if len(values) < expected_count:
        values.extend([""] * (expected_count - len(values)))
    return values[:expected_count]


def get_base_sheet_name(case_ref):
    return str(get_case_meta(case_ref).get("base_sheet_name") or SHEET_BASE)


def _get_drive_service():
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Faltan dependencias para Google Drive. Instala google-api-python-client y google-auth."
        ) from exc

    creds_path = drive_upload._get_credentials_path()
    credentials = Credentials.from_service_account_file(creds_path, scopes=[drive_upload.SCOPE])
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def _get_drive_root_folder_id(service):
    configured_root = drive_upload._get_excel_folder_id()
    return drive_upload._resolve_target_root_id(service, configured_root)


def _find_named_folder(service, parent_id, folder_name):
    safe_name = str(folder_name or "").replace("'", "\\'")
    query = (
        "mimeType='application/vnd.google-apps.folder' "
        f"and name='{safe_name}' and '{parent_id}' in parents and trashed=false"
    )
    result = execute_google_request_with_retry(
        service.files().list(
            q=query,
            fields="files(id,name,webViewLink)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            pageSize=5,
        ),
        operation_name="seguimientos.find_named_folder",
    )
    files = result.get("files", [])
    return files[0] if files else None


def _ensure_seguimientos_folder(service):
    root_id = _get_drive_root_folder_id(service)
    folder = _find_named_folder(service, root_id, SEGUIMIENTOS_FOLDER_NAME)
    if folder:
        return folder["id"]
    return drive_upload._get_or_create_folder(service, root_id, SEGUIMIENTOS_FOLDER_NAME)


def _list_drive_files(service, folder_id):
    result = execute_google_request_with_retry(
        service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id,name,mimeType,webViewLink,modifiedTime,parents,appProperties)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            pageSize=100,
            orderBy="modifiedTime desc",
        ),
        operation_name="seguimientos.list_drive_files",
    )
    return list(result.get("files", []))


def _find_drive_file_by_request_id(service, folder_id, filename, request_id):
    safe_name = str(filename or "").replace("'", "\\'")
    query = (
        f"mimeType='{GOOGLE_SHEETS_MIME}' "
        f"and name='{safe_name}' and '{folder_id}' in parents and trashed=false"
    )
    result = execute_google_request_with_retry(
        service.files().list(
            q=query,
            fields="files(id,name,mimeType,webViewLink,parents,appProperties)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            pageSize=5,
        ),
        operation_name="seguimientos.find_drive_file_by_request_id",
    )
    for item in result.get("files", []):
        app_properties = item.get("appProperties") or {}
        if str(app_properties.get("request_id") or "").strip() == str(request_id or "").strip():
            return item
    return None


def _find_case_folder_drive(service, cedula, nombre_usuario=""):
    seguimientos_folder_id = _ensure_seguimientos_folder(service)
    preferred_name = build_case_folder_name(nombre_usuario, cedula) if nombre_usuario else ""
    candidates = []
    if preferred_name:
        direct = _find_named_folder(service, seguimientos_folder_id, preferred_name)
        if direct:
            return direct
    suffix = f"- {cedula}"
    for item in _list_drive_files(service, seguimientos_folder_id):
        if item.get("mimeType") != "application/vnd.google-apps.folder":
            continue
        name = str(item.get("name") or "").strip()
        if name.endswith(suffix):
            candidates.append(item)
    return candidates[0] if candidates else None


def _pick_case_file(files):
    for item in files:
        mime = str(item.get("mimeType") or "").strip()
        if mime == GOOGLE_SHEETS_MIME:
            return item
    return None


def get_seguimientos_template_id():
    try:
        return get_master_template_id()
    except Exception:
        runtime_env = _load_runtime_env()
        raw = str(
            os.getenv("GOOGLE_SHEETS_SEGUIMIENTOS_TEMPLATE_ID")
            or runtime_env.get("GOOGLE_SHEETS_SEGUIMIENTOS_TEMPLATE_ID")
            or drive_upload._load_config().get("google_sheets_seguimientos_template_id")
            or DEFAULT_SEGUIMIENTOS_TEMPLATE_ID
        ).strip()
        return extract_spreadsheet_id(raw)


def _parse_first_name_lastname(full_name):
    tokens = [t for t in re.split(r"\s+", str(full_name or "").strip()) if t]
    if not tokens:
        return "Usuario", "SinApellido"
    first_name = tokens[0]
    if len(tokens) >= 4:
        first_lastname = tokens[2]
    elif len(tokens) >= 2:
        first_lastname = tokens[1]
    else:
        first_lastname = "SinApellido"
    return first_name, first_lastname


def build_case_folder_name(nombre_usuario, cedula):
    first_name, first_lastname = _parse_first_name_lastname(nombre_usuario)
    base = f"{first_name} {first_lastname} - {cedula}"
    return _sanitize_filename(base)


def find_case_record(cedula, nombre_usuario=""):
    normalized = _normalize_cedula(cedula)
    if not normalized:
        return None
    try:
        service = _get_drive_service()
        folder = _find_case_folder_drive(service, normalized, nombre_usuario=nombre_usuario)
        if folder:
            files = _list_drive_files(service, folder["id"])
            case_file = _pick_case_file(files)
            if case_file:
                record = {
                    "source": "drive",
                    "cedula": normalized,
                    "folder_id": folder.get("id"),
                    "folder_name": folder.get("name"),
                    "file_id": case_file.get("id"),
                    "file_name": case_file.get("name"),
                    "mime_type": case_file.get("mimeType"),
                    "webViewLink": case_file.get("webViewLink"),
                    "modifiedTime": case_file.get("modifiedTime"),
                    "local_path": "",
                }
                app_props = case_file.get("appProperties") or {}
                try:
                    record["max_seguimientos"] = int(app_props.get("max_seguimientos") or 3)
                except Exception:
                    record["max_seguimientos"] = 3
                return record
    except Exception:
        pass
    return None


def _get_str(value):
    if value is None:
        return ""
    return str(value).strip()


def _build_base_payload_from_user_row(user_row):
    discapacidad = _get_str(user_row.get("discapacidad_detalle")) or _get_str(
        user_row.get("discapacidad_usuario")
    )
    payload = {
        "nombre_vinculado": _get_str(user_row.get("nombre_usuario")),
        "cedula": _get_str(user_row.get("cedula_usuario")),
        "telefono_vinculado": _get_str(user_row.get("telefono_oferente")),
        "correo_vinculado": _get_str(user_row.get("correo_oferente")),
        "contacto_emergencia": _get_str(user_row.get("contacto_emergencia")),
        "parentesco": _get_str(user_row.get("parentesco")),
        "telefono_emergencia": _get_str(user_row.get("telefono_emergencia")),
        "cargo_vinculado": _get_str(user_row.get("cargo_oferente")),
        "certificado_discapacidad": _get_str(user_row.get("certificado_discapacidad")),
        "certificado_porcentaje": _get_str(user_row.get("certificado_porcentaje")),
        "discapacidad": discapacidad,
        "tipo_contrato": _get_str(user_row.get("tipo_contrato")),
    }
    payload.update(get_linked_company_for_user(user_row))
    return payload


def _build_ponderado_payload(base_payload):
    payload = dict(base_payload or {})
    fecha_firma = _get_str(payload.get("fecha_firma_contrato"))
    if not fecha_firma:
        fecha_firma = _get_str(payload.get("fecha_inicio_contrato"))
    payload["fecha_firma_contrato"] = fecha_firma
    return payload


def _append_sparse_update(
    updates,
    changes,
    *,
    range_name,
    local_value,
    remote_value,
    label,
    field_id,
):
    local_text = _get_str(local_value)
    remote_text = _get_str(remote_value)
    if not local_text or local_text == remote_text:
        return
    updates.append({"range": range_name, "value": local_text})
    changes.append(
        {
            "field_id": str(field_id or "").strip(),
            "label": _get_str(label),
            "range": str(range_name or "").strip(),
            "previous_value": remote_text,
            "new_value": local_text,
            "change_kind": "overwrite" if remote_text else "new",
        }
    )


def _build_sparse_plan_result(*, spreadsheet_id, sheet_name, updates, changes, remote_payload=None):
    changes_list = [dict(item or {}) for item in list(changes or [])]
    overwrite_fields = [
        dict(item) for item in changes_list if str(item.get("change_kind") or "").strip() == "overwrite"
    ]
    new_fields = [
        dict(item) for item in changes_list if str(item.get("change_kind") or "").strip() == "new"
    ]
    return {
        "spreadsheet_id": str(spreadsheet_id or "").strip(),
        "sheet_name": str(sheet_name or "").strip(),
        "remote_payload": dict(remote_payload or {}),
        "updates": list(updates or []),
        "changes": changes_list,
        "overwrite_fields": overwrite_fields,
        "new_fields": new_fields,
        "has_changes": bool(updates),
        "write_count": len(list(updates or [])),
    }


def _list_value(values, index):
    current = list(values or [])
    return current[index] if index < len(current) else ""


def _attendee_value(values, index, field_id):
    current = list(values or [])
    entry = current[index] if index < len(current) and isinstance(current[index], dict) else {}
    return entry.get(field_id, "")


def _build_base_sheet_sparse_updates(payload, remote_payload, base_sheet_name=SHEET_BASE):
    updates = []
    changes = []
    payload = dict(payload or {})
    remote_payload = dict(remote_payload or {})

    for field_id, cell in BASE_SHEET_FIELD_MAP.items():
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{base_sheet_name}'!{cell}",
            local_value=payload.get(field_id, ""),
            remote_value=remote_payload.get(field_id, ""),
            label=BASE_SHEET_FIELD_LABELS.get(field_id, field_id),
            field_id=field_id,
        )

    for idx, row in enumerate(range(23, 28), start=1):
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{base_sheet_name}'!B{row}",
            local_value=_list_value(payload.get("funciones_1_5"), idx - 1),
            remote_value=_list_value(remote_payload.get("funciones_1_5"), idx - 1),
            label=f"Funcion {idx}",
            field_id=f"funciones_1_5[{idx - 1}]",
        )
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{base_sheet_name}'!N{row}",
            local_value=_list_value(payload.get("funciones_6_10"), idx - 1),
            remote_value=_list_value(remote_payload.get("funciones_6_10"), idx - 1),
            label=f"Funcion {idx + 5}",
            field_id=f"funciones_6_10[{idx - 1}]",
        )

    # El template solo almacena fecha_firma_contrato en PONDERADO FINAL!N18.
    # Mantenemos esta excepción mínima para no perder ese dato al reabrir el caso.
    _append_sparse_update(
        updates,
        changes,
        range_name=f"'{SHEET_FINAL}'!{PONDERADO_USER_MAP['fecha_firma_contrato']}",
        local_value=payload.get("fecha_firma_contrato", ""),
        remote_value=remote_payload.get("fecha_firma_contrato", ""),
        label="Fecha firma contrato",
        field_id="fecha_firma_contrato",
    )

    return updates, changes


def _followup_item_label(payload, remote_payload, index):
    label = _list_value((payload or {}).get("item_labels"), index)
    if label:
        return label
    label = _list_value((remote_payload or {}).get("item_labels"), index)
    if label:
        return label
    return f"Item {index + 1}"


def _followup_company_label(payload, remote_payload, index):
    label = _list_value((payload or {}).get("empresa_item_labels"), index)
    if label:
        return label
    label = _list_value((remote_payload or {}).get("empresa_item_labels"), index)
    if label:
        return label
    return f"Empresa {index + 1}"


def _build_followup_sheet_sparse_updates(
    index,
    payload,
    remote_payload,
    *,
    base_sheet_name=SHEET_BASE,
    base_remote_payload=None,
):
    sheet_name = _get_followup_sheet_name(index)
    updates = []
    changes = []
    payload = dict(payload or {})
    remote_payload = dict(remote_payload or {})
    base_remote_payload = dict(base_remote_payload or {})

    for field_id, cell in FOLLOWUP_SHEET_FIELD_MAP.items():
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{sheet_name}'!{cell}",
            local_value=payload.get(field_id, ""),
            remote_value=remote_payload.get(field_id, ""),
            label=FOLLOWUP_SHEET_FIELD_LABELS.get(field_id, field_id),
            field_id=field_id,
        )

    for i, row in enumerate(range(12, 31)):
        item_label = _followup_item_label(payload, remote_payload, i)
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{sheet_name}'!G{row}",
            local_value=_list_value(payload.get("item_observaciones"), i),
            remote_value=_list_value(remote_payload.get("item_observaciones"), i),
            label=f"Observacion del vinculado: {item_label}",
            field_id=f"item_observaciones[{i}]",
        )
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{sheet_name}'!O{row}",
            local_value=_list_value(payload.get("item_autoevaluacion"), i),
            remote_value=_list_value(remote_payload.get("item_autoevaluacion"), i),
            label=f"Autoevaluacion: {item_label}",
            field_id=f"item_autoevaluacion[{i}]",
        )
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{sheet_name}'!R{row}",
            local_value=_list_value(payload.get("item_eval_empresa"), i),
            remote_value=_list_value(remote_payload.get("item_eval_empresa"), i),
            label=f"Evaluacion de empresa: {item_label}",
            field_id=f"item_eval_empresa[{i}]",
        )

    for i, row in enumerate(range(34, 42)):
        company_label = _followup_company_label(payload, remote_payload, i)
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{sheet_name}'!J{row}",
            local_value=_list_value(payload.get("empresa_eval"), i),
            remote_value=_list_value(remote_payload.get("empresa_eval"), i),
            label=f"Evaluacion empresarial: {company_label}",
            field_id=f"empresa_eval[{i}]",
        )
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{sheet_name}'!L{row}",
            local_value=_list_value(payload.get("empresa_observacion"), i),
            remote_value=_list_value(remote_payload.get("empresa_observacion"), i),
            label=f"Observacion empresarial: {company_label}",
            field_id=f"empresa_observacion[{i}]",
        )

    for i, row in enumerate(range(47, 51), start=1):
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{sheet_name}'!D{row}",
            local_value=_attendee_value(payload.get("asistentes"), i - 1, "nombre"),
            remote_value=_attendee_value(remote_payload.get("asistentes"), i - 1, "nombre"),
            label=f"Asistente {i} nombre",
            field_id=f"asistentes[{i - 1}].nombre",
        )
        _append_sparse_update(
            updates,
            changes,
            range_name=f"'{sheet_name}'!N{row}",
            local_value=_attendee_value(payload.get("asistentes"), i - 1, "cargo"),
            remote_value=_attendee_value(remote_payload.get("asistentes"), i - 1, "cargo"),
            label=f"Asistente {i} cargo",
            field_id=f"asistentes[{i - 1}].cargo",
        )

    followup_date_range = _get_followup_date_cell(base_sheet_name, index)
    if followup_date_range:
        local_date = _get_str(payload.get("fecha_seguimiento"))
        remote_date = _get_followup_date_from_base(base_remote_payload, index)
        if local_date and local_date != remote_date:
            updates.append({"range": followup_date_range, "value": local_date})

    return updates, changes


def _build_base_sheet_updates(payload, base_sheet_name=SHEET_BASE):
    updates = []
    for field_id, cell in BASE_SHEET_FIELD_MAP.items():
        updates.append({"range": f"'{base_sheet_name}'!{cell}", "value": payload.get(field_id, "")})
    for idx, row in enumerate(range(23, 28)):
        updates.append({"range": f"'{base_sheet_name}'!B{row}", "value": (payload.get('funciones_1_5') or [''] * 5)[idx] if idx < len(payload.get('funciones_1_5') or []) else ""})
        updates.append({"range": f"'{base_sheet_name}'!N{row}", "value": (payload.get('funciones_6_10') or [''] * 5)[idx] if idx < len(payload.get('funciones_6_10') or []) else ""})
    for idx, row in enumerate(range(29, 32)):
        updates.append({"range": f"'{base_sheet_name}'!C{row}", "value": (payload.get('seguimiento_fechas_1_3') or [''] * 3)[idx] if idx < len(payload.get('seguimiento_fechas_1_3') or []) else ""})
        updates.append({"range": f"'{base_sheet_name}'!P{row}", "value": (payload.get('seguimiento_fechas_4_6') or [''] * 3)[idx] if idx < len(payload.get('seguimiento_fechas_4_6') or []) else ""})
    ponderado_payload = _build_ponderado_payload(payload)
    for field_id, cell in PONDERADO_COMPANY_MAP.items():
        updates.append({"range": f"'{SHEET_FINAL}'!{cell}", "value": ponderado_payload.get(field_id, "")})
    for field_id, cell in PONDERADO_USER_MAP.items():
        updates.append({"range": f"'{SHEET_FINAL}'!{cell}", "value": ponderado_payload.get(field_id, "")})
    return updates


def _get_followup_date_cell(base_sheet_name, index):
    try:
        idx = int(index)
    except Exception:
        return ""
    if 1 <= idx <= 3:
        return f"'{base_sheet_name}'!C{28 + idx}"
    if 4 <= idx <= 6:
        return f"'{base_sheet_name}'!P{25 + idx}"
    return ""


def _set_followup_date_in_base_payload(payload, index, value):
    payload = dict(payload or {})
    text = _get_str(value)
    try:
        idx = int(index)
    except Exception:
        return payload
    if 1 <= idx <= 3:
        values = list(payload.get("seguimiento_fechas_1_3") or ["", "", ""])
        while len(values) < 3:
            values.append("")
        values[idx - 1] = text
        payload["seguimiento_fechas_1_3"] = values[:3]
        return payload
    if 4 <= idx <= 6:
        values = list(payload.get("seguimiento_fechas_4_6") or ["", "", ""])
        while len(values) < 3:
            values.append("")
        values[idx - 4] = text
        payload["seguimiento_fechas_4_6"] = values[:3]
    return payload


def _build_followup_sheet_updates(index, payload):
    sheet_name = _get_followup_sheet_name(index)
    updates = [
        {"range": f"'{sheet_name}'!E8", "value": payload.get("modalidad", "")},
        {"range": f"'{sheet_name}'!P8", "value": payload.get("seguimiento_numero", index)},
        {"range": f"'{sheet_name}'!{FOLLOWUP_DATE_LABEL_CELL}", "value": FOLLOWUP_DATE_LABEL},
        {"range": f"'{sheet_name}'!{FOLLOWUP_DATE_VALUE_CELL}", "value": payload.get("fecha_seguimiento", "")},
        {"range": f"'{sheet_name}'!J31", "value": payload.get("tipo_apoyo", "")},
        {"range": f"'{sheet_name}'!A43", "value": payload.get("situacion_encontrada", "")},
        {"range": f"'{sheet_name}'!A45", "value": payload.get("estrategias_ajustes", "")},
    ]
    for i, row in enumerate(range(12, 31)):
        updates.append({"range": f"'{sheet_name}'!G{row}", "value": (payload.get('item_observaciones') or [])[i] if i < len(payload.get('item_observaciones') or []) else ""})
        updates.append({"range": f"'{sheet_name}'!O{row}", "value": (payload.get('item_autoevaluacion') or [])[i] if i < len(payload.get('item_autoevaluacion') or []) else ""})
        updates.append({"range": f"'{sheet_name}'!R{row}", "value": (payload.get('item_eval_empresa') or [])[i] if i < len(payload.get('item_eval_empresa') or []) else ""})
    for i, row in enumerate(range(34, 42)):
        updates.append({"range": f"'{sheet_name}'!J{row}", "value": (payload.get('empresa_eval') or [])[i] if i < len(payload.get('empresa_eval') or []) else ""})
        updates.append({"range": f"'{sheet_name}'!L{row}", "value": (payload.get('empresa_observacion') or [])[i] if i < len(payload.get('empresa_observacion') or []) else ""})
    for i, row in enumerate(range(47, 51)):
        asistentes = payload.get("asistentes") or []
        entry = asistentes[i] if i < len(asistentes) else {}
        updates.append({"range": f"'{sheet_name}'!D{row}", "value": entry.get("nombre", "")})
        updates.append({"range": f"'{sheet_name}'!N{row}", "value": entry.get("cargo", "")})
    return updates


def _build_empty_followup_payload(index):
    return {
        "modalidad": "",
        "seguimiento_numero": "",
        "fecha_seguimiento": "",
        "item_observaciones": ["" for _ in range(19)],
        "item_autoevaluacion": ["" for _ in range(19)],
        "item_eval_empresa": ["" for _ in range(19)],
        "tipo_apoyo": "",
        "empresa_eval": ["" for _ in range(8)],
        "empresa_observacion": ["" for _ in range(8)],
        "situacion_encontrada": "",
        "estrategias_ajustes": "",
        "asistentes": [{"nombre": "", "cargo": ""} for _ in range(4)],
    }


def _apply_empty_followup_fields(payload, index):
    payload = dict(payload or {})
    payload.update(_build_empty_followup_payload(index))
    return payload


def _clear_followup_sheet_worksheet(ws, index):
    empty_payload = _build_empty_followup_payload(index)
    ws["E8"].value = empty_payload["modalidad"]
    ws["P8"].value = empty_payload["seguimiento_numero"]
    ws[FOLLOWUP_DATE_LABEL_CELL].value = FOLLOWUP_DATE_LABEL
    ws[FOLLOWUP_DATE_VALUE_CELL].value = empty_payload["fecha_seguimiento"]
    for pos, row in enumerate(range(12, 31)):
        ws[f"G{row}"].value = empty_payload["item_observaciones"][pos]
        ws[f"O{row}"].value = empty_payload["item_autoevaluacion"][pos]
        ws[f"R{row}"].value = empty_payload["item_eval_empresa"][pos]
    ws["J31"].value = empty_payload["tipo_apoyo"]
    for pos, row in enumerate(range(34, 42)):
        ws[f"J{row}"].value = empty_payload["empresa_eval"][pos]
        ws[f"L{row}"].value = empty_payload["empresa_observacion"][pos]
    ws["A43"].value = empty_payload["situacion_encontrada"]
    ws["A45"].value = empty_payload["estrategias_ajustes"]
    for pos, row in enumerate(range(47, 51)):
        ws[f"D{row}"].value = empty_payload["asistentes"][pos]["nombre"]
        ws[f"N{row}"].value = empty_payload["asistentes"][pos]["cargo"]


def ensure_case_workbook(cedula, user_row, is_compensar):
    raise RuntimeError(
        "Seguimientos ya no admite workbooks locales. "
        "Usa ensure_case_record() con una referencia nativa de Google Sheets."
    )
    normalized = _normalize_cedula(cedula)
    if not normalized:
        raise ValueError("Cédula inválida.")
    if not user_row:
        raise ValueError("No se encontró usuario para la cédula indicada.")



def _set_sheet_visibility(spreadsheet_id, max_seguimientos):
    try:
        service = get_google_sheets_service()
        spreadsheet = get_spreadsheet(spreadsheet_id, include_grid_data=False)
    except Exception:
        return
    requests = []
    try:
        base_sheet_name = _get_base_sheet_name_from_spreadsheet(spreadsheet)
    except Exception:
        base_sheet_name = SHEET_BASE
    relevant_titles = {
        base_sheet_name,
        SHEET_FINAL,
        *(f"{SHEET_PREFIX}{i}" for i in range(1, 7)),
    }
    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties", {}) or {}
        title = str(props.get("title") or "")
        try:
            hidden = bool(title not in relevant_titles)
        except Exception:
            hidden = False
        requests.append(
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": props.get("sheetId"),
                        "hidden": hidden,
                    },
                    "fields": "hidden",
                }
            }
        )
    if not requests:
        return
    execute_google_request_with_retry(
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests},
        ),
        operation_name="seguimientos.set_sheet_visibility",
    )


def _create_native_case_record(service, folder_id, folder_name, cedula, user_row, max_seguimientos):
    template_id = get_seguimientos_template_id()
    request_id = uuid.uuid4().hex
    copied = execute_google_create_with_confirmation(
        lambda: service.files().copy(
            fileId=template_id,
            body={
                "name": folder_name,
                "parents": [folder_id],
                "appProperties": {
                    "kind": "seguimiento_il",
                    "request_id": request_id,
                    "cedula": str(cedula),
                    "max_seguimientos": str(max_seguimientos),
                },
            },
            fields="id,name,mimeType,webViewLink,parents,appProperties",
            supportsAllDrives=True,
        ),
        lambda: _find_drive_file_by_request_id(service, folder_id, folder_name, request_id),
        operation_name="seguimientos.copy_native_case",
    )
    record = {
        "source": "drive",
        "cedula": cedula,
        "folder_id": folder_id,
        "folder_name": folder_name,
        "file_id": copied.get("id"),
        "file_name": copied.get("name"),
        "mime_type": copied.get("mimeType") or GOOGLE_SHEETS_MIME,
        "webViewLink": copied.get("webViewLink") or f"https://docs.google.com/spreadsheets/d/{copied.get('id')}/edit",
        "modifiedTime": "",
        "local_path": "",
        "max_seguimientos": max_seguimientos,
        "appProperties": copied.get("appProperties") or {},
    }
    base_payload = _build_base_payload_from_user_row(user_row)
    clear_protected_ranges(record["file_id"])
    spreadsheet = get_spreadsheet(record["file_id"], include_grid_data=False)
    base_sheet_name = _get_base_sheet_name_from_spreadsheet(spreadsheet)
    updates = _build_base_sheet_updates(base_payload, base_sheet_name=base_sheet_name)
    for idx in range(1, 7):
        updates.extend(_build_followup_sheet_updates(idx, _build_empty_followup_payload(idx)))
    batch_write_sheet_updates(record["file_id"], updates)
    _set_sheet_visibility(record["file_id"], max_seguimientos)
    return record


def ensure_case_record(cedula, user_row, is_compensar):
    normalized = _normalize_cedula(cedula)
    if not normalized:
        raise ValueError("Cédula inválida.")
    existing = find_case_record(normalized, user_row.get("nombre_usuario"))
    if existing:
        max_seguimientos = int(existing.get("max_seguimientos") or (6 if bool(is_compensar) else 3))
        save_base_payload(existing, _build_base_payload_from_user_row(user_row), overwrite=False)
        try:
            existing["max_seguimientos"] = int(
                get_case_meta(existing).get("max_seguimientos") or 3
            )
        except Exception:
            existing["max_seguimientos"] = 6 if bool(is_compensar) else 3
        return {
            "record": existing,
            "created": False,
            "max_seguimientos": int(existing.get("max_seguimientos") or (6 if bool(is_compensar) else 3)),
        }

    service = _get_drive_service()
    seguimientos_folder_id = _ensure_seguimientos_folder(service)
    folder_name = build_case_folder_name(user_row.get("nombre_usuario"), normalized)
    case_folder_id = drive_upload._get_or_create_folder(service, seguimientos_folder_id, folder_name)
    max_seguimientos = 6 if bool(is_compensar) else 3
    record = _create_native_case_record(
        service,
        case_folder_id,
        folder_name,
        normalized,
        user_row,
        max_seguimientos,
    )
    return {"record": record, "created": True, "max_seguimientos": max_seguimientos}


def sync_case_record_from_local(record, local_path):
    return None

def _get_base_sheet_name_from_spreadsheet(spreadsheet):
    for sheet in spreadsheet.get("sheets", []):
        title = str(((sheet.get("properties") or {}).get("title")) or "")
        if title in BASE_SHEET_CANDIDATES:
            return title
    raise ValueError("No existe la hoja base de seguimientos en el spreadsheet.")


def _infer_max_seguimientos_from_spreadsheet(spreadsheet):
    visible = 0
    total = 0
    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties") or {}
        title = str(props.get("title") or "")
        if not title.startswith(SHEET_PREFIX):
            continue
        total += 1
        if not bool(props.get("hidden")):
            visible += 1
    if visible >= 6:
        return 6
    if visible >= 3:
        return 3
    if total >= 6:
        return 6
    return 3


def describe_case(case_ref):
    if not case_ref:
        return {
            "empresa": "",
            "seguimientos": [],
            "seguimientos_count": 0,
            "ultimo_seguimiento": "",
            "completed_sheets": [],
            "sheet_progress": [],
            "sheet_progress_summary": [],
        }
    payload = get_base_payload(case_ref)
    workflow = get_workflow_state(case_ref)
    followup_dates = []
    for idx, value in enumerate(payload.get("seguimiento_fechas_1_3") or [], start=1):
        if str(value or "").strip():
            followup_dates.append(f"S{idx}: {str(value).strip()}")
    for offset, value in enumerate(payload.get("seguimiento_fechas_4_6") or [], start=4):
        if str(value or "").strip():
            followup_dates.append(f"S{offset}: {str(value).strip()}")
    progress_summary = []
    for entry in list(workflow.get("sheet_progress") or []):
        label = str(entry.get("label") or entry.get("sheet_name") or "").strip()
        status = str(entry.get("status") or "").strip()
        coverage = int(entry.get("coverage_percent") or 0)
        if status == "review_only":
            progress_summary.append(f"{label}: solo lectura")
        else:
            progress_summary.append(f"{label}: {coverage}% ({status})")
    return {
        "empresa": str(payload.get("nombre_empresa") or "").strip(),
        "profesional_asignado": str(payload.get("profesional_asignado") or "").strip(),
        "seguimientos": followup_dates,
        "seguimientos_count": len(followup_dates),
        "ultimo_seguimiento": followup_dates[-1] if followup_dates else "",
        "completed_sheets": list(workflow.get("completed_sheets") or []),
        "sheet_progress": list(workflow.get("sheet_progress") or []),
        "sheet_progress_summary": progress_summary,
        "suggested_sheet": str(workflow.get("suggested_sheet") or ""),
        "base_sheet_name": str(workflow.get("base_sheet_name") or SHEET_BASE),
        "max_seguimientos": int(workflow.get("max_seguimientos") or 3),
    }


def _cell_has_value(ws, cell):
    value = ws[cell].value
    if value is None:
        return False
    return str(value).strip() != ""


def _is_base_completed(wb):
    try:
        ws = wb[_get_base_sheet_name_from_workbook(wb)]
    except Exception:
        return False
    required = ["A16", "E16", "A18", "N18"]
    return all(_cell_has_value(ws, c) for c in required)


def _is_followup_completed(ws):
    required = ["O12", "R12", "J31"]
    return all(_cell_has_value(ws, c) for c in required)


def _is_base_payload_completed(payload):
    payload = payload or {}
    required = ["nombre_vinculado", "cedula", "cargo_vinculado", "discapacidad"]
    return all(str(payload.get(item) or "").strip() for item in required)


def _is_followup_payload_completed(payload):
    payload = payload or {}
    auto_vals = payload.get("item_autoevaluacion") or []
    emp_vals = payload.get("item_eval_empresa") or []
    first_auto = str(auto_vals[0] or "").strip() if auto_vals else ""
    first_emp = str(emp_vals[0] or "").strip() if emp_vals else ""
    tipo_apoyo = str(payload.get("tipo_apoyo") or "").strip()
    return bool(first_auto and first_emp and tipo_apoyo)


def _get_followup_date_from_base(base_payload, index):
    try:
        idx = int(index)
    except Exception:
        return ""
    if 1 <= idx <= 3:
        values = base_payload.get("seguimiento_fechas_1_3") or []
        pos = idx - 1
    elif 4 <= idx <= 6:
        values = base_payload.get("seguimiento_fechas_4_6") or []
        pos = idx - 4
    else:
        return ""
    return str(values[pos] or "").strip() if pos < len(values) else ""


def _has_followup_meaningful_content(payload):
    payload = payload or {}
    if str(payload.get("situacion_encontrada") or "").strip():
        return True
    if str(payload.get("estrategias_ajustes") or "").strip():
        return True
    for value in (payload.get("item_observaciones") or []):
        if str(value or "").strip():
            return True
    for value in (payload.get("empresa_observacion") or []):
        if str(value or "").strip():
            return True
    return False


def _normalize_empty_followup_payload(payload, index, base_payload=None):
    payload = dict(payload or {})
    has_followup_date = bool(_get_followup_date_from_base(base_payload or {}, index))
    explicit_date = _get_str(payload.get("fecha_seguimiento"))
    if not explicit_date and has_followup_date:
        payload["fecha_seguimiento"] = _get_followup_date_from_base(base_payload or {}, index)
        explicit_date = _get_str(payload.get("fecha_seguimiento"))
    if explicit_date or has_followup_date or _has_followup_meaningful_content(payload):
        return payload
    payload["modalidad"] = ""
    payload["seguimiento_numero"] = str(index)
    payload["fecha_seguimiento"] = ""
    payload["item_observaciones"] = ["" for _ in (payload.get("item_observaciones") or [])]
    payload["item_autoevaluacion"] = ["" for _ in (payload.get("item_autoevaluacion") or [])]
    payload["item_eval_empresa"] = ["" for _ in (payload.get("item_eval_empresa") or [])]
    payload["tipo_apoyo"] = ""
    payload["empresa_eval"] = ["" for _ in (payload.get("empresa_eval") or [])]
    payload["empresa_observacion"] = ["" for _ in (payload.get("empresa_observacion") or [])]
    payload["situacion_encontrada"] = ""
    payload["estrategias_ajustes"] = ""
    payload["asistentes"] = [{"nombre": "", "cargo": ""} for _ in (payload.get("asistentes") or [])]
    return payload


def _normalize_base_payload(payload):
    payload = dict(payload or {})
    has_started = any(
        [
            str(payload.get("fecha_visita") or "").strip(),
            str(payload.get("modalidad") or "").strip(),
            str(payload.get("apoyos_ajustes") or "").strip(),
        ]
    )
    for value in (payload.get("seguimiento_fechas_1_3") or []):
        if str(value or "").strip():
            has_started = True
            break
    if not has_started:
        for value in (payload.get("seguimiento_fechas_4_6") or []):
            if str(value or "").strip():
                has_started = True
                break
    if not has_started:
        payload["fecha_fin_contrato"] = ""
    return payload


def _coverage_percent(filled, total):
    try:
        filled_int = int(filled or 0)
        total_int = int(total or 0)
    except Exception:
        return 0
    if total_int <= 0:
        return 0
    return int((filled_int * 100) / total_int)


def _build_progress_snapshot(*, values):
    flags = [bool(item) for item in list(values or [])]
    total = len(flags)
    filled = sum(1 for item in flags if item)
    coverage = _coverage_percent(filled, total)
    if filled <= 0:
        status = "not_started"
    elif coverage >= SHEET_COVERAGE_THRESHOLD:
        status = "completed"
    else:
        status = "in_progress"
    return {
        "filled": filled,
        "total": total,
        "coverage_percent": coverage,
        "status": status,
        "is_completed": status == "completed",
    }


def _has_value(value):
    return bool(_get_str(value))


def _build_base_progress_snapshot(payload):
    payload = dict(payload or {})
    flags = [_has_value(payload.get(field_id)) for field_id in BASE_PROGRESS_SCALAR_FIELDS]
    flags.extend(_has_value(item) for item in list(payload.get("funciones_1_5") or [])[:5])
    flags.extend(_has_value(item) for item in list(payload.get("funciones_6_10") or [])[:5])
    return _build_progress_snapshot(values=flags)


def _build_followup_progress_snapshot(payload):
    payload = dict(payload or {})
    flags = [_has_value(payload.get(field_id)) for field_id in FOLLOWUP_PROGRESS_SCALAR_FIELDS]
    flags.extend(_has_value(item) for item in list(payload.get("item_autoevaluacion") or [])[:19])
    flags.extend(_has_value(item) for item in list(payload.get("item_eval_empresa") or [])[:19])
    flags.extend(_has_value(item) for item in list(payload.get("empresa_eval") or [])[:8])
    return _build_progress_snapshot(values=flags)


def _sheet_stage_title(sheet_name, base_sheet_name):
    current = str(sheet_name or "").strip()
    if current == str(base_sheet_name or "").strip():
        return "Ficha inicial del proceso"
    if current == SHEET_FINAL:
        return "Resultado final"
    match = re.search(r"(\d+)$", current)
    if match:
        return f"Seguimiento {int(match.group(1))}"
    return current


def _sheet_stage_id(sheet_name, base_sheet_name):
    current = str(sheet_name or "").strip()
    if current == str(base_sheet_name or "").strip():
        return "base_process"
    if current == SHEET_FINAL:
        return "final_result"
    match = re.search(r"(\d+)$", current)
    if match:
        return f"followup_{int(match.group(1))}"
    return re.sub(r"[^a-z0-9]+", "_", current.lower()).strip("_") or "sheet_stage"


def _sheet_stage_helper_text(*, sheet_name, base_sheet_name, status, is_suggested, is_editable):
    current = str(sheet_name or "").strip()
    current_status = str(status or "").strip()
    if current == SHEET_FINAL:
        return "Consolidado automático del proceso. Solo lectura."
    if current == str(base_sheet_name or "").strip():
        if is_suggested and current_status != "completed":
            return "Empieza aquí para dejar lista la ficha inicial del caso."
        if current_status == "completed":
            return "La ficha inicial está completa y lista para soportar los seguimientos."
        if current_status == "in_progress":
            return "Completa la información inicial del proceso y los apoyos requeridos."
        return "Registra los datos base del caso, la visita y los apoyos requeridos."
    match = re.search(r"(\d+)$", current)
    followup_number = int(match.group(1)) if match else 0
    if current_status == "completed":
        return (
            f"Seguimiento {followup_number} registrado. Puedes revisarlo o corregirlo si hace falta."
        )
    if is_suggested and is_editable:
        return f"Esta es la etapa sugerida para continuar con Seguimiento {followup_number}."
    if current_status == "in_progress":
        return f"Seguimiento {followup_number} en curso."
    return f"Seguimiento {followup_number} pendiente por diligenciar."


def _build_sheet_stage_entry(
    sheet_name,
    base_sheet_name,
    *,
    status,
    coverage_percent,
    is_completed=False,
    is_suggested=False,
    is_editable=False,
):
    title = _sheet_stage_title(sheet_name, base_sheet_name)
    return {
        "stage_id": _sheet_stage_id(sheet_name, base_sheet_name),
        "title": title,
        "label": title,
        "sheet_name": str(sheet_name or "").strip(),
        "status": str(status or "").strip(),
        "coverage_percent": int(coverage_percent or 0),
        "is_completed": bool(is_completed),
        "is_suggested": bool(is_suggested),
        "is_editable": bool(is_editable),
        "helper_text": _sheet_stage_helper_text(
            sheet_name=sheet_name,
            base_sheet_name=base_sheet_name,
            status=status,
            is_suggested=is_suggested,
            is_editable=is_editable,
        ),
    }


def _sheet_label(sheet_name, base_sheet_name):
    return _sheet_stage_title(sheet_name, base_sheet_name)


def _legacy_get_workflow_state(case_ref):
    if not case_ref:
        return {
            "base_sheet_name": SHEET_BASE,
            "max_seguimientos": 3,
            "base_completed": False,
            "completed_followups": [],
            "next_followup": 1,
            "editable_sheet": SHEET_BASE,
            "suggested_sheet": SHEET_BASE,
            "visible_sheets": [SHEET_BASE, f"{SHEET_PREFIX}1"],
            "message": "Completa primero la ficha inicial del proceso.",
        }

    meta = get_case_meta(case_ref)
    base_sheet_name = str(meta.get("base_sheet_name") or SHEET_BASE)
    max_seguimientos = 6 if int(meta.get("max_seguimientos") or 3) >= 6 else 3
    base_payload = get_base_payload(case_ref)
    base_completed = _is_base_payload_completed(base_payload)

    completed_followups = []
    next_followup = 1
    if base_completed:
        next_followup = None
        for idx in range(1, max_seguimientos + 1):
            has_followup_date = bool(_get_followup_date_from_base(base_payload, idx))
            followup_payload = get_followup_payload(case_ref, idx)
            followup_completed = has_followup_date and _is_followup_payload_completed(
                followup_payload
            )
            if followup_completed:
                completed_followups.append(idx)
                continue
            next_followup = idx
            break

    if not base_completed:
        editable_sheet = base_sheet_name
        suggested_sheet = base_sheet_name
        visible_followups = 1
        message = "Completa primero la ficha inicial del proceso."
    elif next_followup is not None and next_followup <= max_seguimientos:
        editable_sheet = f"{SHEET_PREFIX}{next_followup}"
        suggested_sheet = base_sheet_name if next_followup == 1 else editable_sheet
        visible_followups = next_followup
        if next_followup == 1:
            message = "Hoja base y seguimiento 1 habilitados hasta diligenciar el seguimiento 1."
        else:
            message = (
                f"Seguimiento {next_followup} habilitado. "
                f"En la ficha inicial solo verás la fecha alimentada desde Seguimiento {next_followup}."
            )
    else:
        editable_sheet = ""
        suggested_sheet = f"{SHEET_PREFIX}{max_seguimientos}"
        visible_followups = max_seguimientos
        message = "Todos los seguimientos están diligenciados."

    visible_sheets = [base_sheet_name] + [
        f"{SHEET_PREFIX}{idx}" for idx in range(1, visible_followups + 1)
    ]

    return {
        "base_sheet_name": base_sheet_name,
        "max_seguimientos": max_seguimientos,
        "base_completed": base_completed,
        "completed_followups": completed_followups,
        "next_followup": next_followup,
        "editable_sheet": editable_sheet,
        "suggested_sheet": suggested_sheet,
        "visible_sheets": visible_sheets,
        "message": message,
    }


def _legacy_suggest_next_step(case_ref):
    if not case_ref:
        return {"sheet": SHEET_BASE, "message": "Inicia con la ficha inicial del proceso.", "max_seguimientos": 3}
    workflow = get_workflow_state(case_ref)
    return {
        "sheet": workflow.get("suggested_sheet") or workflow.get("base_sheet_name") or SHEET_BASE,
        "message": workflow.get("message") or "",
        "max_seguimientos": int(workflow.get("max_seguimientos") or 3),
    }


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
            "certificado_discapacidad",
            "certificado_porcentaje",
            "telefono_oferente",
            "correo_oferente",
            "cargo_oferente",
            "contacto_emergencia",
            "parentesco",
            "telefono_emergencia",
            "fecha_firma_contrato",
            "tipo_contrato",
            "fecha_fin",
            "empresa_nit",
            "empresa_nombre",
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
    if not nit:
        return None
    nit = "".join(str(nit).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {
        "select": select_cols,
        "nit_empresa": f"eq.{nit}",
        "limit": 1,
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    return (
        _merge_company_row_with_cache(
            _map_company_row(data[0]),
            field_map=SECTION_1_SUPABASE_MAP,
            nit=nit,
        )
        if data
        else None
    )


def get_empresa_by_nombre(nombre, env_path=".env"):
    if not nombre:
        return None
    nombre = " ".join(str(nombre).split())
    select_cols = ",".join(sorted(set(SECTION_1_SUPABASE_MAP.values()) | {"nit_empresa"}))
    params = {
        "select": select_cols,
        "nombre_empresa": f"ilike.{nombre}",
        "limit": 2,
    }
    data = _supabase_get("empresas", params, env_path=env_path)
    if not data:
        return None
    if len(data) > 1:
        raise ValueError("Hay más de una empresa con ese nombre. Usa el NIT.")
    return _merge_company_row_with_cache(
        _map_company_row(data[0]),
        field_map=SECTION_1_SUPABASE_MAP,
        nombre=nombre,
    )


def get_empresas_by_nombre_prefix(prefix, env_path=".env", limit=50):
    text = str(prefix or "").strip()
    if not text:
        return []
    params = {
        "select": "nombre_empresa,nit_empresa",
        "nombre_empresa": f"ilike.{text}%",
        "order": "nombre_empresa.asc",
        "limit": int(limit),
    }
    return _supabase_get("empresas", params, env_path=env_path)


def get_empresas_by_nit_prefix(prefix, env_path=".env", limit=10):
    text = "".join(str(prefix or "").split())
    if not text:
        return []
    params = {
        "select": "nombre_empresa,nit_empresa",
        "nit_empresa": f"like.{text}%",
        "order": "nit_empresa.asc",
        "limit": int(limit),
    }
    return _supabase_get("empresas", params, env_path=env_path)


def _ensure_sheet_exists(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"No existe la hoja '{sheet_name}' en el archivo.")
    return wb[sheet_name]


def _cell_value(ws, address):
    value = ws[address].value
    if value is None:
        return ""
    return str(value).strip()


def get_case_meta(case_ref):
    spreadsheet_id = _get_spreadsheet_id_from_case_ref(case_ref)
    spreadsheet = get_spreadsheet(spreadsheet_id, include_grid_data=False)
    props = case_ref.get("appProperties") if isinstance(case_ref, dict) else {}
    try:
        max_seg = int(
            ((case_ref or {}).get("max_seguimientos") if isinstance(case_ref, dict) else 0)
            or props.get("max_seguimientos")
            or _infer_max_seguimientos_from_spreadsheet(spreadsheet)
        )
    except Exception:
        max_seg = _infer_max_seguimientos_from_spreadsheet(spreadsheet)
    max_seg = 6 if max_seg >= 6 else 3
    return {
        "cedula": _get_str(
            ((case_ref or {}).get("cedula") if isinstance(case_ref, dict) else "")
            or props.get("cedula")
        ),
        "nombre_usuario": _get_str(
            (case_ref or {}).get("folder_name") if isinstance(case_ref, dict) else ""
        ),
        "is_compensar": bool(max_seg >= 6),
        "max_seguimientos": max_seg,
        "base_sheet_name": _get_base_sheet_name_from_spreadsheet(spreadsheet),
    }


def get_base_payload(case_ref):
    meta = get_case_meta(case_ref)
    base_sheet_name = str(meta.get("base_sheet_name") or SHEET_BASE)
    ranges = [
        f"'{base_sheet_name}'!D8",
        f"'{base_sheet_name}'!R8",
        f"'{base_sheet_name}'!D9",
        f"'{base_sheet_name}'!R9",
        f"'{base_sheet_name}'!D10",
        f"'{base_sheet_name}'!R10",
        f"'{base_sheet_name}'!D11",
        f"'{base_sheet_name}'!R11",
        f"'{base_sheet_name}'!D12",
        f"'{base_sheet_name}'!R12",
        f"'{base_sheet_name}'!D13",
        f"'{base_sheet_name}'!R13",
        f"'{base_sheet_name}'!A16",
        f"'{base_sheet_name}'!E16",
        f"'{base_sheet_name}'!I16",
        f"'{base_sheet_name}'!K16",
        f"'{base_sheet_name}'!P16",
        f"'{base_sheet_name}'!S16",
        f"'{base_sheet_name}'!U16",
        f"'{base_sheet_name}'!A18",
        f"'{base_sheet_name}'!E18",
        f"'{base_sheet_name}'!I18",
        f"'{base_sheet_name}'!N18",
        f"'{base_sheet_name}'!C20",
        f"'{base_sheet_name}'!M20",
        f"'{base_sheet_name}'!T20",
        f"'{base_sheet_name}'!E21",
        f"'{base_sheet_name}'!B23:B27",
        f"'{base_sheet_name}'!N23:N27",
        f"'{base_sheet_name}'!C29:C31",
        f"'{base_sheet_name}'!P29:P31",
        f"'{SHEET_FINAL}'!D11",
        f"'{SHEET_FINAL}'!Q12",
        f"'{SHEET_FINAL}'!N18",
    ]
    values = _batch_read_sheet_values(_get_spreadsheet_id_from_case_ref(case_ref), ranges)
    payload = {
        "fecha_visita": _first_batch_value(values, f"'{base_sheet_name}'!D8"),
        "modalidad": _first_batch_value(values, f"'{base_sheet_name}'!R8"),
        "nombre_empresa": _first_batch_value(values, f"'{base_sheet_name}'!D9"),
        "ciudad_empresa": _first_batch_value(values, f"'{base_sheet_name}'!R9"),
        "direccion_empresa": _first_batch_value(values, f"'{base_sheet_name}'!D10"),
        "nit_empresa": _first_batch_value(values, f"'{base_sheet_name}'!R10"),
        "correo_1": _first_batch_value(values, f"'{base_sheet_name}'!D11"),
        "telefono_empresa": _first_batch_value(values, f"'{base_sheet_name}'!R11"),
        "contacto_empresa": _first_batch_value(values, f"'{base_sheet_name}'!D12"),
        "cargo": _first_batch_value(values, f"'{base_sheet_name}'!R12"),
        "asesor": _first_batch_value(values, f"'{base_sheet_name}'!D13"),
        "sede_empresa": _first_batch_value(values, f"'{base_sheet_name}'!R13"),
        "caja_compensacion": _first_batch_value(values, f"'{SHEET_FINAL}'!D11"),
        "profesional_asignado": _first_batch_value(values, f"'{SHEET_FINAL}'!Q12"),
        "nombre_vinculado": _first_batch_value(values, f"'{base_sheet_name}'!A16"),
        "cedula": _first_batch_value(values, f"'{base_sheet_name}'!E16"),
        "telefono_vinculado": _first_batch_value(values, f"'{base_sheet_name}'!I16"),
        "correo_vinculado": _first_batch_value(values, f"'{base_sheet_name}'!K16"),
        "contacto_emergencia": _first_batch_value(values, f"'{base_sheet_name}'!P16"),
        "parentesco": _first_batch_value(values, f"'{base_sheet_name}'!S16"),
        "telefono_emergencia": _first_batch_value(values, f"'{base_sheet_name}'!U16"),
        "cargo_vinculado": _first_batch_value(values, f"'{base_sheet_name}'!A18"),
        "certificado_discapacidad": _first_batch_value(values, f"'{base_sheet_name}'!E18"),
        "certificado_porcentaje": _first_batch_value(values, f"'{base_sheet_name}'!I18"),
        "discapacidad": _first_batch_value(values, f"'{base_sheet_name}'!N18"),
        "tipo_contrato": _first_batch_value(values, f"'{base_sheet_name}'!C20"),
        "fecha_inicio_contrato": _first_batch_value(values, f"'{base_sheet_name}'!M20"),
        "fecha_fin_contrato": _first_batch_value(values, f"'{base_sheet_name}'!T20"),
        "fecha_firma_contrato": _first_batch_value(values, f"'{SHEET_FINAL}'!N18"),
        "apoyos_ajustes": _first_batch_value(values, f"'{base_sheet_name}'!E21"),
        "funciones_1_5": _column_batch_values(values, f"'{base_sheet_name}'!B23:B27", 5),
        "funciones_6_10": _column_batch_values(values, f"'{base_sheet_name}'!N23:N27", 5),
        "seguimiento_fechas_1_3": _column_batch_values(values, f"'{base_sheet_name}'!C29:C31", 3),
        "seguimiento_fechas_4_6": _column_batch_values(values, f"'{base_sheet_name}'!P29:P31", 3),
    }
    return _normalize_base_payload(payload)

def build_base_save_plan(case_ref, payload, *, overwrite=True, save_mode="full"):
    spreadsheet_id = _get_spreadsheet_id_from_case_ref(case_ref)
    base_sheet_name = str(get_case_meta(case_ref).get("base_sheet_name") or SHEET_BASE)
    if str(save_mode or "full").strip().lower() != "diff":
        if not overwrite:
            existing = get_base_payload(case_ref)
            merged = dict(existing)
            for key, value in (payload or {}).items():
                if isinstance(value, list):
                    current_list = list(existing.get(key) or [])
                    merged_list = []
                    for idx, item in enumerate(value):
                        current_value = current_list[idx] if idx < len(current_list) else ""
                        merged_list.append(current_value if str(current_value or "").strip() else item)
                    merged[key] = merged_list
                else:
                    current_value = existing.get(key)
                    merged[key] = current_value if str(current_value or "").strip() else value
            payload = merged
        updates = _build_base_sheet_updates(payload, base_sheet_name=base_sheet_name)
        return _build_sparse_plan_result(
            spreadsheet_id=spreadsheet_id,
            sheet_name=base_sheet_name,
            updates=updates,
            changes=[],
            remote_payload={},
        )

    remote_payload = get_base_payload(case_ref)
    updates, changes = _build_base_sheet_sparse_updates(
        payload,
        remote_payload,
        base_sheet_name=base_sheet_name,
    )
    return _build_sparse_plan_result(
        spreadsheet_id=spreadsheet_id,
        sheet_name=base_sheet_name,
        updates=updates,
        changes=changes,
        remote_payload=remote_payload,
    )


def save_base_payload(case_ref, payload, overwrite=True, *, save_mode="full"):
    plan = build_base_save_plan(
        case_ref,
        payload,
        overwrite=overwrite,
        save_mode=save_mode,
    )
    if plan.get("updates"):
        batch_write_sheet_updates(plan["spreadsheet_id"], plan["updates"])
    return plan


def _get_followup_sheet_name(index):
    idx = int(index)
    if idx < 1 or idx > 6:
        raise ValueError("El seguimiento debe estar entre 1 y 6.")
    return f"{SHEET_PREFIX}{idx}"


def get_followup_payload(case_ref, index):
    sheet_name = _get_followup_sheet_name(index)
    base_payload = get_base_payload(case_ref)
    ranges = [
        f"'{sheet_name}'!E8",
        f"'{sheet_name}'!P8",
        f"'{sheet_name}'!{FOLLOWUP_DATE_VALUE_CELL}",
        f"'{sheet_name}'!A12:A30",
        f"'{sheet_name}'!G12:G30",
        f"'{sheet_name}'!O12:O30",
        f"'{sheet_name}'!R12:R30",
        f"'{sheet_name}'!J31",
        f"'{sheet_name}'!A34:A41",
        f"'{sheet_name}'!J34:J41",
        f"'{sheet_name}'!L34:L41",
        f"'{sheet_name}'!A43",
        f"'{sheet_name}'!A45",
        f"'{sheet_name}'!D47:D50",
        f"'{sheet_name}'!N47:N50",
    ]
    values = _batch_read_sheet_values(_get_spreadsheet_id_from_case_ref(case_ref), ranges)
    payload = {
        "modalidad": _first_batch_value(values, f"'{sheet_name}'!E8"),
        "seguimiento_numero": _first_batch_value(values, f"'{sheet_name}'!P8"),
        "fecha_seguimiento": _first_batch_value(values, f"'{sheet_name}'!{FOLLOWUP_DATE_VALUE_CELL}"),
        "item_labels": _column_batch_values(values, f"'{sheet_name}'!A12:A30", 19),
        "item_observaciones": _column_batch_values(values, f"'{sheet_name}'!G12:G30", 19),
        "item_autoevaluacion": _column_batch_values(values, f"'{sheet_name}'!O12:O30", 19),
        "item_eval_empresa": _column_batch_values(values, f"'{sheet_name}'!R12:R30", 19),
        "tipo_apoyo": _first_batch_value(values, f"'{sheet_name}'!J31"),
        "empresa_item_labels": _column_batch_values(values, f"'{sheet_name}'!A34:A41", 8),
        "empresa_eval": _column_batch_values(values, f"'{sheet_name}'!J34:J41", 8),
        "empresa_observacion": _column_batch_values(values, f"'{sheet_name}'!L34:L41", 8),
        "situacion_encontrada": _first_batch_value(values, f"'{sheet_name}'!A43"),
        "estrategias_ajustes": _first_batch_value(values, f"'{sheet_name}'!A45"),
        "asistentes": [
            {
                "nombre": (_column_batch_values(values, f"'{sheet_name}'!D47:D50", 4)[i]),
                "cargo": (_column_batch_values(values, f"'{sheet_name}'!N47:N50", 4)[i]),
            }
            for i in range(4)
        ],
    }
    return _normalize_empty_followup_payload(payload, index, base_payload=base_payload)


def build_followup_save_plan(case_ref, index, payload, *, save_mode="full"):
    base_sheet_name = str(get_case_meta(case_ref).get("base_sheet_name") or SHEET_BASE)
    spreadsheet_id = _get_spreadsheet_id_from_case_ref(case_ref)
    if str(save_mode or "full").strip().lower() != "diff":
        updates = _build_followup_sheet_updates(index, payload)
        followup_date_range = _get_followup_date_cell(base_sheet_name, index)
        if followup_date_range:
            updates.append({"range": followup_date_range, "value": _get_str(payload.get("fecha_seguimiento"))})
        return _build_sparse_plan_result(
            spreadsheet_id=spreadsheet_id,
            sheet_name=_get_followup_sheet_name(index),
            updates=updates,
            changes=[],
            remote_payload={},
        )

    remote_payload = get_followup_payload(case_ref, index)
    base_remote_payload = get_base_payload(case_ref)
    updates, changes = _build_followup_sheet_sparse_updates(
        index,
        payload,
        remote_payload,
        base_sheet_name=base_sheet_name,
        base_remote_payload=base_remote_payload,
    )
    return _build_sparse_plan_result(
        spreadsheet_id=spreadsheet_id,
        sheet_name=_get_followup_sheet_name(index),
        updates=updates,
        changes=changes,
        remote_payload=remote_payload,
    )


def save_followup_payload(case_ref, index, payload, *, save_mode="full"):
    plan = build_followup_save_plan(case_ref, index, payload, save_mode=save_mode)
    if plan.get("updates"):
        batch_write_sheet_updates(plan["spreadsheet_id"], plan["updates"])
    return plan


def get_workflow_state(case_ref):
    if not case_ref:
        visible_sheets = [SHEET_BASE] + [f"{SHEET_PREFIX}{idx}" for idx in range(1, 4)] + [SHEET_FINAL]
        base_entry = _build_sheet_stage_entry(
            SHEET_BASE,
            SHEET_BASE,
            status="not_started",
            coverage_percent=0,
            is_completed=False,
            is_suggested=True,
            is_editable=True,
        )
        followup_entries = [
            _build_sheet_stage_entry(
                f"{SHEET_PREFIX}{idx}",
                SHEET_BASE,
                status="not_started",
                coverage_percent=0,
                is_completed=False,
                is_suggested=False,
                is_editable=True,
            )
            for idx in range(1, 4)
        ]
        final_entry = _build_sheet_stage_entry(
            SHEET_FINAL,
            SHEET_BASE,
            status="review_only",
            coverage_percent=0,
            is_completed=False,
            is_suggested=False,
            is_editable=False,
        )
        stage_model = [base_entry] + followup_entries + [final_entry]
        return {
            "base_sheet_name": SHEET_BASE,
            "max_seguimientos": 3,
            "base_completed": False,
            "base_coverage_percent": 0,
            "completed_followups": [],
            "completed_sheets": [],
            "next_followup": 1,
            "editable_sheet": SHEET_BASE,
            "suggested_sheet": SHEET_BASE,
            "visible_sheets": visible_sheets,
            "sheet_progress": stage_model,
            "stage_model": stage_model,
            "message": "Empieza por la ficha inicial del proceso.",
        }

    meta = get_case_meta(case_ref)
    base_sheet_name = str(meta.get("base_sheet_name") or SHEET_BASE)
    max_seguimientos = 6 if int(meta.get("max_seguimientos") or 3) >= 6 else 3
    base_payload = get_base_payload(case_ref)
    base_progress = _build_base_progress_snapshot(base_payload)

    sheet_progress = []
    completed_followups = []
    completed_sheets = []
    next_followup = 1

    base_entry = _build_sheet_stage_entry(
        base_sheet_name,
        base_sheet_name,
        status=base_progress["status"],
        coverage_percent=base_progress["coverage_percent"],
        is_completed=bool(base_progress["is_completed"]),
        is_suggested=False,
        is_editable=True,
    )
    sheet_progress.append(base_entry)
    if base_progress["is_completed"]:
        completed_sheets.append(base_entry["label"])

    if base_progress["is_completed"]:
        next_followup = None
        for idx in range(1, max_seguimientos + 1):
            followup_sheet = _get_followup_sheet_name(idx)
            followup_payload = get_followup_payload(case_ref, idx)
            followup_progress = _build_followup_progress_snapshot(followup_payload)
            entry = _build_sheet_stage_entry(
                followup_sheet,
                base_sheet_name,
                status=followup_progress["status"],
                coverage_percent=followup_progress["coverage_percent"],
                is_completed=bool(followup_progress["is_completed"]),
                is_suggested=False,
                is_editable=True,
            )
            sheet_progress.append(entry)
            if followup_progress["is_completed"]:
                completed_followups.append(idx)
                completed_sheets.append(entry["label"])
                continue
            if next_followup is None:
                next_followup = idx
        if next_followup is None:
            suggested_sheet = SHEET_FINAL
            message = "Todos los seguimientos visibles ya alcanzaron el 90%."
        else:
            suggested_sheet = _get_followup_sheet_name(next_followup)
            if next_followup == 1:
                message = "La ficha inicial está completa. Continúa con Seguimiento 1."
            else:
                message = f"Continúa con Seguimiento {next_followup}."
    else:
        for idx in range(1, max_seguimientos + 1):
            followup_sheet = _get_followup_sheet_name(idx)
            sheet_progress.append(
                _build_sheet_stage_entry(
                    followup_sheet,
                    base_sheet_name,
                    status="not_started",
                    coverage_percent=0,
                    is_completed=False,
                    is_suggested=False,
                    is_editable=True,
                )
            )
        suggested_sheet = base_sheet_name
        message = "Empieza por la ficha inicial del proceso."

    sheet_progress.append(
        _build_sheet_stage_entry(
            SHEET_FINAL,
            base_sheet_name,
            status="review_only",
            coverage_percent=0,
            is_completed=False,
            is_suggested=False,
            is_editable=False,
        )
    )

    for entry in sheet_progress:
        if str(entry.get("sheet_name") or "").strip() == str(suggested_sheet or "").strip():
            entry["is_suggested"] = True
            entry["helper_text"] = _sheet_stage_helper_text(
                sheet_name=entry.get("sheet_name"),
                base_sheet_name=base_sheet_name,
                status=entry.get("status"),
                is_suggested=True,
                is_editable=entry.get("is_editable"),
            )
            break

    visible_sheets = [base_sheet_name] + [
        _get_followup_sheet_name(idx) for idx in range(1, max_seguimientos + 1)
    ] + [SHEET_FINAL]

    return {
        "base_sheet_name": base_sheet_name,
        "max_seguimientos": max_seguimientos,
        "base_completed": bool(base_progress["is_completed"]),
        "base_coverage_percent": base_progress["coverage_percent"],
        "completed_followups": completed_followups,
        "completed_sheets": completed_sheets,
        "next_followup": next_followup,
        "editable_sheet": suggested_sheet,
        "suggested_sheet": suggested_sheet,
        "visible_sheets": visible_sheets,
        "sheet_progress": sheet_progress,
        "stage_model": sheet_progress,
        "message": message,
    }


def suggest_next_step(case_ref):
    if not case_ref:
        return {"sheet": SHEET_BASE, "message": "Inicia con la ficha inicial del proceso.", "max_seguimientos": 3}
    workflow = get_workflow_state(case_ref)
    return {
        "sheet": workflow.get("suggested_sheet") or workflow.get("base_sheet_name") or SHEET_BASE,
        "message": workflow.get("message") or "",
        "max_seguimientos": int(workflow.get("max_seguimientos") or 3),
    }


def _normalize_export_date_string(value):
    text = _get_str(value)
    if not text:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except Exception:
            continue
    return ""


def _build_pdf_participants(base_payload):
    participant = {
        "nombre": _get_str((base_payload or {}).get("nombre_vinculado")),
        "cedula": _get_str((base_payload or {}).get("cedula")),
        "cargo": _get_str((base_payload or {}).get("cargo_vinculado")),
    }
    if not any(participant.values()):
        return []
    return [participant]


def _normalize_followup_attendees(payload):
    rows = []
    for entry in list((payload or {}).get("asistentes") or []):
        if not isinstance(entry, dict):
            continue
        normalized = {
            "nombre": _get_str(entry.get("nombre")),
            "cargo": _get_str(entry.get("cargo")),
        }
        if normalized["nombre"] or normalized["cargo"]:
            rows.append(normalized)
    return rows


def list_pdf_followup_candidates(case_ref):
    if not case_ref:
        return []
    workflow = get_workflow_state(case_ref)
    base_payload = get_base_payload(case_ref)
    max_seguimientos = int(workflow.get("max_seguimientos") or 3)
    stage_entries = {
        str((entry or {}).get("sheet_name") or "").strip(): dict(entry or {})
        for entry in list(workflow.get("sheet_progress") or [])
    }
    candidates = []
    for idx in range(1, max_seguimientos + 1):
        sheet_name = _get_followup_sheet_name(idx)
        entry = stage_entries.get(sheet_name) or {}
        status = str(entry.get("status") or "").strip()
        if status not in {"in_progress", "completed"}:
            continue
        payload = get_followup_payload(case_ref, idx)
        candidates.append(
            {
                "followup_index": idx,
                "sheet_name": sheet_name,
                "title": str(entry.get("title") or entry.get("label") or f"Seguimiento {idx}").strip(),
                "status": status,
                "fecha_seguimiento": _get_str(
                    payload.get("fecha_seguimiento") or _get_followup_date_from_base(base_payload, idx)
                ),
            }
        )
    return candidates


def build_pdf_export_bundle(case_ref, followup_index=None):
    if not case_ref:
        raise RuntimeError("No hay un caso válido para exportar el PDF de seguimiento.")

    meta = get_case_meta(case_ref)
    base_payload = get_base_payload(case_ref)
    base_sheet_name = str(meta.get("base_sheet_name") or SHEET_BASE).strip() or SHEET_BASE
    temp_parent_folder_id = ""
    if isinstance(case_ref, dict):
        temp_parent_folder_id = _get_str(case_ref.get("folder_id"))

    participants = _build_pdf_participants(base_payload)
    included_sheet_names = [base_sheet_name]
    nombre_empresa = _get_str(base_payload.get("nombre_empresa"))
    nit_empresa = _get_str(base_payload.get("nit_empresa"))
    cargo_objetivo = _get_str(base_payload.get("cargo_vinculado"))
    nombre_profesional = _get_str(base_payload.get("profesional_asignado") or base_payload.get("asesor"))

    common_metadata = {
        "tipo_acta": "seguimiento",
        "nit_empresa": nit_empresa,
        "nombre_empresa": nombre_empresa,
        "nombre_profesional": nombre_profesional,
        "participantes": participants,
        "cargo_objetivo": cargo_objetivo,
    }

    if followup_index is None:
        fecha_servicio = _normalize_export_date_string(base_payload.get("fecha_visita"))
        if not fecha_servicio:
            raise RuntimeError("La ficha inicial no tiene una fecha válida para generar el PDF.")
        acta_metadata = {
            **common_metadata,
            "document_variant": "base_only",
            "fecha_servicio": fecha_servicio,
            "modalidad_servicio": _get_str(base_payload.get("modalidad")),
            "asistentes": [],
            "included_sheet_names": included_sheet_names,
            "included_followup_index": None,
        }
        return {
            "tipo_acta": "seguimiento",
            "fecha_servicio": fecha_servicio,
            "extra_name": "Ficha inicial",
            "acta_metadata": acta_metadata,
            "selected_sheet_names": included_sheet_names,
            "temp_parent_folder_id": temp_parent_folder_id,
        }

    idx = int(followup_index)
    followup_sheet_name = _get_followup_sheet_name(idx)
    followup_payload = get_followup_payload(case_ref, idx)
    included_sheet_names = [base_sheet_name, followup_sheet_name]
    fecha_seguimiento = _get_str(
        followup_payload.get("fecha_seguimiento") or _get_followup_date_from_base(base_payload, idx)
    )
    fecha_servicio = _normalize_export_date_string(fecha_seguimiento) or _normalize_export_date_string(
        base_payload.get("fecha_visita")
    )
    if not fecha_servicio:
        raise RuntimeError(f"Seguimiento {idx} no tiene una fecha válida para generar el PDF.")
    asistentes = _normalize_followup_attendees(followup_payload)
    acta_metadata = {
        **common_metadata,
        "document_variant": "base_plus_followup",
        "fecha_servicio": fecha_servicio,
        "modalidad_servicio": _get_str(
            followup_payload.get("modalidad") or base_payload.get("modalidad")
        ),
        "included_sheet_names": included_sheet_names,
        "included_followup_index": idx,
        "numero_seguimiento": idx,
        "seguimiento_numero": _get_str(followup_payload.get("seguimiento_numero") or idx),
        "fecha_seguimiento": fecha_seguimiento,
        "tipo_apoyo": _get_str(followup_payload.get("tipo_apoyo")),
        "situacion_encontrada": _get_str(followup_payload.get("situacion_encontrada")),
        "estrategias_ajustes": _get_str(followup_payload.get("estrategias_ajustes")),
        "asistentes": asistentes,
    }
    return {
        "tipo_acta": "seguimiento",
        "fecha_servicio": fecha_servicio,
        "extra_name": f"Seguimiento {idx}",
        "acta_metadata": acta_metadata,
        "selected_sheet_names": included_sheet_names,
        "temp_parent_folder_id": temp_parent_folder_id,
    }
