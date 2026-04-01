import json
import os
import re
import sys
from functools import lru_cache
from pathlib import Path
from openpyxl.utils.cell import range_boundaries
from formularios.common import _load_env_file


DEFAULT_CONFIG_PATH = "config.json"
SCOPE = "https://www.googleapis.com/auth/spreadsheets"
SPREADSHEET_ID_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")


def _get_bundle_dir():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def _load_config():
    bundle_path = os.path.join(_get_bundle_dir(), DEFAULT_CONFIG_PATH)
    cwd_path = DEFAULT_CONFIG_PATH
    for candidate in (bundle_path, cwd_path):
        if not os.path.exists(candidate):
            continue
        try:
            with open(candidate, "r", encoding="utf-8") as handle:
                return json.load(handle) or {}
        except (OSError, json.JSONDecodeError):
            return {}
    return {}


def _load_runtime_env():
    try:
        return _load_env_file(".env") or {}
    except Exception:
        return {}


def _get_credentials_path():
    runtime_env = _load_runtime_env()
    env_path = str(
        os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
        or os.getenv("GOOGLE_SHEETS_SA_JSON")
        or runtime_env.get("GOOGLE_SERVICE_ACCOUNT_FILE")
        or runtime_env.get("GOOGLE_SHEETS_SA_JSON")
        or ""
    ).strip()
    if env_path:
        path = env_path
    else:
        config = _load_config()
        path = (
            config.get("google_service_account_file")
            or config.get("google_sheets_sa_json")
            or config.get("google_drive_sa_json")
            or ""
        )
    if not path:
        raise RuntimeError(
            "Falta GOOGLE_SERVICE_ACCOUNT_FILE/GOOGLE_SHEETS_SA_JSON o config.json con "
            "google_service_account_file/google_sheets_sa_json/google_drive_sa_json."
        )
    if not os.path.isabs(path):
        path = os.path.join(_get_bundle_dir(), path)
    if not os.path.exists(path):
        raise RuntimeError(f"No existe el JSON de credenciales: {path}")
    return path


def extract_spreadsheet_id(value):
    text = str(value or "").strip()
    if not text:
        raise RuntimeError("Debe indicar spreadsheet_id o URL de Google Sheets.")

    match = SPREADSHEET_ID_RE.search(text)
    if match:
        return match.group(1)
    return text


def get_default_spreadsheet_id():
    runtime_env = _load_runtime_env()
    env_value = str(
        os.getenv("GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID")
        or runtime_env.get("GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID")
        or ""
    ).strip()
    if env_value:
        return extract_spreadsheet_id(env_value)

    config = _load_config()
    config_value = str(config.get("google_sheets_default_spreadsheet_id") or "").strip()
    if config_value:
        return extract_spreadsheet_id(config_value)

    raise RuntimeError(
        "Falta GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID o config.json con "
        "google_sheets_default_spreadsheet_id."
    )


def get_evaluacion_accesibilidad_template_id():
    runtime_env = _load_runtime_env()
    env_value = str(
        os.getenv("GOOGLE_SHEETS_EVALUACION_ACCESIBILIDAD_TEMPLATE_ID")
        or runtime_env.get("GOOGLE_SHEETS_EVALUACION_ACCESIBILIDAD_TEMPLATE_ID")
        or ""
    ).strip()
    if env_value:
        return extract_spreadsheet_id(env_value)

    config = _load_config()
    config_value = str(
        config.get("google_sheets_evaluacion_accesibilidad_template_id") or ""
    ).strip()
    if config_value:
        return extract_spreadsheet_id(config_value)

    raise RuntimeError(
        "Falta GOOGLE_SHEETS_EVALUACION_ACCESIBILIDAD_TEMPLATE_ID o config.json con "
        "google_sheets_evaluacion_accesibilidad_template_id."
    )


@lru_cache(maxsize=1)
def get_google_sheets_service():
    creds_path = _get_credentials_path()
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Faltan dependencias de Google Sheets. "
            "Instala google-api-python-client y google-auth."
        ) from exc

    credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def clear_google_sheets_service_cache():
    get_google_sheets_service.cache_clear()


def get_spreadsheet(spreadsheet_id_or_url, include_grid_data=False):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    service = get_google_sheets_service()
    return (
        service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, includeGridData=include_grid_data)
        .execute()
    )


def read_sheet_values(spreadsheet_id_or_url, range_name):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    service = get_google_sheets_service()
    response = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_name)
        .execute()
    )
    return list(response.get("values", []))


def write_sheet_values(
    spreadsheet_id_or_url,
    range_name,
    values,
    value_input_option="USER_ENTERED",
):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    service = get_google_sheets_service()
    body = {"values": values}
    return (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption=value_input_option,
            body=body,
        )
        .execute()
    )


def clear_sheet_ranges(spreadsheet_id_or_url, ranges):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    cleaned_ranges = [str(item or "").strip() for item in (ranges or []) if str(item or "").strip()]
    if not cleaned_ranges:
        return {"clearedRanges": []}
    service = get_google_sheets_service()
    return (
        service.spreadsheets()
        .values()
        .batchClear(
            spreadsheetId=spreadsheet_id,
            body={"ranges": cleaned_ranges},
        )
        .execute()
    )


def list_protected_ranges(spreadsheet_id_or_url):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    spreadsheet = get_spreadsheet(spreadsheet_id, include_grid_data=False)
    rows = []
    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties", {}) or {}
        sheet_id = props.get("sheetId")
        sheet_title = str(props.get("title") or "")
        for protected_range in sheet.get("protectedRanges", []) or []:
            if not isinstance(protected_range, dict):
                continue
            protected_range_id = protected_range.get("protectedRangeId")
            if protected_range_id is None:
                continue
            rows.append(
                {
                    "protectedRangeId": int(protected_range_id),
                    "sheetId": sheet_id,
                    "sheetTitle": sheet_title,
                    "warningOnly": bool(protected_range.get("warningOnly")),
                    "description": str(protected_range.get("description") or ""),
                }
            )
    return rows


def clear_protected_ranges(spreadsheet_id_or_url, protected_range_ids=None):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    if protected_range_ids is None:
        target_ids = [item["protectedRangeId"] for item in list_protected_ranges(spreadsheet_id)]
    else:
        target_ids = []
        for protected_range_id in protected_range_ids or []:
            if protected_range_id is None:
                continue
            target_ids.append(int(protected_range_id))
    if not target_ids:
        return {"deletedProtectedRangeIds": [], "deletedProtectedRangeCount": 0}

    # Preserve order while deduplicating to avoid invalid duplicate delete requests.
    unique_target_ids = list(dict.fromkeys(target_ids))
    requests = [
        {
            "deleteProtectedRange": {
                "protectedRangeId": protected_range_id,
            }
        }
        for protected_range_id in unique_target_ids
    ]
    service = get_google_sheets_service()
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()
    return {
        "deletedProtectedRangeIds": unique_target_ids,
        "deletedProtectedRangeCount": len(unique_target_ids),
    }


def _split_a1_range(range_name):
    text = str(range_name or "").strip()
    if not text:
        raise RuntimeError("Debe indicar un rango A1 de Google Sheets.")
    if "!" in text:
        sheet_name, cell_range = text.rsplit("!", 1)
        sheet_name = sheet_name.strip()
        if sheet_name.startswith("'") and sheet_name.endswith("'"):
            sheet_name = sheet_name[1:-1].replace("''", "'")
    else:
        sheet_name = ""
        cell_range = text
    return sheet_name, cell_range.replace("$", "").strip()


def set_sheet_ranges_bold(spreadsheet_id_or_url, ranges, *, bold):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    cleaned_ranges = [str(item or "").strip() for item in (ranges or []) if str(item or "").strip()]
    if not cleaned_ranges:
        return {"updatedRanges": [], "updatedRangeCount": 0}

    spreadsheet = get_spreadsheet(spreadsheet_id, include_grid_data=False)
    sheets = spreadsheet.get("sheets", []) or []
    default_sheet_id = None
    sheet_ids = {}
    for sheet in sheets:
        props = sheet.get("properties", {}) or {}
        sheet_id = props.get("sheetId")
        sheet_title = str(props.get("title") or "")
        if default_sheet_id is None:
            default_sheet_id = sheet_id
        if sheet_title:
            sheet_ids[sheet_title] = sheet_id

    requests = []
    applied_ranges = []
    for range_name in list(dict.fromkeys(cleaned_ranges)):
        sheet_name, cell_range = _split_a1_range(range_name)
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        sheet_id = sheet_ids.get(sheet_name, default_sheet_id)
        if sheet_id is None:
            raise RuntimeError(f"No se encontró la hoja para el rango: {range_name}")
        requests.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": min_row - 1,
                        "endRowIndex": max_row,
                        "startColumnIndex": min_col - 1,
                        "endColumnIndex": max_col,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "bold": bool(bold),
                            }
                        }
                    },
                    "fields": "userEnteredFormat.textFormat.bold",
                }
            }
        )
        applied_ranges.append(range_name)

    service = get_google_sheets_service()
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()
    return {
        "updatedRanges": applied_ranges,
        "updatedRangeCount": len(applied_ranges),
        "bold": bool(bold),
    }


def batch_write_sheet_updates(
    spreadsheet_id_or_url,
    updates,
    value_input_option="USER_ENTERED",
):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    rows = []
    for update in updates or []:
        if not isinstance(update, dict):
            continue
        range_name = str(update.get("range") or "").strip()
        if not range_name:
            continue
        value = update.get("value", "")
        if value is None:
            value = ""
        rows.append(
            {
                "range": range_name,
                "majorDimension": "ROWS",
                "values": [[value]],
            }
        )
    if not rows:
        return {"totalUpdatedCells": 0, "totalUpdatedRows": 0, "responses": []}

    service = get_google_sheets_service()
    return (
        service.spreadsheets()
        .values()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "valueInputOption": value_input_option,
                "data": rows,
            },
        )
        .execute()
    )


def export_spreadsheet_to_excel(spreadsheet_id_or_url, destination):
    spreadsheet = get_spreadsheet(spreadsheet_id_or_url, include_grid_data=False)
    destination_path = Path(destination).expanduser()
    destination_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise RuntimeError("openpyxl no esta instalado.") from exc

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    for sheet in spreadsheet.get("sheets", []):
        props = sheet.get("properties", {})
        title = str(props.get("title") or "Sheet")
        ws = wb.create_sheet(title=title[:31] or "Sheet")
        rows = read_sheet_values(spreadsheet["spreadsheetId"], f"'{title}'")
        for row_idx, row in enumerate(rows, start=1):
            for col_idx, value in enumerate(row, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        grid_props = props.get("gridProperties", {}) or {}
        frozen_rows = int(grid_props.get("frozenRowCount", 0) or 0)
        frozen_cols = int(grid_props.get("frozenColumnCount", 0) or 0)
        if frozen_rows > 0 or frozen_cols > 0:
            ws.freeze_panes = ws.cell(row=frozen_rows + 1, column=frozen_cols + 1)

    if not wb.worksheets:
        wb.create_sheet("Sheet")
    wb.save(destination_path)
    return destination_path
