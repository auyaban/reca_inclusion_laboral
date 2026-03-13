import json
import os
import re
import sys
from functools import lru_cache
from pathlib import Path


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


def _get_credentials_path():
    env_path = str(
        os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
        or os.getenv("GOOGLE_SHEETS_SA_JSON")
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
            "Falta GOOGLE_SERVICE_ACCOUNT_FILE o config.json con "
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
    env_value = str(os.getenv("GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID") or "").strip()
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
    env_value = str(
        os.getenv("GOOGLE_SHEETS_EVALUACION_ACCESIBILIDAD_TEMPLATE_ID") or ""
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
