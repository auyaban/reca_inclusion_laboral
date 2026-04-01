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


def get_template_id(config_key, env_key=None):
    """Read a Google Sheets template ID from env var or config.json.

    Parameters
    ----------
    config_key : str
        Key inside ``config.json`` (e.g. ``"google_sheets_master_template_id"``).
    env_key : str | None
        Optional environment-variable name.  When *None* the *config_key* is
        upper-cased and used as env-var name (e.g. ``GOOGLE_SHEETS_MASTER_TEMPLATE_ID``).
    """
    if env_key is None:
        env_key = config_key.upper()

    env_value = str(os.getenv(env_key) or "").strip()
    if env_value:
        return extract_spreadsheet_id(env_value)

    config = _load_config()
    config_value = str(config.get(config_key) or "").strip()
    if config_value:
        return extract_spreadsheet_id(config_value)

    raise RuntimeError(
        f"Falta {env_key} o config.json con {config_key}."
    )


def get_master_template_id():
    """Shortcut – returns the master Google Sheets template spreadsheet ID."""
    return get_template_id("google_sheets_master_template_id")


def get_evaluacion_accesibilidad_template_id():
    return get_template_id(
        "google_sheets_evaluacion_accesibilidad_template_id",
        "GOOGLE_SHEETS_EVALUACION_ACCESIBILIDAD_TEMPLATE_ID",
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


def get_sheet_titles(spreadsheet_id_or_url):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    service = get_google_sheets_service()
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets.properties.title",
    ).execute()
    titles = []
    for sheet in meta.get("sheets", []):
        title = str((sheet.get("properties") or {}).get("title") or "").strip()
        if title:
            titles.append(title)
    return titles


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


def batch_read_sheet_values(spreadsheet_id_or_url, ranges):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    cleaned = [str(item or "").strip() for item in (ranges or []) if str(item or "").strip()]
    if not cleaned:
        return {}
    service = get_google_sheets_service()
    response = (
        service.spreadsheets()
        .values()
        .batchGet(spreadsheetId=spreadsheet_id, ranges=cleaned)
        .execute()
    )
    result = {}
    for item in response.get("valueRanges", []):
        range_name = str(item.get("range") or "").strip()
        if not range_name:
            continue
        result[range_name] = list(item.get("values", []))
    return result


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


def unmerge_cells_in_area(spreadsheet_id_or_url, sheet_name, start_row, end_row, start_col=0, end_col=21):
    """Unmerge all merged cells overlapping the given area (0-indexed rows/cols)."""
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    service = get_google_sheets_service()
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title),merges)",
    ).execute()
    sheet_id = None
    for s in meta.get("sheets", []):
        if s.get("properties", {}).get("title") == sheet_name:
            sheet_id = s["properties"]["sheetId"]
            break
    if sheet_id is None:
        return
    requests = []
    for s in meta.get("sheets", []):
        if s.get("properties", {}).get("sheetId") != sheet_id:
            continue
        for m in s.get("merges", []):
            sr, er = m.get("startRowIndex", 0), m.get("endRowIndex", 0)
            sc, ec = m.get("startColumnIndex", 0), m.get("endColumnIndex", 0)
            if sr < end_row and er > start_row and sc < end_col and ec > start_col:
                requests.append({"unmergeCells": {"range": {
                    "sheetId": sheet_id,
                    "startRowIndex": sr, "endRowIndex": er,
                    "startColumnIndex": sc, "endColumnIndex": ec,
                }}})
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests},
        ).execute()


def remove_sheet_protection(spreadsheet_id_or_url):
    """Remove all protectedRanges from every sheet in the spreadsheet.

    When a template is copied via files().copy(), any sheet protection from
    the template is inherited by the copy.  Professionals who receive the copy
    are not the owner, so the inherited protection blocks cells that should be
    editable.  Calling this right after the copy is created strips all
    protected ranges so the copy is fully editable.
    """
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    service = get_google_sheets_service()
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(protectedRanges(protectedRangeId))",
    ).execute()

    requests = []
    for sheet in meta.get("sheets", []):
        for pr in sheet.get("protectedRanges", []):
            pr_id = pr.get("protectedRangeId")
            if pr_id is not None:
                requests.append({"deleteProtectedRange": {"protectedRangeId": pr_id}})

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests},
        ).execute()


def copy_sheet_to_spreadsheet(
    source_spreadsheet_id_or_url,
    source_sheet_name,
    destination_spreadsheet_id_or_url,
    *,
    new_sheet_name=None,
):
    source_spreadsheet_id = extract_spreadsheet_id(source_spreadsheet_id_or_url)
    destination_spreadsheet_id = extract_spreadsheet_id(destination_spreadsheet_id_or_url)
    service = get_google_sheets_service()

    meta = service.spreadsheets().get(
        spreadsheetId=source_spreadsheet_id,
        fields="sheets.properties(sheetId,title)",
    ).execute()
    source_sheet_id = None
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {}) or {}
        if str(props.get("title") or "").strip() == str(source_sheet_name or "").strip():
            source_sheet_id = props.get("sheetId")
            break
    if source_sheet_id is None:
        raise RuntimeError(f"No existe la hoja '{source_sheet_name}' en el archivo maestro.")

    copied = (
        service.spreadsheets()
        .sheets()
        .copyTo(
            spreadsheetId=source_spreadsheet_id,
            sheetId=source_sheet_id,
            body={"destinationSpreadsheetId": destination_spreadsheet_id},
        )
        .execute()
    )
    copied_sheet_id = copied.get("sheetId")
    copied_title = str(copied.get("title") or "").strip()
    target_title = str(new_sheet_name or copied_title or source_sheet_name or "").strip()

    if copied_sheet_id is not None and target_title and target_title != copied_title:
        service.spreadsheets().batchUpdate(
            spreadsheetId=destination_spreadsheet_id,
            body={
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": copied_sheet_id,
                                "title": target_title,
                            },
                            "fields": "title",
                        }
                    }
                ]
            },
        ).execute()
        copied_title = target_title

    return {
        "sheetId": copied_sheet_id,
        "title": copied_title or target_title,
    }


def insert_template_rows(
    spreadsheet_id_or_url,
    sheet_name,
    *,
    insert_at_row,
    template_row,
    count,
    paste_type="PASTE_NORMAL",
):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    total_rows = max(0, int(count or 0))
    if total_rows <= 0:
        return {"insertedRows": 0}

    service = get_google_sheets_service()
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets.properties(sheetId,title)",
    ).execute()

    sheet_id = None
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {}) or {}
        if str(props.get("title") or "").strip() == str(sheet_name or "").strip():
            sheet_id = props.get("sheetId")
            break
    if sheet_id is None:
        raise RuntimeError(f"No existe la hoja '{sheet_name}' en la spreadsheet destino.")

    insert_index = max(0, int(insert_at_row or 1) - 1)
    template_index = max(0, int(template_row or 1) - 1)

    requests = [
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": insert_index,
                    "endIndex": insert_index + total_rows,
                },
                "inheritFromBefore": insert_index > 0,
            }
        }
    ]
    if paste_type:
        requests.append(
            {
                "copyPaste": {
                    "source": {
                        "sheetId": sheet_id,
                        "startRowIndex": template_index,
                        "endRowIndex": template_index + 1,
                    },
                    "destination": {
                        "sheetId": sheet_id,
                        "startRowIndex": insert_index,
                        "endRowIndex": insert_index + total_rows,
                    },
                    "pasteType": str(paste_type),
                    "pasteOrientation": "NORMAL",
                }
            }
        )

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()
    return {"sheetId": sheet_id, "insertedRows": total_rows}


def insert_template_block_rows(
    spreadsheet_id_or_url,
    sheet_name,
    *,
    insert_at_row,
    template_start_row,
    template_end_row,
    repeat_count=1,
    paste_type="PASTE_NORMAL",
):
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    total_blocks = max(0, int(repeat_count or 0))
    if total_blocks <= 0:
        return {"insertedRows": 0, "insertedBlocks": 0}

    start_row = int(template_start_row or 0)
    end_row = int(template_end_row or 0)
    if start_row <= 0 or end_row < start_row:
        raise RuntimeError("template_start_row/template_end_row invalidos para insertar bloques.")

    block_height = (end_row - start_row) + 1
    total_rows = block_height * total_blocks

    service = get_google_sheets_service()
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets.properties(sheetId,title)",
    ).execute()

    sheet_id = None
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {}) or {}
        if str(props.get("title") or "").strip() == str(sheet_name or "").strip():
            sheet_id = props.get("sheetId")
            break
    if sheet_id is None:
        raise RuntimeError(f"No existe la hoja '{sheet_name}' en la spreadsheet destino.")

    insert_index = max(0, int(insert_at_row or 1) - 1)
    source_start_index = start_row - 1
    source_end_index = end_row

    requests = [
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": insert_index,
                    "endIndex": insert_index + total_rows,
                },
                "inheritFromBefore": insert_index > 0,
            }
        }
    ]

    if paste_type:
        for block_index in range(total_blocks):
            destination_start = insert_index + (block_index * block_height)
            requests.append(
                {
                    "copyPaste": {
                        "source": {
                            "sheetId": sheet_id,
                            "startRowIndex": source_start_index,
                            "endRowIndex": source_end_index,
                        },
                        "destination": {
                            "sheetId": sheet_id,
                            "startRowIndex": destination_start,
                            "endRowIndex": destination_start + block_height,
                        },
                        "pasteType": str(paste_type),
                        "pasteOrientation": "NORMAL",
                    }
                }
            )

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()
    return {
        "sheetId": sheet_id,
        "insertedRows": total_rows,
        "insertedBlocks": total_blocks,
        "blockHeight": block_height,
    }


def hide_sheets(spreadsheet_id_or_url, sheet_names_to_keep):
    """Hide all sheets in the spreadsheet except those in *sheet_names_to_keep*.

    At least one sheet must remain visible (Google Sheets requirement).
    Sheets whose title matches any entry in *sheet_names_to_keep* stay visible;
    every other sheet is hidden.

    Parameters
    ----------
    spreadsheet_id_or_url : str
        Spreadsheet ID or full URL.
    sheet_names_to_keep : list[str]
        Sheet titles that should remain visible.
    """
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    keep = {s.strip() for s in (sheet_names_to_keep or []) if s and s.strip()}
    if not keep:
        return

    service = get_google_sheets_service()
    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties"
    ).execute()

    unhide_requests = []
    hide_requests = []
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {})
        title = props.get("title", "")
        sheet_id = props.get("sheetId")
        is_hidden = props.get("hidden", False)
        should_keep = title in keep
        if should_keep and is_hidden:
            # Unhide sheets that should be visible (important for file reuse)
            unhide_requests.append({
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "hidden": False,
                    },
                    "fields": "hidden",
                }
            })
        elif not should_keep and not is_hidden:
            hide_requests.append({
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "hidden": True,
                    },
                    "fields": "hidden",
                }
            })
    requests = [*unhide_requests, *hide_requests]
    if not requests:
        return
    # Safety: never hide ALL sheets
    total_sheets = len(meta.get("sheets", []))
    if len(hide_requests) >= total_sheets:
        requests.remove(hide_requests[-1])

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()


def _col_letter_to_index(col_str):
    """Convert column letters (A, B, ..., Z, AA, ...) to 0-indexed column number."""
    result = 0
    for ch in col_str.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def _parse_a1_cell(a1_range):
    """Parse "'SheetName'!B5" into (sheet_name, row_0idx, col_0idx)."""
    import re
    m = re.match(r"^'([^']+)'!([A-Z]+)(\d+)$", a1_range)
    if not m:
        raise ValueError(f"Cannot parse A1 range: {a1_range}")
    sheet_name = m.group(1)
    col_idx = _col_letter_to_index(m.group(2))
    row_idx = int(m.group(3)) - 1
    return sheet_name, row_idx, col_idx


def set_native_checkboxes(spreadsheet_id_or_url, checkbox_cells):
    """Set native Google Sheets checkboxes on the given cells.

    Parameters
    ----------
    spreadsheet_id_or_url : str
        Spreadsheet ID or full URL.
    checkbox_cells : list[dict]
        Each dict: {"range": "'Sheet'!A1", "value": True/False}
    """
    if not checkbox_cells:
        return
    spreadsheet_id = extract_spreadsheet_id(spreadsheet_id_or_url)
    service = get_google_sheets_service()

    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties"
    ).execute()
    name_to_id = {}
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {})
        name_to_id[props.get("title", "")] = props.get("sheetId")

    requests = []
    for cell in checkbox_cells:
        sheet_name, row_idx, col_idx = _parse_a1_cell(cell["range"])
        sheet_id = name_to_id.get(sheet_name)
        if sheet_id is None:
            continue
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_idx,
                    "endRowIndex": row_idx + 1,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1,
                },
                "cell": {
                    "dataValidation": {
                        "condition": {"type": "BOOLEAN"},
                    },
                    "userEnteredValue": {
                        "boolValue": bool(cell.get("value", False)),
                    },
                },
                "fields": "dataValidation,userEnteredValue",
            }
        })

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests},
        ).execute()


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
