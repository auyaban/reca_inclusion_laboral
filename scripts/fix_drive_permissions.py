"""fix_drive_permissions.py

Quita todos los permisos de la carpeta "2026" (y su contenido recursivo)
dentro del Shared Drive, conservando sólo las cuentas de la allow-list.
También activa copyRequiresWriterPermission=True en todos los archivos.

Uso:
    python scripts/fix_drive_permissions.py              # DRY-RUN (sin cambios)
    python scripts/fix_drive_permissions.py --execute    # aplica los cambios
"""

import argparse
import json
import os
import random
import socket
import sys
import time
from datetime import datetime
from urllib.error import URLError

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

DEFAULT_TARGET_FOLDER_NAME = "2026"
SCOPE = "https://www.googleapis.com/auth/drive"
GOOGLE_FOLDER_MIME = "application/vnd.google-apps.folder"

RETRYABLE_CODES = {429, 500, 502, 503, 504}
NON_RETRYABLE_CODES = {400, 401, 403, 404}


# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

def _log(message: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {message}", flush=True)


# ---------------------------------------------------------------------------
# Credenciales y servicio
# ---------------------------------------------------------------------------

def _get_script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _load_config() -> dict:
    candidates = [
        os.path.join(os.getenv("APPDATA") or "", "RECA Inclusion Laboral", "config.json"),
        os.path.join(_get_script_dir(), "..", "config.json"),
        os.path.join(os.getcwd(), "config.json"),
    ]
    for path in candidates:
        path = os.path.normpath(path)
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    return {}


def _get_config_or_env(config_key: str, env_key: str, default=""):
    config = _load_config()
    value = os.getenv(env_key)
    if value is None:
        value = config.get(config_key, default)
    return value


def _get_shared_drive_id() -> str:
    value = str(
        _get_config_or_env(
            "fix_drive_permissions_shared_drive_id",
            "FIX_DRIVE_PERMISSIONS_SHARED_DRIVE_ID",
        )
        or ""
    ).strip()
    if not value:
        raise RuntimeError("Falta FIX_DRIVE_PERMISSIONS_SHARED_DRIVE_ID o config.json con fix_drive_permissions_shared_drive_id.")
    return value


def _get_target_folder_name() -> str:
    value = str(
        _get_config_or_env(
            "fix_drive_permissions_target_folder_name",
            "FIX_DRIVE_PERMISSIONS_TARGET_FOLDER_NAME",
            DEFAULT_TARGET_FOLDER_NAME,
        )
        or ""
    ).strip()
    if not value:
        raise RuntimeError("Falta FIX_DRIVE_PERMISSIONS_TARGET_FOLDER_NAME o config.json con fix_drive_permissions_target_folder_name.")
    return value


def _get_service_account_email() -> str:
    return str(
        _get_config_or_env(
            "fix_drive_permissions_service_account_email",
            "FIX_DRIVE_PERMISSIONS_SERVICE_ACCOUNT_EMAIL",
        )
        or ""
    ).strip().lower()


def _get_allowed_emails() -> set[str]:
    raw = _get_config_or_env(
        "fix_drive_permissions_allowed_emails",
        "FIX_DRIVE_PERMISSIONS_ALLOWED_EMAILS",
        [],
    )
    emails = set()
    if isinstance(raw, str):
        for token in raw.replace(";", ",").split(","):
            value = str(token or "").strip().lower()
            if value:
                emails.add(value)
    elif isinstance(raw, (list, tuple, set)):
        for token in raw:
            value = str(token or "").strip().lower()
            if value:
                emails.add(value)
    service_account_email = _get_service_account_email()
    if service_account_email:
        emails.add(service_account_email)
    if not emails:
        raise RuntimeError(
            "Falta FIX_DRIVE_PERMISSIONS_ALLOWED_EMAILS o config.json con fix_drive_permissions_allowed_emails."
        )
    return emails


def _get_credentials_path() -> str:
    raw = (
        os.getenv("GOOGLE_DRIVE_SA_JSON")
        or os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
    )
    if not raw:
        raise RuntimeError(
            "No se encontró la ruta al JSON de la service account. "
            "Configure GOOGLE_SERVICE_ACCOUNT_FILE en %APPDATA%\\RECA Inclusion Laboral\\.env "
            "o exporte GOOGLE_DRIVE_SA_JSON."
        )
    if not os.path.isabs(raw):
        project_root = os.path.normpath(os.path.join(_get_script_dir(), ".."))
        raw = os.path.join(project_root, raw)
    raw = os.path.normpath(raw)
    if not os.path.isfile(raw):
        raise RuntimeError("El archivo de credenciales configurado no es vÃ¡lido.")
    return raw


def _build_drive_service(creds_path: str):
    credentials = Credentials.from_service_account_file(creds_path, scopes=[SCOPE])
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


# ---------------------------------------------------------------------------
# Retry
# ---------------------------------------------------------------------------

def _get_status_code(exc) -> int | None:
    resp = getattr(exc, "resp", None)
    for src in [resp, getattr(resp, "headers", None)]:
        for attr in ("status", "status_code"):
            val = getattr(src, attr, None) if src and not isinstance(src, dict) else (src or {}).get(attr)
            try:
                if val is not None:
                    return int(val)
            except (TypeError, ValueError):
                pass
    return None


def _is_retryable(exc) -> bool:
    code = _get_status_code(exc)
    if code is not None:
        if code in NON_RETRYABLE_CODES:
            return False
        if code in RETRYABLE_CODES or code >= 500:
            return True
    if isinstance(exc, (TimeoutError, socket.timeout, ConnectionError, URLError, OSError)):
        return True
    reason = getattr(exc, "reason", None)
    if isinstance(reason, (TimeoutError, socket.timeout, ConnectionError, URLError, OSError)):
        return True
    return False


def _retry_delay(exc, attempt: int) -> float:
    resp = getattr(exc, "resp", None)
    for src in [resp, getattr(resp, "headers", None)]:
        if src and hasattr(src, "get"):
            raw = src.get("retry-after") or src.get("Retry-After")
            if raw:
                try:
                    return max(0.0, float(raw))
                except (TypeError, ValueError):
                    pass
    delay = min(16.0, 1.0 * (2 ** (attempt - 1)))
    delay *= 1.0 + random.random() * 0.25
    return delay


def _execute_with_retry(request, operation: str = "api_call", max_attempts: int = 5):
    for attempt in range(1, max_attempts + 1):
        try:
            return request.execute()
        except Exception as exc:
            if attempt >= max_attempts or not _is_retryable(exc):
                raise
            delay = _retry_delay(exc, attempt)
            _log(f"  [RETRY] {operation} intento {attempt}/{max_attempts} espera {delay:.1f}s — {exc}")
            time.sleep(delay)


# ---------------------------------------------------------------------------
# Búsqueda de la carpeta objetivo
# ---------------------------------------------------------------------------

def _find_target_folder(service, shared_drive_id: str, folder_name: str) -> dict:
    _log(f"Buscando carpeta '{folder_name}' en Shared Drive {shared_drive_id}...")
    q = (
        f"name='{folder_name}' "
        f"and mimeType='{GOOGLE_FOLDER_MIME}' "
        f"and '{shared_drive_id}' in parents "
        f"and trashed=false"
    )
    result = _execute_with_retry(
        service.files().list(
            q=q,
            fields="files(id,name,mimeType)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            corpora="drive",
            driveId=shared_drive_id,
            pageSize=10,
        ),
        operation="find_target_folder",
    )
    files = result.get("files", [])
    if not files:
        raise RuntimeError(
            f"No se encontró la carpeta '{folder_name}' directamente bajo el Shared Drive {shared_drive_id}."
        )
    folder = files[0]
    _log(f"Carpeta encontrada: id={folder['id']} name={folder['name']}")
    return folder


# ---------------------------------------------------------------------------
# Traversal recursivo
# ---------------------------------------------------------------------------

def _list_children(service, parent_id: str, shared_drive_id: str) -> list:
    items = []
    page_token = None
    while True:
        kwargs = dict(
            q=f"'{parent_id}' in parents and trashed=false",
            fields="nextPageToken,files(id,name,mimeType)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            corpora="drive",
            driveId=shared_drive_id,
            pageSize=200,
        )
        if page_token:
            kwargs["pageToken"] = page_token
        result = _execute_with_retry(
            service.files().list(**kwargs),
            operation=f"list_children:{parent_id}",
        )
        items.extend(result.get("files", []))
        page_token = result.get("nextPageToken")
        if not page_token:
            break
    return items


def _collect_all_items(service, root_folder: dict, shared_drive_id: str) -> list:
    _log("Recolectando todos los ítems recursivamente...")
    all_items = [root_folder]
    stack = [root_folder]
    while stack:
        current = stack.pop()
        if current["mimeType"] == GOOGLE_FOLDER_MIME:
            children = _list_children(service, current["id"], shared_drive_id)
            all_items.extend(children)
            stack.extend(children)
    _log(f"Total ítems a procesar: {len(all_items)}")
    return all_items


# ---------------------------------------------------------------------------
# Permisos
# ---------------------------------------------------------------------------

def _list_all_permissions(service, file_id: str) -> list:
    perms = []
    page_token = None
    while True:
        kwargs = dict(
            fileId=file_id,
            fields="nextPageToken,permissions(id,type,emailAddress,role,displayName,domain)",
            supportsAllDrives=True,
            pageSize=100,
        )
        if page_token:
            kwargs["pageToken"] = page_token
        result = _execute_with_retry(
            service.permissions().list(**kwargs),
            operation=f"list_permissions:{file_id}",
        )
        perms.extend(result.get("permissions", []))
        page_token = result.get("nextPageToken")
        if not page_token:
            break
    return perms


def _should_keep_permission(perm: dict) -> bool:
    allowed_emails = _get_allowed_emails()
    ptype = perm.get("type", "")
    if ptype in ("anyone", "domain", "group"):
        return False
    if ptype == "user":
        email = (perm.get("emailAddress") or "").lower()
        if not email:
            # Sin email: conservar si es owner para no romper nada
            if perm.get("role") == "owner":
                return True
            return False
        return email in allowed_emails
    return False


def _delete_permission(service, file_id: str, perm_id: str, dry_run: bool, stats: dict) -> None:
    if dry_run:
        _log(f"    [DRY-RUN] BORRARÍA permiso id={perm_id}")
        stats["permissions_deleted"] += 1
        return
    try:
        _execute_with_retry(
            service.permissions().delete(
                fileId=file_id,
                permissionId=perm_id,
                supportsAllDrives=True,
            ),
            operation=f"delete_permission:{file_id}:{perm_id}",
        )
        _log(f"    [EXECUTE] Permiso eliminado id={perm_id}")
        stats["permissions_deleted"] += 1
    except HttpError as exc:
        if _get_status_code(exc) == 404:
            _log(f"    [SKIP] Permiso ya no existe id={perm_id}")
            stats["permissions_deleted"] += 1
        else:
            _log(f"    [ERROR] No se pudo eliminar permiso id={perm_id}: {exc}")
            stats["errors"] += 1
    except Exception as exc:
        _log(f"    [ERROR] No se pudo eliminar permiso id={perm_id}: {exc}")
        stats["errors"] += 1


def _set_copy_restriction(service, file_id: str, dry_run: bool, stats: dict) -> None:
    if dry_run:
        _log(f"    [DRY-RUN] ACTIVARÍA copyRequiresWriterPermission=True")
        stats["copy_restrictions_set"] += 1
        return
    try:
        _execute_with_retry(
            service.files().update(
                fileId=file_id,
                body={"copyRequiresWriterPermission": True},
                fields="id,copyRequiresWriterPermission",
                supportsAllDrives=True,
            ),
            operation=f"set_copy_restriction:{file_id}",
        )
        _log(f"    [EXECUTE] copyRequiresWriterPermission=True activado")
        stats["copy_restrictions_set"] += 1
    except Exception as exc:
        _log(f"    [ERROR] No se pudo activar restricción de copia: {exc}")
        stats["errors"] += 1


# ---------------------------------------------------------------------------
# Procesamiento por ítem
# ---------------------------------------------------------------------------

def _process_item(service, item: dict, dry_run: bool, stats: dict) -> None:
    is_folder = item["mimeType"] == GOOGLE_FOLDER_MIME
    kind = "CARPETA" if is_folder else "ARCHIVO"
    _log(f"\n--- {kind}: {item['name']}  id={item['id']}")

    try:
        perms = _list_all_permissions(service, item["id"])
    except Exception as exc:
        _log(f"  [ERROR] No se pudieron listar permisos: {exc}")
        stats["errors"] += 1
        stats["items_processed"] += 1
        return

    for perm in perms:
        ptype = perm.get("type", "?")
        email = perm.get("emailAddress") or perm.get("domain") or f"type={ptype}"
        role = perm.get("role", "?")
        perm_id = perm.get("id", "?")

        if _should_keep_permission(perm):
            _log(f"  CONSERVAR : {email:<40} role={role}")
            stats["permissions_kept"] += 1
        else:
            _log(f"  ELIMINAR  : {email:<40} role={role}  id={perm_id}")
            _delete_permission(service, item["id"], perm_id, dry_run, stats)

    if not is_folder:
        _set_copy_restriction(service, item["id"], dry_run, stats)

    stats["items_processed"] += 1


# ---------------------------------------------------------------------------
# Resumen
# ---------------------------------------------------------------------------

def _print_summary(stats: dict, dry_run: bool) -> None:
    mode = "DRY-RUN" if dry_run else "EJECUTADO"
    print("\n" + "=" * 50)
    print(f"RESUMEN ({mode})")
    print("=" * 50)
    print(f"  Ítems procesados         : {stats['items_processed']}")
    print(f"  Permisos conservados     : {stats['permissions_kept']}")
    print(f"  Permisos eliminados      : {stats['permissions_deleted']}")
    print(f"  Restricciones de copia   : {stats['copy_restrictions_set']}")
    print(f"  Errores                  : {stats['errors']}")
    print("=" * 50)
    if dry_run:
        print("\nEsto fue un DRY-RUN. Para aplicar cambios use: --execute")


# ---------------------------------------------------------------------------
# Entrada principal
# ---------------------------------------------------------------------------

def main() -> None:
    default_shared_drive_id = _get_shared_drive_id()
    default_folder_name = _get_target_folder_name()
    allowed_emails = _get_allowed_emails()
    parser = argparse.ArgumentParser(
        description="Limpia permisos en la carpeta 2026 del Shared Drive."
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        default=False,
        help="Aplica los cambios reales. Sin este flag corre en modo DRY-RUN.",
    )
    parser.add_argument(
        "--shared-drive-id",
        default=default_shared_drive_id,
        help=f"ID del Shared Drive (default: {default_shared_drive_id})",
    )
    parser.add_argument(
        "--folder-name",
        default=default_folder_name,
        help=f"Nombre de la carpeta objetivo (default: {default_folder_name})",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.05,
        help="Pausa en segundos entre ítems para respetar rate limits (default: 0.05)",
    )
    args = parser.parse_args()

    dry_run = not args.execute
    mode_label = "DRY-RUN (sin cambios)" if dry_run else "EXECUTE (se aplicarán cambios reales)"
    print(f"\n{'=' * 50}")
    print(f"MODO: {mode_label}")
    print(f"Shared Drive : {args.shared_drive_id}")
    print(f"Carpeta      : {args.folder_name}")
    print(f"Allow-list   : {', '.join(sorted(allowed_emails))}")
    print(f"{'=' * 50}\n")

    try:
        creds_path = _get_credentials_path()
        _log(f"Credenciales: {creds_path}")
        service = _build_drive_service(creds_path)
        _log("Servicio Drive v3 construido correctamente.")
    except Exception as exc:
        _log(f"[FATAL] Error al inicializar: {exc}")
        sys.exit(1)

    try:
        root_folder = _find_target_folder(service, args.shared_drive_id, args.folder_name)
    except RuntimeError as exc:
        _log(f"[FATAL] {exc}")
        sys.exit(1)

    try:
        all_items = _collect_all_items(service, root_folder, args.shared_drive_id)
    except Exception as exc:
        _log(f"[FATAL] Error al recolectar ítems: {exc}")
        sys.exit(1)

    stats = {
        "items_processed": 0,
        "permissions_kept": 0,
        "permissions_deleted": 0,
        "copy_restrictions_set": 0,
        "errors": 0,
    }

    for item in all_items:
        _process_item(service, item, dry_run, stats)
        if args.delay > 0:
            time.sleep(args.delay)

    _print_summary(stats, dry_run)


if __name__ == "__main__":
    main()
