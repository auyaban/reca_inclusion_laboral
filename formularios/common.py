import os
import re
import time
import unicodedata
import json
import sys
import ctypes
import threading
import uuid
import sqlite3
import hashlib
import base64
import urllib.parse
import urllib.request
import urllib.error

_SUPABASE_SESSION_LOCK = threading.Lock()
_SUPABASE_SESSION = {
    "access_token": "",
    "refresh_token": "",
    "expires_at": 0.0,
}


def _resolve_env_candidates(env_path=".env"):
    if os.path.isabs(env_path):
        return [env_path]
    candidates = []
    # 1) executable/script directory (installed app priority)
    try:
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        candidates.append(os.path.join(exe_dir, env_path))
    except Exception:
        pass
    # 2) roaming appdata fallback
    appdata = os.getenv("APPDATA")
    if appdata:
        candidates.append(os.path.join(appdata, "RECA Inclusion Laboral", env_path))
    # 3) project root (when running from source)
    try:
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        candidates.append(os.path.join(project_root, env_path))
    except Exception:
        pass
    # 4) current working directory (last resort)
    candidates.append(os.path.abspath(env_path))
    # preserve order and uniqueness
    uniq = []
    seen = set()
    for path in candidates:
        key = os.path.normcase(path)
        if key in seen:
            continue
        seen.add(key)
        uniq.append(path)
    return uniq


def _load_env_file(env_path=".env"):
    chosen = None
    for candidate in _resolve_env_candidates(env_path):
        if os.path.exists(candidate):
            chosen = candidate
            break
    if not chosen:
        return {}
    env = {}
    with open(chosen, "r", encoding="utf-8") as handle:
        for raw in handle:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            clean_key = key.strip().lstrip("\ufeff")
            env[clean_key] = value.strip().strip('"').strip("'")
    return env


def _load_supabase_credentials(env_path=".env"):
    checked = []
    for candidate in _resolve_env_candidates(env_path):
        if not os.path.exists(candidate):
            continue
        checked.append(candidate)
        env = _load_env_file(candidate)
        supabase_url = (env.get("SUPABASE_URL") or "").strip()
        supabase_key = (env.get("SUPABASE_KEY") or "").strip()
        if supabase_url and supabase_key:
            return supabase_url, supabase_key
    if checked:
        joined = " | ".join(checked)
        raise RuntimeError(
            f"Missing SUPABASE_URL or SUPABASE_KEY. Revisa .env en: {joined}"
        )
    raise RuntimeError("Missing SUPABASE_URL or SUPABASE_KEY")


def _supabase_headers(api_key, bearer_token=None):
    token = (bearer_token or "").strip() or api_key
    return {
        "apikey": api_key,
        "Authorization": f"Bearer {token}",
    }


def _coerce_expires_at(expires_at=None, expires_in=None):
    now = time.time()
    try:
        if expires_at is not None:
            return float(expires_at)
    except Exception:
        pass
    try:
        if expires_in is not None:
            return now + max(0.0, float(expires_in))
    except Exception:
        pass
    return now + 3600.0


def _set_supabase_session(access_token, refresh_token=None, expires_at=None, expires_in=None):
    with _SUPABASE_SESSION_LOCK:
        _SUPABASE_SESSION["access_token"] = str(access_token or "").strip()
        _SUPABASE_SESSION["refresh_token"] = str(refresh_token or "").strip()
        _SUPABASE_SESSION["expires_at"] = _coerce_expires_at(
            expires_at=expires_at,
            expires_in=expires_in,
        )


def _clear_supabase_session():
    with _SUPABASE_SESSION_LOCK:
        _SUPABASE_SESSION["access_token"] = ""
        _SUPABASE_SESSION["refresh_token"] = ""
        _SUPABASE_SESSION["expires_at"] = 0.0


def _get_supabase_session():
    with _SUPABASE_SESSION_LOCK:
        return dict(_SUPABASE_SESSION)


def _extract_error_message(exc):
    if isinstance(exc, urllib.error.HTTPError):
        try:
            body = exc.read().decode("utf-8", errors="replace")
            payload = json.loads(body) if body else {}
            if isinstance(payload, dict):
                for key in ("msg", "message", "error_description", "error"):
                    value = payload.get(key)
                    if value:
                        return str(value)
        except Exception:
            pass
    return str(exc)


def _decode_jwt_payload(token):
    raw = str(token or "").strip()
    if not raw:
        return {}
    parts = raw.split(".")
    if len(parts) < 2:
        return {}
    segment = parts[1]
    padding = "=" * ((4 - len(segment) % 4) % 4)
    try:
        decoded = base64.urlsafe_b64decode((segment + padding).encode("utf-8"))
        payload = json.loads(decoded.decode("utf-8", errors="replace"))
        return payload if isinstance(payload, dict) else {}
    except Exception:
        return {}


def _get_cache_scope(token):
    payload = _decode_jwt_payload(token)
    uid = str(payload.get("sub") or "").strip()
    role = str(payload.get("role") or "").strip()
    if uid:
        return f"{role or 'authenticated'}:{uid}"
    if role:
        return role
    return "anon"


def _supabase_refresh_session(env_path=".env"):
    session = _get_supabase_session()
    refresh_token = (session.get("refresh_token") or "").strip()
    if not refresh_token:
        return False
    supabase_url, supabase_key = _load_supabase_credentials(env_path)
    url = f"{supabase_url.rstrip('/')}/auth/v1/token?grant_type=refresh_token"
    payload = {"refresh_token": refresh_token}
    body = json.dumps(payload).encode("utf-8")
    headers = _supabase_headers(supabase_key)
    headers["Content-Type"] = "application/json"
    request = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            raw = response.read().decode("utf-8")
        data = json.loads(raw) if raw else {}
    except Exception:
        _clear_supabase_session()
        return False
    access_token = (data.get("access_token") or "").strip()
    if not access_token:
        return False
    _set_supabase_session(
        access_token=access_token,
        refresh_token=data.get("refresh_token") or refresh_token,
        expires_at=data.get("expires_at"),
        expires_in=data.get("expires_in"),
    )
    return True


def _supabase_get_access_token(env_path=".env"):
    session = _get_supabase_session()
    access_token = (session.get("access_token") or "").strip()
    expires_at = float(session.get("expires_at") or 0.0)
    if not access_token:
        return ""
    if expires_at <= (time.time() + 60.0):
        if _supabase_refresh_session(env_path=env_path):
            refreshed = _get_supabase_session()
            return (refreshed.get("access_token") or "").strip()
        return ""
    return access_token


def _supabase_auth_password_login(email, password, env_path=".env"):
    supabase_url, supabase_key = _load_supabase_credentials(env_path)
    url = f"{supabase_url.rstrip('/')}/auth/v1/token?grant_type=password"
    payload = {"email": str(email or "").strip(), "password": str(password or "")}
    body = json.dumps(payload).encode("utf-8")
    headers = _supabase_headers(supabase_key)
    headers["Content-Type"] = "application/json"
    request = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=45) as response:
            raw = response.read().decode("utf-8")
        data = json.loads(raw) if raw else {}
    except urllib.error.HTTPError as exc:
        code = int(getattr(exc, "code", 0) or 0)
        if code in {400, 401}:
            raise RuntimeError("Usuario y contraseña incorrectos.") from exc
        raise RuntimeError(_format_supabase_error("No se pudo autenticar con Supabase", exc)) from exc
    except Exception as exc:
        raise RuntimeError(_format_supabase_error("No se pudo autenticar con Supabase", exc)) from exc

    access_token = (data.get("access_token") or "").strip()
    if not access_token:
        raise RuntimeError("No se recibió access token de Supabase Auth.")
    _set_supabase_session(
        access_token=access_token,
        refresh_token=data.get("refresh_token") or "",
        expires_at=data.get("expires_at"),
        expires_in=data.get("expires_in"),
    )
    return data


def _supabase_auth_update_password(new_password, env_path=".env"):
    supabase_url, supabase_key = _load_supabase_credentials(env_path)
    token = _supabase_get_access_token(env_path=env_path)
    if not token:
        raise RuntimeError("Sesion no valida para actualizar contraseña.")
    url = f"{supabase_url.rstrip('/')}/auth/v1/user"
    payload = {"password": str(new_password or "")}
    body = json.dumps(payload).encode("utf-8")
    attempted_refresh = False
    last_exc = None
    for _ in range(2):
        headers = _supabase_headers(supabase_key, bearer_token=token)
        headers["Content-Type"] = "application/json"
        request = urllib.request.Request(url, data=body, headers=headers, method="PUT")
        try:
            with urllib.request.urlopen(request, timeout=45) as response:
                raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as exc:
            last_exc = exc
            if int(getattr(exc, "code", 0) or 0) == 401 and not attempted_refresh:
                attempted_refresh = True
                if _supabase_refresh_session(env_path=env_path):
                    token = _supabase_get_access_token(env_path=env_path)
                    continue
            break
        except Exception as exc:
            last_exc = exc
            break
    raise RuntimeError(_format_supabase_error("No se pudo actualizar la contraseña en Auth", last_exc)) from last_exc


def _supabase_rpc(function_name, params=None, env_path=".env", use_session=True):
    supabase_url, supabase_key = _load_supabase_credentials(env_path)
    url = f"{supabase_url.rstrip('/')}/rest/v1/rpc/{function_name}"
    body = json.dumps(params or {}).encode("utf-8")
    attempted_refresh = False
    last_exc = None
    for _ in range(3):
        token = _supabase_get_access_token(env_path=env_path) if use_session else ""
        headers = _supabase_headers(supabase_key, bearer_token=token)
        headers["Content-Type"] = "application/json"
        request = urllib.request.Request(url, data=body, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else None
        except urllib.error.HTTPError as exc:
            last_exc = exc
            if (
                use_session
                and int(getattr(exc, "code", 0) or 0) == 401
                and not attempted_refresh
                and _supabase_refresh_session(env_path=env_path)
            ):
                attempted_refresh = True
                continue
            break
        except Exception as exc:
            last_exc = exc
            break
    raise RuntimeError(_format_supabase_error(f"No se pudo ejecutar RPC {function_name}", last_exc)) from last_exc


def _supabase_get(table, params, env_path=".env"):
    supabase_url, supabase_key = _load_supabase_credentials(env_path)
    query = urllib.parse.urlencode(params)
    url = f"{supabase_url.rstrip('/')}/rest/v1/{table}?{query}"
    last_error = None
    attempted_refresh = False
    for _ in range(3):
        token = _supabase_get_access_token(env_path=env_path)
        cache_scope = _get_cache_scope(token)
        request = urllib.request.Request(
            url,
            headers=_supabase_headers(supabase_key, bearer_token=token),
        )
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = response.read().decode("utf-8")
            data = json.loads(payload)
            try:
                if _can_cache_supabase_response(table, params):
                    _cache_supabase_get_response(
                        table,
                        params,
                        _sanitize_payload_for_cache(data),
                        scope=cache_scope,
                    )
            except Exception:
                pass
            return data
        except urllib.error.HTTPError as exc:
            last_error = exc
            if (
                int(getattr(exc, "code", 0) or 0) == 401
                and not attempted_refresh
                and _supabase_refresh_session(env_path=env_path)
            ):
                attempted_refresh = True
                continue
        except Exception as exc:
            last_error = exc
    try:
        cached = _load_supabase_get_cached_response(
            table,
            params,
            scope=_get_cache_scope(_supabase_get_access_token(env_path=env_path)),
        )
    except Exception:
        cached = None
    if cached is not None:
        return cached
    raise RuntimeError(_format_supabase_error("Supabase no esta disponible", last_error)) from last_error


def _supabase_get_paged(table, params=None, env_path=".env", page_size=1000, max_pages=200):
    """
    Obtiene registros de forma paginada usando limit/offset.
    """
    base = dict(params or {})
    try:
        page_size_int = max(1, int(page_size))
    except Exception:
        page_size_int = 1000
    try:
        max_pages_int = max(1, int(max_pages))
    except Exception:
        max_pages_int = 200

    offset = 0
    all_rows = []
    for _ in range(max_pages_int):
        query = dict(base)
        query["limit"] = page_size_int
        query["offset"] = offset
        rows = _supabase_get(table, query, env_path=env_path)
        if not isinstance(rows, list):
            break
        all_rows.extend(rows)
        if len(rows) < page_size_int:
            break
        offset += page_size_int
    return all_rows


def _format_supabase_error(prefix, exc):
    detail = ""
    if isinstance(exc, urllib.error.HTTPError):
        try:
            body = exc.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        detail = body.strip()
        code = getattr(exc, "code", None)
        if code:
            prefix = f"{prefix} (HTTP {code})"
    elif exc:
        detail = str(exc)
    return f"{prefix}: {detail}" if detail else prefix


_WRITE_QUEUE_LOCK = threading.Lock()
_WRITE_QUEUE = []
_WRITE_WORKER_STARTED = False
_FAILED_WRITE_QUEUE = []
_SENSITIVE_CACHE_KEYS = {
    "usuario_pass",
    "usuario_pass_hash",
    "password",
    "pass",
    "passwd",
    "token",
    "access_token",
    "refresh_token",
    "api_key",
    "apikey",
    "authorization",
}
_OFFLINE_DB_LOCK = threading.Lock()
_OFFLINE_DB_READY = False


def _get_cache_dir():
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        base = os.path.join(local_app_data, "RECA", "cache")
    else:
        base = os.path.join(os.getcwd(), ".cache")
    os.makedirs(base, exist_ok=True)
    return base


def _get_supabase_queue_path():
    return os.path.join(_get_cache_dir(), "supabase_write_queue.json")


def _get_supabase_failed_queue_path():
    return os.path.join(_get_cache_dir(), "supabase_write_failed.json")


def _get_offline_db_path():
    return os.path.join(_get_cache_dir(), "offline_store.db")


def _offline_connect():
    return sqlite3.connect(_get_offline_db_path(), timeout=15)


def _ensure_offline_db():
    global _OFFLINE_DB_READY
    if _OFFLINE_DB_READY:
        return
    with _OFFLINE_DB_LOCK:
        if _OFFLINE_DB_READY:
            return
        conn = _offline_connect()
        try:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS supabase_get_cache (
                    table_name TEXT NOT NULL,
                    query_hash TEXT NOT NULL,
                    query_json TEXT NOT NULL,
                    payload_json TEXT NOT NULL,
                    updated_at REAL NOT NULL,
                    PRIMARY KEY (table_name, query_hash)
                )
                """
            )
            conn.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_supabase_get_cache_table_updated
                ON supabase_get_cache (table_name, updated_at DESC)
                """
            )
            conn.commit()
            _OFFLINE_DB_READY = True
        finally:
            conn.close()


def _serialize_query_for_cache(params, scope=""):
    clean = {}
    for key in sorted((params or {}).keys()):
        value = (params or {}).get(key)
        if isinstance(value, (list, tuple)):
            clean[str(key)] = [str(v) for v in value]
        elif value is None:
            clean[str(key)] = ""
        else:
            clean[str(key)] = str(value)
    clean["__scope"] = str(scope or "")
    query_json = json.dumps(clean, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    query_hash = hashlib.sha256(query_json.encode("utf-8")).hexdigest()
    return query_hash, query_json


def _can_cache_supabase_response(table, params):
    table_name = str(table or "").strip().lower()
    select = str((params or {}).get("select") or "").lower()
    if table_name == "profesionales" and (
        "usuario_pass" in select or "usuario_pass_hash" in select
    ):
        return False
    return True


def _sanitize_payload_for_cache(payload):
    if isinstance(payload, list):
        return [_sanitize_payload_for_cache(item) for item in payload]
    if isinstance(payload, dict):
        clean = {}
        for key, value in payload.items():
            key_str = str(key or "").strip().lower()
            if key_str in _SENSITIVE_CACHE_KEYS:
                clean[key] = None
            else:
                clean[key] = _sanitize_payload_for_cache(value)
        return clean
    return payload


def _cache_supabase_get_response(table, params, payload, scope=""):
    _ensure_offline_db()
    query_hash, query_json = _serialize_query_for_cache(params, scope=scope)
    payload_json = json.dumps(payload, ensure_ascii=False)
    now = time.time()
    with _OFFLINE_DB_LOCK:
        conn = _offline_connect()
        try:
            conn.execute(
                """
                INSERT INTO supabase_get_cache (table_name, query_hash, query_json, payload_json, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(table_name, query_hash) DO UPDATE SET
                    query_json=excluded.query_json,
                    payload_json=excluded.payload_json,
                    updated_at=excluded.updated_at
                """,
                (str(table), query_hash, query_json, payload_json, now),
            )
            conn.commit()
        finally:
            conn.close()


def _load_supabase_get_cached_response(table, params, scope=""):
    _ensure_offline_db()
    query_hash, _ = _serialize_query_for_cache(params, scope=scope)
    with _OFFLINE_DB_LOCK:
        conn = _offline_connect()
        try:
            row = conn.execute(
                """
                SELECT payload_json
                FROM supabase_get_cache
                WHERE table_name = ? AND query_hash = ?
                LIMIT 1
                """,
                (str(table), query_hash),
            ).fetchone()
        finally:
            conn.close()
    if not row:
        return None
    try:
        return json.loads(row[0])
    except Exception:
        return None


def _clear_supabase_get_cache():
    _ensure_offline_db()
    with _OFFLINE_DB_LOCK:
        conn = _offline_connect()
        try:
            conn.execute("DELETE FROM supabase_get_cache")
            conn.commit()
        finally:
            conn.close()


def _atomic_write_json(path, payload):
    tmp_path = f"{path}.{uuid.uuid4().hex}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _persist_write_queue_locked():
    path = _get_supabase_queue_path()
    _atomic_write_json(path, _WRITE_QUEUE)


def _persist_failed_write_queue_locked():
    path = _get_supabase_failed_queue_path()
    _atomic_write_json(path, _FAILED_WRITE_QUEUE)


def _load_write_queue_once():
    path = _get_supabase_queue_path()
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        if isinstance(data, list):
            with _WRITE_QUEUE_LOCK:
                for item in data:
                    if isinstance(item, dict) and item.get("id"):
                        _WRITE_QUEUE.append(item)
    except Exception:
        return


def _load_failed_write_queue_once():
    path = _get_supabase_failed_queue_path()
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        if isinstance(data, list):
            with _WRITE_QUEUE_LOCK:
                _FAILED_WRITE_QUEUE[:] = [item for item in data if isinstance(item, dict)]
    except Exception:
        return


def _get_supabase_write_queue_snapshot(limit=200):
    path = _get_supabase_queue_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    rows = [item for item in data if isinstance(item, dict)]
    rows.sort(key=lambda r: float(r.get("next_try_at") or 0))
    if limit and limit > 0:
        rows = rows[: int(limit)]
    return rows


def _get_supabase_failed_writes_snapshot(limit=200):
    path = _get_supabase_failed_queue_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    rows = [item for item in data if isinstance(item, dict)]
    rows.sort(key=lambda r: float(r.get("failed_at") or 0), reverse=True)
    if limit and limit > 0:
        rows = rows[: int(limit)]
    return rows


def _clear_supabase_failed_writes():
    with _WRITE_QUEUE_LOCK:
        _FAILED_WRITE_QUEUE.clear()
        _persist_failed_write_queue_locked()


def _get_supabase_write_queue_stats():
    rows = _get_supabase_write_queue_snapshot(limit=0)
    failed = _get_supabase_failed_writes_snapshot(limit=0)
    pending = len(rows)
    if not rows:
        return {
            "pending": 0,
            "failed": len(failed),
            "max_attempts": 0,
            "oldest_next_try_at": None,
        }
    max_attempts = max(int(r.get("attempts") or 0) for r in rows)
    oldest_next_try_at = min(float(r.get("next_try_at") or 0) for r in rows)
    return {
        "pending": pending,
        "failed": len(failed),
        "max_attempts": max_attempts,
        "oldest_next_try_at": oldest_next_try_at,
    }


def _supabase_retry_all_queued_writes():
    """
    Fuerza reintento inmediato de todos los jobs en cola.
    """
    _ensure_write_worker()
    with _WRITE_QUEUE_LOCK:
        if not _WRITE_QUEUE:
            return 0
        now = time.time()
        for idx, item in enumerate(_WRITE_QUEUE):
            item["next_try_at"] = now
            _WRITE_QUEUE[idx] = item
        _persist_write_queue_locked()
        return len(_WRITE_QUEUE)


def _next_retry_delay_seconds(attempts):
    # 1s, 2s, 4s, ... capped at 5 minutes
    tries = max(1, int(attempts))
    return min(300, 2 ** min(tries, 8))


def _supabase_write_worker_loop():
    while True:
        job = None
        with _WRITE_QUEUE_LOCK:
            now = time.time()
            for item in _WRITE_QUEUE:
                if float(item.get("next_try_at") or 0) <= now:
                    job = dict(item)
                    break

        if not job:
            time.sleep(0.6)
            continue

        try:
            if job.get("op") == "upsert":
                _supabase_upsert(
                    job["table"],
                    job.get("rows") or [],
                    env_path=job.get("env_path") or ".env",
                    on_conflict=job.get("on_conflict"),
                )
            elif job.get("op") == "patch":
                _supabase_patch(
                    job["table"],
                    job.get("filters") or {},
                    job.get("values") or {},
                    env_path=job.get("env_path") or ".env",
                )
            else:
                raise RuntimeError(f"Operacion de cola no soportada: {job.get('op')}")
        except Exception as exc:
            with _WRITE_QUEUE_LOCK:
                if not _is_transient_supabase_exception(exc):
                    _FAILED_WRITE_QUEUE.append(
                        {
                            "id": job.get("id"),
                            "op": job.get("op"),
                            "table": job.get("table"),
                            "attempts": int(job.get("attempts") or 0),
                            "failed_at": time.time(),
                            "error": str(exc),
                            "payload": {
                                "rows": job.get("rows"),
                                "filters": job.get("filters"),
                                "values": job.get("values"),
                                "on_conflict": job.get("on_conflict"),
                            },
                        }
                    )
                    if len(_FAILED_WRITE_QUEUE) > 2000:
                        _FAILED_WRITE_QUEUE[:] = _FAILED_WRITE_QUEUE[-2000:]
                    _persist_failed_write_queue_locked()
                    _WRITE_QUEUE[:] = [item for item in _WRITE_QUEUE if item.get("id") != job.get("id")]
                    _persist_write_queue_locked()
                    time.sleep(0.2)
                    continue
                for idx, item in enumerate(_WRITE_QUEUE):
                    if item.get("id") != job.get("id"):
                        continue
                    item["attempts"] = int(item.get("attempts") or 0) + 1
                    item["last_error"] = str(exc)
                    item["next_try_at"] = time.time() + _next_retry_delay_seconds(item["attempts"])
                    _WRITE_QUEUE[idx] = item
                    _persist_write_queue_locked()
                    break
            time.sleep(0.4)
            continue

        with _WRITE_QUEUE_LOCK:
            _WRITE_QUEUE[:] = [item for item in _WRITE_QUEUE if item.get("id") != job.get("id")]
            _persist_write_queue_locked()


def _ensure_write_worker():
    global _WRITE_WORKER_STARTED
    if _WRITE_WORKER_STARTED:
        return
    _load_write_queue_once()
    _load_failed_write_queue_once()
    worker = threading.Thread(target=_supabase_write_worker_loop, daemon=True)
    worker.start()
    _WRITE_WORKER_STARTED = True


def _enqueue_write_job(job):
    _ensure_write_worker()
    record = {
        "id": str(uuid.uuid4()),
        "attempts": 0,
        "next_try_at": time.time(),
        "last_error": "",
    }
    record.update(job or {})
    with _WRITE_QUEUE_LOCK:
        _WRITE_QUEUE.append(record)
        _persist_write_queue_locked()
    return record["id"]


def _supabase_enqueue_upsert(table, rows, env_path=".env", on_conflict=None):
    if not rows:
        return None
    return _enqueue_write_job(
        {
            "op": "upsert",
            "table": table,
            "rows": rows,
            "env_path": env_path,
            "on_conflict": on_conflict,
        }
    )


def _supabase_enqueue_patch(table, filters, values, env_path=".env"):
    if not values:
        return None
    return _enqueue_write_job(
        {
            "op": "patch",
            "table": table,
            "filters": filters or {},
            "values": values or {},
            "env_path": env_path,
        }
    )


def _is_transient_supabase_exception(exc):
    if exc is None:
        return False
    root = exc
    if isinstance(root, RuntimeError) and getattr(root, "__cause__", None) is not None:
        root = root.__cause__

    if isinstance(root, urllib.error.HTTPError):
        code = int(getattr(root, "code", 0) or 0)
        # 5xx + 429 are typically transient.
        return code >= 500 or code == 429
    if isinstance(root, urllib.error.URLError):
        return True
    if isinstance(root, TimeoutError):
        return True
    if isinstance(root, OSError):
        return True
    return False


def _supabase_ping(env_path=".env", timeout=4):
    """
    Verifica conectividad básica con Supabase sin depender de una tabla específica.
    Devuelve True si hay conexión alcanzable, False en caso contrario.
    """
    try:
        supabase_url, supabase_key = _load_supabase_credentials(env_path)
    except Exception:
        return False
    url = f"{supabase_url.rstrip('/')}/rest/v1/"
    request = urllib.request.Request(
        url,
        headers=_supabase_headers(supabase_key),
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout):
            return True
    except urllib.error.HTTPError as exc:
        # 401/403 indican que el host está alcanzable.
        code = int(getattr(exc, "code", 0) or 0)
        return code in {401, 403}
    except Exception:
        return False


def _supabase_upsert_with_queue(table, rows, env_path=".env", on_conflict=None):
    if not rows:
        return {"status": "skipped", "queued": False, "rows": 0, "data": []}
    try:
        data = _supabase_upsert(
            table,
            rows,
            env_path=env_path,
            on_conflict=on_conflict,
        )
        return {"status": "synced", "queued": False, "rows": len(rows), "data": data}
    except Exception as exc:
        if not _is_transient_supabase_exception(exc):
            raise
        try:
            _supabase_enqueue_upsert(
                table,
                rows,
                env_path=env_path,
                on_conflict=on_conflict,
            )
        except Exception as enqueue_exc:
            raise RuntimeError(
                f"No se pudo guardar ni encolar {table}: {enqueue_exc}"
            ) from enqueue_exc
        return {
            "status": "queued",
            "queued": True,
            "rows": len(rows),
            "data": [],
            "error": str(exc),
        }


def _supabase_patch_with_queue(table, filters, values, env_path=".env"):
    if not values:
        return {"status": "skipped", "queued": False, "rows": 0, "data": []}
    try:
        data = _supabase_patch(
            table,
            filters,
            values,
            env_path=env_path,
        )
        return {"status": "synced", "queued": False, "rows": 1, "data": data}
    except Exception as exc:
        if not _is_transient_supabase_exception(exc):
            raise
        try:
            _supabase_enqueue_patch(
                table,
                filters,
                values,
                env_path=env_path,
            )
        except Exception as enqueue_exc:
            raise RuntimeError(
                f"No se pudo actualizar ni encolar {table}: {enqueue_exc}"
            ) from enqueue_exc
        return {
            "status": "queued",
            "queued": True,
            "rows": 1,
            "data": [],
            "error": str(exc),
        }


def _supabase_upsert(table, rows, env_path=".env", on_conflict=None):
    supabase_url, supabase_key = _load_supabase_credentials(env_path)
    if not rows:
        return []
    conflict_query = f"?on_conflict={on_conflict}" if on_conflict else ""
    url = f"{supabase_url.rstrip('/')}/rest/v1/{table}{conflict_query}"
    body = json.dumps(rows, ensure_ascii=False).encode("utf-8")
    last_exc = None
    attempted_refresh = False
    for delay in (0, 0.6, 1.5):
        if delay:
            time.sleep(delay)
        token = _supabase_get_access_token(env_path=env_path)
        headers = _supabase_headers(supabase_key, bearer_token=token)
        headers["Content-Type"] = "application/json"
        headers["Prefer"] = "resolution=merge-duplicates"
        request = urllib.request.Request(url, data=body, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = response.read().decode("utf-8")
            return json.loads(payload) if payload else []
        except urllib.error.HTTPError as exc:
            last_exc = exc
            if (
                int(getattr(exc, "code", 0) or 0) == 401
                and not attempted_refresh
                and _supabase_refresh_session(env_path=env_path)
            ):
                attempted_refresh = True
                continue
        except Exception as exc:
            last_exc = exc
            continue
    raise RuntimeError(
        _format_supabase_error(f"No se pudo guardar en {table}", last_exc)
    ) from last_exc


def _supabase_patch(table, filters, values, env_path=".env"):
    supabase_url, supabase_key = _load_supabase_credentials(env_path)
    if not values:
        return []
    params = {}
    for key, val in (filters or {}).items():
        params[key] = f"eq.{val}"
    query = urllib.parse.urlencode(params)
    url = f"{supabase_url.rstrip('/')}/rest/v1/{table}"
    if query:
        url = f"{url}?{query}"
    body = json.dumps(values, ensure_ascii=False).encode("utf-8")
    last_exc = None
    attempted_refresh = False
    for delay in (0, 0.6, 1.5):
        if delay:
            time.sleep(delay)
        token = _supabase_get_access_token(env_path=env_path)
        headers = _supabase_headers(supabase_key, bearer_token=token)
        headers["Content-Type"] = "application/json"
        headers["Prefer"] = "return=representation"
        request = urllib.request.Request(url, data=body, headers=headers, method="PATCH")
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = response.read().decode("utf-8")
            return json.loads(payload) if payload else []
        except urllib.error.HTTPError as exc:
            last_exc = exc
            if (
                int(getattr(exc, "code", 0) or 0) == 401
                and not attempted_refresh
                and _supabase_refresh_session(env_path=env_path)
            ):
                attempted_refresh = True
                continue
        except Exception as exc:
            last_exc = exc
            continue
    raise RuntimeError(
        _format_supabase_error(f"No se pudo actualizar {table}", last_exc)
    ) from last_exc


def _normalize_text(value):
    if value is None:
        return ""
    normalized = unicodedata.normalize("NFD", str(value))
    text = "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", text).strip().lower()


def _normalize_cedula(value):
    if value is None:
        return ""
    return re.sub(r"\D+", "", str(value))


def _parse_date_value(value):
    if not value:
        return None
    raw = str(value).strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return time.strftime("%Y-%m-%d", time.strptime(raw, fmt))
        except ValueError:
            continue
    return None


def _get_desktop_dir():
    def _validate(path):
        candidate = str(path or "").strip()
        if not candidate:
            return None
        try:
            os.makedirs(candidate, exist_ok=True)
            return candidate
        except Exception:
            return None

    # 0) Windows User Shell Folders: soporta redirecciones/GPO.
    if os.name == "nt":
        try:
            import winreg

            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
            ) as key:
                desktop_raw, _ = winreg.QueryValueEx(key, "Desktop")
            resolved = _validate(os.path.expandvars(str(desktop_raw or "").strip()))
            if resolved:
                return resolved
        except Exception:
            pass

    # 1) OneDrive.
    for env_key in ("OneDrive", "OneDriveConsumer"):
        base = os.getenv(env_key)
        if base:
            resolved = _validate(os.path.join(base, "Desktop"))
            if resolved:
                return resolved

    # 2) USERPROFILE.
    userprofile = os.getenv("USERPROFILE")
    if userprofile:
        for name in ("Desktop", "Escritorio"):
            resolved = _validate(os.path.join(userprofile, name))
            if resolved:
                return resolved

    # 3) HOME.
    home = os.path.expanduser("~")
    for name in ("Desktop", "Escritorio"):
        resolved = _validate(os.path.join(home, name))
        if resolved:
            return resolved

    # 4) Fallback estable.
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        resolved = _validate(os.path.join(local_app_data, "RECA", "outputs"))
        if resolved:
            return resolved
    return os.getcwd()


def _sanitize_filename(value):
    safe = re.sub(r"[\\/:*?\"<>|]+", " ", str(value or ""))
    safe = re.sub(r"\s+", " ", safe).strip()
    return safe or "Empresa"


def format_checkbox_symbol(value):
    return "☑" if bool(value) else "☐"


def sanitize_logo_error_cells(workbook_or_sheet):
    """
    Limpia celdas A1 con error literal #VALUE! en hojas de salida para evitar
    que el usuario vea el error en encabezados cuando el logo es una capa visual.
    """
    if workbook_or_sheet is None:
        return

    try:
        sheets = workbook_or_sheet.Worksheets
        is_com = True
    except Exception:
        sheets = None
        is_com = False

    if is_com:
        try:
            total = int(sheets.Count)
        except Exception:
            total = 0
        for idx in range(1, total + 1):
            try:
                ws = sheets(idx)
                cell = ws.Range("A1")
                value = cell.Value
                formula = ""
                try:
                    formula = str(cell.Formula or "").strip()
                except Exception:
                    formula = ""
                if formula:
                    continue
                try:
                    text_value = str(cell.Text or "").strip().upper()
                except Exception:
                    text_value = ""
                value_str = str(value or "").strip().upper()
                if value_str == "#VALUE!" or text_value == "#VALUE!" or value == 2015:
                    cell.Value = ""
            except Exception:
                continue
        return

    sheetnames = getattr(workbook_or_sheet, "sheetnames", None)
    if sheetnames is not None:
        for name in list(sheetnames):
            try:
                ws = workbook_or_sheet[name]
                value = ws["A1"].value
                if str(value or "").strip().upper() == "#VALUE!":
                    ws["A1"].value = ""
            except Exception:
                continue
        return

    try:
        value = workbook_or_sheet["A1"].value
        if str(value or "").strip().upper() == "#VALUE!":
            workbook_or_sheet["A1"].value = ""
    except Exception:
        return
