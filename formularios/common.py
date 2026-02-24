import os
import re
import time
import unicodedata
import json
import sys
import threading
import uuid
import urllib.parse
import urllib.request
import urllib.error


def _resolve_env_candidates(env_path=".env"):
    if os.path.isabs(env_path):
        return [env_path]
    candidates = []
    # 1) current working directory
    candidates.append(os.path.abspath(env_path))
    # 2) executable/script directory
    try:
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        candidates.append(os.path.join(exe_dir, env_path))
    except Exception:
        pass
    # 3) project root (when running from source)
    try:
        project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        candidates.append(os.path.join(project_root, env_path))
    except Exception:
        pass
    # 4) roaming appdata fallback
    appdata = os.getenv("APPDATA")
    if appdata:
        candidates.append(os.path.join(appdata, "RECA Inclusion Laboral", env_path))
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
            env[key.strip()] = value.strip().strip('"').strip("'")
    return env


def _supabase_headers(api_key):
    return {
        "apikey": api_key,
        "Authorization": f"Bearer {api_key}",
    }


def _supabase_get(table, params, env_path=".env"):
    env = _load_env_file(env_path)
    supabase_url = env.get("SUPABASE_URL")
    supabase_key = env.get("SUPABASE_KEY")
    if not supabase_url or not supabase_key:
        raise RuntimeError("Missing SUPABASE_URL or SUPABASE_KEY")
    query = urllib.parse.urlencode(params)
    url = f"{supabase_url.rstrip('/')}/rest/v1/{table}?{query}"
    request = urllib.request.Request(url, headers=_supabase_headers(supabase_key))
    last_error = None
    for _ in range(3):
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = response.read().decode("utf-8")
            return json.loads(payload)
        except Exception as exc:
            last_error = exc
    raise RuntimeError(_format_supabase_error("Supabase no esta disponible", last_error)) from last_error


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


def _persist_write_queue_locked():
    path = _get_supabase_queue_path()
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(_WRITE_QUEUE, handle, ensure_ascii=False, indent=2)


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


def _supabase_upsert(table, rows, env_path=".env", on_conflict=None):
    env = _load_env_file(env_path)
    supabase_url = env.get("SUPABASE_URL")
    supabase_key = env.get("SUPABASE_KEY")
    if not supabase_url or not supabase_key:
        raise RuntimeError("Missing SUPABASE_URL or SUPABASE_KEY")
    if not rows:
        return []
    conflict_query = f"?on_conflict={on_conflict}" if on_conflict else ""
    url = f"{supabase_url.rstrip('/')}/rest/v1/{table}{conflict_query}"
    body = json.dumps(rows, ensure_ascii=False).encode("utf-8")
    headers = _supabase_headers(supabase_key)
    headers["Content-Type"] = "application/json"
    headers["Prefer"] = "resolution=merge-duplicates"
    request = urllib.request.Request(url, data=body, headers=headers, method="POST")
    last_exc = None
    for delay in (0, 0.6, 1.5):
        if delay:
            time.sleep(delay)
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = response.read().decode("utf-8")
            return json.loads(payload) if payload else []
        except Exception as exc:
            last_exc = exc
            continue
    raise RuntimeError(
        _format_supabase_error(f"No se pudo guardar en {table}", last_exc)
    ) from last_exc


def _supabase_patch(table, filters, values, env_path=".env"):
    env = _load_env_file(env_path)
    supabase_url = env.get("SUPABASE_URL")
    supabase_key = env.get("SUPABASE_KEY")
    if not supabase_url or not supabase_key:
        raise RuntimeError("Missing SUPABASE_URL or SUPABASE_KEY")
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
    headers = _supabase_headers(supabase_key)
    headers["Content-Type"] = "application/json"
    headers["Prefer"] = "return=representation"
    request = urllib.request.Request(url, data=body, headers=headers, method="PATCH")
    last_exc = None
    for delay in (0, 0.6, 1.5):
        if delay:
            time.sleep(delay)
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                payload = response.read().decode("utf-8")
            return json.loads(payload) if payload else []
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
    for env_key in ("OneDrive", "OneDriveConsumer"):
        base = os.getenv(env_key)
        if base:
            desktop = os.path.join(base, "Desktop")
            if os.path.isdir(desktop):
                return desktop
    userprofile = os.getenv("USERPROFILE")
    if userprofile:
        desktop = os.path.join(userprofile, "Desktop")
        if os.path.isdir(desktop):
            return desktop
    return os.getcwd()


def _sanitize_filename(value):
    safe = re.sub(r"[\\/:*?\"<>|]+", " ", str(value or ""))
    safe = re.sub(r"\s+", " ", safe).strip()
    return safe or "Empresa"
